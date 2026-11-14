"""Sort incoming emails with an LLM, extract the key fields, draft a reply, and log it all to Excel.
Drafts are written to the workbook only. Nothing is ever sent.

usage: python inbox.py samples/*.eml -o inbox_report.xlsx    # demo: a folder of .eml files
       python inbox.py --imap -o inbox_report.xlsx           # real inbox, read-only (IMAP_* env vars)
"""
import argparse
import email
import imaplib
import json
import os
from email import policy
from email.utils import parseaddr
from pathlib import Path

import openpyxl
from openai import OpenAI
from openpyxl.styles import Alignment, Font, PatternFill

MODEL = os.getenv("LLM_MODEL", "gpt-4o-mini")
CATEGORIES = ["sales lead", "support", "invoice/billing", "spam", "other"]
URGENCY = ["high", "medium", "low"]
COLOURS = {"high": "F8D7DA", "medium": "FFF3CD", "low": "D4EDDA"}
BUSINESS = os.getenv("BUSINESS_NAME", "Copperleaf Furniture Studio")  # fictional demo business

PROMPT = """You triage the inbox of {business}. Read the email and reply with JSON with exactly these keys:
category: one of {categories}
urgency: one of high, medium, low (high = deadline or upset customer in the next few days)
name: sender's name, company: sender's company or null,
request: what they want, in one short sentence,
summary: one sentence for the team,
reply: a short, polite draft reply signed "{business}". Do not promise prices, dates or refunds; say we will confirm.
For spam, reply must be "".

From: {sender}
Subject: {subject}

{body}"""


def parse(raw: bytes) -> dict:
    m = email.message_from_bytes(raw, policy=policy.default)
    part = m.get_body(preferencelist=("plain", "html"))
    return {"from": m["from"] or "", "subject": m["subject"] or "", "date": m["date"] or "",
            "body": part.get_content().strip() if part else ""}


def from_imap(limit: int) -> list[dict]:
    """Read-only: the folder is opened with readonly=True and fetched with PEEK, so nothing is marked read."""
    with imaplib.IMAP4_SSL(os.environ["IMAP_HOST"]) as imap:
        imap.login(os.environ["IMAP_USER"], os.environ["IMAP_PASSWORD"])
        imap.select(os.getenv("IMAP_FOLDER", "INBOX"), readonly=True)
        ids = imap.search(None, "UNSEEN")[1][0].split()[-limit:]
        return [parse(imap.fetch(i, "(BODY.PEEK[])")[1][0][1]) for i in ids]


def classify(msg: dict, llm) -> dict:
    r = llm.chat.completions.create(model=MODEL, temperature=0, response_format={"type": "json_object"},
                                    messages=[{"role": "user", "content": PROMPT.format(
                                        business=BUSINESS, categories=", ".join(CATEGORIES), sender=msg["from"],
                                        subject=msg["subject"], body=msg["body"][:6000])}])
    return clean(json.loads(r.choices[0].message.content), msg)


def clean(ai: dict, msg: dict) -> dict:
    """Never trust the model's labels blindly: anything off-list falls back to a safe value."""
    cat = str(ai.get("category", "")).lower().strip()
    cat = "invoice/billing" if cat in ("invoice", "billing", "invoice-billing") else cat
    urg = str(ai.get("urgency", "")).lower().strip()
    name = ai.get("name") or parseaddr(msg["from"])[0] or None
    return {**msg, "category": cat if cat in CATEGORIES else "other", "urgency": urg if urg in URGENCY else "medium",
            "name": name, "company": ai.get("company") or None, "request": ai.get("request") or "",
            "summary": ai.get("summary") or "", "reply": "" if cat == "spam" else (ai.get("reply") or "")}


HEAD = ["date", "from", "subject", "category", "urgency", "name", "company", "request", "summary", "draft reply (not sent)"]
KEYS = ["date", "from", "subject", "category", "urgency", "name", "company", "request", "summary", "reply"]


def write_xlsx(rows: list[dict], out: Path):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Inbox"
    ws.append(HEAD)
    for r in rows:
        ws.append([r.get(k) for k in KEYS])
        ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor=COLOURS[r["urgency"]])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, width in zip("ABCDEFGHIJ", [24, 30, 34, 15, 9, 16, 18, 40, 45, 70]):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    wb.save(out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("emls", nargs="*", type=Path)
    ap.add_argument("--imap", action="store_true", help="read unread mail via IMAP (read-only)")
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("-o", "--out", type=Path, default=Path("inbox_report.xlsx"))
    a = ap.parse_args()
    llm = OpenAI(base_url=os.getenv("LLM_BASE_URL", "https://api.openai.com/v1"),
                 api_key=os.getenv("LLM_API_KEY", "not-needed-for-local"))
    msgs = from_imap(a.limit) if a.imap else [parse(p.read_bytes()) for p in a.emls]
    rows = []
    for m in msgs:
        r = classify(m, llm)
        print(f"{r['category']:<16} {r['urgency']:<7} {r['subject'][:60]}")
        rows.append(r)
    write_xlsx(rows, a.out)
    print("wrote", a.out, "- drafts are in the workbook, nothing was sent")
