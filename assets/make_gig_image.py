"""Render the gig images from the real demo output (inbox_report.xlsx): 1280x769 (Fiverr) and 1024x768 (Upwork/Contra)."""
from pathlib import Path
import openpyxl
from PIL import Image, ImageDraw, ImageFont

root = Path(__file__).resolve().parent.parent
rows = list(openpyxl.load_workbook(root / "inbox_report.xlsx")["Inbox"].iter_rows(values_only=True))[1:]
H = "/System/Library/Fonts/HelveticaNeue.ttc"
f = lambda s, b=False: ImageFont.truetype(H, s, index=1 if b else 0)
BG, GREEN, WHITE, INK = (15, 42, 36), (110, 231, 183), (255, 255, 255), (50, 50, 50)
URG = {"high": ((248, 215, 218), (170, 30, 40)), "medium": ((255, 243, 205), (140, 100, 0)), "low": ((212, 237, 218), (30, 110, 50))}


def fit(d, text, font, w):
    while d.textlength(text, font=font) > w:
        text = text[:-2].rstrip() + "…"
    return text


def render(W, Hh, out, scale):
    s = lambda v: int(v * scale)
    im = Image.new("RGB", (W, Hh), BG)
    d = ImageDraw.Draw(im)
    y0 = (Hh - s(560)) // 2
    d.text((s(60), y0), "AI that sorts", font=f(s(60), True), fill=WHITE)
    d.text((s(60), y0 + s(70)), "your inbox", font=f(s(60), True), fill=GREEN)
    d.text((s(60), y0 + s(145)), "Automate busywork with AI", font=f(s(28)), fill=(200, 220, 212))
    for i, t in enumerate(["Sorts emails & forms automatically", "Drafts replies for you to approve",
                           "Logs everything to Sheets/Excel", "Python or n8n"]):
        yy = y0 + s(215) + i * s(44)
        d.line([(s(62), yy + s(15)), (s(70), yy + s(23)), (s(84), yy + s(7))], fill=GREEN, width=max(3, s(4)))
        d.text((s(98), yy), t, font=f(s(25)), fill=WHITE)
    x = s(60)
    for t, c in [("EMAIL", (66, 133, 244)), ("AI", (124, 77, 255)), ("SHEETS", (33, 150, 83))]:
        w = d.textlength(t, font=f(s(18), True)) + s(28)
        d.rounded_rectangle((x, y0 + s(420), x + w, y0 + s(456)), 8, fill=c)
        d.text((x + s(14), y0 + s(428)), t, font=f(s(18), True), fill=WHITE)
        x += w + s(14)

    # spreadsheet card built from real output: subject, category, urgency, name/company
    X, PW = s(560), W - s(560) - s(40)
    cols = [("Email", 0.36), ("Category", 0.25), ("Urgency", 0.16), ("From", 0.23)]
    d.rounded_rectangle((X, y0 - s(10), X + PW, y0 + s(560)), 14, fill=WHITE)
    d.text((X + s(24), y0 + s(10)), "inbox_report.xlsx", font=f(s(20), True), fill=(30, 30, 30))
    ty = y0 + s(55)
    d.rectangle((X + s(16), ty, X + PW - s(16), ty + s(36)), fill=(232, 240, 236))
    inner = PW - s(48)
    cx = X + s(26)
    for name, frac in cols:
        d.text((cx, ty + s(9)), name, font=f(s(16), True), fill=(40, 40, 40)); cx += int(inner * frac)
    for r in rows:
        ty += s(46)
        subj, cat, urg, who = r[2], r[3], r[4], r[6] or r[5]
        cx = X + s(26)
        for v, (_, frac) in zip([subj, cat, None, who], cols):
            cw = int(inner * frac) - s(10)
            if v is None:  # urgency pill, same colours as the workbook
                bg, fg = URG[urg]
                pw = d.textlength(urg, font=f(s(15), True)) + s(18)
                d.rounded_rectangle((cx, ty + s(3), cx + pw, ty + s(29)), 6, fill=bg)
                d.text((cx + s(9), ty + s(7)), urg, font=f(s(15), True), fill=fg)
            else:
                d.text((cx, ty + s(6)), fit(d, str(v), f(s(16)), cw), font=f(s(16)), fill=INK)
            cx += int(inner * frac)
        d.line([(X + s(16), ty + s(38)), (X + PW - s(16), ty + s(38))], fill=(235, 235, 235))
    lead = rows[0]
    by = ty + s(56)
    d.rounded_rectangle((X + s(16), by, X + PW - s(16), by + s(118)), 10, outline=GREEN, width=2)
    d.text((X + s(32), by + s(12)), f"Draft reply to {lead[5]} (not sent):", font=f(s(16), True), fill=(20, 110, 80))
    words, lines, cur = lead[9].split(), [], ""
    for wd in words:
        if d.textlength(cur + " " + wd, font=f(s(15))) > PW - s(64):
            lines.append(cur); cur = wd
        else:
            cur = (cur + " " + wd).strip()
    lines.append(cur)
    for i, ln in enumerate(lines[:3]):
        d.text((X + s(32), by + s(40) + i * s(22)), ln, font=f(s(15)), fill=(60, 60, 60))
    d.text((X + s(24), y0 + s(530)), "Real output from the demo on sample emails", font=f(s(14)), fill=(110, 110, 110))
    im.save(out)


render(1280, 769, root / "assets" / "gig-image.png", 1.0)
render(1024, 768, root / "assets" / "gig-4x3.png", 0.82)
