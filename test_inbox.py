"""Offline test of parsing, label cleaning and Excel writing with a stubbed LLM (no network). Run: python test_inbox.py"""
import json
import tempfile
from pathlib import Path
from types import SimpleNamespace

import openpyxl
from inbox import classify, parse, write_xlsx

raw = Path("samples/02_damaged_delivery.eml").read_bytes()
msg = parse(raw)
assert msg["subject"].startswith("Table arrived with a cracked leg"), msg
assert "07700 900123" in msg["body"]


def stub(answer: dict):
    content = json.dumps(answer)
    create = lambda **kw: SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=content))])
    return SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)))


good = classify(msg, stub(dict(category="Support", urgency="HIGH", name="Tom Becker", company=None,
                                request="Replace cracked table leg", summary="Damaged delivery", reply="Sorry Tom...")))
assert (good["category"], good["urgency"], good["reply"]) == ("support", "high", "Sorry Tom..."), good
odd = classify(msg, stub(dict(category="complaint", urgency="asap", reply="x")))  # off-list labels
assert (odd["category"], odd["urgency"], odd["name"]) == ("other", "medium", "Tom Becker"), odd
spam = classify(msg, stub(dict(category="spam", urgency="low", reply="Thanks!")))
assert spam["reply"] == "", "spam must never get a draft"
assert classify(msg, stub(dict(category="billing")))["category"] == "invoice/billing"

out = Path(tempfile.mkdtemp()) / "t.xlsx"
write_xlsx([good, spam], out)
ws = openpyxl.load_workbook(out)["Inbox"]
assert ws["D2"].value == "support" and ws["E2"].fill.fgColor.rgb.endswith("F8D7DA")
assert ws["J2"].value == "Sorry Tom..." and ws.max_row == 3
print("ok")
