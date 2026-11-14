# AI Inbox Automation

Read incoming emails, let an LLM **sort them, pull out the key details and draft a reply**, and log everything to a spreadsheet for a person to check.

- Sorts each email into: sales lead / support / invoice-billing / spam / other
- Extracts name, company, what they want, and urgency (high / medium / low, colour-coded)
- Drafts a polite reply for you to approve. **Nothing is ever sent**: drafts only go into the spreadsheet
- Real inbox via IMAP is **read-only** (folder opened read-only, messages fetched without marking them read)
- The model's labels are checked in code: anything off-list falls back to a safe value, and spam never gets a draft
- Any LLM: OpenAI, Groq, or a local model through Ollama (emails never leave your machine)

```
$ python inbox.py samples/*.eml -o inbox_report.xlsx
sales lead       high    Quote for 40 lobby chairs
support          high    Table arrived with a cracked leg - order #CL-2291
invoice/billing  high    Invoice INV-1042 - charged twice?
spam             low     Rank #1 on Google in 7 days GUARANTEED!!!
sales lead       medium  Wholesale / stockist enquiry
other            low     Student visit to your workshop?
wrote inbox_report.xlsx - drafts are in the workbook, nothing was sent
```

`inbox_report.xlsx` in this repo is the real output of that run, on six sample emails to a fictional furniture workshop, using a local 7B model (Qwen2.5-7B via Ollama). All six were categorised correctly. The drafts are usable starting points, not final copy: that is why a person approves them.

## Run it

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env && set -a; source .env; set +a   # choose your LLM provider
python inbox.py samples/*.eml -o inbox_report.xlsx    # demo on the sample emails
python inbox.py --imap -o inbox_report.xlsx           # your real inbox (IMAP_* in .env), unread only
python test_inbox.py                                  # offline test, stubbed LLM, no network
```

## n8n version

`n8n-workflow.json` is the same flow as an n8n **template**: IMAP Email trigger → HTTP Request to an OpenAI-compatible chat endpoint (JSON mode) → Code node that validates the labels → Google Sheets append. Import it (Workflows → Import from file), then add your IMAP, API-key (header `Authorization: Bearer ...`) and Google Sheets credentials and the sheet URL. The sheet needs an `Inbox` tab with a header row matching the columns in the Code node. It has not been run against a live n8n instance here, so test it on a spare mailbox first.

## Next steps for production

Web forms and document attachments as inputs, pushing leads into a CRM (HubSpot, Pipedrive), saving approved drafts as Gmail/Outlook drafts, and a per-business prompt with your real prices and policies.
