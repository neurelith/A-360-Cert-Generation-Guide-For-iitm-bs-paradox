# Saavan '26 — Certificate Pipeline

Batch certificate generation and email delivery system for Saavan '26 (IIT Madras cultural fest).

Takes a spreadsheet of candidate data, stamps it onto designed PowerPoint templates, compiles vector PDFs, and mails them out via Amazon SES — all automated.

---

## How It Works (4 steps)

```
Spreadsheet (Excel/Sheets)          Python Pipeline               Amazon SES
┌─────────────────────┐       ┌──────────────────────┐       ┌──────────────┐
│ 1. Generate UUIDs   │       │ 3. Build certs       │       │ 4. Send      │
│    (Excel formula)  │──CSV──│    PPTX → PDF        │──PDF──│    emails    │
│ 2. Build verify URLs│       │    (PowerPoint COM)  │       │    (SMTP)    │
└─────────────────────┘       └──────────────────────┘       └──────────────┘
```

**Step 1** — Generate a unique ID for each candidate using an Excel formula.  
**Step 2** — Build a verification URL from that ID (also an Excel formula).  
**Step 3** — Run `generate_certificates.py` to fill templates and export PDFs.  
**Step 4** — Run `send_emails.py` to deliver PDFs with styled HTML emails.

---

## Quickstart

```bash
cd saavan26/dummy

# generate all certificates (PPTX + PDF)
python generate_certificates.py

# send 1 test email to check everything works
python send_emails.py --count 1

# send to everyone
python send_emails.py --all
```

---

## Folder Layout

```
saavan26/
├── README.md                  ← you are here
├── docs/
│   ├── 01-data-prep.md        ← Excel formulas, CSV schema, how to prep your data
│   ├── 02-certificate-gen.md  ← how the PPTX engine works under the hood
│   ├── 03-email-dispatch.md   ← HTML templates, SES setup, sending emails
│   ├── 04-config.md           ← config.py reference (paths, creds, coordinates)
│   └── 05-troubleshooting.md  ← common failures and how to fix them
│
└── dummy/                     ← working directory (code + data + templates)
    ├── config.py              ← central config (SMTP creds, paths, QR coords)
    ├── generate_certificates.py  ← cert generation engine
    ├── send_emails.py         ← email dispatch engine
    ├── inspect_pptx.py        ← template inspection utility
    ├── dummy_data.csv         ← sample candidate data (5 rows, 3 types)
    ├── ses-smtp-user.*.csv    ← AWS SES SMTP credentials (DO NOT COMMIT)
    │
    ├── template/              ← master PPTX decks (designed by the team)
    │   ├── General_Event_Winner_Template.pptx
    │   ├── General_Event_Participant_Template.pptx
    │   └── Guest_Template.pptx
    │
    ├── email html/            ← responsive HTML email bodies
    │   ├── saavan26_certificate_winner(1).html
    │   ├── saavan26_certificate_particiation.html
    │   └── saavan26_certificate_guest.html
    │
    ├── generated_certificates/
    │   ├── pptx/              ← intermediate PPTX files
    │   └── pdf/               ← final PDF output (this gets emailed)
    │
    └── logs/
        └── email_delivery_log.csv  ← audit trail of every sent email
```

---

## Requirements

- **Python 3.9+**
- **Microsoft PowerPoint** installed (used for PPTX → PDF conversion via COM)
- **Windows** (COM automation is Windows-only)
- **AWS SES credentials** (for email delivery)

### Python packages

```
pip install python-pptx pandas qrcode pillow pywin32
```

---

## Docs

| Doc | What it covers |
|-----|---------------|
| [01-data-prep.md](docs/01-data-prep.md) | Excel formulas for UUID + verify URL, CSV column schema, export rules |
| [02-certificate-gen.md](docs/02-certificate-gen.md) | Template system, placeholder replacement, QR codes, PDF compilation |
| [03-email-dispatch.md](docs/03-email-dispatch.md) | HTML template matching, MIME structure, SES config, CLI options |
| [04-config.md](docs/04-config.md) | Every setting in config.py explained |
| [05-troubleshooting.md](docs/05-troubleshooting.md) | Zombie PowerPoint processes, run splitting, UTF-8 encoding, SES limits |

---

## Why PowerPoint instead of HTML-to-PDF?

Short version: the design team works in PowerPoint. Their templates use custom fonts (Roxborough-CF, Crimson Pro), precise kerning, and vector artwork. Converting that to HTML/CSS and then to PDF with Puppeteer or WeasyPrint introduces rendering bugs — misaligned text, wrong font metrics, rasterized vectors.

By keeping PPTX as the source of truth, we get pixel-perfect output that matches what the designers signed off on. PowerPoint's own PDF export preserves vectors, embedded fonts, and alpha transparency natively.
