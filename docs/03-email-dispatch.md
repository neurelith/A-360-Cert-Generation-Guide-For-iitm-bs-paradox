# Step 4 — Email Dispatch

Once the PDFs are generated and visually validated, `send_emails.py` delivers them via Amazon SES.

```bash
# test with 1 email first
python send_emails.py --count 1

# send to everyone in the CSV
python send_emails.py --all
```

---

## What Happens When You Run It

1. Loads `dummy_data.csv` and the generated PDFs from `generated_certificates/pdf/`.
2. For each row, picks the right HTML email template based on the `type` column.
3. Injects candidate data into the HTML placeholders.
4. Builds a multi-part MIME email with the HTML body + attached PDF.
5. Sends it through Amazon SES over TLS.
6. Logs the result (success/fail + timestamp) to `logs/email_delivery_log.csv`.

---

## HTML Template Matching

Each certificate type gets a different email design. The templates live in `email html/` and are selected automatically:

| Candidate type | HTML template | Key differences |
|---------------|---------------|-----------------|
| **Winner** | `saavan26_certificate_winner(1).html` | Shows prize amount, rank/position, achievement-focused styling |
| **Participant** | `saavan26_certificate_particiation.html` | Participation recognition, event details |
| **Judge / Guest** | `saavan26_certificate_guest.html` | Formal appreciation tone, no position/rank fields |

### Placeholders in the HTML

The templates use `{{DoubleHandlebar}}` placeholders (not the same as the `<<angle bracket>>` ones in the PPTX). The script does simple string replacement:

| Placeholder | Replaced with | Source column |
|------------|---------------|---------------|
| `{{StudentName}}` | Candidate name | `name` |
| `{{GuestName}}` | Same as above (used in judge template) | `name` |
| `{{EventName}}` | Event name | `event_name` |
| `{{Position}}` | Rank or role | `position` |
| `{{CertificateType}}` | Display-friendly type | Derived from `position` |
| `{{WinningPrize}}` | Prize description | `winning_prize` |
| `{{CertificateID}}` | Certificate code | `cert_id` |
| `{{CertificateLink}}` | Verification URL | `qr_code` |
| `{{FeedbackLink}}` | Feedback form | `feedback_link` |
| `{{DiscrepancyFormLink}}` | Correction request form | `discrepancy_link` |

---

## Email Structure (MIME)

The email is built as a nested MIME structure:

```
MIMEMultipart("mixed")
├── MIMEMultipart("alternative")
│   ├── MIMEText (plain text fallback)
│   └── MIMEText (HTML body)
└── MIMEApplication (PDF attachment)
```

The `alternative` container lets email clients choose between plain text (for ancient clients) and HTML (for everything else). The PDF is attached at the top level of the `mixed` container so it shows as a downloadable file.

The attachment filename is sanitized from the candidate name:
```
Certificate_SAAVAN26-W-0001_Krishna.pdf
```

---

## Subject Lines

Subjects are generated dynamically based on certificate type:

- **Winner:** `Saavan '26 - Certificate of Achievement (Winner) - Stand-Up Comedy Showdown - Krishna`
- **Participant:** `Saavan '26 - Certificate of Participation - Rhythm & Rhapsody - Aritri Roy`
- **Judge/Guest:** `Saavan '26 - Certificate of Appreciation (Judge) - Western Music Solo - Dr. Ananya Sen`

---

## Amazon SES Setup

The email goes through Amazon Simple Email Service (Mumbai region: `ap-south-1`).

**Connection details** (configured in `config.py`):
- Host: `email-smtp.ap-south-1.amazonaws.com`
- Port: `587` (STARTTLS)
- Sender: `certificates@iitmparadox.org` (must be verified in SES)

**Credentials** are read from `ses-smtp-user.*.csv` — this is the IAM user credentials file you download from the AWS console when creating SES SMTP credentials. It has two columns: `SMTP user name` and `SMTP password`.

### SES sandbox limitations

New SES accounts start in sandbox mode. You can only send to **verified email addresses**. That's why the test pipeline delivers to two pre-verified staging inboxes by default:
- `24f2006473@ds.study.iitm.ac.in`
- `24f3004704@ds.study.iitm.ac.in`

To send to real candidates, request production access from the AWS console.

---

## CLI Options

| Flag | What it does | Example |
|------|-------------|---------|
| `--all` | Send to every row in the CSV | `python send_emails.py --all` |
| `--count N` | Send only the first N rows | `python send_emails.py --count 3` |
| `--to EMAIL` | Override recipient address (comma-separated for multiple) | `python send_emails.py --count 1 --to "me@gmail.com"` |

### Typical workflow

```bash
# 1. smoke test — send 1 email to yourself
python send_emails.py --count 1 --to "youremail@gmail.com"

# 2. check the email on desktop and mobile
#    - does the HTML render correctly?
#    - is the PDF attached?
#    - does the verify link in the email work?

# 3. send to staging addresses
python send_emails.py --count 1

# 4. full batch
python send_emails.py --all
```

---

## Delivery Logging

Every email attempt is appended to `logs/email_delivery_log.csv` with:

| Field | Example |
|-------|---------|
| `timestamp` | `2026-10-01T14:30:22.123456` |
| `cert_id` | `SAAVAN26-W-0001` |
| `type` | `Winner` |
| `name` | `Krishna` |
| `event_name` | `Stand-Up Comedy Showdown` |
| `position` | `Winner` |
| `original_email` | `25f2007055@ds.study.iitm.ac.in` |
| `delivered_to` | `24f2006473@ds.study.iitm.ac.in, 24f3004704@ds.study.iitm.ac.in` |
| `status` | `SUCCESS` or `FAILED` |
| `error` | Empty on success, error message on failure |

If a batch fails halfway through, use this log to figure out which emails went out and which didn't. Filter by `status == FAILED`, fix the issue, and re-run with only those entries.

---

## Custom Headers

Every email includes traceability headers that don't show in the email body but can be found in the raw email source:

```
X-Original-Recipient: 25f2007055@ds.study.iitm.ac.in
X-Certificate-ID: SAAVAN26-W-0001
X-Certificate-Type: Winner
X-Test-Mode: True
```

Useful for debugging delivery issues with AWS support or verifying which candidate an email was meant for.
