# Config Reference

Everything configurable lives in `config.py`. This doc explains what each setting does and when you'd change it.

---

## SMTP / Email Settings

```python
SMTP_HOST = "email-smtp.ap-south-1.amazonaws.com"
SMTP_PORT = 587
USE_TLS = True
```

Amazon SES endpoint for the Mumbai region. Change the region prefix if your SES is provisioned elsewhere (e.g., `us-east-1`, `eu-west-1`).

```python
SENDER_EMAIL = "certificates@iitmparadox.org"
SENDER_NAME = "Saavan 26 - Team Paradox"
```

The "From" address and display name. `SENDER_EMAIL` must be a verified identity in your SES account — either a verified domain or a verified individual address.

```python
CREDENTIALS_FILE = os.path.join(BASE_DIR, "ses-smtp-user.20260930-234217_credentials.csv")
```

Path to the AWS IAM SMTP credentials CSV. This file has two columns: `SMTP user name` and `SMTP password`. The script reads them on import. If the file is missing or the columns are wrong, `SMTP_USERNAME` and `SMTP_PASSWORD` will be `None` and `send_emails.py` will crash with a clear error.

---

## Test Recipients

```python
PERMANENT_TEST_EMAIL = "24f2006473@ds.study.iitm.ac.in"
SECOND_TEST_EMAIL = "24f3004704@ds.study.iitm.ac.in"
RECIPIENT_EMAILS = [PERMANENT_TEST_EMAIL, SECOND_TEST_EMAIL]
```

Default delivery targets. When running `send_emails.py` without `--to`, emails go to both of these simultaneously. Change them to your own test addresses during development.

---

## File Paths

```python
DUMMY_DATA_CSV = os.path.join(BASE_DIR, "dummy_data.csv")
TEMPLATES_DIR = os.path.join(BASE_DIR, "template")
EMAIL_HTML_DIR = os.path.join(BASE_DIR, "email html")
GENERATED_CERTS_DIR = os.path.join(BASE_DIR, "generated_certificates")
PPTX_OUTPUT_DIR = os.path.join(GENERATED_CERTS_DIR, "pptx")
PDF_OUTPUT_DIR = os.path.join(GENERATED_CERTS_DIR, "pdf")
LOGS_DIR = os.path.join(BASE_DIR, "logs")
```

All paths are relative to `config.py`'s own directory (`BASE_DIR`). No hardcoded absolute paths — the whole `dummy/` folder is portable.

---

## Template Mappings

### PPTX Templates (for certificate generation)

```python
PPTX_TEMPLATES = {
    "winner":      "template/General_Event_Winner_Template.pptx",
    "participant": "template/General_Event_Participant_Template.pptx",
    "judge":       "template/Guest_Template.pptx",
    "guest":       "template/Guest_Template.pptx",
}
```

The key is the lowercase `type` column from the CSV. `judge` and `guest` share the same template.

### HTML Email Templates (for email body)

```python
EMAIL_TEMPLATES = {
    "winner":      "email html/saavan26_certificate_winner(1).html",
    "participant": "email html/saavan26_certificate_particiation.html",
    "judge":       "email html/saavan26_certificate_guest.html",
    "guest":       "email html/saavan26_certificate_guest.html",
}
```

Same routing logic. `judge` and `guest` share the same email template too.

---

## QR Code Coordinates

```python
QR_CONFIGS = {
    "winner":      {"left": 683.15, "top": 127.55, "size": 49.25},
    "participant": {"left": 683.15, "top": 145.74, "size": 49.25},
    "judge":       {"left": 683.15, "top": 89.08,  "size": 49.25},
    "guest":       {"left": 683.15, "top": 89.08,  "size": 49.25},
}
```

All values are in **typographical points** (1 pt = 1/72 inch). The script converts them to inches internally when placing the QR image on the slide:

```python
Inches(left_pt / 72.0)
```

These coordinates were calibrated by opening each template in PowerPoint, checking the exact position where the QR needs to land, and measuring in points using `inspect_pptx.py` and helper scripts.

### Why the `top` value differs per type

Each template has a different vertical layout. The winner cert has more text (position/rank line), so the QR sits higher. The participant cert pushes the QR lower. The guest cert has less content, so the QR goes near the top.

---

## Adding a New Certificate Type

1. Get the designed `.pptx` template from the design team.
2. Run `inspect_pptx.py` on it to find placeholder tags and shape positions.
3. Add the type key to all four dictionaries: `PPTX_TEMPLATES`, `EMAIL_TEMPLATES`, `QR_CONFIGS`.
4. Create an HTML email template in `email html/` with the appropriate `{{placeholders}}`.
5. Add rows with the new type in your CSV.
6. Run `generate_certificates.py` and `send_emails.py` as usual.
