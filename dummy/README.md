# saavan26/dummy

Working directory for the Saavan '26 certificate pipeline. Contains all code, data, templates, and output.

For full documentation, see [../README.md](../README.md) and [../docs/](../docs/).

---

## Quick Commands

```bash
# pre-flight data audit & firestore sync payload export
python validate_and_verify.py --csv dummy_data.csv --export-firestore firestore_payload.json

# inspect templates for placeholder tags
python inspect_pptx.py

# generate certificates (PPTX + PDF)
python generate_certificates.py

# reconcile compiled PDFs against CSV
python validate_and_verify.py --csv dummy_data.csv --check-pdfs

# send 1 test email
python send_emails.py --count 1

# send to a specific address
python send_emails.py --count 1 --to "you@example.com"

# send all
python send_emails.py --all
```

## Files

| File | Role |
|------|------|
| `validate_and_verify.py` | Audits dataset schema, RFC 4122 UUIDs, QR links & compiles Firestore sync payloads |
| `config.py` | Central config — SMTP credentials, paths, template mappings, QR coordinates |
| `generate_certificates.py` | Reads CSV → fills PPTX templates → inserts QR codes → exports PDFs via COM |
| `send_emails.py` | Reads CSV + PDFs → matches HTML email template → sends via Amazon SES |
| `inspect_pptx.py` | Diagnostic — dumps all text shapes from template files |
| `dummy_data.csv` | Sample data with 5 test entries across Winner/Participant/Judge |
| `ses-smtp-user.*.csv` | AWS SES SMTP credentials (gitignored) |

## Directories

| Directory | Contents |
|-----------|----------|
| `template/` | Master PPTX decks from the design team |
| `email html/` | Responsive HTML email templates (1 per certificate type) |
| `generated_certificates/pptx/` | Intermediate PPTX output |
| `generated_certificates/pdf/` | Final PDF output (gets attached to emails) |
| `logs/` | Email delivery audit log |
