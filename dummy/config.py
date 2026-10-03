import os
import csv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, "ses-smtp-user.20260930-234217_credentials.csv")

# SMTP Configuration (Amazon SES ap-south-1)
SMTP_HOST = "email-smtp.ap-south-1.amazonaws.com"
SMTP_PORT = 587
USE_TLS = True

# Read credentials from CSV
SMTP_USERNAME = None
SMTP_PASSWORD = None

if os.path.exists(CREDENTIALS_FILE):
    with open(CREDENTIALS_FILE, mode="r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            SMTP_USERNAME = row.get("SMTP user name", "").strip()
            SMTP_PASSWORD = row.get("SMTP password", "").strip()
            break

# Sender identity (Verified on AWS SES)
SENDER_EMAIL = "certificates@iitmparadox.org"
SENDER_NAME = "Saavan 26 - Team Paradox"


# Test Email Delivery Addresses (both at once)
PERMANENT_TEST_EMAIL = "24f2006473@ds.study.iitm.ac.in"
SECOND_TEST_EMAIL = "24f3004704@ds.study.iitm.ac.in"
RECIPIENT_EMAILS = [PERMANENT_TEST_EMAIL, SECOND_TEST_EMAIL]

# File & Directory Paths
DUMMY_DATA_CSV = os.path.join(BASE_DIR, "dummy_data.csv")
TEMPLATES_DIR = os.path.join(BASE_DIR, "template")
EMAIL_HTML_DIR = os.path.join(BASE_DIR, "email html")

PPTX_TEMPLATES = {
    "winner": os.path.join(TEMPLATES_DIR, "General_Event_Winner_Template.pptx"),
    "participant": os.path.join(TEMPLATES_DIR, "General_Event_Participant_Template.pptx"),
    "judge": os.path.join(TEMPLATES_DIR, "Guest_Template.pptx"),
    "guest": os.path.join(TEMPLATES_DIR, "Guest_Template.pptx"),
}

EMAIL_TEMPLATES = {
    "winner": os.path.join(EMAIL_HTML_DIR, "saavan26_certificate_winner(1).html"),
    "participant": os.path.join(EMAIL_HTML_DIR, "saavan26_certificate_particiation.html"),
    "judge": os.path.join(EMAIL_HTML_DIR, "saavan26_certificate_guest.html"),
    "guest": os.path.join(EMAIL_HTML_DIR, "saavan26_certificate_guest.html"),
}

QR_CONFIGS = {
    "winner": {"left": 683.15, "top": 127.55, "size": 49.25},
    "participant": {"left": 683.15, "top": 145.74, "size": 49.25},
    "judge": {"left": 683.15, "top": 89.08, "size": 49.25},
    "guest": {"left": 683.15, "top": 89.08, "size": 49.25},
}

PPTX_TEMPLATE_FILE = PPTX_TEMPLATES["winner"]
GENERATED_CERTS_DIR = os.path.join(BASE_DIR, "generated_certificates")
PPTX_OUTPUT_DIR = os.path.join(GENERATED_CERTS_DIR, "pptx")
PDF_OUTPUT_DIR = os.path.join(GENERATED_CERTS_DIR, "pdf")
LOGS_DIR = os.path.join(BASE_DIR, "logs")

# QR Code layout coordinates (Pt)
QR_LEFT = 683.15
QR_TOP = 127.55
QR_SIZE = 49.25
