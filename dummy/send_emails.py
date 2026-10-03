import os
import sys
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
from email.header import Header
from datetime import datetime
import pandas as pd

from config import (
    SMTP_HOST,
    SMTP_PORT,
    SMTP_USERNAME,
    SMTP_PASSWORD,
    SENDER_EMAIL,
    SENDER_NAME,
    PERMANENT_TEST_EMAIL,
    SECOND_TEST_EMAIL,
    RECIPIENT_EMAILS,
    DUMMY_DATA_CSV,
    EMAIL_TEMPLATES,
    PDF_OUTPUT_DIR,
    LOGS_DIR
)

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        pass

def render_email_html(row):
    """Renders the appropriate HTML email template with candidate placeholders."""
    cert_type = str(row.get('type', 'winner')).strip().lower()
    tmpl_path = EMAIL_TEMPLATES.get(cert_type, EMAIL_TEMPLATES['winner'])
    
    if not os.path.exists(tmpl_path):
        raise FileNotFoundError(f"HTML template not found: {tmpl_path}")
        
    with open(tmpl_path, "r", encoding="utf-8") as f:
        html = f.read()
        
    name = str(row.get('name', ''))
    event_name = str(row.get('event_name', ''))
    position = str(row.get('position', ''))
    cert_id = str(row.get('cert_id', ''))
    qr_code = str(row.get('qr_code', 'https://saavan.iitmparadox.org/'))
    feedback_link = str(row.get('feedback_link', 'https://forms.gle/saavan26feedback'))
    discrepancy_link = str(row.get('discrepancy_link', 'https://forms.gle/saavan26discrepancy'))
    winning_prize = str(row.get('winning_prize', 'N/A')) if pd.notna(row.get('winning_prize')) else 'N/A'
    
    cert_type_display = "Participation" if position.lower() == "participant" else position

    replacements = {
        '{{StudentName}}': name,
        '{{GuestName}}': name,
        '{{EventName}}': event_name,
        '{{Position}}': position,
        '{{CertificateType}}': cert_type_display,
        '{{WinningPrize}}': winning_prize,
        '{{CertificateID}}': cert_id,
        '{{CertificateLink}}': qr_code,
        '{{FeedbackLink}}': feedback_link,
        '{{DiscrepancyFormLink}}': discrepancy_link
    }
    
    for placeholder, val in replacements.items():
        html = html.replace(placeholder, val)
        
    return html

def build_email_message(row, pdf_path, recipients):
    """Builds a multipart email message with HTML body, plain text fallback, and PDF certificate attachment."""
    msg = MIMEMultipart("mixed")
    
    student_name = row.get("name", "")
    event_name = row.get("event_name", "")
    position = row.get("position", "")
    cert_id = row.get("cert_id", "")
    original_email = row.get("email", "")
    cert_type = str(row.get("type", "winner")).strip().lower()

    if cert_type == "winner":
        subject = f"Saavan '26 - Certificate of Achievement ({position}) - {event_name} - {student_name}"
    elif cert_type == "participant":
        subject = f"Saavan '26 - Certificate of Participation - {event_name} - {student_name}"
    elif cert_type in ["judge", "guest"]:
        subject = f"Saavan '26 - Certificate of Appreciation (Judge) - {event_name} - {student_name}"
    else:
        subject = f"Saavan '26 - Certificate - {event_name} - {student_name}"

    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL))
    
    if isinstance(recipients, list):
        msg["To"] = ", ".join(recipients)
    else:
        msg["To"] = str(recipients).strip()
    
    # Custom headers for traceability
    msg["X-Original-Recipient"] = original_email
    msg["X-Certificate-ID"] = cert_id
    msg["X-Certificate-Type"] = cert_type
    msg["X-Test-Mode"] = "True"

    # Alternative part (plain text + html)
    alt_part = MIMEMultipart("alternative")

    plain_text = (
        f"Greetings, {student_name}!\n\n"
        f"Thank you for being part of Saavan '26! It was truly wonderful to have you with us, "
        f"bringing your energy, passion, and enthusiasm to make this edition an unforgettable celebration.\n\n"
        f"✦ Event: {event_name}\n"
        f"✦ Position / Role: {position}\n"
        f"✦ Certificate ID: {cert_id}\n\n"
        f"Please find your official certificate attached to this email as a PDF.\n\n"
        f"Certificate Verification: {row.get('qr_code', 'https://saavan.iitmparadox.org/')}\n"
        f"Feedback Form: {row.get('feedback_link', 'https://forms.gle/saavan26feedback')}\n\n"
        f"Need help or notice a discrepancy?\n"
        f"Report form: {row.get('discrepancy_link', 'https://forms.gle/saavan26discrepancy')}\n"
        f"Support: support@iitmparadox.org\n\n"
        f"With warmth & celebration,\n"
        f"Team Paradox, IIT Madras BS Degree"
    )
    alt_part.attach(MIMEText(plain_text, "plain", "utf-8"))

    html_content = render_email_html(row)
    alt_part.attach(MIMEText(html_content, "html", "utf-8"))

    msg.attach(alt_part)

    # Attach PDF Certificate
    if os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
            pdf_attachment = MIMEApplication(pdf_data, _subtype="pdf")
            safe_name = "".join(c for c in student_name if c.isalnum() or c in (" ", "_", "-")).replace(" ", "_")
            filename = f"Certificate_{cert_id}_{safe_name}.pdf"
            pdf_attachment.add_header("Content-Disposition", "attachment", filename=filename)
            msg.attach(pdf_attachment)
    else:
        print(f"⚠️ Warning: PDF file '{pdf_path}' not found. Email will be sent without attachment.")

    return msg

def send_all_test_emails(limit=None, to_emails_override=None):
    """Sends certificate emails to both test email addresses simultaneously or an override."""
    if to_emails_override:
        if isinstance(to_emails_override, str):
            delivery_recipients = [e.strip() for e in to_emails_override.split(",") if e.strip()]
        else:
            delivery_recipients = to_emails_override
    else:
        delivery_recipients = RECIPIENT_EMAILS

    print("=" * 70)
    print("🚀 INITIALIZING SAAVAN '26 DUAL-RECIPIENT EMAIL DISPATCH")
    print(f"📌 Target Delivery Recipients: {', '.join(delivery_recipients)}")
    print(f"📡 SMTP Server Endpoint      : {SMTP_HOST}:{SMTP_PORT} (TLS)")
    print(f"✉️  Verified Sender Identity   : {SENDER_NAME} <{SENDER_EMAIL}>")
    print("=" * 70)

    if not SMTP_USERNAME or not SMTP_PASSWORD:
        raise ValueError("SMTP credentials missing! Check ses-smtp-user credentials file.")

    if not os.path.exists(DUMMY_DATA_CSV):
        raise FileNotFoundError(f"Data file '{DUMMY_DATA_CSV}' not found.")

    df = pd.read_csv(DUMMY_DATA_CSV)
    if limit:
        df = df.head(limit)
    print(f"📋 Loaded {len(df)} candidate records to dispatch.")

    os.makedirs(LOGS_DIR, exist_ok=True)
    log_csv_path = os.path.join(LOGS_DIR, "email_delivery_log.csv")

    logs = []

    print("\n🔌 Establishing secure TLS connection to Amazon SES...")
    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20)
    server.ehlo()
    server.starttls(context=ssl.create_default_context())
    server.ehlo()
    server.login(SMTP_USERNAME, SMTP_PASSWORD)
    print("✅ Authenticated successfully with Amazon SES SMTP.\n")

    try:
        for idx, row in df.iterrows():
            cert_id = row['cert_id']
            name = row['name']
            event_name = row['event_name']
            position = row['position']
            original_email = row['email']
            cert_type = str(row['type']).strip()
            
            pdf_path = os.path.join(PDF_OUTPUT_DIR, f"{cert_id}.pdf")
            
            email_msg = build_email_message(row, pdf_path, delivery_recipients)
            
            print(f"📤 [{idx+1}/{len(df)}] Sending certificate email:")
            print(f"   Candidate    : {name} (Type: {cert_type}, Orig: {original_email})")
            print(f"   Event & Pos  : {event_name} - {position}")
            print(f"   Certificate  : {cert_id} ({os.path.basename(pdf_path)})")
            print(f"   Template     : {os.path.basename(EMAIL_TEMPLATES.get(cert_type.lower(), ''))}")
            print(f"   Delivered To : {', '.join(delivery_recipients)}")

            start_t = datetime.now()
            status = "SUCCESS"
            error_msg = ""
            try:
                server.sendmail(SENDER_EMAIL, delivery_recipients, email_msg.as_string())
                print("   Status       : ✅ Sent successfully to both addresses!")
            except Exception as e:
                status = "FAILED"
                error_msg = str(e)
                print(f"   Status       : ❌ Delivery failed: {e}")

            logs.append({
                "timestamp": datetime.now().isoformat(),
                "cert_id": cert_id,
                "type": cert_type,
                "name": name,
                "event_name": event_name,
                "position": position,
                "original_email": original_email,
                "delivered_to": ", ".join(delivery_recipients),
                "status": status,
                "error": error_msg
            })

    finally:
        server.quit()
        print("\n🔒 Closed SMTP connection.")

    # Save log to CSV
    log_df = pd.DataFrame(logs)
    if os.path.exists(log_csv_path):
        log_df.to_csv(log_csv_path, mode="a", header=False, index=False)
    else:
        log_df.to_csv(log_csv_path, index=False)
    print(f"📊 Delivery log saved to: '{log_csv_path}'")
    print("=" * 70)
    print("🎉 Email dispatch routine completed!")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Dispatch Saavan '26 certificate emails via Amazon SES")
    parser.add_argument("--count", type=int, default=None, help="Number of test emails to send")
    parser.add_argument("--all", action="store_true", help="Send to all entries in dummy dataset")
    parser.add_argument("--to", type=str, default=None, help="Comma-separated recipient override (default: both test emails)")
    args = parser.parse_args()

    limit = None if args.all else args.count
    send_all_test_emails(limit=limit, to_emails_override=args.to)
