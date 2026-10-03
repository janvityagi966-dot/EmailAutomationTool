# email_automation.py
"""
Email Automation Tool (Windows-friendly)
- No Unix shebang (works in PowerShell/CMD)
- Graceful fallback if jinja2 or python-dotenv are missing
- Send single or CSV-driven bulk emails via SMTP (TLS/SSL)
- Supports HTML templates (Jinja2 if available), attachments, rate limiting, retries, dry-run
"""

import csv
import os
import time
import smtplib
import ssl
import logging
import argparse
from email.message import EmailMessage
from email.utils import make_msgid
from typing import List, Dict, Optional

# Optional dependencies: jinja2 and python-dotenv
HAS_JINJA2 = True
HAS_DOTENV = True
try:
    from importlib import import_module

    _jinja2 = import_module("jinja2")
    Environment = _jinja2.Environment
    FileSystemLoader = _jinja2.FileSystemLoader
    select_autoescape = _jinja2.select_autoescape
except Exception:
    HAS_JINJA2 = False

try:
    _dotenv = import_module("dotenv")
    load_dotenv = _dotenv.load_dotenv
except Exception:
    HAS_DOTENV = False

# If dotenv is available, load .env
if HAS_DOTENV:
    load_dotenv()

# Configuration (environment variables or defaults)
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.example.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")
USE_SSL = os.getenv("USE_SSL", "false").lower() in ("1", "true", "yes")
FROM_NAME = os.getenv("FROM_NAME", "Sender Name")
FROM_EMAIL = os.getenv("FROM_EMAIL", SMTP_USER or "no-reply@example.com")
RATE_LIMIT = float(os.getenv("RATE_LIMIT", "1.0"))  # seconds between sends
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "3"))
RETRY_DELAY = float(os.getenv("RETRY_DELAY", "5.0"))  # seconds

# Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("email_automation")

# Jinja2 environment if available
TEMPLATE_DIR = os.getenv("TEMPLATE_DIR", "templates")
if HAS_JINJA2:
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=select_autoescape(["html", "xml"]))
else:
    env = None  # fallback will be used


def create_message(subject: str,
                   to_email: str,
                   to_name: Optional[str] = None,
                   plain_text: Optional[str] = None,
                   html_template: Optional[str] = None,
                   template_vars: Optional[Dict] = None,
                   attachments: Optional[List[str]] = None,
                   inline_images: Optional[Dict[str, str]] = None) -> EmailMessage:
    """
    Build an EmailMessage with optional HTML template and attachments.
    If Jinja2 is not available, html_template is treated as a raw HTML filename read from TEMPLATE_DIR.
    inline_images: dict mapping cid_name -> file_path
    """
    msg = EmailMessage()
    from_header = f"{FROM_NAME} <{FROM_EMAIL}>"
    msg["From"] = from_header
    msg["To"] = f"{to_name} <{to_email}>" if to_name else to_email
    msg["Subject"] = subject

    # Plain text body
    if plain_text:
        msg.set_content(plain_text)

    # HTML body via Jinja2 if available, else read file and do simple replacement
    if html_template:
        template_vars = template_vars or {}
        html_body = None
        if HAS_JINJA2 and env:
            try:
                template = env.get_template(html_template)
                html_body = template.render(**template_vars)
            except Exception as e:
                logger.warning("Jinja2 template render failed: %s. Falling back to raw file read.", e)
        if html_body is None:
            # Fallback: try to read the file and perform simple {var} replacement
            path = os.path.join(TEMPLATE_DIR, html_template)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    raw = f.read()
                # simple Python format replacement
                try:
                    html_body = raw.format(**template_vars)
                except Exception:
                    html_body = raw
            except Exception as e:
                logger.warning("Could not read HTML template %s: %s", path, e)
                html_body = None

        if html_body:
            if not plain_text:
                msg.set_content("This message contains HTML content. Please view in an HTML-capable client.")
            msg.add_alternative(html_body, subtype="html")

            # Attach inline images if provided
            if inline_images:
                for cid_name, img_path in inline_images.items():
                    try:
                        with open(img_path, "rb") as f:
                            img_data = f.read()
                        ext = os.path.splitext(img_path)[1].lstrip(".").lower()
                        subtype = "jpeg" if ext in ("jpg", "jpeg") else ext or "octet-stream"
                        cid = make_msgid(domain="example.com")
                        # Attach to the last (HTML) part
                        msg.get_payload()[-1].add_related(img_data, maintype="image", subtype=subtype, cid=cid)
                        # expose cid to template_vars if needed
                        template_vars[cid_name] = cid[1:-1]
                    except Exception as e:
                        logger.warning("Could not attach inline image %s: %s", img_path, e)

    # Attach files
    if attachments:
        for path in attachments:
            if not path:
                continue
            try:
                with open(path, "rb") as f:
                    data = f.read()
                filename = os.path.basename(path)
                ext = os.path.splitext(filename)[1].lower()
                if ext in (".jpg", ".jpeg"):
                    maintype, subtype = "image", "jpeg"
                elif ext == ".png":
                    maintype, subtype = "image", "png"
                elif ext == ".pdf":
                    maintype, subtype = "application", "pdf"
                else:
                    maintype, subtype = "application", "octet-stream"
                msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
            except Exception as e:
                logger.warning("Failed to attach %s: %s", path, e)

    return msg


def send_message(msg: EmailMessage, dry_run: bool = False) -> None:
    """
    Send a single EmailMessage with retries and logging.
    """
    if dry_run:
        logger.info("Dry-run enabled. Email not sent. Subject: %s To: %s", msg["Subject"], msg["To"])
        return

    attempt = 0
    while attempt <= MAX_RETRIES:
        try:
            if USE_SSL:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
                    if SMTP_USER and SMTP_PASS:
                        server.login(SMTP_USER, SMTP_PASS)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                    server.ehlo()
                    server.starttls(context=ssl.create_default_context())
                    server.ehlo()
                    if SMTP_USER and SMTP_PASS:
                        server.login(SMTP_USER, SMTP_PASS)
                    server.send_message(msg)
            logger.info("Email sent: Subject=%s To=%s", msg["Subject"], msg["To"])
            return
        except Exception as e:
            attempt += 1
            logger.error("Send failed (attempt %d/%d): %s", attempt, MAX_RETRIES, e)
            if attempt > MAX_RETRIES:
                logger.exception("Max retries reached. Giving up on this message.")
                raise
            time.sleep(RETRY_DELAY)


def send_bulk_from_csv(csv_path: str,
                       subject_template: str,
                       html_template: Optional[str] = None,
                       plain_template: Optional[str] = None,
                       attachments_column: Optional[str] = None,
                       dry_run: bool = False) -> None:
    """
    CSV expected columns: email, name (optional), plus any template variables.
    attachments_column: name of CSV column that contains semicolon-separated file paths.
    """
    if not os.path.exists(csv_path):
        logger.error("CSV file not found: %s", csv_path)
        return

    with open(csv_path, newline="", encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            to_email = row.get("email")
            if not to_email:
                logger.warning("Skipping row without email: %s", row)
                continue
            to_name = row.get("name") or None
            template_vars = dict(row)

            attachments = None
            if attachments_column and row.get(attachments_column):
                attachments = [p.strip() for p in row[attachments_column].split(";") if p.strip()]

            # Render subject and plain text using Python format with safe fallback
            try:
                subject = subject_template.format(**template_vars)
            except Exception:
                subject = subject_template

            plain_text = None
            if plain_template:
                try:
                    plain_text = plain_template.format(**template_vars)
                except Exception:
                    plain_text = plain_template

            msg = create_message(
                subject=subject,
                to_email=to_email,
                to_name=to_name,
                plain_text=plain_text,
                html_template=html_template,
                template_vars=template_vars,
                attachments=attachments
            )

            try:
                send_message(msg, dry_run=dry_run)
            except Exception:
                logger.exception("Failed to send to %s", to_email)

            time.sleep(RATE_LIMIT)


def send_single_example(dry_run: bool = True):
    subject = "Hello from Email Automation"
    to_email = "recipient@example.com"
    to_name = "Recipient Name"
    plain_text = "Hi there,\n\nThis is a test email from the automation tool.\n\nBest,\nTeam"
    html_template = None  # set to a filename in templates/ if you have one
    template_vars = {"first_name": "Recipient", "promo_code": "WELCOME10"}

    msg = create_message(
        subject=subject,
        to_email=to_email,
        to_name=to_name,
        plain_text=plain_text,
        html_template=html_template,
        template_vars=template_vars,
        attachments=None
    )
    send_message(msg, dry_run=dry_run)


def check_optional_dependencies():
    missing = []
    if not HAS_JINJA2:
        missing.append("jinja2")
    if not HAS_DOTENV:
        missing.append("python-dotenv")
    if missing:
        logger.info("Optional packages missing: %s", ", ".join(missing))
        logger.info("Install them with: python -m pip install %s", " ".join(missing))


def main():
    parser = argparse.ArgumentParser(description="Email Automation Tool (Windows-friendly)")
    parser.add_argument("--csv", help="Path to CSV file for bulk send")
    parser.add_argument("--subject", help="Subject template (use {var} placeholders)", required=False)
    parser.add_argument("--html", help="HTML template filename in templates/ (Jinja2 or raw file)", required=False)
    parser.add_argument("--plain", help="Plain text template string (use {var})", required=False)
    parser.add_argument("--attachments-column", help="CSV column name for attachments", required=False)
    parser.add_argument("--dry-run", action="store_true", help="Do not actually send emails")
    parser.add_argument("--single", action="store_true", help="Send a single example email (edit code to change recipient)")
    args = parser.parse_args()

    check_optional_dependencies()

    if args.single:
        send_single_example(dry_run=args.dry_run)
        return

    if args.csv:
        if not args.subject:
            parser.error("--subject is required when using --csv")
        send_bulk_from_csv(
            csv_path=args.csv,
            subject_template=args.subject,
            html_template=args.html,
            plain_template=args.plain,
            attachments_column=args.attachments_column,
            dry_run=args.dry_run
        )
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
