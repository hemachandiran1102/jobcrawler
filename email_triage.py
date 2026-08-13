#!/usr/bin/env python3
"""
email_triage.py — Intelligent Email Triage & Next Steps Pipeline
================================================================
Connects securely to your inbox (Gmail, Outlook, Yahoo, IMAP), scans incoming
recruiter & hiring platform emails, and extracts actionable "Proceed to Next Steps"
responses (interview invitations, screening calls, coding assessments, availability requests).

Outputs:
  - Master Excel: interview_pipeline.xlsx (Action Required + Per-Category Tabs)
  - Master CSV  : interview_pipeline.csv
  - Master JSON : interview_pipeline.json (for web dashboard rendering)

Usage:
  python email_triage.py
  python email_triage.py --days 14 --provider gmail
  python email_triage.py --email your.email@gmail.com --password "xxxx xxxx xxxx xxxx"
"""

import os
import re
import sys
import json
import time
import email
import email.message
import email.utils
from email.header import decode_header
import imaplib
import argparse
import urllib.parse
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False


# ══════════════════════════════════════════════════════════════════════
# PATHS & CONFIGURATION
# ══════════════════════════════════════════════════════════════════════
_SCRIPT_DIR = Path(__file__).parent.resolve()
WORK_DIR = Path(os.environ.get("WORK_DIR", str(_SCRIPT_DIR)))
CONFIG_PATH = WORK_DIR / "email_config.json"

MASTER_EXCEL_PATH = WORK_DIR / "interview_pipeline.xlsx"
MASTER_CSV_PATH   = WORK_DIR / "interview_pipeline.csv"
MASTER_JSON_PATH  = WORK_DIR / "interview_pipeline.json"

# Scheduling & Meeting link domains
MEETING_LINK_PATTERNS = [
    r'https?://(?:www\.)?calendly\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?goodtime\.io/[^\s"\'<>]+',
    r'https?://(?:www\.)?savvycal\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?chilipiper\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?hubspot\.com/meetings/[^\s"\'<>]+',
    r'https?://(?:www\.)?tidycal\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?youcanbook\.me/[^\s"\'<>]+',
    r'https?://(?:[a-zA-Z0-9-]+\.)?greenhouse\.io/[^\s"\'<>]+',
    r'https?://(?:[a-zA-Z0-9-]+\.)?lever\.co/[^\s"\'<>]+',
    r'https?://(?:[a-zA-Z0-9-]+\.)?smartrecruiters\.com/[^\s"\'<>]+',
    r'https?://(?:[a-zA-Z0-9-]+\.)?workday\.com/[^\s"\'<>]+',
    r'https?://(?:[a-zA-Z0-9-]+\.)?ashbyhq\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?hackerrank\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?codility\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?testgorilla\.com/[^\s"\'<>]+',
    r'https?://(?:www\.)?hirevue\.com/[^\s"\'<>]+',
    r'https?://meet\.google\.com/[a-z]{3}-[a-z]{4}-[a-z]{3}',
    r'https?://(?:[a-zA-Z0-9-]+\.)?zoom\.us/j/[0-9]+[^\s"\'<>]*',
    r'https?://teams\.microsoft\.com/l/meetup-join/[^\s"\'<>]+',
]

KNOWN_TECH_ROLES = [
    "DevOps Engineer", "Senior DevOps Engineer", "Lead DevOps Engineer",
    "Cloud Engineer", "Senior Cloud Engineer", "Cloud Infrastructure Engineer", "Cloud Architect",
    "Site Reliability Engineer", "Senior SRE", "SRE Lead",
    "Platform Engineer", "Senior Platform Engineer", "Staff Platform Engineer",
    "Infrastructure Engineer", "Senior Infrastructure Engineer",
    "AWS DevOps Engineer", "Azure DevOps Engineer", "Kubernetes Engineer",
    "Software Engineer", "Senior Software Engineer", "Full Stack Engineer",
    "Systems Engineer", "Security Engineer", "Solutions Architect",
]


def log(msg: str, level: str = "INFO"):
    """Formatted timestamped console logger."""
    ts = datetime.now().strftime("%H:%M:%S")
    icons = {
        "INFO": "[i]", "OK": "[OK]", "WARN": "[!]",
        "ERROR": "[ERR]", "NEXT": "[NEXT]", "TEST": "[TEST]"
    }
    icon = icons.get(level, "[i]")
    safe = str(msg).encode("ascii", errors="replace").decode("ascii")
    print(f"[{ts}] {icon}  {safe}", flush=True)


def load_email_config() -> dict:
    """Load configuration from email_config.json or environment variables."""
    cfg = {
        "email": os.environ.get("EMAIL_USER", ""),
        "password": os.environ.get("EMAIL_PASSWORD", ""),
        "provider": "gmail",
        "imap_server": os.environ.get("IMAP_SERVER", "imap.gmail.com"),
        "imap_port": int(os.environ.get("IMAP_PORT", 993)),
        "use_ssl": True,
        "folders": ["INBOX"],
        "days_back": 14,
        "max_emails": 200,
        "mark_as_read": False,
    }

    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                disk_cfg = json.load(f)
                cfg.update({k: v for k, v in disk_cfg.items() if v is not None and v != ""})
        except Exception as e:
            log(f"Warning reading {CONFIG_PATH}: {e}", "WARN")

    # Resolve provider preset if server not explicitly specified
    provider = cfg.get("provider", "gmail").lower()
    presets = {
        "gmail": "imap.gmail.com",
        "outlook": "outlook.office365.com",
        "hotmail": "outlook.office365.com",
        "office365": "outlook.office365.com",
        "yahoo": "imap.mail.yahoo.com",
        "icloud": "imap.mail.me.com",
    }
    if not cfg.get("imap_server") and provider in presets:
        cfg["imap_server"] = presets[provider]

    return cfg


# ══════════════════════════════════════════════════════════════════════
# EMAIL DECODING & PARSING
# ══════════════════════════════════════════════════════════════════════
def decode_mime_header(header_val: str) -> str:
    """Decode RFC 2047 MIME encoded headers safely."""
    if not header_val:
        return ""
    decoded_parts = []
    try:
        parts = decode_header(header_val)
        for text, charset in parts:
            if isinstance(text, bytes):
                try:
                    charset = charset or "utf-8"
                    decoded_parts.append(text.decode(charset, errors="replace"))
                except Exception:
                    decoded_parts.append(text.decode("latin-1", errors="replace"))
            else:
                decoded_parts.append(str(text))
    except Exception:
        return str(header_val)
    return " ".join(decoded_parts).strip()


def extract_email_body(msg: email.message.Message) -> tuple[str, str]:
    """Extract plain text and HTML bodies from a multi-part email message."""
    plain_text = ""
    html_text = ""

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))

            if "attachment" in content_disposition:
                continue

            payload = part.get_payload(decode=True)
            if not payload:
                continue

            charset = part.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except Exception:
                decoded = payload.decode("latin-1", errors="replace")

            if content_type == "text/plain" and not plain_text:
                plain_text = decoded
            elif content_type == "text/html" and not html_text:
                html_text = decoded
    else:
        content_type = msg.get_content_type()
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except Exception:
                decoded = payload.decode("latin-1", errors="replace")

            if content_type == "text/html":
                html_text = decoded
            else:
                plain_text = decoded

    # Clean HTML to plain text fallback if text/plain is missing
    if not plain_text and html_text:
        # Strip HTML tags
        clean = re.sub(r'<style.*?>.*?</style>', '', html_text, flags=re.DOTALL | re.IGNORECASE)
        clean = re.sub(r'<script.*?>.*?</script>', '', clean, flags=re.DOTALL | re.IGNORECASE)
        clean = re.sub(r'<br\s*/?>', '\n', clean, flags=re.IGNORECASE)
        clean = re.sub(r'</p>', '\n\n', clean, flags=re.IGNORECASE)
        clean = re.sub(r'<[^>]+>', ' ', clean)
        clean = re.sub(r'&nbsp;', ' ', clean)
        clean = re.sub(r'&amp;', '&', clean)
        clean = re.sub(r'&lt;', '<', clean)
        clean = re.sub(r'&gt;', '>', clean)
        clean = re.sub(r'\n{3,}', '\n\n', clean)
        plain_text = clean.strip()

    return plain_text.strip(), html_text.strip()


def extract_action_links(plain_text: str, html_text: str) -> list[str]:
    """Extract scheduling, interview, and assessment links from email."""
    combined = f"{plain_text}\n{html_text}"
    links = []

    for pattern in MEETING_LINK_PATTERNS:
        matches = re.findall(pattern, combined, flags=re.IGNORECASE)
        for m in matches:
            clean_link = m.rstrip('.,;)"\'\\]>')
            if clean_link not in links:
                links.append(clean_link)

    # Also search for hrefs in HTML containing calendar keywords
    if html_text:
        href_matches = re.findall(r'href=[\'"](https?://[^\'"]+)[\'"]', html_text, flags=re.IGNORECASE)
        for href in href_matches:
            if any(k in href.lower() for k in ["schedule", "calendly", "goodtime", "interview", "assessment", "booking", "meet.google", "teams.microsoft", "zoom.us"]):
                if href not in links and not any(href.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".gif", ".css", ".js"]):
                    links.append(href)

    return links


# ══════════════════════════════════════════════════════════════════════
# CLASSIFICATION & TRIAGE ENGINE
# ══════════════════════════════════════════════════════════════════════
def extract_company_name(from_header: str, subject: str, body: str) -> str:
    """Extract hiring company name from email headers, domain, or subject."""
    # 1. From ATS Subject Patterns: "Interview with [Company]", "Your application at [Company]"
    subject_patterns = [
        r'(?:interview|chat|invitation|speaking)\s+(?:with|at)\s+([A-Z0-9][A-Za-z0-9\s&.,-]{2,30}?)(?:\s+for|\s+-|\s+:|\s+team|\s*$)',
        r'(?:your\s+application\s+(?:to|at|with))\s+([A-Z0-9][A-Za-z0-9\s&.,-]{2,30}?)(?:\s+for|\s+-|\s+:|\s*$)',
        r'(?:welcome\s+to)\s+([A-Z0-9][A-Za-z0-9\s&.,-]{2,30}?)(?:\s+team|\s+careers|\s+hiring|\s*$)',
        r'\[([A-Za-z0-9\s&.,-]{2,25})\]\s+(?:interview|application|next\s+step)',
        r'([A-Z0-9][A-Za-z0-9\s&.,-]{2,25})\s+(?:is\s+inviting\s+you|team\s+wants\s+to\s+connect)',
    ]
    for pat in subject_patterns:
        m = re.search(pat, subject, flags=re.IGNORECASE)
        if m:
            c = m.group(1).strip(" -:;,.")
            if len(c) >= 2 and not any(w in c.lower() for w in ["interview", "application", "update", "reminder", "job", "career"]):
                return c

    # 2. Extract from Sender Display Name: "John Doe from Stripe", "Google Careers"
    if "<" in from_header:
        display_name = from_header.split("<")[0].strip("\"' ")
        email_addr = from_header.split("<")[1].rstrip(">").strip()
    else:
        display_name = from_header
        email_addr = from_header

    # Pattern: "... from [Company]" or "... at [Company]"
    m_from = re.search(r'(?:from|at|@)\s+([A-Z0-9][A-Za-z0-9\s&.,-]{2,30})', display_name, flags=re.IGNORECASE)
    if m_from:
        return m_from.group(1).strip(" -:;,.")

    # 3. Extract from Domain Name (if not generic provider)
    domain_match = re.search(r'@([a-zA-Z0-9.-]+)\.([a-zA-Z]{2,})', email_addr)
    if domain_match:
        domain_name = domain_match.group(1).lower()
        generic_domains = ["gmail", "yahoo", "hotmail", "outlook", "live", "icloud", "mail", "greenhouse", "lever", "workday", "smartrecruiters", "ashbyhq", "breezy", "jazzhr", "recruitee"]
        if domain_name not in generic_domains and len(domain_name) >= 3:
            return domain_name.capitalize()

    return "Recruiter / Hiring Team"


def extract_job_title(subject: str, body: str) -> str:
    """Identify the target job title from subject or body snippet."""
    combined = f"{subject}\n{body[:500]}"
    for role in KNOWN_TECH_ROLES:
        if re.search(rf'\b{re.escape(role)}\b', combined, flags=re.IGNORECASE):
            return role

    # Regex fallback for "... for [Role Title] position"
    m = re.search(r'(?:for\s+(?:the\s+)?(?:position\s+of\s+|role\s+of\s+)?)([A-Z][A-Za-z0-9\s/&-]{3,35}?)(?:\s+role|\s+position|\s+at|\s+with|\s+-|\s*$|\n)', subject)
    if m:
        t = m.group(1).strip(" -:;,.")
        if len(t) >= 4 and not any(w in t.lower() for w in ["your", "next", "update", "application", "schedule"]):
            return t

    return "DevOps / Cloud Role"


def classify_email(subject: str, body: str, links: list[str]) -> dict:
    """
    Classify email intent and extract actionable next steps.
    Categories:
      - 🟢 Interview Invitation
      - 🔵 Technical Assessment
      - 🟡 Availability / Inquiry Request
      - ⚪ Application Confirmation
      - 🔴 Rejection
    """
    text = f"{subject}\n{body}".lower()

    # 1. High Priority: Interview Invitation / Scheduling Next Steps
    interview_signals = [
        "invitation to interview", "invite you to interview", "pleased to invite you",
        "schedule a call", "schedule an interview", "schedule a time",
        "book a time", "select a time slot", "choose a time",
        "like to speak with you", "like to chat with you", "love to connect",
        "next step in our process", "moving forward to the next round",
        "next round of interviews", "first round interview", "screening interview",
        "introductory call", "hiring manager interview", "technical interview round",
        "calendly.com", "goodtime.io", "savvycal.com", "schedule your interview",
    ]
    interview_score = sum(2 if sig in text else 0 for sig in interview_signals)
    if any(l for l in links if any(k in l.lower() for k in ["calendly", "goodtime", "savvycal", "chilipiper", "schedule", "meet.google", "teams.microsoft", "zoom.us"])):
        interview_score += 4

    # 2. Technical Assessment / Coding Test
    assessment_signals = [
        "technical assessment", "coding challenge", "coding test",
        "hackerrank", "codility", "testgorilla", "byteboard",
        "take-home test", "take home assessment", "practical test",
        "online assessment", "complete the assessment before", "assessment deadline",
        "hirevue", "coderbyte", "codewars",
    ]
    assessment_score = sum(2 if sig in text else 0 for sig in assessment_signals)
    if any(l for l in links if any(k in l.lower() for k in ["hackerrank", "codility", "testgorilla", "byteboard", "hirevue"])):
        assessment_score += 4

    # 3. Recruiter Inquiry / Availability Request
    inquiry_signals = [
        "please let us know your availability", "share your availability",
        "provide your availability", "what is your notice period",
        "expected salary", "salary expectations", "right to work",
        "visa sponsorship required", "share your updated cv",
        "could you confirm your current location", "share a few time slots",
    ]
    inquiry_score = sum(2 if sig in text else 0 for sig in inquiry_signals)

    # 4. Rejections
    rejection_signals = [
        "unfortunately", "not moving forward", "other candidates",
        "pursue other candidates", "decided not to proceed",
        "not selected for this role", "carefully reviewed your application",
        "at this time we have decided", "wish you all the best",
        "will not be progressing", "regret to inform",
    ]
    rejection_score = sum(2 if sig in text else 0 for sig in rejection_signals)

    # 5. Application Confirmation
    confirmation_signals = [
        "thank you for applying", "application received", "we have received your application",
        "application submitted", "thanks for applying", "your application has been received",
        "application confirmation", "successfully submitted",
    ]
    confirmation_score = sum(1 if sig in text else 0 for sig in confirmation_signals)

    # Determine Winner
    if rejection_score >= 3 and interview_score < 2 and assessment_score < 2:
        return {
            "category": "Rejection",
            "priority": "🔴 Low - Archived",
            "action_required": "Application not moving forward",
            "is_next_step": False,
            "badge": "🔴 Rejection"
        }

    if interview_score >= 2 or (links and interview_score >= 1):
        action = "Book interview slot via scheduling link" if links else "Reply with your availability to schedule interview"
        return {
            "category": "Interview Invitation",
            "priority": "🟢 High - Immediate Action",
            "action_required": action,
            "is_next_step": True,
            "badge": "🟢 Interview Invite"
        }

    if assessment_score >= 2:
        return {
            "category": "Technical Assessment",
            "priority": "🔵 High - Assessment",
            "action_required": "Complete online technical assessment before deadline",
            "is_next_step": True,
            "badge": "🔵 Assessment"
        }

    if inquiry_score >= 2:
        return {
            "category": "Availability / Inquiry",
            "priority": "🟡 Action Required",
            "action_required": "Reply to recruiter with requested availability / details",
            "is_next_step": True,
            "badge": "🟡 Availability Request"
        }

    if confirmation_score >= 2:
        return {
            "category": "Application Confirmation",
            "priority": "⚪ Info Only",
            "action_required": "Application acknowledged (pending review)",
            "is_next_step": False,
            "badge": "⚪ Confirmation"
        }

    # Default general message
    return {
        "category": "General Update",
        "priority": "⚪ General",
        "action_required": "Review email",
        "is_next_step": False,
        "badge": "⚪ General"
    }


# ══════════════════════════════════════════════════════════════════════
# INBOX SCANNER (IMAP)
# ══════════════════════════════════════════════════════════════════════
def scan_inbox(cfg: dict) -> list[dict]:
    """Connect to IMAP server and fetch relevant job application emails."""
    email_user = cfg.get("email", "").strip()
    email_pass = cfg.get("password", "").strip().replace(" ", "")
    imap_server = cfg.get("imap_server", "imap.gmail.com").strip()
    imap_port = int(cfg.get("imap_port", 993))
    days_back = int(cfg.get("days_back", 14))
    max_emails = int(cfg.get("max_emails", 200))
    folders = cfg.get("folders", ["INBOX"])

    if not email_user or not email_pass:
        log("Email credentials not configured.", "WARN")
        log(f"Please set your email and App Password in {CONFIG_PATH} or run with:", "INFO")
        log("  python email_triage.py --email your.email@gmail.com --password 'xxxx xxxx xxxx xxxx'", "INFO")
        return []

    log(f"Connecting to {imap_server}:{imap_port} as {email_user} …", "INFO")
    results = []

    try:
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        mail.login(email_user, email_pass)
        log("IMAP Login Successful!", "OK")
    except Exception as e:
        log(f"IMAP Login Failed: {e}", "ERROR")
        log("Tip: For Gmail, generate a 16-character App Password at https://myaccount.google.com/apppasswords", "INFO")
        return []

    since_date = (datetime.now() - timedelta(days=days_back)).strftime("%d-%b-%Y")

    for folder in folders:
        try:
            status, _ = mail.select(f'"{folder}"', readonly=True)
            if status != "OK":
                status, _ = mail.select(folder, readonly=True)
            if status != "OK":
                log(f"Could not open folder '{folder}' -- skipping.", "WARN")
                continue

            log(f"Scanning folder '{folder}' (emails since {since_date}) …", "INFO")

            # Search queries targeting job, interview, application keywords
            search_query = f'(SINCE "{since_date}")'
            typ, data = mail.search(None, search_query)

            if typ != "OK" or not data[0]:
                log(f"No emails found in {folder} since {since_date}.", "INFO")
                continue

            mail_ids = data[0].split()
            log(f"Found {len(mail_ids)} total emails in {folder}. Triaging latest {min(len(mail_ids), max_emails)} …", "INFO")

            # Process in reverse (newest first)
            for mid in reversed(mail_ids[-max_emails:]):
                try:
                    res, msg_data = mail.fetch(mid, "(RFC822)")
                    if res != "OK" or not msg_data or not msg_data[0]:
                        continue

                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)

                    subject = decode_mime_header(msg.get("Subject", ""))
                    from_header = decode_mime_header(msg.get("From", ""))
                    date_header = msg.get("Date", "")
                    message_id = msg.get("Message-ID", str(mid))

                    # Parse Date
                    try:
                        date_parsed = email.utils.parsedate_to_datetime(date_header)
                        date_str = date_parsed.strftime("%Y-%m-%d %H:%M")
                        date_sort = date_parsed.strftime("%Y-%m-%d")
                    except Exception:
                        date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                        date_sort = datetime.now().strftime("%Y-%m-%d")

                    plain_body, html_body = extract_email_body(msg)
                    links = extract_action_links(plain_body, html_body)

                    # Triage & Classify
                    classification = classify_email(subject, plain_body, links)
                    company = extract_company_name(from_header, subject, plain_body)
                    role = extract_job_title(subject, plain_body)

                    # Extract primary action link
                    primary_link = links[0] if links else ""

                    # Create clean snippet
                    snippet = plain_body[:300].replace('\n', ' ').strip()

                    record = {
                        "Message ID": message_id,
                        "Date Received": date_str,
                        "Date Sort": date_sort,
                        "Company": company,
                        "Job Title": role,
                        "Category": classification["category"],
                        "Priority": classification["priority"],
                        "Badge": classification["badge"],
                        "Action Required": classification["action_required"],
                        "Is Next Step": classification["is_next_step"],
                        "Action URL": primary_link,
                        "All Links": " | ".join(links),
                        "Sender": from_header,
                        "Subject": subject,
                        "Email Snippet": snippet,
                        "Replied Status": "No",
                        "Notes": "",
                    }

                    results.append(record)

                    if classification["is_next_step"]:
                        log(f"   {classification['badge']} | {company} — '{role}'", "OK")
                        if primary_link:
                            log(f"      🔗 Booking Link: {primary_link[:65]}...", "NEXT")

                except Exception as parse_err:
                    continue

        except Exception as folder_err:
            log(f"Error scanning {folder}: {folder_err}", "WARN")

    try:
        mail.close()
        mail.logout()
    except Exception:
        pass

    log(f"\nTriage Complete: Scanned {len(results)} emails.", "OK")
    return results


# ══════════════════════════════════════════════════════════════════════
# MASTER SPREADSHEETS & EXPORT
# ══════════════════════════════════════════════════════════════════════
def deduplicate_pipeline(new_records: list[dict]) -> pd.DataFrame:
    """Merge newly scanned records with existing pipeline and deduplicate."""
    if not new_records and not MASTER_CSV_PATH.exists():
        return pd.DataFrame()

    incoming_df = pd.DataFrame(new_records)

    if MASTER_CSV_PATH.exists():
        try:
            master_df = pd.read_csv(MASTER_CSV_PATH, on_bad_lines="skip")
            combined_df = pd.concat([master_df, incoming_df], ignore_index=True)
        except Exception:
            combined_df = incoming_df
    else:
        combined_df = incoming_df

    if combined_df.empty:
        return combined_df

    # Deduplicate by (Message ID) or (Company + Subject + Date)
    combined_df["_dedup"] = combined_df["Message ID"].fillna("")
    mask_empty = combined_df["_dedup"] == ""
    combined_df.loc[mask_empty, "_dedup"] = (
        combined_df.loc[mask_empty, "Company"].astype(str) + "|" +
        combined_df.loc[mask_empty, "Subject"].astype(str) + "|" +
        combined_df.loc[mask_empty, "Date Sort"].astype(str)
    )

    clean_df = combined_df.drop_duplicates(subset=["_dedup"], keep="first").copy()
    clean_df.drop(columns=["_dedup"], errors="ignore", inplace=True)

    # Sort by Date descending
    if "Date Received" in clean_df.columns:
        clean_df.sort_values(by="Date Received", ascending=False, inplace=True)

    return clean_df


def save_master_excel(df: pd.DataFrame):
    """Save clean multi-sheet Excel with Action Required + Per-Category Tabs."""
    if not HAS_EXCEL or df is None or df.empty:
        return

    log("Generating Master Excel: interview_pipeline.xlsx …", "INFO")
    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    cols = [
        "#", "Priority", "Category", "Company", "Job Title",
        "Action Required", "Action URL", "Date Received",
        "Sender", "Subject", "Email Snippet", "Replied Status", "Notes"
    ]

    hf = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hfnt = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))

    # Priority fills
    green_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
    blue_fill  = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")

    def populate_sheet(ws, sub_df):
        for ci, col in enumerate(cols, 1):
            c = ws.cell(row=1, column=ci, value=col)
            c.fill = hf; c.font = hfnt
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin

        for ri, (_, row) in enumerate(sub_df.iterrows(), 2):
            vals = [
                ri - 1,
                str(row.get("Priority", "")),
                str(row.get("Category", "")),
                str(row.get("Company", "")),
                str(row.get("Job Title", "")),
                str(row.get("Action Required", "")),
                str(row.get("Action URL", "")),
                str(row.get("Date Received", "")),
                str(row.get("Sender", "")),
                str(row.get("Subject", "")),
                str(row.get("Email Snippet", "")),
                str(row.get("Replied Status", "No")),
                str(row.get("Notes", ""))
            ]
            cat = str(row.get("Category", ""))
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.font = Font(name="Calibri", size=9)
                c.border = thin

                # Highlight action categories
                if cat == "Interview Invitation":
                    c.fill = green_fill
                elif cat == "Technical Assessment":
                    c.fill = blue_fill
                elif "Availability" in cat:
                    c.fill = yellow_fill

        # Auto column widths
        widths = [4, 20, 20, 22, 26, 35, 45, 18, 25, 35, 50, 14, 20]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        if len(sub_df) > 0:
            ws.auto_filter.ref = ws.dimensions

    # 1. Sheet: Action Required / Next Steps (High Priority)
    next_steps_df = df[df["Is Next Step"] == True] if "Is Next Step" in df.columns else df
    ws_action = wb.create_sheet(title="Action Required (Next Steps)")
    populate_sheet(ws_action, next_steps_df)

    # 2. Sheet: All Inbound Responses
    ws_all = wb.create_sheet(title="All Responses")
    populate_sheet(ws_all, df)

    # 3. Sheet: Interview Invites
    if "Category" in df.columns:
        invites_df = df[df["Category"] == "Interview Invitation"]
        if not invites_df.empty:
            ws_inv = wb.create_sheet(title="Interview Invites")
            populate_sheet(ws_inv, invites_df)

        # 4. Sheet: Assessments
        assess_df = df[df["Category"] == "Technical Assessment"]
        if not assess_df.empty:
            ws_ass = wb.create_sheet(title="Assessments")
            populate_sheet(ws_ass, assess_df)

    try:
        wb.save(MASTER_EXCEL_PATH)
        log(f"Master Excel saved -> {MASTER_EXCEL_PATH}", "OK")
    except PermissionError:
        alt = WORK_DIR / f"interview_pipeline_{int(time.time())}.xlsx"
        wb.save(alt)
        log(f"Excel file locked -- saved to {alt}", "WARN")


def save_outputs(clean_df: pd.DataFrame):
    """Save CSV, Excel, and JSON outputs for the web dashboard."""
    if clean_df.empty:
        return

    # 1. Master CSV
    clean_df.to_csv(MASTER_CSV_PATH, index=False)
    log(f"Master CSV updated  -> {MASTER_CSV_PATH} ({len(clean_df)} total triaged emails)", "OK")

    # 2. Master Excel
    save_master_excel(clean_df)

    # 3. Master JSON (for UI Dashboard)
    records = clean_df.to_dict(orient="records")
    with open(MASTER_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)
    log(f"Dashboard JSON saved -> {MASTER_JSON_PATH}", "OK")


# ══════════════════════════════════════════════════════════════════════
# CLI ENTRY POINT
# ══════════════════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description="Email Triage & Interview Next Steps Pipeline")
    parser.add_argument("--email", type=str, default="", help="Your email address")
    parser.add_argument("--password", type=str, default="", help="Email App Password")
    parser.add_argument("--provider", type=str, default="", help="Email provider: gmail, outlook, yahoo, icloud")
    parser.add_argument("--server", type=str, default="", help="Custom IMAP server host")
    parser.add_argument("--port", type=int, default=993, help="IMAP SSL Port (default: 993)")
    parser.add_argument("--days", type=int, default=14, help="Number of past days to scan (default: 14)")
    parser.add_argument("--max", type=int, default=200, help="Max emails to scan (default: 200)")
    parser.add_argument("--folder", type=str, default="INBOX", help="IMAP Folder (default: INBOX)")
    args = parser.parse_args()

    cfg = load_email_config()

    # CLI overrides
    if args.email: cfg["email"] = args.email
    if args.password: cfg["password"] = args.password
    if args.provider: cfg["provider"] = args.provider
    if args.server: cfg["imap_server"] = args.server
    if args.port: cfg["imap_port"] = args.port
    if args.days: cfg["days_back"] = args.days
    if args.max: cfg["max_emails"] = args.max
    if args.folder: cfg["folders"] = [args.folder]

    log("=" * 65)
    log("  EMAIL TRIAGE & INTERVIEW PIPELINE — Job Application Next Steps")
    log(f"   Provider    : {cfg.get('provider', 'gmail').upper()} ({cfg.get('imap_server', '')})")
    log(f"   User Email  : {cfg.get('email', '(Not set in config)')}")
    log(f"   Time Window : Past {cfg.get('days_back', 14)} Days")
    log(f"   Folder      : {', '.join(cfg.get('folders', ['INBOX']))}")
    log("=" * 65)

    raw_records = scan_inbox(cfg)
    clean_df = deduplicate_pipeline(raw_records)

    if not clean_df.empty:
        save_outputs(clean_df)

        next_steps = clean_df[clean_df["Is Next Step"] == True] if "Is Next Step" in clean_df.columns else clean_df
        invites = clean_df[clean_df["Category"] == "Interview Invitation"]
        assessments = clean_df[clean_df["Category"] == "Technical Assessment"]
        inquiries = clean_df[clean_df["Category"] == "Availability / Inquiry"]

        log("\n" + "=" * 65)
        log("  EMAIL TRIAGE SUMMARY", "OK")
        log(f"   🟢 Interview Invitations : {len(invites)}")
        log(f"   🔵 Technical Assessments : {len(assessments)}")
        log(f"   🟡 Availability Requests : {len(inquiries)}")
        log(f"   ⭐ Total Next Steps Ready: {len(next_steps)}")
        log(f"   📊 Master Excel Workbook : {MASTER_EXCEL_PATH}")
        log(f"   📄 Master CSV Database   : {MASTER_CSV_PATH}")
        log("=" * 65)
    else:
        log("No email records to save. Configure your credentials in email_config.json to run a live scan.", "WARN")


if __name__ == "__main__":
    main()
