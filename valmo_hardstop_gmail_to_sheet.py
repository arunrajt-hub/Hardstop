"""
Valmo Hardstop Gmail → Google Sheet Automation

Runs when the daily "[IMP] Valmo Control Tower!!! DD-MM-YYYY" email arrives from
lsn-meesho-central@loadshare.net. Downloads the hardstop attachment, filters
columns and locations, and pushes to Google Sheets.

Email: [IMP] Valmo Control Tower!!! 23-03-2026 (date varies daily)
From: lsn-meesho-central@loadshare.net

Attachments:
  1. hardstop_lsn-meesho-central@loadshare.net → Hardstop worksheet
     Columns: Date, awb, shipment_type, current_movement_type, current_status, location,
              shipment_value, location_ageing, location_age_bucket
  2. lost_lsn-meesho-central@loadshare.net → LostMarked worksheet
     Columns: Date, lost_date, awd, current_movement_type, loss_value, location

Locations to filter: MQR, MQE, YLG, YLZ, MHK

Output: Same Google Sheet, worksheets "Hardstop" and "LostMarked"
        https://docs.google.com/spreadsheets/d/1qnqzVf-S41F4S6DN8CRtXVgk-BcsaW377aVVEyFrnzg

Usage:
    python valmo_hardstop_gmail_to_sheet.py

    # Manual run with local files (skip Gmail):
    python valmo_hardstop_gmail_to_sheet.py --file path/to/hardstop.xlsx
    python valmo_hardstop_gmail_to_sheet.py --lost-file path/to/lost.xlsx

Environment:
    GMAIL_EMAIL        - Gmail address to read from (default: arunraj@loadshare.net)
    GMAIL_APP_PASSWORD - Gmail App Password (create at myaccount.google.com/apppasswords)
    WHAPI_TOKEN        - WhatsApp API token (same as 4D Active / whatsapp_sheet_image)
    WHATSAPP_PHONE     - WhatsApp recipient(s), comma-separated
    .env file          - Place next to script; load_dotenv() loads WHAPI_TOKEN etc.

Scheduling (run daily after the Valmo email arrives):
    Windows Task Scheduler: Run run_valmo_hardstop.bat at 9:00 AM daily
    Or: python valmo_hardstop_gmail_to_sheet.py (via cron/Task Scheduler)
"""

import argparse
import email
import imaplib
import logging
import os
import re
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None

try:
    import whatsapp_sheet_image as _wsi
    send_sheet_range_to_whatsapp = _wsi.send_sheet_range_to_whatsapp
except ImportError:
    send_sheet_range_to_whatsapp = None

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# Configuration
SCRIPT_DIR = Path(__file__).resolve().parent
SERVICE_ACCOUNT_FILE = SCRIPT_DIR / "service_account_key.json"

GMAIL_EMAIL = os.getenv("GMAIL_EMAIL", "arunraj@loadshare.net")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")

SPREADSHEET_ID = "1qnqzVf-S41F4S6DN8CRtXVgk-BcsaW377aVVEyFrnzg"
HARDSTOP_WORKSHEET = "Hardstop"
LOSTMARKED_WORKSHEET = "LostMarked"

# Email filter
SENDER_EMAIL = "lsn-meesho-central@loadshare.net"
SUBJECT_PATTERN = re.compile(r"\[IMP\]\s*Valmo\s*Control\s*Tower!!!\s*\d{2}-\d{2}-\d{4}", re.I)
SUBJECT_DATE_PATTERN = re.compile(r"(\d{2}-\d{2}-\d{4})\s*$")
HARDSTOP_ATTACHMENT_PATTERN = re.compile(r"hardstop_lsn-meesho-central@loadshare\.net", re.I)
LOST_ATTACHMENT_PATTERN = re.compile(r"lost_lsn-meesho-central@loadshare\.net", re.I)

# Columns to keep (case-insensitive match)
COLUMNS_TO_KEEP = [
    "awb",
    "shipment_type",
    "current_movement_type",
    "current_status",
    "location",
    "shipment_value",
    "location_ageing",
    "location_age_bucket",
]

# Lost attachment columns (Date is added from email subject, not from file)
# Output order: Date, lost_date, awd (3rd data column), then movement / loss / location
LOST_COLUMNS_TO_KEEP = [
    "lost_date",
    "awd",
    "current_movement_type",
    "loss_value",
    "location",
]

# Map output column name → source header names to try (case-insensitive via find_column)
LOST_COLUMN_ALIASES = {
    "awd": ("awd", "awb"),
}

# Expected output columns for LostMarked (Date first, from email subject)
LOSTMARKED_HEADERS = ["Date"] + LOST_COLUMNS_TO_KEEP

# Locations to filter
TARGET_LOCATIONS = {"MQR", "MQE", "YLG", "YLZ", "MHK"}


def get_gmail_connection():
    """Connect to Gmail via IMAP."""
    if not GMAIL_APP_PASSWORD:
        logger.error("GMAIL_APP_PASSWORD not set. Create one at: https://myaccount.google.com/apppasswords")
        return None
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
        return mail
    except Exception as e:
        logger.error(f"Gmail IMAP login failed: {e}")
        return None


def find_matching_email(mail, target_date: str | None = None) -> list:
    """
    Search for Valmo Control Tower emails from the sender.
    If target_date (DD-MM-YYYY) given: prefer emails with that date in subject.
    Returns ids, newest first (reversed).
    """
    from datetime import datetime

    mail.select("INBOX")
    date_str = target_date or datetime.now().strftime("%d-%m-%Y")

    # First try: subject + today's/report date (e.g. "[IMP] Valmo Control Tower!!! 23-03-2026")
    query_with_date = f'FROM "{SENDER_EMAIL}" SUBJECT "[IMP] Valmo Control Tower!!! {date_str}"'
    status, data = mail.search(None, query_with_date)
    if status == "OK" and data[0].strip():
        ids = data[0].split()
        if ids:
            logger.info("Found %d email(s) for subject + date %s", len(ids), date_str)
            return list(reversed(ids))  # newest first

    # Fallback: subject only, get latest
    query = f'FROM "{SENDER_EMAIL}" SUBJECT "[IMP] Valmo Control Tower!!!"'
    status, data = mail.search(None, query)
    if status != "OK":
        logger.warning("Gmail search failed")
        return []
    ids = data[0].split()
    if ids:
        logger.info("Found %d email(s) for subject (any date), using latest", len(ids))
    return list(reversed(ids))  # newest first


def extract_date_from_subject(subject: str | None) -> str | None:
    """Extract DD-MM-YYYY from subject e.g. '[IMP] Valmo Control Tower!!! 23-03-2026'."""
    if not subject:
        return None
    m = SUBJECT_DATE_PATTERN.search(subject)
    return m.group(1) if m else None


def get_attachments_from_message(mail, msg_id) -> tuple[str | None, dict]:
    """
    Fetch hardstop and lost attachments from an email.
    Returns (date_from_subject, {"hardstop": (data, filename), "lost": (data, filename)}).
    """
    status, data = mail.fetch(msg_id, "(RFC822)")
    if status != "OK" or not data:
        return None, {}

    raw = data[0][1]
    msg = email.message_from_bytes(raw)
    subject = msg.get("Subject") or ""
    date_str = extract_date_from_subject(subject)
    attachments = {}

    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        if HARDSTOP_ATTACHMENT_PATTERN.search(filename):
            attachments["hardstop"] = (payload, filename)
        elif LOST_ATTACHMENT_PATTERN.search(filename):
            attachments["lost"] = (payload, filename)

    return date_str, attachments


def load_dataframe_from_bytes(data: bytes, filename: str) -> "pd.DataFrame | None":
    """Load attachment bytes into a DataFrame (Excel or CSV)."""
    if not pd:
        logger.error("pandas not installed. pip install pandas openpyxl")
        return None

    import io
    buf = io.BytesIO(data)
    filename_lower = filename.lower()
    try:
        if filename_lower.endswith(".xlsx") or filename_lower.endswith(".xls"):
            return pd.read_excel(buf, sheet_name=0, engine="openpyxl")
        if filename_lower.endswith(".csv"):
            return pd.read_csv(buf, encoding="utf-8", on_bad_lines="skip")
        # Try Excel first, then CSV
        try:
            buf.seek(0)
            return pd.read_excel(buf, sheet_name=0, engine="openpyxl")
        except Exception:
            buf.seek(0)
            return pd.read_csv(buf, encoding="utf-8", on_bad_lines="skip")
    except Exception as e:
        logger.error("Failed to parse attachment: %s", e)
        return None


def find_column(df: "pd.DataFrame", name: str) -> str | None:
    """Find column by case-insensitive match."""
    name_lower = name.lower()
    for col in df.columns:
        if str(col).strip().lower() == name_lower:
            return col
    return None


def _filter_by_location(df: "pd.DataFrame") -> "pd.DataFrame":
    """Filter rows by TARGET_LOCATIONS. Expects a 'location' column."""
    loc_col = None
    for candidate in ["location", "Location", "location_code", "hub", "hub_code"]:
        if candidate in df.columns:
            loc_col = candidate
            break
        c = find_column(df, candidate)
        if c:
            df = df.rename(columns={c: "location"})
            loc_col = "location"
            break

    if loc_col:
        df[loc_col] = df[loc_col].astype(str).str.strip().str.upper()
        df = df[df[loc_col].isin(TARGET_LOCATIONS)].copy()
    else:
        logger.warning("No location column found. Keeping all rows.")
    return df


def filter_and_transform(
    df: "pd.DataFrame",
    columns: list[str],
    column_aliases: dict[str, tuple[str, ...]] | None = None,
) -> "pd.DataFrame":
    """Keep only required columns and filter by target locations."""
    if df.empty:
        return df

    col_map = {}
    for want in columns:
        found = None
        if column_aliases and want in column_aliases:
            for alias in column_aliases[want]:
                found = find_column(df, alias)
                if found:
                    break
        if not found:
            found = find_column(df, want)
        if found:
            col_map[found] = want

    if not col_map:
        logger.warning("None of the required columns found. Available: %s", list(df.columns))
        return pd.DataFrame()

    df = df[[c for c in col_map]].copy()
    df.columns = [col_map[c] for c in df.columns]
    return _filter_by_location(df)


def filter_and_transform_hardstop(df: "pd.DataFrame") -> "pd.DataFrame":
    """Filter hardstop data (columns + locations)."""
    return filter_and_transform(df, COLUMNS_TO_KEEP)


def filter_and_transform_lost(df: "pd.DataFrame") -> "pd.DataFrame":
    """Filter lost data (columns + locations)."""
    return filter_and_transform(df, LOST_COLUMNS_TO_KEEP, LOST_COLUMN_ALIASES)


def add_date_column(df: "pd.DataFrame", date_str: str | None) -> "pd.DataFrame":
    """Insert Date column as first column. Uses date_str or today if None."""
    from datetime import datetime
    if date_str is None:
        date_str = datetime.now().strftime("%d-%m-%Y")
    df = df.copy()
    df.insert(0, "Date", date_str)
    return df


def finalize_lostmarked_frame(df: "pd.DataFrame") -> "pd.DataFrame":
    """Ensure LostMarked columns match LOSTMARKED_HEADERS (fill missing with empty string)."""
    df = df.copy()
    for h in LOSTMARKED_HEADERS:
        if h not in df.columns:
            df[h] = ""
    return df[LOSTMARKED_HEADERS]


def _get_sheet_client():
    """Get gspread client and spreadsheet. Returns (gc, sh) or (None, None)."""
    if not gspread or not SERVICE_ACCOUNT_FILE.exists():
        return None, None
    try:
        creds = Credentials.from_service_account_file(
            str(SERVICE_ACCOUNT_FILE),
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SPREADSHEET_ID)
        return gc, sh
    except Exception as e:
        logger.error("Failed to get sheet client: %s", e)
        return None, None


def _normalize_date_for_match(val: str) -> str:
    """Normalize date string for comparison (strip, handle formats)."""
    if not val:
        return ""
    s = str(val).strip()
    if len(s) == 10 and s[2] in "-/" and s[5] in "-/":  # DD-MM-YYYY or DD/MM/YYYY
        return s.replace("/", "-")
    return s


def send_hardstop_to_whatsapp(date_str: str | None, last_row: int | None = None) -> bool:
    """
    Send Hardstop worksheet as WhatsApp image — same pattern as conversion_trend_analyzer
    (github.com/arunrajt-hub/conversion_trend_analyzer): fixed A1 range, _wh_log, caption;
    no auto_detect_rows. Range is A1:J{last_row} where last_row comes from push_to_google_sheet.
    """
    if not send_sheet_range_to_whatsapp:
        logger.warning("whatsapp_sheet_image not available - skip WhatsApp send")
        return False

    _, sh = _get_sheet_client()
    if not sh:
        return False

    def _wh_log(msg, level):
        if level == "ERROR":
            logger.error(msg)
        elif level == "WARNING":
            logger.warning(msg)
        else:
            logger.info(msg)

    try:
        from datetime import datetime

        # Explicit range like conversion_trend_analyzer (A1:S25); Hardstop table is columns A–J
        end_row = last_row if last_row and last_row >= 2 else 50
        end_row = min(max(end_row, 2), 500)

        ws = sh.worksheet(HARDSTOP_WORKSHEET)
        send_sheet_range_to_whatsapp(
            ws,
            range=f"A1:J{end_row}",
            caption=(
                f"Hardstop - {date_str or 'Report'} - "
                f"{datetime.now().strftime('%d-%b-%Y %H:%M')}"
            ),
            log_func=_wh_log,
        )
        return True
    except Exception as e:
        logger.warning("WhatsApp send failed (non-fatal): %s", e)
        return False


def push_to_google_sheet(
    df: "pd.DataFrame", worksheet_name: str, date_str: str | None = None
) -> tuple[bool, int | None]:
    """
    Push data to worksheet. If date_str exists: replace those rows. If new: append.
    Returns (success, last_row).
    """
    if not gspread:
        logger.error("gspread not installed. pip install gspread google-auth")
        return False, None

    if not SERVICE_ACCOUNT_FILE.exists():
        logger.error("service_account_key.json not found")
        return False, None

    try:
        creds = Credentials.from_service_account_file(
            str(SERVICE_ACCOUNT_FILE),
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SPREADSHEET_ID)
        try:
            ws = sh.worksheet(worksheet_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=worksheet_name, rows=10000, cols=20)

        # Prepare headers and data rows
        headers = df.columns.tolist()
        data_rows = df.fillna("").astype(str).values.tolist()
        num_cols = len(headers)

        def col_letter(n: int) -> str:
            """1->A, 2->B, ..., 27->AA"""
            s = ""
            while n > 0:
                n, r = divmod(n - 1, 26)
                s = chr(65 + r) + s
            return s

        existing = ws.get_all_values()
        target_date = _normalize_date_for_match(date_str or "") if date_str else ""

        if not existing or len(existing) <= 1:
            # Empty or headers-only: append
            combined = [headers] + data_rows if data_rows else [headers]
            last_row = len(combined)
            logger.info("Appended %d rows to %s (initial)", len(data_rows), worksheet_name)
        else:
            # If date exists: replace (remove old rows for that date). If new: append.
            data_existing = existing[1:]
            if target_date:
                other_rows = [
                    row for row in data_existing
                    if _normalize_date_for_match(row[0] if row else "") != target_date
                ]
                had_matching = len(other_rows) < len(data_existing)
                data_existing = other_rows
            else:
                had_matching = False
            combined_data = data_existing + data_rows
            combined = [headers] + combined_data
            last_row = len(combined)
            if had_matching:
                logger.info("Replaced rows for date %s in %s", date_str, worksheet_name)
            else:
                logger.info("Appended %d rows to %s (new date)", len(data_rows), worksheet_name)

        ws.clear()
        ws.update(range_name="A1", values=combined, value_input_option="USER_ENTERED")
        logger.info("Wrote %d rows to %s", last_row, worksheet_name)
        return True, last_row
    except Exception as e:
        logger.error("Failed to push to Google Sheet: %s", e)
        return False, None


def run_from_gmail(target_date: str | None = None) -> bool:
    """Fetch email, download attachments, process, and push to both worksheets."""
    mail = get_gmail_connection()
    if not mail:
        return False

    try:
        ids = find_matching_email(mail, target_date=target_date)
        if not ids:
            logger.warning("No Valmo Control Tower email found in inbox")
            return False

        # Process newest first (ids already newest-first from find_matching_email)
        for msg_id in ids:
            date_str, attachments = get_attachments_from_message(mail, msg_id)
            if not attachments:
                continue

            any_ok = False

            # Process hardstop
            if "hardstop" in attachments:
                data, filename = attachments["hardstop"]
                logger.info("Found hardstop: %s (date: %s)", filename, date_str or "N/A")
                df = load_dataframe_from_bytes(data, filename)
                if df is not None and not df.empty:
                    df = filter_and_transform_hardstop(df)
                    if not df.empty:
                        df = add_date_column(df, date_str)
                        ok, last_row = push_to_google_sheet(df, HARDSTOP_WORKSHEET, date_str)
                        if ok:
                            any_ok = True
                            # WhatsApp: same as conversion_trend_analyzer — fixed A1:J{last_row} image
                            send_hardstop_to_whatsapp(date_str, last_row=last_row)
                    else:
                        logger.warning("No rows after hardstop filtering")

            # Process lost
            if "lost" in attachments:
                data, filename = attachments["lost"]
                logger.info("Found lost: %s (date: %s)", filename, date_str or "N/A")
                df = load_dataframe_from_bytes(data, filename)
                if df is not None and not df.empty:
                    df = filter_and_transform_lost(df)
                    if not df.empty:
                        df = add_date_column(df, date_str)
                        df = finalize_lostmarked_frame(df)
                        ok, _ = push_to_google_sheet(df, LOSTMARKED_WORKSHEET, date_str)
                        if ok:
                            any_ok = True
                    else:
                        logger.warning("No rows after lost filtering")

            if any_ok:
                return True

        logger.warning("No valid attachments found in matching emails")
        return False
    finally:
        try:
            mail.logout()
        except Exception:
            pass


def run_from_file(file_path: str, date_str: str | None = None, is_lost: bool = False) -> bool:
    """Process a local file and push to sheet (for testing/manual runs)."""
    if not pd:
        logger.error("pandas required. pip install pandas openpyxl")
        return False

    path = Path(file_path)
    if not path.exists():
        logger.error("File not found: %s", path)
        return False

    try:
        if path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
        else:
            df = pd.read_csv(path, encoding="utf-8", on_bad_lines="skip")
    except Exception as e:
        logger.error("Failed to read file: %s", e)
        return False

    df = filter_and_transform_lost(df) if is_lost else filter_and_transform_hardstop(df)
    if df.empty:
        logger.warning("No rows after filtering")
        return False

    df = add_date_column(df, date_str)
    if is_lost:
        df = finalize_lostmarked_frame(df)
    ws = LOSTMARKED_WORKSHEET if is_lost else HARDSTOP_WORKSHEET
    ok, last_row = push_to_google_sheet(df, ws, date_str)
    if ok and not is_lost:
        send_hardstop_to_whatsapp(date_str, last_row=last_row)
    return ok


def main():
    ap = argparse.ArgumentParser(
        description="Valmo Hardstop: Gmail attachment → Google Sheet (Hardstop)"
    )
    ap.add_argument(
        "--file",
        "-f",
        help="Use local hardstop file instead of Gmail (e.g. hardstop.xlsx)",
    )
    ap.add_argument(
        "--lost-file",
        "-l",
        help="Use local lost file instead of Gmail (e.g. lost.xlsx)",
    )
    ap.add_argument(
        "--date",
        "-d",
        help="Date DD-MM-YYYY. For Gmail: search subject with this date. For --file/--lost-file: use as report date. Default: today",
    )
    args = ap.parse_args()

    if args.file:
        ok = run_from_file(args.file, date_str=args.date, is_lost=False)
    elif args.lost_file:
        ok = run_from_file(args.lost_file, date_str=args.date, is_lost=True)
    else:
        ok = run_from_gmail(target_date=args.date)

    if ok:
        logger.info("Done. Sheet: https://docs.google.com/spreadsheets/d/%s/edit#gid=1469155675", SPREADSHEET_ID)
    else:
        logger.error("Automation failed")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
