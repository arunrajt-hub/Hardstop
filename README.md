# Valmo Hardstop Automation

Automation to process the daily **Valmo Control Tower** email, extract hardstop and lost attachments, and push filtered data to Google Sheets. Optionally sends Hardstop summary to WhatsApp.

## Email Source

- **From:** lsn-meesho-central@loadshare.net  
- **Subject:** `[IMP] Valmo Control Tower!!! DD-MM-YYYY` (date varies daily)
- **Attachments:**
  - `hardstop_lsn-meesho-central@loadshare.net` → Hardstop worksheet
  - `lost_lsn-meesho-central@loadshare.net` → LostMarked worksheet

## Output

- **Google Sheet:** [Meesho Reports](https://docs.google.com/spreadsheets/d/1qnqzVf-S41F4S6DN8CRtXVgk-BcsaW377aVVEyFrnzg)
- **Worksheets:** Hardstop, LostMarked  
- **LostMarked columns:** Date, lost_date, awd (from source `awd` or `awb`), current_movement_type, loss_value, location  
- **Locations filtered:** MQR, MQE, YLG, YLZ, MHK

## Logic

- Searches for the **latest** email with that subject (and today's date when possible)
- **Same date:** Replace existing rows for that date
- **New date:** Append rows
- **Hardstop "Remarks" column:** If present as the last column, values are **preserved** for rows kept when updating other dates. New rows from the attachment get an empty Remarks cell.
- **WhatsApp (Hardstop):** Image shows data columns **excluding** Remarks, and **only rows where Remarks is empty** (rows with any text in Remarks are omitted from the image).
- **WhatsApp:** Uses `whatsapp_sheet_image` when configured (`WHAPI_TOKEN`, etc.)

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Add credentials

- **`service_account_key.json`** – Google service account for Sheets (share the sheet with the service account email)
- **Gmail App Password** – [Create one](https://myaccount.google.com/apppasswords)

### 3. Environment variables

| Variable | Description |
|----------|-------------|
| `GMAIL_EMAIL` | Gmail address that receives the Valmo email |
| `GMAIL_APP_PASSWORD` | Gmail App Password |
| `WHAPI_TOKEN` | WhatsApp API token (optional, for Hardstop image) |
| `WHATSAPP_PHONE` | WhatsApp recipient(s) (optional) |

## Usage

```bash
# Run from Gmail (fetches today's report)
python valmo_hardstop_gmail_to_sheet.py

# Search for a specific date
python valmo_hardstop_gmail_to_sheet.py --date 24-03-2026

# Manual run with local files
python valmo_hardstop_gmail_to_sheet.py --file hardstop.xlsx --date 24-03-2026
python valmo_hardstop_gmail_to_sheet.py --lost-file lost.xlsx --date 24-03-2026
```

## Scheduling

### Windows (4 PM & 10 PM IST)

```bash
# One-time setup
schedule_valmo_hardstop.bat

# Remove schedule
unschedule_valmo_hardstop.bat
```

### GitHub Actions (optional)

The workflow `.github/workflows/valmo-hardstop.yml` runs on **GitHub-hosted Ubuntu** at **4 PM** and **10 PM IST**. It installs **Chrome** for HTML→image (WhatsApp) the same way as your other automations. Push this repo to GitHub, then add secrets below.

Add these repository secrets in Settings → Secrets and variables → Actions:

| Secret | Description |
|--------|-------------|
| `GMAIL_EMAIL` | Gmail address |
| `GMAIL_APP_PASSWORD` | Gmail App Password |
| `SERVICE_ACCOUNT_JSON` | Full contents of `service_account_key.json` (copy-paste the entire JSON) |
| `WHAPI_TOKEN` | Optional – for WhatsApp |
| `WHATSAPP_PHONE` | Optional – recipient(s) |

## Files

| File | Purpose |
|------|---------|
| `valmo_hardstop_gmail_to_sheet.py` | Main script |
| `whatsapp_sheet_image.py` | WhatsApp image sender (optional) |
| `html_table_to_image.py` | Sheet HTML → PNG (Chrome/Selenium; required for WhatsApp image) |
| `run_valmo_hardstop.bat` | Batch runner |
| `schedule_valmo_hardstop.bat` | Create Windows scheduled tasks |
| `schedule_valmo_hardstop.ps1` | PowerShell scheduler |
| `unschedule_valmo_hardstop.bat` | Remove scheduled tasks |
