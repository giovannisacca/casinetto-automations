"""
Power BI Report Export, Excel Pivot Refresh, and Email Automation.

Exports a Power BI report as Excel (.xlsx) and PDF, marks any pivot tables
in the Excel file to refresh on open, then emails both files via Gmail SMTP.

Usage:
    1. Copy .env.example to .env and fill in credentials
    2. pip install -r requirements.txt
    3. python refresh_powerbi_report.py
"""

import os
import sys
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from datetime import datetime

import msal
import requests
from openpyxl import load_workbook
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
AZURE_CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
AZURE_TENANT_ID = os.environ["AZURE_TENANT_ID"]
POWERBI_USERNAME = os.environ["POWERBI_USERNAME"]
POWERBI_PASSWORD = os.environ["POWERBI_PASSWORD"]

WORKSPACE_ID = os.environ.get(
    "POWERBI_WORKSPACE_ID", "720de077-9471-4923-80f1-3352618c131d"
)
REPORT_ID = os.environ.get(
    "POWERBI_REPORT_ID", "febba574-7c8f-4352-ae59-ebeb2b0f96de"
)

GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
RECIPIENT_EMAIL = os.environ["RECIPIENT_EMAIL"]

POWERBI_API = "https://api.powerbi.com/v1.0/myorg"
POWERBI_SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

TODAY = datetime.now().strftime("%Y-%m-%d")


# ---------------------------------------------------------------------------
# 1. Authenticate with Power BI (ROPC / Master User)
# ---------------------------------------------------------------------------
def get_powerbi_token() -> str:
    """Acquire an access token using Resource Owner Password Credentials."""
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    app = msal.PublicClientApplication(AZURE_CLIENT_ID, authority=authority)

    result = app.acquire_token_by_username_password(
        username=POWERBI_USERNAME,
        password=POWERBI_PASSWORD,
        scopes=POWERBI_SCOPE,
    )

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        print(f"Authentication failed: {error}", file=sys.stderr)
        sys.exit(1)

    print("Authenticated with Power BI successfully.")
    return result["access_token"]


# ---------------------------------------------------------------------------
# 2. Export report from Power BI
# ---------------------------------------------------------------------------
def export_report(token: str, file_format: str) -> Path:
    """
    Trigger an export of the report and poll until complete.

    file_format: "XLSX" or "PDF"
    Returns the local file path of the downloaded export.
    """
    headers = {"Authorization": f"Bearer {token}"}
    export_url = (
        f"{POWERBI_API}/groups/{WORKSPACE_ID}/reports/{REPORT_ID}/ExportTo"
    )

    body = {"format": file_format}
    print(f"Requesting {file_format} export...")
    resp = requests.post(export_url, headers=headers, json=body, timeout=60)
    resp.raise_for_status()
    export_id = resp.json()["id"]

    # Poll for completion
    status_url = (
        f"{POWERBI_API}/groups/{WORKSPACE_ID}/reports/{REPORT_ID}"
        f"/exports/{export_id}"
    )
    for attempt in range(60):
        time.sleep(5)
        poll = requests.get(status_url, headers=headers, timeout=30)
        poll.raise_for_status()
        state = poll.json()
        status = state.get("status")
        pct = state.get("percentComplete", 0)
        print(f"  Export {file_format}: {status} ({pct}%)")

        if status == "Succeeded":
            break
        if status in ("Failed", "Cancelled"):
            print(f"Export failed: {state}", file=sys.stderr)
            sys.exit(1)
    else:
        print("Export timed out after 5 minutes.", file=sys.stderr)
        sys.exit(1)

    # Download the file
    file_url = (
        f"{POWERBI_API}/groups/{WORKSPACE_ID}/reports/{REPORT_ID}"
        f"/exports/{export_id}/file"
    )
    download = requests.get(file_url, headers=headers, timeout=120)
    download.raise_for_status()

    ext = "xlsx" if file_format == "XLSX" else "pdf"
    out_path = OUTPUT_DIR / f"report_{TODAY}.{ext}"
    out_path.write_bytes(download.content)
    print(f"Saved {out_path} ({len(download.content):,} bytes)")
    return out_path


# ---------------------------------------------------------------------------
# 3. Refresh pivot tables in the Excel file
# ---------------------------------------------------------------------------
def refresh_pivot_tables(xlsx_path: Path) -> Path:
    """
    Mark every pivot table in the workbook to refresh when the file is opened
    in Excel. Also refreshes the pivot cache so Excel recalculates on open.
    """
    wb = load_workbook(xlsx_path)

    pivot_count = 0
    for sheet in wb.worksheets:
        for pivot in sheet._pivots:
            pivot.cache.refreshOnLoad = True
            pivot_count += 1

    if pivot_count:
        print(f"Marked {pivot_count} pivot table(s) to refresh on open.")
    else:
        print("No pivot tables found in the workbook.")

    # Also set the workbook to recalculate formulas on open
    wb.calculation.calcMode = "auto"

    wb.save(xlsx_path)
    wb.close()
    print(f"Saved updated workbook: {xlsx_path}")
    return xlsx_path


# ---------------------------------------------------------------------------
# 4. Email the files via Gmail SMTP
# ---------------------------------------------------------------------------
def send_email(attachments: list[Path]) -> None:
    """Send an email with the given file attachments via Gmail SMTP."""
    msg = MIMEMultipart()
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = RECIPIENT_EMAIL
    msg["Subject"] = f"Power BI Report - {TODAY}"

    body = (
        f"Hi,\n\n"
        f"Please find attached the Power BI report exported on {TODAY}.\n\n"
        f"The Excel file's pivot tables are set to refresh automatically "
        f"when you open it in Excel.\n\n"
        f"Best regards"
    )
    msg.attach(MIMEText(body, "plain"))

    for filepath in attachments:
        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={filepath.name}",
        )
        msg.attach(part)

    print(f"Sending email to {RECIPIENT_EMAIL}...")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, RECIPIENT_EMAIL, msg.as_string())

    print("Email sent successfully.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    token = get_powerbi_token()

    xlsx_path = export_report(token, "XLSX")
    pdf_path = export_report(token, "PDF")

    refresh_pivot_tables(xlsx_path)

    send_email([xlsx_path, pdf_path])

    print("\nDone. All steps completed successfully.")


if __name__ == "__main__":
    main()
