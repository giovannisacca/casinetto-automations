"""
Power BI Report Export, Excel Pivot Refresh, and Email Automation.

Uses Selenium to open Chrome, log into Power BI, and download the report
as Excel (.xlsx) and PDF. Then marks pivot tables to refresh on open and
emails both files via Gmail SMTP.

Usage:
    1. Copy .env.example to .env and fill in credentials
    2. pip install -r requirements.txt
    3. python refresh_powerbi_report.py
"""

import os
import shutil
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
POWERBI_USERNAME = os.environ["POWERBI_USERNAME"]
POWERBI_PASSWORD = os.environ["POWERBI_PASSWORD"]

REPORT_URL = os.environ.get(
    "POWERBI_REPORT_URL",
    "https://app.powerbi.com/groups/720de077-9471-4923-80f1-3352618c131d"
    "/reports/febba574-7c8f-4352-ae59-ebeb2b0f96de"
    "/ReportSection84fb5c2a3382b283c39f?experience=power-bi",
)

GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
RECIPIENT_EMAIL = os.environ["RECIPIENT_EMAIL"]

HEADLESS = os.environ.get("HEADLESS", "false").lower() == "true"

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

DOWNLOAD_DIR = Path(__file__).parent / "downloads_temp"
DOWNLOAD_DIR.mkdir(exist_ok=True)

TODAY = datetime.now().strftime("%Y-%m-%d")

WAIT_TIMEOUT = 60  # seconds for explicit waits


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _clear_download_dir() -> None:
    """Remove stale files from the temp download directory."""
    for f in DOWNLOAD_DIR.iterdir():
        f.unlink()


def _wait_for_download(extension: str, timeout: int = 120) -> Path:
    """
    Poll the download directory until a file with the given extension appears
    and Chrome has finished writing it (no .crdownload temp files).
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        files = list(DOWNLOAD_DIR.glob(f"*.{extension}"))
        crdownloads = list(DOWNLOAD_DIR.glob("*.crdownload"))
        if files and not crdownloads:
            # Return the most recently modified matching file
            return max(files, key=lambda p: p.stat().st_mtime)
        time.sleep(1)
    raise TimeoutError(
        f"Download did not complete within {timeout}s "
        f"(looking for *.{extension} in {DOWNLOAD_DIR})"
    )


# ---------------------------------------------------------------------------
# 1. Set up Chrome with Selenium
# ---------------------------------------------------------------------------
def create_driver() -> webdriver.Chrome:
    """Create a Chrome WebDriver configured for file downloads."""
    chrome_options = Options()

    if HEADLESS:
        chrome_options.add_argument("--headless=new")

    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--start-maximized")

    # Configure Chrome to download to our temp directory without prompts
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)

    # Enable downloads in headless mode
    if HEADLESS:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": str(DOWNLOAD_DIR.resolve()),
            },
        )

    return driver


# ---------------------------------------------------------------------------
# 2. Log into Power BI via Microsoft login
# ---------------------------------------------------------------------------
def login_to_powerbi(driver: webdriver.Chrome) -> None:
    """Navigate to the report URL and handle Microsoft OAuth login."""
    print("Navigating to Power BI report...")
    driver.get(REPORT_URL)

    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    # --- Enter email ---
    print("Entering username...")
    email_input = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="email"]'))
    )
    email_input.clear()
    email_input.send_keys(POWERBI_USERNAME)

    next_btn = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"]'))
    )
    next_btn.click()

    # --- Enter password ---
    print("Entering password...")
    password_input = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'input[type="password"]')
        )
    )
    password_input.clear()
    password_input.send_keys(POWERBI_PASSWORD)

    sign_in_btn = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"]'))
    )
    sign_in_btn.click()

    # --- Handle "Stay signed in?" prompt ---
    try:
        stay_signed_in_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[value="No"]'))
        )
        stay_signed_in_btn.click()
        print("Dismissed 'Stay signed in?' prompt.")
    except TimeoutException:
        # Some tenants skip this prompt
        pass

    # --- Wait for Power BI to fully load the report ---
    print("Waiting for Power BI report to load...")
    wait.until(EC.url_contains("app.powerbi.com"))
    # Wait for the main report canvas to render
    time.sleep(10)
    print("Power BI report loaded.")


# ---------------------------------------------------------------------------
# 3. Export report as Excel (underlying data) via browser
# ---------------------------------------------------------------------------
def export_excel(driver: webdriver.Chrome) -> Path:
    """
    Click Export → 'Analyze in Excel' or use the Export Data flow
    to download the report data as an Excel workbook.

    Power BI's export-to-Excel flow:
      File menu → Export → Analyze in Excel  (downloads .xlsx)
    """
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    _clear_download_dir()

    print("Exporting report as Excel...")

    # Click the "File" menu in the Power BI toolbar
    file_menu = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'button[aria-label="File menu"]')
        )
    )
    file_menu.click()
    time.sleep(1)

    # Hover/click "Export" submenu
    export_item = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//span[text()="Export"]/..')
        )
    )
    export_item.click()
    time.sleep(1)

    # Click "Analyze in Excel" to download the .xlsx
    analyze_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[contains(text(), "Analyze in Excel")]')
        )
    )
    analyze_btn.click()
    time.sleep(2)

    # Handle any confirmation dialog
    try:
        confirm = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(text(), "Open") or contains(text(), "Download") or contains(text(), "OK")]')
            )
        )
        confirm.click()
    except TimeoutException:
        pass

    # Wait for the .xlsx download to finish
    print("Waiting for Excel download to complete...")
    downloaded = _wait_for_download("xlsx")

    # Move to output directory with standardised name
    out_path = OUTPUT_DIR / f"report_{TODAY}.xlsx"
    shutil.move(str(downloaded), str(out_path))
    print(f"Saved {out_path} ({out_path.stat().st_size:,} bytes)")
    return out_path


# ---------------------------------------------------------------------------
# 4. Export report as PDF via browser
# ---------------------------------------------------------------------------
def export_pdf(driver: webdriver.Chrome) -> Path:
    """
    Use Power BI's built-in Export to PDF.

    Flow: File menu → Export → PDF
    """
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    _clear_download_dir()

    print("Exporting report as PDF...")

    # Click the "File" menu
    file_menu = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'button[aria-label="File menu"]')
        )
    )
    file_menu.click()
    time.sleep(1)

    # Click "Export" submenu
    export_item = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//span[text()="Export"]/..')
        )
    )
    export_item.click()
    time.sleep(1)

    # Click "PDF"
    pdf_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[contains(text(), "PDF")]')
        )
    )
    pdf_btn.click()
    time.sleep(2)

    # Power BI shows a dialog to configure pages — click "Export"
    try:
        export_confirm = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(text(), "Export")]')
            )
        )
        export_confirm.click()
    except TimeoutException:
        pass

    # PDF generation can take a while server-side
    print("Waiting for PDF generation and download...")
    downloaded = _wait_for_download("pdf", timeout=180)

    out_path = OUTPUT_DIR / f"report_{TODAY}.pdf"
    shutil.move(str(downloaded), str(out_path))
    print(f"Saved {out_path} ({out_path.stat().st_size:,} bytes)")
    return out_path


# ---------------------------------------------------------------------------
# 5. Refresh pivot tables in the Excel file
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
# 6. Email the files via Gmail SMTP
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
    driver = create_driver()
    try:
        login_to_powerbi(driver)

        xlsx_path = export_excel(driver)
        pdf_path = export_pdf(driver)
    finally:
        driver.quit()
        print("Browser closed.")

    refresh_pivot_tables(xlsx_path)

    send_email([xlsx_path, pdf_path])

    # Clean up temp download dir
    shutil.rmtree(DOWNLOAD_DIR, ignore_errors=True)

    print("\nDone. All steps completed successfully.")


if __name__ == "__main__":
    main()
