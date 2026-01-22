import os
import re
import sys
import json
import shutil
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, Page
from app.paths import internal_path, external_path
from app.irmas_scan_outdated_version import run_outdated_scan
from app.irmas_generate_paginated_json import IrmasReportMerger
from app.irmas_report_generator import IrmasReportGenerator
from app.cht_sso_login import ChtSsoLogin
from app.address_book_exporter import AddressBookExporter

IRMAS_OUTPUT_DIR = "./output/irmas"
# Create directory if not exists
os.makedirs(IRMAS_OUTPUT_DIR, exist_ok=True)

# List of websites to visit
ldap_site = "https://ntpe.cht.com.tw/ldap/eo.aspx"
irmas_site = "https://irmas.cht.com.tw"
detail_results = []

# -----------------------------
# 🔐 Load credentials
# -----------------------------
load_dotenv()

def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(".")

BASE_DIR = get_base_dir()

CHROMIUM_PATH = internal_path("chromium/chrome.exe")
CHROMIUM_PROFILE = os.path.join(BASE_DIR, "chromium_profile")

os.makedirs(CHROMIUM_PROFILE, exist_ok=True)

if not os.path.exists(CHROMIUM_PATH):
    raise RuntimeError(f"Chromium not found: {CHROMIUM_PATH}")

def run_software_query(page: Page, cat_label: str, values: list[str]):
    """
    Run a 軟體資料 query with given Cat label and input values
    """
    page.goto("https://irmas.cht.com.tw/90102_00.php")
    page.wait_for_load_state("networkidle")
    page.click("a[href='./90102_01.php?SubInfo=11']")
    page.wait_for_load_state("networkidle")

    print("Selecting 軟體資料...")
    page.select_option("select[name='Item']", label="軟體資料")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(300)

    print(f"Selecting Cat: {cat_label}")
    page.select_option("select[name='Cat']", label=cat_label)
    page.wait_for_timeout(300)

    for value in values:
        print(f"Adding {cat_label}: {value}")
        page.fill("input[name='s101_data1']", value)
        page.click("input[name='s101_button']")
        page.wait_for_timeout(200)

    with page.expect_request_finished(timeout=120000) as finished:
        page.locator("input[name='submit_data']").click(no_wait_after=True)

    print(f"{cat_label} request completed:", finished.value.url)
    page.wait_for_load_state("networkidle")

def select_irmas_role(page: Page):
    """
    Handle role selection for different users:
    1. Navigate to IRMAS
    2. Click "角色切換" button
    3. Check if redirected to 0_auth2.php
    4. If yes, select specific role based on 管轄範圍
    """
    # ---------------------------------------------
    # 1️⃣ Navigate to IRMAS first
    # ---------------------------------------------
    page.goto("https://irmas.cht.com.tw")
    page.wait_for_load_state("networkidle")
    print("Navigated to IRMAS")
    
    # ---------------------------------------------
    # 2️⃣ Click the "角色切換" button
    # ---------------------------------------------
    page.click("input[value='角色切換']")
    print("Clicked 角色切換")

    page.wait_for_load_state("networkidle")

    # ---------------------------------------------
    # 2️⃣ Check if redirected to role selection page
    # ---------------------------------------------
    if "0_auth2.php" in page.url:
        print("On role selection page...")
        # Find the row where 管轄範圍 contains "/中華電信公司/新北營運處"
        rows = page.locator("table tr")
        row_count = rows.count()
        
        button_clicked = False
        for i in range(row_count):
            row = rows.nth(i)
            cells = row.locator("td")
            
            # Check if row has 3 cells (button, 功能權限, 管轄範圍)
            if cells.count() >= 3:
                scope_cell = cells.nth(2)  # Third column is 管轄範圍
                scope_text = scope_cell.inner_text().strip()
                
                if "/中華電信公司/新北營運處" in scope_text:
                    print(f"Found matching role with scope: {scope_text}")
                    button = cells.nth(0).locator("input[type='button']")
                    if button.count() > 0:
                        button.click()
                        print("Clicked role selection button")
                        page.wait_for_load_state("networkidle")
                        button_clicked = True
                        break
        
        if not button_clicked:
            print("Role selection button for '新北營運處' not found, continuing...")
    else:
        print("Not on role selection page")

def banned_software_finding_procedure(page: Page):
    # ---------------------------------------------
    # 1️⃣ Click the "電腦資料" menu item
    # ---------------------------------------------
    # <font id="STMtubtehr_0__5___TX">電腦資料</font>
    page.click("font#STMtubtehr_0__5___TX")
    print("Clicked 電腦資料")

    page.goto("https://irmas.cht.com.tw/90102_00.php")
    page.wait_for_load_state("networkidle")

    # Click 主機資訊 link
    page.click("a[href='./90102_01.php?SubInfo=11']")
    page.wait_for_load_state("networkidle")
    print("Opened 主機資訊 page")

    # Load config
    banned_software_config_path = internal_path("config/banned_software.json")
    with open(banned_software_config_path, "r", encoding="utf-8") as f:
        banned_softwares = json.load(f)

    # ---------------------------------------------
    # 5️⃣ Search by 名稱
    # ---------------------------------------------
    run_software_query(
        page,
        cat_label="名稱",
        values=banned_softwares["keywords"]
    )

    results_by_name = extract_table(page)

    # ---------------------------------------------
    # 6️⃣ Search by 廠商
    # ---------------------------------------------
    run_software_query(
        page,
        cat_label="廠商",
        values=banned_softwares["manufacturers"]
    )

    results_by_manufacturer = extract_table(page)
    results = [*results_by_name, *results_by_manufacturer]

    # Save to JSON
    with open(os.path.join(IRMAS_OUTPUT_DIR, "banned_softwares_report.json"), "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    for row in results:

        # skip rows that have no detail link
        if "detail_link" not in row:
            print(f"Skipping (no detail link): {row['value']}")
            continue

        url = irmas_site + "/" + row["detail_link"]
        value = row["value"]

        print(f"Visiting detail page for {value}: {url}")

        page.goto(url)
        page.wait_for_load_state("networkidle")

        detail_json = extract_detail_table(page, value)
        detail_results.append(detail_json)

    with open(os.path.join(IRMAS_OUTPUT_DIR, "banned_softwares_detail_report.json"), "w", encoding="utf-8") as f:
        json.dump(detail_results, f, ensure_ascii=False, indent=2)
        page.wait_for_timeout(800)

# -----------------------------
# Table Extraction Function
# -----------------------------
def extract_table(page: Page):
    table = page.locator("div#container table")
    rows = table.locator("tr")
    row_count = rows.count()

    data = []

    for i in range(row_count):
        cells = rows.nth(i).locator("td")
        if cells.count() == 0:
            continue  # skip headers

        raw_text = cells.nth(0).inner_text().strip()
        if "：" in raw_text:
            label, value = raw_text.split("：", 1)
        else:
            continue

        count_cell = cells.nth(1)
        link_el = count_cell.locator("a")

        # Count > 0 → include detail_link
        if link_el.count() > 0:
            count = int(link_el.inner_text().strip())
            onclick = link_el.get_attribute("onclick")
            link = onclick.split("'")[1]

            data.append({
                "label": label,
                "value": value,
                "count": count,
                "detail_link": link
            })

        # Count = 0 → do NOT include detail_link key
        else:
            count = int(count_cell.inner_text().strip())
            data.append({
                "label": label,
                "value": value,
                "count": count
            })

    return data

def extract_detail_table(page: Page, value_name: str):
    """
    Extract device table from 90102_03.php?ArgVal=xxxx
    using CHINESE KEYS exactly as displayed.
    """

    # Wait until page has at least one table
    page.wait_for_selector("body > table")

    table = page.locator("body > table")
    rows = table.locator("tr")
    row_count = rows.count()

    result = {
        "value": value_name,
        "items": []
    }

    # Find the header row (col_title)
    header_index = None
    for i in range(row_count):
        cells = rows.nth(i).locator("td")
        if cells.count() == 0:
            continue
        if cells.nth(0).get_attribute("class") == "col_title":
            header_index = i
            break

    if header_index is None:
        print("⚠️ No column header found!")
        return result

    # Data begins from header_index + 1
    for i in range(header_index + 1, row_count):
        cells = rows.nth(i).locator("td")

        if cells.count() < 6:
            continue

        # IP + link
        ip_cell = cells.nth(0)
        ip_text = ip_cell.inner_text().strip()

        link_el = ip_cell.locator("a[onclick]")
        pc_detail_link = None
        if link_el.count() > 0:
            onclick = link_el.get_attribute("onclick")
            if onclick:
                pc_detail_link = onclick.split("'")[1]

        # OS + 場域名稱
        os_html = cells.nth(4).inner_html().strip()
        os_parts = [x.strip() for x in os_html.split("<br>")]
        作業系統 = os_parts[0]
        場域名稱 = os_parts[1] if len(os_parts) > 1 else ""

        ip_text = ip_text.strip()
        # Filter out IP addresses starting with 10.28 (not part of the allowed network range)
        if (ip_text.startswith("10.28")):
            print(f"Skipping IP not in allowed range: {ip_text}")
            continue

        item = {
            "IP位址": ip_text,
            "電腦名稱": cells.nth(1).inner_text().strip(),
            "資產ID": cells.nth(2).inner_text().strip(),
            "使用者": cells.nth(3).inner_text().strip(),
            "作業系統": 作業系統,
            "場域名稱": 場域名稱,
            "更新時間": cells.nth(5).inner_text().strip(),
            "PC明細連結": pc_detail_link
        }

        result["items"].append(item)
    return result

def query_antivirus_server_ip_range(page: Page, start=1, end=12):
    """
    防毒軟體資料 → 連線伺服器IP → 批次輸入 IP 範圍並提交
    範例: 10.173.105.1 ~ 10.173.105.12 （不含 .3）
    """
    page.goto(irmas_site)
    # ---------------------------------------------
    # 2️⃣ Click the “電腦資料” menu item
    # ---------------------------------------------
    # <font id="STMtubtehr_0__5___TX">電腦資料</font>
    page.click("font#STMtubtehr_0__5___TX")
    print("Clicked 電腦資料")

    # ---------------------------------------------
    # 3️⃣ Go to 主機資訊頁面
    # ---------------------------------------------
    print("Opening 主機資訊...")

    page.goto(irmas_site + "/90102_00.php")
    page.wait_for_load_state("networkidle")

    # Click 主機資訊 link
    page.click("a[href='./90102_01.php?SubInfo=11']")
    page.wait_for_load_state("networkidle")
    print("Opened 主機資訊 page")

    print("Selecting 防毒軟體資料 ...")
    page.select_option("select[name='Item']", label="防毒軟體資料")
    page.wait_for_timeout(300)

    print("Selecting 連線伺服器IP ...")
    page.select_option("select[name='Cat']", value="1101405")
    page.wait_for_timeout(300)

    # Generate IPs but EXCLUDE .3
    ips = [
        f"10.173.105.{i}"
        for i in range(start, end + 1)
        if i != 3
    ]

    print("Adding IPs:")
    for ip in ips:
        print(" -", ip)
        page.fill("input[name='s101_data1']", ip)
        page.click("input[name='s101_button']")  # 新增
        page.wait_for_timeout(150)

    print("Submitting query...")
    page.click("input[name='submit_data']")
    page.wait_for_load_state("networkidle")
    antivirus_summary = extract_antivirus_summary(page)

    details = {}

    for row in antivirus_summary:
        value = row["value"]
        count = row["count"]

        details[value] = {"count": count}

        if "detail_link" not in row:
            continue

        url = irmas_site + "/" + row["detail_link"]
        print(f"Fetching detail page: {url}")

        page.goto(url)
        page.wait_for_load_state("networkidle")

        detail_json = extract_detail_table(page, value)
        details[value]["detail"] = detail_json

    with open(os.path.join(IRMAS_OUTPUT_DIR, "antivirus_summary.json"), "w", encoding="utf-8") as f:
        json.dump(antivirus_summary, f, ensure_ascii=False, indent=2)
    with open(os.path.join(IRMAS_OUTPUT_DIR, "antivirus_detail_report.json"), "w", encoding="utf-8") as f:
        json.dump(details, f, ensure_ascii=False, indent=2)

    print("Antivirus Server IP query completed.")

def extract_detail_antivirus_table(page: Page, value_name: str):
    """
    Extract detail table for 防毒軟體資料_連線伺服器IP
    or any 90102_03.php detail page.
    Chinese keys.
    """

    # Wait for table to load
    page.wait_for_selector("body > table")

    table = page.locator("body > table")
    rows = table.locator("tr")
    row_count = rows.count()

    result = {
        "value": value_name,
        "items": []
    }

    # Step 1: locate header row
    header_index = None
    for i in range(row_count):
        cells = rows.nth(i).locator("td")
        if cells.count() > 0 and cells.nth(0).get_attribute("class") == "col_title":
            header_index = i
            break

    if header_index is None:
        return result  # No data, return empty

    # Step 2: iterate data rows
    for i in range(header_index + 1, row_count):
        cells = rows.nth(i).locator("td")
        if cells.count() < 6:
            continue

        # IP + detail link
        ip_cell = cells.nth(0)
        ip_text = ip_cell.inner_text().strip()

        # Only match <a onclick=""> (ignore empty <a></a>)
        link_el = ip_cell.locator("a[onclick]")

        pc_detail_link = None
        if link_el.count() > 0:
            onclick = link_el.first.get_attribute("onclick")
            pc_detail_link = onclick.split("'")[1] if onclick else None

        # 作業系統 + 場域名稱
        os_html = cells.nth(4).inner_html().strip()
        os_parts = [p.strip() for p in os_html.split("<br>")]

        作業系統 = os_parts[0].rstrip(".")  # remove trailing dot if exists
        場域名稱 = os_parts[1] if len(os_parts) > 1 else ""

        # Build item
        item = {
            "IP位址": ip_text,
            "電腦名稱": cells.nth(1).inner_text().strip(),
            "資產ID": cells.nth(2).inner_text().strip(),
            "使用者": cells.nth(3).inner_text().strip(),
            "作業系統": 作業系統,
            "場域名稱": 場域名稱,
            "更新時間": cells.nth(5).inner_text().strip(),
            "PC明細連結": pc_detail_link
        }

        result["items"].append(item)

    return result

def extract_antivirus_summary(page: Page):
    """
    Extracts 防毒軟體資料_連線伺服器IP summary table.
    Same logic as extract_table() but for antivirus mode.
    """

    page.wait_for_selector("table")

    table = page.locator("table")
    rows = table.locator("tr")
    row_count = rows.count()

    results = []

    # Skip first two rows (header + label)
    for i in range(2, row_count):
        cells = rows.nth(i).locator("td")

        if cells.count() < 2:
            continue

        raw_text = cells.nth(0).inner_text().strip()

        # Expected format: 防毒軟體資料_連線伺服器IP：10.173.105.1
        if "：" in raw_text:
            _, value = raw_text.split("：", 1)
        else:
            continue

        count_cell = cells.nth(1)
        link_el = count_cell.locator("a")

        # Case 1: Count > 0 and <a> exists
        if link_el.count() > 0:
            count = int(link_el.inner_text().strip())
            onclick = link_el.get_attribute("onclick")
            detail = onclick.split("'")[1]

            results.append({
                "value": value,
                "count": count,
                "detail_link": detail
            })

        # Case 2: Count = 0 (no link)
        else:
            count = int(count_cell.inner_text().strip())
            results.append({
                "value": value,
                "count": count
            })

    return results

def sanitize_filename(name: str) -> str:
    """Remove illegal characters from filename."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def audit_specific_software(page: Page):
    page.goto(irmas_site + "/90303_00.php")

    # Create placeholder dict for config example
    software_version_map = {}

    # ────────────────────────────────────────────────
    # 1) Get folder title "特定軟體清查"
    # ────────────────────────────────────────────────
    folder_title = page.locator("th.black_title").inner_text().strip()

    # Build final save path
    save_dir = f"./output/irmas/{folder_title}"
    os.makedirs(save_dir, exist_ok=True)
    print("Saving downloads to:", save_dir)

    # ────────────────────────────────────────────────
    # 2) Get all checkboxes
    # ────────────────────────────────────────────────
    checkboxes = page.locator("input[type='checkbox'][name='ApArr[]']")
    total = checkboxes.count()

    print(f"Found {total} rows")
    # 2) Collect all checkbox + name column elements
    rows = page.locator("tr")  # all table rows
    total_rows = rows.count()

    export_button = page.locator("input[value='匯出Excel']")

    # ────────────────────────────────────────────────
    # 3) Loop each row → download Excel
    # ────────────────────────────────────────────────
    for i in range(total):
        print(f"=== Processing row {i+1}/{total} ===")

        cb = checkboxes.nth(i)
        cb.scroll_into_view_if_needed()
        cb.check()

        with page.expect_download() as download_info:
            export_button.click()

        # find the row (<tr>) that contains this checkbox
        row = cb.locator("xpath=ancestor::tr")
        download = download_info.value
        # extract software name from 2nd t_text column
        software_name = row.locator("td.t_text").nth(1).inner_text().strip()
        software_name = software_name.replace("\u00a0", "")  # remove &nbsp;
        safe_name = sanitize_filename(software_name)
        # Extract the software name (2nd <td> with class t_text)
        safe_name = sanitize_filename(software_name)     
        filename = f"{safe_name}.xls"
        # Add to config map (value always "0")
        software_version_map.setdefault(software_name, {})
        software_version_map[software_name]["version"] = "0"
        save_path = os.path.join(save_dir, filename)
        download.save_as(save_path)

        print(f"Downloaded → {save_path}")

        cb.uncheck()

    # ────────────────────────────────────────────────
    # 4) Generate config/software_latest_version_example.json
    # ────────────────────────────────────────────────
    config_dir = external_path("config")
    os.makedirs(config_dir, exist_ok=True)
    example_config_path = config_dir / "software_latest_version_example.json"

    with open(example_config_path, "w", encoding="utf-8") as f:
        json.dump(software_version_map, f, ensure_ascii=False, indent=4)

    print(f"\nGenerated example config → {example_config_path}")
    print("\nAll rows processed.")

def is_onedrive_dispatch_enabled() -> bool:
    return os.getenv("IRMAS_ONEDRIVE_DISPATCH") == "1"

def get_onedrive_path():
    """
    Prefer OneDrive for Business, fallback to personal OneDrive.
    """
    return (
        os.environ.get("OneDriveCommercial")
        or os.environ.get("OneDrive")
    )

def dispatch_outputs_to_onedrive(enable_dispatch: bool = False):
    if not enable_dispatch:
        print("ℹ️ OneDrive dispatch disabled. Skipping.")
        return

    onedrive = get_onedrive_path()
    if not onedrive:
        print("⚠️ OneDrive not found. Skipping dispatch.")
        return

    dst_dir = Path(onedrive) / "IrmasAutomate"

    # Only create folder when dispatch is enabled
    dst_dir.mkdir(parents=True, exist_ok=True)

    sources = [
        Path("output/irmas/ready_for_dispatch/irmas_messages.json"),
        Path("output/irmas/ready_for_dispatch/irmas_messages_missing_contacts.json"),
        Path("output/irmas/irmas_report_searchable.html"),
    ]

    for src in sources:
        if src.exists():
            shutil.copy2(src, dst_dir / src.name)
            print(f"✅ Copied: {src.name}")
        else:
            print(f"❌ Missing: {src}")
    
    print("✅ OneDrive dispatch completed.")

# -----------------------------
# ▶️ Main runner
# -----------------------------
def run(playwright, enable_onedrive_dispatch: bool = True):
    browser = playwright.chromium.launch(
        executable_path=CHROMIUM_PATH,
        headless=False,
        args=[
                "--disable-dev-shm-usage",
                "--no-sandbox",
            ]
    )
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=CHROMIUM_PROFILE,
        executable_path=CHROMIUM_PATH,
        headless=False,
        args=[
            "--disable-dev-shm-usage",
            "--no-sandbox",
        ]
    )

    page = context.new_page()
    # Login once
    login = ChtSsoLogin()
    # Step 1: Ensure login for IRMAS first
    print("🔐 Logging in to IRMAS...")
    login.ensure_login(page, irmas_site)
    # Step 2: Then navigate to LDAP
    print(f"🌐 Navigating to LDAP...")
    address_book_exporter = AddressBookExporter(page)
    address_book_exporter.run()
    select_irmas_role(page)
    audit_specific_software(page)
    # after all crawling jobs complete:
    run_outdated_scan()
    banned_software_finding_procedure(page)
    query_antivirus_server_ip_range(page)    
    merger = IrmasReportMerger(base_dir=IRMAS_OUTPUT_DIR)
    merger.load_reports()
    merger.load_address_book("./output/contacts/address_book.json")
    merger.process()
    merger.export_name_message_list("irmas_messages.json")
    merger.export_missing_contacts("irmas_messages_missing_contacts.json")
    # 🔽 New: generate Excel + HTML reports from the 4 JSON files
    gen = IrmasReportGenerator(IRMAS_OUTPUT_DIR)
    gen.generate_reports()

    if is_onedrive_dispatch_enabled():
        dispatch_outputs_to_onedrive(enable_dispatch=True)
    else:
        print("ℹ️ OneDrive dispatch disabled via environment")

    print("All tasks done successfully!")
    page.wait_for_timeout(3000)
    browser.close()

# -----------------------------
# 🚀 Start Playwright
# -----------------------------
with sync_playwright() as playwright:
    run(playwright)
