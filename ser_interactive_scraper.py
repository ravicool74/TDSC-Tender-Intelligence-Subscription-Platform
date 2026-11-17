import time, os, sys, re

# Force UTF-8 stdout for Windows consoles
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def create_styled_excel(contractors, filename="SER_Contractors_Sample_5.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SER Contractors"

    headers = [
        "SL. No.", "ApplicationID", "Applicant/Firm/Company Name", "Applicant Father/Director Name",
        "DoB/DoA", "Address", "District", "Cast", "Gender", "Email ID", "Entity Type", "Mobile No", "PAN No"
    ]

    ws.append(headers)

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for r_idx, c in enumerate(contractors, start=2):
        row_vals = [
            c.get("sl_no", r_idx - 1),
            c.get("application_id", ""),
            c.get("company_name", ""),
            c.get("director_name", ""),
            c.get("dob_doa", ""),
            c.get("address", ""),
            c.get("district", ""),
            c.get("cast", ""),
            c.get("gender", ""),
            c.get("email", ""),
            c.get("entity_type", ""),
            c.get("mobile", ""),
            c.get("pan_no", "")
        ]
        ws.append(row_vals)
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=col_num)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if col_num in [1, 2, 5, 7, 8, 9, 11, 12, 13]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # Column auto-width
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    wb.save(filename)
    return os.path.abspath(filename)


def extract_fields_from_current_page(driver):
    """
    Extracts the 12 fields from whichever view is currently open (modal, popup, or page)
    """
    fields = {
        "application_id": "",
        "company_name": "",
        "director_name": "",
        "dob_doa": "",
        "address": "",
        "district": "",
        "cast": "",
        "gender": "",
        "email": "",
        "entity_type": "",
        "mobile": "",
        "pan_no": ""
    }

    try:
        body_elem = driver.find_element(By.TAG_NAME, "body")
        soup_text = body_elem.text
    except Exception:
        soup_text = driver.page_source

    # Strategy 1: Find by table label pairs (Label in td[0], Value in td[1])
    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        for tbl in tables:
            rows = tbl.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                for idx in range(len(cells) - 1):
                    lbl = cells[idx].text.strip().lower()
                    val = cells[idx + 1].text.strip()
                    if not val:
                        continue

                    if any(k in lbl for k in ["application id", "application no", "reg no", "registration no"]):
                        fields["application_id"] = val
                    elif any(k in lbl for k in ["applicant name", "firm name", "company name", "contractor name"]):
                        fields["company_name"] = val
                    elif any(k in lbl for k in ["father", "director", "husband"]):
                        fields["director_name"] = val
                    elif any(k in lbl for k in ["dob", "doa", "date of birth", "date of agreement"]):
                        fields["dob_doa"] = val
                    elif "address" in lbl:
                        fields["address"] = val
                    elif "district" in lbl:
                        fields["district"] = val
                    elif any(k in lbl for k in ["cast", "caste", "category"]):
                        fields["cast"] = val
                    elif "gender" in lbl or "sex" in lbl:
                        fields["gender"] = val
                    elif "email" in lbl:
                        fields["email"] = val
                    elif any(k in lbl for k in ["entity type", "firm type", "constitution"]):
                        fields["entity_type"] = val
                    elif any(k in lbl for k in ["mobile", "contact no", "phone"]):
                        fields["mobile"] = val
                    elif "pan" in lbl:
                        fields["pan_no"] = val
    except Exception as e:
        print(f"Table inspection warning: {e}")

    # Strategy 2: Regex scanning on whole text if any fields remain empty
    if not fields["application_id"]:
        m = re.search(r'\b(\d{6}[A-Za-z]{2}\d{8})\b', soup_text)
        if m: fields["application_id"] = m.group(1)

    if not fields["mobile"]:
        m = re.search(r'\b([6-9]\d{9})\b', soup_text)
        if m: fields["mobile"] = m.group(1)

    if not fields["pan_no"]:
        m = re.search(r'\b([A-Z]{5}\d{4}[A-Z])\b', soup_text)
        if m: fields["pan_no"] = m.group(1)

    if not fields["email"]:
        m = re.search(r'\b([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)\b', soup_text)
        if m: fields["email"] = m.group(1)

    return fields


def run_interactive_scraper(max_records=5):
    print("=" * 65)
    print("  SER CONTRACTOR CLASS 4 - BROWSER ASSISTED SCRAPER")
    print("=" * 65)
    print("\n[1/4] Launching Google Chrome window...")

    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=options)
    login_url = "https://www.ser.org.in/ContractorClass4/login.aspx"
    driver.get(login_url)

    print(f"\n[2/4] Chrome opened at: {login_url}")
    print("\n[*] KRIPYA BROWSER ME APNA USER ID AUR PASSWORD DAALKAR 'Sign In' KAREIN.")
    print("    (Website slow hai, araam se login karein. Script aapka wait kar rahi hai...)\n")

    # Wait for user to log in (URL leaves login.aspx)
    logged_in = False
    for sec in range(300): # 5 minutes timeout
        try:
            current_url = driver.current_url.lower()
            if "login.aspx" not in current_url:
                logged_in = True
                print(f"\n[+] Login Successful! Current URL: {driver.current_url}")
                break
        except Exception:
            pass
        time.sleep(2)

    if not logged_in:
        print("\n[-] Login timeout or browser closed. Please run the script again when ready.")
        try: driver.quit()
        except: pass
        return

    # Check if we are on ContractorSearchByName.aspx
    target_report_url = "https://www.ser.org.in/ContractorClass4/Report/ContractorSearchByName.aspx"
    try:
        if "contractorsearchbyname.aspx" not in driver.current_url.lower():
            print(f"\n[3/4] Navigating to Contractor Search Report: {target_report_url}")
            driver.get(target_report_url)
            time.sleep(5)
    except Exception as e:
        print(f"Navigation note: {e}")

    print("\n[*] BROWSER ME APNA DISTRICT CHUNEIN AUR 'Search' / 'Show' BUTTON DABAYEIN.")
    print("    Jab table screen par load ho jaye, tab script automatic 'Detail' button par click karegi...")
    print("    (Script waiting for contractor results table to appear...)\n")

    # Wait for the results table to appear
    detail_buttons = []
    
    for wait_count in range(120): # up to 4 minutes waiting for user to search
        try:
            # Look for rows with "Detail" or "View" button/link
            btns = driver.find_elements(By.XPATH, "//a[contains(translate(text(), 'DETAIL', 'detail'), 'detail') or contains(@id, 'btnDetail') or contains(@href, 'Detail')] | //input[@type='button' or @type='submit'][contains(translate(@value, 'DETAIL', 'detail'), 'detail')]")
            if btns and len(btns) >= 1:
                detail_buttons = btns
                print(f"[+] Contractor Table Found with {len(detail_buttons)} records!")
                break
        except Exception:
            pass
        time.sleep(2)

    if not detail_buttons:
        print("[-] Could not find 'Detail' buttons in the table automatically. Will inspect visible table rows...")

    main_window = driver.current_window_handle
    extracted_contractors = []

    limit = min(max_records, len(detail_buttons) if detail_buttons else 5)
    print(f"\n[4/4] Starting sequential extraction of top {limit} contractors (SL No. preserved)...")

    for i in range(limit):
        sl_no = i + 1
        print(f"\n--- Processing Record #{sl_no} ---")

        time.sleep(2)
        try:
            btns = driver.find_elements(By.XPATH, "//a[contains(translate(text(), 'DETAIL', 'detail'), 'detail') or contains(@id, 'btnDetail') or contains(@href, 'Detail')] | //input[contains(translate(@value, 'DETAIL', 'detail'), 'detail')]")
            if i >= len(btns):
                print(f"No detail button found at index {i}.")
                break
            target_btn = btns[i]

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_btn)
            time.sleep(1)

            before_windows = driver.window_handles
            driver.execute_script("arguments[0].click();", target_btn)

            print(f"Clicked 'Detail' button #{sl_no}. Waiting for detail page to load...")
            time.sleep(5) # wait for slow government portal to load detail

            after_windows = driver.window_handles
            if len(after_windows) > len(before_windows):
                new_win = [w for w in after_windows if w != main_window][-1]
                driver.switch_to.window(new_win)
                time.sleep(3)
                fields = extract_fields_from_current_page(driver)
                driver.close()
                driver.switch_to.window(main_window)
            else:
                time.sleep(3)
                fields = extract_fields_from_current_page(driver)
                
                # Check for modal close button
                close_btns = driver.find_elements(By.XPATH, "//button[contains(text(), 'Close') or contains(text(), '×')] | //a[contains(text(), 'Close') or contains(text(), 'Back')] | //input[@value='Close' or @value='Back']")
                if close_btns:
                    driver.execute_script("arguments[0].click();", close_btns[0])
                    time.sleep(2)

            fields["sl_no"] = sl_no
            print(f"  [+] Extracted SL #{sl_no}:")
            print(f"      AppID:    {fields.get('application_id') or 'N/A'}")
            print(f"      Firm:     {fields.get('company_name') or 'N/A'}")
            print(f"      Director: {fields.get('director_name') or 'N/A'}")
            print(f"      Mobile:   {fields.get('mobile') or 'N/A'}")
            print(f"      District: {fields.get('district') or 'N/A'}")
            print(f"      PAN:      {fields.get('pan_no') or 'N/A'}")

            extracted_contractors.append(fields)
        except Exception as e:
            print(f"Error on record #{sl_no}: {e}")

        time.sleep(2)

    # Save to Excel
    excel_path = create_styled_excel(extracted_contractors, "SER_Contractors_Sample_5.xlsx")
    print("\n" + "=" * 65)
    print(f"[+] EXTRACTION COMPLETE! Extracted {len(extracted_contractors)} records.")
    print(f"[+] Excel file saved at: {excel_path}")
    print("=" * 65)

    print("\nSummary Table:")
    for c in extracted_contractors:
        print(f"  SL #{c['sl_no']} | AppID: {c.get('application_id')} | Firm: {c.get('company_name')} | Mobile: {c.get('mobile')}")

    print("\nBrowser will remain open for your inspection.")
    time.sleep(60)

if __name__ == '__main__':
    run_interactive_scraper(max_records=5)
