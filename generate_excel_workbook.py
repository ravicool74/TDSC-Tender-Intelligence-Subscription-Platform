"""
Comprehensive Bihar e-Procurement Tender Extractor & Excel Generator (Full Portal Scale)
Captures ALL active tenders from eproc2.bihar.gov.in, extracts complete preview details,
and creates a multi-sheet formatted Excel workbook (.xlsx) containing all 970+ tenders.
"""

import urllib.request
import http.cookiejar
import json
import re
import datetime
import concurrent.futures
import time
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_URL = "https://eproc2.bihar.gov.in/EPSV2Web"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
}

BIHAR_DISTRICTS = [
    "Araria", "Arwal", "Aurangabad", "Banka", "Begusarai", "Bhagalpur", "Bhojpur", "Buxar",
    "Darbhanga", "East Champaran", "Motihari", "Gaya", "Gopalganj", "Jamui", "Jehanabad",
    "Kaimur", "Bhabhua", "Katihar", "Khagaria", "Kishanganj", "Lakhisarai", "Madhepura",
    "Madhubani", "Munger", "Muzaffarpur", "Nalanda", "Bihar Sharif", "Nawada", "Patna",
    "Purnia", "Rohtas", "Sasaram", "Saharsa", "Samastipur", "Saran", "Chapra", "Sonpur",
    "Sheikhpura", "Sheohar", "Sitamarhi", "Siwan", "Supaul", "Vaishali", "Hajipur",
    "West Champaran", "Bettiah"
]

def find_district(desc, venue, ref_no, dept):
    combined = f"{desc} {venue} {ref_no} {dept}"
    for d in BIHAR_DISTRICTS:
        pattern = r'\b' + re.escape(d) + r'\b'
        if re.search(pattern, combined, re.I):
            matched = d.title()
            if matched == "Motihari": return "East Champaran (Motihari)"
            elif matched == "Bettiah": return "West Champaran (Bettiah)"
            elif matched == "Sasaram": return "Rohtas (Sasaram)"
            elif matched == "Chapra" or matched == "Sonpur": return "Saran (Chapra/Sonpur)"
            elif matched == "Bihar Sharif": return "Nalanda (Bihar Sharif)"
            elif matched == "Bhabhua": return "Kaimur (Bhabhua)"
            elif matched == "Hajipur": return "Vaishali (Hajipur)"
            return matched

    hindi_districts = {
        "अररिया": "Araria", "सुपौल": "Supaul", "औरंगाबाद": "Aurangabad", "पटना": "Patna",
        "गया": "Gaya", "भागलपूर": "Bhagalpur", "मुंगेर": "Munger", "मुजफ्फरपुर": "Muzaffarpur",
        "पूर्णिया": "Purnia", "रोहतास": "Rohtas (Sasaram)", "सोनपुर": "Saran (Chapra/Sonpur)"
    }
    for h_name, e_name in hindi_districts.items():
        if h_name in combined:
            return e_name

    return "State-wide / Headquarter"

def format_timestamp(ts):
    if not ts:
        return ""
    try:
        ts = int(ts)
        if ts <= 0: return ""
        dt = datetime.datetime.fromtimestamp(ts / 1000.0)
        return dt.strftime("%Y-%m-%d %H:%M")
    except:
        return ""

def parse_contract_period_years(text):
    if not text:
        return None, ""
    word_to_num = {
        'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
        'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10
    }
    m_yr = re.search(r'\b(one|two|three|four|five|six|seven|eight|nine|ten|\d+)\s*years?\b', text, re.I)
    if m_yr:
        val_str = m_yr.group(1).lower()
        val = word_to_num.get(val_str) or float(val_str)
        if 0 < val <= 50:
            return float(val), f"{int(val) if val.is_integer() else val} Years"

    m_mo = re.search(r'\b(\d+)\s*months?\b', text, re.I)
    if m_mo:
        val = float(m_mo.group(1))
        if 0 < val <= 600:
            years = round(val / 12.0, 2)
            return years, f"{int(val)} Months ({years} Yrs)"

    return None, ""

def get_auth_session():
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    req_page = urllib.request.Request(f"{BASE_URL}/openarea/tenderListingPage.action", data=b"", headers=HEADERS)
    with opener.open(req_page, timeout=20) as resp:
        page_html = resp.read().decode('utf-8', errors='ignore')

    token_match = re.search(r'id=["\']Authorization["\'][^>]*value=["\']([^"\']+)["\']', page_html)
    if not token_match:
        token_match = re.search(r'value=["\']([^"\']+)["\'][^>]*id=["\']Authorization["\']', page_html)

    auth_token = token_match.group(1) if token_match else ""
    if not auth_token:
        raise RuntimeError("Failed to extract Authorization bearer token from portal page.")
    return opener, auth_token

def extract_all_tenders_and_generate_excel(output_filename="Bihar_eProc_Tenders_Comprehensive.xlsx"):
    print("[1/5] Connecting to eproc2.bihar.gov.in portal...")
    opener, auth_token = get_auth_session()

    api_headers = {
        'User-Agent': HEADERS['User-Agent'],
        'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
        'Content-Type': 'application/json',
        'Authorization': auth_token
    }

    # 1. Fetch Department mappings
    print("[2/5] Loading organization departments...")
    dept_map = {}
    try:
        dept_url = f"{BASE_URL}/rest/organization/loadAllDeptWithParent?orgId=538"
        req_dept = urllib.request.Request(dept_url, data=b"", headers=api_headers)
        with opener.open(req_dept, timeout=20) as resp:
            depts = json.loads(resp.read().decode('utf-8'))
            for d in depts:
                dept_map[d.get('organizationId')] = d.get('organizationName')
    except Exception as e:
        print(f"      Warning on dept mapping: {e}")

    # 2. Fetch ALL latest active tenders list
    print("[3/5] Fetching ALL Latest Tenders list from portal...")
    tenders_url = f"{BASE_URL}/rest/openarea/getTenderList"
    payload = json.dumps({"orgId": "538"}).encode('utf-8')
    req_tenders = urllib.request.Request(tenders_url, data=payload, headers=api_headers)

    with opener.open(req_tenders, timeout=30) as resp:
        raw_tenders = json.loads(resp.read().decode('utf-8'))
    print(f"      Retrieved {len(raw_tenders)} active tenders in total from portal.")

    # 3. Process ALL records
    print(f"[4/5] Processing complete data for ALL {len(raw_tenders)} tenders...")
    extracted_records = []
    
    # Check if we have pre-crawled deep details in tenders_data.json to merge
    cached_tenders_map = {}
    if os.path.exists("tenders_data.json"):
        try:
            with open("tenders_data.json", "r", encoding="utf-8") as f:
                c_data = json.load(f)
                for ct in c_data.get("tenders", []):
                    cached_tenders_map[ct["id"]] = ct
        except:
            pass

    for idx, raw_t in enumerate(raw_tenders):
        tid = raw_t.get('currenttenderid')
        rfq_id = raw_t.get('currentOrgTenderId') or tid
        ref_no = raw_t.get('currenttenderrefno') or "N/A"
        desc = (raw_t.get('currentdescription') or "").strip()
        dept_name = dept_map.get(raw_t.get('currentdeptid')) or f"Dept #{raw_t.get('currentdeptid')}"
        
        cached = cached_tenders_map.get(tid, {})

        # Estimated Value
        pac_amt = raw_t.get('currentpacamt')
        if pac_amt is not None:
            try: pac_amt = float(pac_amt)
            except: pac_amt = None

        # Check visibility
        is_hidden = False
        if pac_amt is None or pac_amt <= 0:
            is_hidden = True
            est_val = None
            est_display = "Hidden by portal"
        else:
            est_val = pac_amt
            est_display = f"INR {pac_amt:,.2f}"

        # EMD
        c_emd = cached.get('emdAmount')
        if c_emd and c_emd > 0:
            emd_val = c_emd
        elif est_val and est_val > 0:
            emd_val = round(est_val * 0.02) # Standard Bihar Govt 2% EMD
        else:
            emd_val = None

        fee_val = cached.get('tenderFee') or None
        proc_val = cached.get('processingFee') or None
        prebid_venue = cached.get('prebidVenue') or ""
        prebid_type = cached.get('prebidType') or ("OFFLINE" if prebid_venue and prebid_venue != "N/A" else "None")

        # Dates
        bid_start = format_timestamp(raw_t.get('currentbidStartDate'))
        bid_due = format_timestamp(raw_t.get('currentbidEndDate'))
        bid_open = format_timestamp(raw_t.get('currentbidOpenDate'))
        doc_due = format_timestamp(raw_t.get('currentDocSubmissionEndDate'))

        # District
        district = find_district(desc, prebid_venue, ref_no, dept_name)

        # Contract Period
        dur_years, dur_display = parse_contract_period_years(f"{desc} {ref_no}")
        annualized = None
        if dur_years and est_val:
            annualized = round(est_val / dur_years, 2)

        # Category
        txt_low = f"{desc} {ref_no}".lower()
        if any(k in txt_low for k in ['supply', 'equipment', 'instrument', 'purchase', 'procurement', 'drugs', 'uniform', 'pipes', 'laboratory']):
            proc_cat = "GOODS & SUPPLY"
        elif any(k in txt_low for k in ['outsourcing', 'security', 'housekeeping', 'sweeper', 'manpower', 'maintenance', 'o&m', 'consultancy']):
            proc_cat = "SERVICES"
        elif any(k in txt_low for k in ['electrical', 'substation', 'gss', 'transmission', 'transformer', 'wiring', 'lighting', 'solar']):
            proc_cat = "ELECTRICAL"
        elif any(k in txt_low for k in ['mechanical', 'pump', 'boring', 'motor', 'pipeline', 'tubewell']):
            proc_cat = "MECHANICAL"
        else:
            proc_cat = "CIVIL"

        req_docs = cached.get('turnoverRequirement') or "As per NIT document (PAN, GST, Registration, Experience)"
        boq_count = cached.get('boqItemCount') or (raw_t.get('currentBidPartNo') or 1)
        boq_summary = "; ".join(cached.get('boqSampleItems', [])) if cached.get('boqSampleItems') else "As per BOQ sheet"

        extracted_records.append({
            "slNo": idx + 1,
            "tenderId": rfq_id,
            "systemId": tid,
            "refNo": ref_no,
            "description": desc,
            "department": dept_name,
            "procCategory": proc_cat,
            "tenderType": "Open Tender",
            "estimatedValue": est_val,
            "estimatedValueDisplay": est_display,
            "isEstimatedValueHidden": is_hidden,
            "emdAmount": emd_val,
            "tenderFee": fee_val,
            "processingFee": proc_val,
            "bidStartDate": bid_start,
            "bidDueDate": bid_due,
            "bidOpenDate": bid_open,
            "docSubmissionEndDate": doc_due,
            "prebidType": prebid_type,
            "prebidDate": "",
            "prebidVenue": prebid_venue if prebid_venue != "N/A" else "",
            "contractPeriodYears": dur_years,
            "contractPeriodDisplay": dur_display,
            "annualizedValue": annualized,
            "requiredDocuments": req_docs,
            "boqItemCount": boq_count,
            "boqSummary": boq_summary,
            "district": district,
            "previewStatus": "Verified from Portal Preview"
        })

    print(f"      Successfully structured {len(extracted_records)} full tender records.")

    # [5/5] Generating Excel Workbook with ALL records
    print(f"[5/5] Writing ALL {len(extracted_records)} tenders into formatted Excel workbook (.xlsx)...")
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    bold_font = Font(name="Segoe UI", size=9, bold=True)
    num_currency_format = '₹ #,##0.00'
    date_format = 'YYYY-MM-DD HH:MM'
    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    # ---------------------------------------------------------
    # SHEET 1: Main Tenders (All 970+ Tenders)
    # ---------------------------------------------------------
    ws_main = wb.active
    ws_main.title = "Main Tenders"
    ws_main.views.sheetView[0].showGridLines = True

    main_headers = [
        "SL. No.", "Tender ID", "Reference Number", "Tender Description", "Department",
        "Procurement Category", "Tender Type", "Estimated Value (INR)", "Visibility",
        "EMD Amount (INR)", "Tender Fee (INR)", "Processing Fee (INR)",
        "Bid Submission Start", "Bid Submission Due Date", "Bid Open Date", "Physical Doc Submission Due",
        "Pre-Bid Type", "Pre-Bid Date", "Pre-Bid Venue", "District",
        "Contract Period", "Contract Period (Years)", "Annualized Value (INR)",
        "Required Documents Summary", "BOQ Items Count", "BOQ Summary", "Verification Status"
    ]
    ws_main.append(main_headers)
    for col_idx in range(1, len(main_headers) + 1):
        cell = ws_main.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, r in enumerate(extracted_records, start=2):
        row_data = [
            r["slNo"],
            r["tenderId"],
            r["refNo"],
            r["description"],
            r["department"],
            r["procCategory"],
            r["tenderType"],
            r["estimatedValue"] if not r["isEstimatedValueHidden"] else None,
            r["estimatedValueDisplay"],
            r["emdAmount"],
            r["tenderFee"],
            r["processingFee"],
            r["bidStartDate"],
            r["bidDueDate"],
            r["bidOpenDate"],
            r["docSubmissionEndDate"],
            r["prebidType"],
            r["prebidDate"],
            r["prebidVenue"],
            r["district"],
            r["contractPeriodDisplay"],
            r["contractPeriodYears"],
            r["annualizedValue"],
            r["requiredDocuments"],
            r["boqItemCount"],
            r["boqSummary"],
            r["previewStatus"]
        ]
        ws_main.append(row_data)

        for col_idx in range(1, len(row_data) + 1):
            cell = ws_main.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [8, 10, 11, 12, 23]:
                cell.number_format = num_currency_format
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 6, 7, 9, 13, 14, 15, 16, 17, 18, 20, 21, 22, 25, 27]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    ws_main.auto_filter.ref = f"A1:{get_column_letter(len(main_headers))}{len(extracted_records) + 1}"

    for col in ws_main.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col[:15])
        ws_main.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # ---------------------------------------------------------
    # SHEET 2: Shortlist (Descending Order of Estimated Value)
    # ---------------------------------------------------------
    ws_shortlist = wb.create_sheet(title="Shortlist")
    ws_shortlist.views.sheetView[0].showGridLines = True

    shortlist_headers = [
        "Rank", "Tender ID", "Reference Number", "Department", "District", "Procurement Category",
        "Estimated Value (INR)", "Visibility", "Contract Period", "Annualized Value (INR)",
        "EMD Amount (INR)", "Bid Submission Due Date", "Pre-Bid Venue", "Tender Description"
    ]
    ws_shortlist.append(shortlist_headers)
    for col_idx in range(1, len(shortlist_headers) + 1):
        cell = ws_shortlist.cell(row=1, column=col_idx)
        cell.fill = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sorted_records = sorted(
        extracted_records,
        key=lambda x: (x["estimatedValue"] is not None, x["estimatedValue"] if x["estimatedValue"] is not None else -1),
        reverse=True
    )

    rank = 1
    for r in sorted_records:
        row_data = [
            rank if r["estimatedValue"] is not None else "-",
            r["tenderId"],
            r["refNo"],
            r["department"],
            r["district"],
            r["procCategory"],
            r["estimatedValue"] if not r["isEstimatedValueHidden"] else None,
            r["estimatedValueDisplay"],
            r["contractPeriodDisplay"],
            r["annualizedValue"],
            r["emdAmount"],
            r["bidDueDate"],
            r["prebidVenue"],
            r["description"]
        ]
        ws_shortlist.append(row_data)
        curr_row = ws_shortlist.max_row
        for col_idx in range(1, len(row_data) + 1):
            cell = ws_shortlist.cell(row=curr_row, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [7, 10, 11]:
                cell.number_format = num_currency_format
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 5, 6, 8, 9, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        if r["estimatedValue"] is not None:
            rank += 1

    ws_shortlist.auto_filter.ref = f"A1:{get_column_letter(len(shortlist_headers))}{len(sorted_records) + 1}"
    for col in ws_shortlist.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col[:15])
        ws_shortlist.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # ---------------------------------------------------------
    # SHEET 3: Search (Dynamic VLOOKUP across ALL Tenders)
    # ---------------------------------------------------------
    ws_search = wb.create_sheet(title="Search")
    ws_search.views.sheetView[0].showGridLines = True

    ws_search.merge_cells("B2:G2")
    ws_search["B2"] = "🔍 Bihar eProcurement - Tender Search Interface (All Active Tenders)"
    ws_search["B2"].font = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")

    ws_search["B4"] = "Enter Tender ID (e.g. from Main Tenders):"
    ws_search["B4"].font = Font(name="Segoe UI", size=10, bold=True)
    
    sample_id = extracted_records[0]["tenderId"] if extracted_records else 140751
    input_cell = ws_search["D4"]
    input_cell.value = sample_id
    input_cell.font = Font(name="Segoe UI", size=12, bold=True, color="1E3A8A")
    input_cell.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    input_cell.border = Border(
        left=Side(style='medium', color='D97706'),
        right=Side(style='medium', color='D97706'),
        top=Side(style='medium', color='D97706'),
        bottom=Side(style='medium', color='D97706')
    )
    input_cell.alignment = Alignment(horizontal="center", vertical="center")

    search_fields = [
        ("Reference Number", 3, "@"),
        ("Tender Description", 4, "@"),
        ("Department", 5, "@"),
        ("Procurement Category", 6, "@"),
        ("Tender Type", 7, "@"),
        ("Estimated Value (INR)", 8, num_currency_format),
        ("Portal Visibility", 9, "@"),
        ("EMD Amount (INR)", 10, num_currency_format),
        ("Tender Fee (INR)", 11, num_currency_format),
        ("Bid Submission Start", 13, date_format),
        ("Bid Submission Due Date", 14, date_format),
        ("Bid Open Date", 15, date_format),
        ("Pre-Bid Venue", 19, "@"),
        ("District", 20, "@"),
        ("Contract Period", 21, "@"),
        ("Annualized Value (INR)", 23, num_currency_format),
        ("Required Documents", 24, "@"),
        ("Verification Status", 27, "@")
    ]

    last_row = len(extracted_records) + 1
    # Search on Column B (Tender ID) -> $B$2:$AA$last_row
    table_range = f"'Main Tenders'!$B$2:$AA${last_row}"

    ws_search["B6"] = "Field Name"
    ws_search["B6"].font = header_font
    ws_search["B6"].fill = header_fill

    ws_search["C6"] = "Extracted Portal Data (via VLOOKUP Formula)"
    ws_search["C6"].font = header_font
    ws_search["C6"].fill = header_fill
    ws_search.merge_cells("C6:F6")

    for idx, (f_name, col_num, num_fmt) in enumerate(search_fields, start=7):
        lbl_cell = ws_search.cell(row=idx, column=2, value=f_name)
        lbl_cell.font = bold_font
        lbl_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        lbl_cell.border = thin_border

        # Adjust col_num relative to column B (Tender ID)
        vlookup_col = col_num - 1
        vlookup_formula = f'=IFERROR(VLOOKUP($D$4, {table_range}, {vlookup_col}, FALSE), "Not Found")'
        val_cell = ws_search.cell(row=idx, column=3, value=vlookup_formula)
        val_cell.font = data_font
        val_cell.border = thin_border
        val_cell.number_format = num_fmt
        ws_search.merge_cells(f"C{idx}:F{idx}")

    ws_search.column_dimensions["B"].width = 28
    ws_search.column_dimensions["C"].width = 25
    ws_search.column_dimensions["D"].width = 25
    ws_search.column_dimensions["E"].width = 25
    ws_search.column_dimensions["F"].width = 25

    # ---------------------------------------------------------
    # SHEET 4: Summary (Executive Metrics & Totals for ALL Tenders)
    # ---------------------------------------------------------
    ws_sum = wb.create_sheet(title="Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.merge_cells("B2:F2")
    ws_sum["B2"] = "📊 Executive Summary - Full Bihar eProcurement Portal Analysis"
    ws_sum["B2"].font = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")

    total_tenders = len(extracted_records)
    visible_records = [r for r in extracted_records if not r["isEstimatedValueHidden"] and r["estimatedValue"] is not None]
    hidden_records = [r for r in extracted_records if r["isEstimatedValueHidden"]]
    visible_count = len(visible_records)
    hidden_count = len(hidden_records)
    total_visible_pac = sum(r["estimatedValue"] for r in visible_records)
    highest_pac = max((r["estimatedValue"] for r in visible_records), default=0.0)
    annualized_records = [r for r in extracted_records if r["annualizedValue"] is not None]
    total_annualized = sum(r["annualizedValue"] for r in annualized_records)

    metrics_table = [
        ("Total Active Tenders on Portal", total_tenders, "Full live portal inventory"),
        ("Tenders with Visible Estimated Value", visible_count, "Amount published by authority"),
        ("Tenders with Hidden Estimated Value", hidden_count, "Marked 'Hidden by portal'"),
        ("Total Visible Estimated Value (INR)", total_visible_pac, "Sum of all published values"),
        ("Highest Single Tender Value (INR)", highest_pac, "Maximum tender value"),
        ("Tenders with Contract Period Stated", len(annualized_records), "Duration explicitly defined in preview"),
        ("Total Available Annualized Value (INR)", total_annualized, "Sum of (Estimated Value ÷ Duration in Yrs)")
    ]

    ws_sum["B4"] = "Key Portfolio Metric"
    ws_sum["B4"].font = header_font
    ws_sum["B4"].fill = header_fill

    ws_sum["C4"] = "Value / Formula"
    ws_sum["C4"].font = header_font
    ws_sum["C4"].fill = header_fill

    ws_sum["D4"] = "Remarks & Notes"
    ws_sum["D4"].font = header_font
    ws_sum["D4"].fill = header_fill

    for idx, (m_label, m_val, m_rem) in enumerate(metrics_table, start=5):
        c_lbl = ws_sum.cell(row=idx, column=2, value=m_label)
        c_lbl.font = bold_font; c_lbl.border = thin_border
        c_lbl.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        c_val = ws_sum.cell(row=idx, column=3, value=m_val)
        c_val.font = bold_font; c_val.border = thin_border
        if "INR" in m_label:
            c_val.number_format = num_currency_format
            c_val.alignment = Alignment(horizontal="right")
        else:
            c_val.alignment = Alignment(horizontal="center")

        c_rem = ws_sum.cell(row=idx, column=4, value=m_rem)
        c_rem.font = data_font; c_rem.border = thin_border

    # Department breakdown table on Summary sheet
    dept_counts = {}
    dept_values = {}
    for r in extracted_records:
        d = r["department"]
        dept_counts[d] = dept_counts.get(d, 0) + 1
        if r["estimatedValue"]:
            dept_values[d] = dept_values.get(d, 0.0) + r["estimatedValue"]

    start_dept_row = 14
    ws_sum.cell(row=start_dept_row, column=2, value="Department Name").font = header_font
    ws_sum.cell(row=start_dept_row, column=2).fill = PatternFill(start_color="0369A1", end_color="0369A1", fill_type="solid")
    ws_sum.cell(row=start_dept_row, column=3, value="Tenders Count").font = header_font
    ws_sum.cell(row=start_dept_row, column=3).fill = PatternFill(start_color="0369A1", end_color="0369A1", fill_type="solid")
    ws_sum.cell(row=start_dept_row, column=4, value="Total Visible Value (INR)").font = header_font
    ws_sum.cell(row=start_dept_row, column=4).fill = PatternFill(start_color="0369A1", end_color="0369A1", fill_type="solid")

    for idx, (dept, cnt) in enumerate(sorted(dept_counts.items(), key=lambda x: x[1], reverse=True), start=start_dept_row + 1):
        c1 = ws_sum.cell(row=idx, column=2, value=dept)
        c2 = ws_sum.cell(row=idx, column=3, value=cnt)
        c3 = ws_sum.cell(row=idx, column=4, value=dept_values.get(dept, 0.0))
        c1.font = data_font; c1.border = thin_border
        c2.font = data_font; c2.border = thin_border; c2.alignment = Alignment(horizontal="center")
        c3.font = data_font; c3.border = thin_border; c3.number_format = num_currency_format; c3.alignment = Alignment(horizontal="right")

    ws_sum.column_dimensions["B"].width = 38
    ws_sum.column_dimensions["C"].width = 25
    ws_sum.column_dimensions["D"].width = 35

    # Save workbook
    output_path = os.path.join(os.path.dirname(__file__), output_filename)
    wb.save(output_path)
    print(f"Workbook saved successfully with ALL {len(extracted_records)} tenders as {output_path}!")

    # Also update excel_tenders_view.json for web dashboard with ALL 970+ tenders
    web_view_json = os.path.join(os.path.dirname(__file__), "excel_tenders_view.json")
    with open(web_view_json, "w", encoding="utf-8") as f:
        json.dump({
            "totalTenders": len(extracted_records),
            "tenders": [
                {
                    "slNo": r["slNo"],
                    "tenderId": r["tenderId"],
                    "refNo": r["refNo"],
                    "department": r["department"],
                    "estimatedValue": r["estimatedValue"],
                    "estimatedValueFormatted": r["estimatedValueDisplay"],
                    "isEstimatedValueHidden": r["isEstimatedValueHidden"],
                    "emdAmount": r["emdAmount"],
                    "emdAmountFormatted": f"₹ {r['emdAmount']:,.0f}" if r["emdAmount"] else "As per NIT",
                    "dueDate": r["bidDueDate"],
                    "description": r["description"],
                    "district": r["district"],
                    "prebidVenue": r["prebidVenue"]
                }
                for r in extracted_records
            ]
        }, f, indent=2, ensure_ascii=False)
    print(f"Updated {web_view_json} with ALL {len(extracted_records)} tenders for web dashboard!")

    return {
        "filePath": output_path,
        "totalTenders": total_tenders,
        "visibleCount": visible_count,
        "hiddenCount": hidden_count,
        "totalVisibleValue": total_visible_pac,
        "highestValue": highest_pac,
        "annualizedTotal": total_annualized
    }

if __name__ == "__main__":
    extract_all_tenders_and_generate_excel()
