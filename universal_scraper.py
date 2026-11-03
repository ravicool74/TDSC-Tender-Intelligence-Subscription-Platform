"""
Universal Tender & Webpage Scraper Webhook Engine
Extracts tabular and structured data from any website URL or e-procurement portal,
normalizes it into the standard comprehensive fields, generates an Excel workbook,
and makes it available for shortlisting and dashboard analysis.
"""

import os
import re
import sys
import json
import time
import urllib.request
import urllib.parse
import ssl
from datetime import datetime
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DIRECTORY = os.path.dirname(os.path.abspath(__file__))

BIHAR_DISTRICTS = [
    "Araria", "Arwal", "Aurangabad", "Banka", "Begusarai", "Bhagalpur", "Bhojpur",
    "Buxar", "Darbhanga", "East Champaran", "Gaya", "Gopalganj", "Jamui", "Jehanabad",
    "Kaimur", "Katihar", "Khagaria", "Kishanganj", "Lakhisarai", "Madhepura",
    "Madhubani", "Munger", "Muzaffarpur", "Nalanda", "Nawada", "Patna", "Purnia",
    "Rohtas", "Saharsa", "Samastipur", "Saran", "Sheikhpura", "Sheohar",
    "Sitamarhi", "Siwan", "Supaul", "Vaishali", "West Champaran"
]

DISTRICT_ALIASES = {
    "motihari": "East Champaran",
    "bettiah": "West Champaran",
    "biharsharif": "Nalanda",
    "bihar sharif": "Nalanda",
    "ara": "Bhojpur",
    "arrah": "Bhojpur",
    "chhapra": "Saran",
    "chapra": "Saran",
    "sasaram": "Rohtas",
    "bhabua": "Kaimur",
    "hajipur": "Vaishali",
    "purnea": "Purnia",
    "monghyr": "Munger"
}

def clean_text(text):
    if not text:
        return ""
    text = re.sub(r'[\r\n\t]+', ' ', str(text))
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()

def parse_currency_amount(val_str):
    if not val_str:
        return None, "Hidden by portal", True
    
    clean = clean_text(val_str).lower()
    if any(k in clean for k in ['hidden', 'not disclosed', 'na', 'n/a', 'to be quoted', 'refer nit', 'as per tender']):
        return None, "Hidden by portal", True
    
    cr_match = re.search(r'([\d,\.]+)\s*(?:cr|crore|crores)', clean)
    if cr_match:
        try:
            num = float(cr_match.group(1).replace(',', ''))
            val = num * 10000000.0
            return val, f"₹ {val:,.2f}", False
        except:
            pass

    lakh_match = re.search(r'([\d,\.]+)\s*(?:lakh|lakhs|lac|lacs)', clean)
    if lakh_match:
        try:
            num = float(lakh_match.group(1).replace(',', ''))
            val = num * 100000.0
            return val, f"₹ {val:,.2f}", False
        except:
            pass

    num_match = re.search(r'(?:rs\.?|inr|₹)?\s*([\d,]+(?:\.\d+)?)', clean)
    if num_match:
        try:
            raw_num_str = num_match.group(1).replace(',', '')
            if raw_num_str:
                val = float(raw_num_str)
                if val > 0:
                    return val, f"₹ {val:,.2f}", False
        except:
            pass

    return None, "Hidden by portal", True

def detect_district(text):
    if not text:
        return "State-wide / Headquarter"
    
    text_lower = text.lower()
    for alias, standard in DISTRICT_ALIASES.items():
        if re.search(rf'\b{re.escape(alias)}\b', text_lower):
            return standard

    for d in BIHAR_DISTRICTS:
        if re.search(rf'\b{re.escape(d.lower())}\b', text_lower):
            return d

    loc_match = re.search(r'(?:district|dist\.?|at|in|location)\s*[:\-]?\s*([A-Za-z]+)', text, re.IGNORECASE)
    if loc_match:
        cand = loc_match.group(1).capitalize()
        if cand.lower() not in ['the', 'bihar', 'india', 'all', 'any']:
            return cand

    return "State-wide / Headquarter"

def fetch_url_content(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8',
        'Connection': 'keep-alive',
        'Cache-Control': 'max-age=0'
    }
    
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=25) as response:
            content_type = response.headers.get('content-type', '').lower()
            charset = 'utf-8'
            if 'charset=' in content_type:
                charset = content_type.split('charset=')[-1].split(';')[0].strip()
            raw_data = response.read()
            try:
                return raw_data.decode(charset, errors='replace'), content_type
            except:
                return raw_data.decode('utf-8', errors='replace'), content_type
    except Exception as e_ssl:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        with urllib.request.urlopen(req, context=ctx, timeout=30) as response:
            content_type = response.headers.get('content-type', '').lower()
            raw_data = response.read()
            return raw_data.decode('utf-8', errors='replace'), content_type

def parse_html_tables(soup, source_url):
    tables = soup.find_all('table')
    if not tables:
        return []

    keywords = ['tender', 'nit', 'work', 'bid', 'department', 'estimated', 'amount', 'due date', 'closing', 'emd', 'ref', 'sl', 'no']
    best_table = None
    best_score = -1

    for table in tables:
        rows = table.find_all('tr')
        if len(rows) < 2:
            continue
        table_text = table.get_text().lower()
        score = sum(1 for kw in keywords if kw in table_text) * 10 + len(rows)
        if score > best_score:
            best_score = score
            best_table = table

    if not best_table:
        return []

    rows = best_table.find_all('tr')
    if not rows:
        return []

    header_row = rows[0]
    headers = [clean_text(th.get_text()).lower() for th in header_row.find_all(['th', 'td'])]

    data_start_idx = 1
    if not any(any(k in h for k in ['tender', 'sl', 's.no', 'nit', 'work', 'date', 'amount', 'dept']) for h in headers) and len(rows) > 1:
        headers = [clean_text(th.get_text()).lower() for th in rows[1].find_all(['th', 'td'])]
        data_start_idx = 2

    col_map = {
        'sl': -1, 'id': -1, 'ref': -1, 'dept': -1, 'desc': -1,
        'val': -1, 'emd': -1, 'date': -1, 'dist': -1
    }

    for idx, h in enumerate(headers):
        if col_map['sl'] == -1 and any(k in h for k in ['sl', 's.no', 'sr.no', 'serial', 'क्रमांक', '#']):
            col_map['sl'] = idx
        elif col_map['id'] == -1 and any(k in h for k in ['tender id', 'bid id', 'id', 'tender no']):
            col_map['id'] = idx
        elif col_map['ref'] == -1 and any(k in h for k in ['ref', 'nit', 'notice no', 'tender ref']):
            col_map['ref'] = idx
        elif col_map['dept'] == -1 and any(k in h for k in ['dept', 'department', 'organization', 'authority', 'buyer', 'office', 'ministry']):
            col_map['dept'] = idx
        elif col_map['desc'] == -1 and any(k in h for k in ['desc', 'title', 'work', 'name of work', 'item', 'subject', 'particular']):
            col_map['desc'] = idx
        elif col_map['val'] == -1 and any(k in h for k in ['value', 'amount', 'cost', 'estimated', 'pac', 'ecv', 'price', 'राशि']):
            col_map['val'] = idx
        elif col_map['emd'] == -1 and any(k in h for k in ['emd', 'earnest', 'security deposit', 'धरोहर']):
            col_map['emd'] = idx
        elif col_map['date'] == -1 and any(k in h for k in ['due date', 'closing', 'submission', 'last date', 'end date', 'opening']):
            col_map['date'] = idx
        elif col_map['dist'] == -1 and any(k in h for k in ['district', 'location', 'place', 'city', 'site', 'जिला']):
            col_map['dist'] = idx

    num_cols = len(headers)
    if col_map['desc'] == -1:
        longest_col = 0
        longest_len = 0
        for c in range(min(num_cols, 10)):
            sample_len = sum(len(clean_text(r.find_all(['td', 'th'])[c].get_text())) for r in rows[data_start_idx:min(data_start_idx+5, len(rows))] if len(r.find_all(['td', 'th'])) > c)
            if sample_len > longest_len:
                longest_len = sample_len
                longest_col = c
        col_map['desc'] = longest_col

    records = []
    sl_counter = 1

    for row in rows[data_start_idx:]:
        cells = row.find_all(['td', 'th'])
        if len(cells) < 2:
            continue

        def get_cell(col_idx):
            if col_idx >= 0 and col_idx < len(cells):
                return clean_text(cells[col_idx].get_text())
            return ""

        raw_sl = get_cell(col_map['sl'])
        raw_id = get_cell(col_map['id'])
        raw_ref = get_cell(col_map['ref'])
        raw_dept = get_cell(col_map['dept'])
        raw_desc = get_cell(col_map['desc'])
        raw_val = get_cell(col_map['val'])
        raw_emd = get_cell(col_map['emd'])
        raw_date = get_cell(col_map['date'])
        raw_dist = get_cell(col_map['dist'])

        if not raw_id:
            for c in cells:
                c_text = clean_text(c.get_text())
                m = re.search(r'\b\d{5,9}\b', c_text)
                if m:
                    raw_id = m.group(0)
                    break
        if not raw_id:
            raw_id = f"TND-{sl_counter:04d}"

        if not raw_ref:
            raw_ref = f"REF/{raw_id}"

        if not raw_dept:
            raw_dept = urllib.parse.urlparse(source_url).netloc

        if not raw_desc:
            raw_desc = " ".join(clean_text(c.get_text()) for c in cells[:4])

        val_num, val_fmt, is_val_hidden = parse_currency_amount(raw_val)
        emd_num, emd_fmt, is_emd_hidden = parse_currency_amount(raw_emd)
        if is_emd_hidden:
            emd_fmt = "As per NIT" if not raw_emd else raw_emd

        if not raw_date:
            for c in cells:
                c_text = clean_text(c.get_text())
                d_match = re.search(r'\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}(?:\s+\d{1,2}:\d{2})?\b', c_text)
                if d_match:
                    raw_date = d_match.group(0)
                    break
        if not raw_date:
            raw_date = "Refer Notice"

        final_dist = raw_dist if raw_dist else detect_district(raw_desc + " " + raw_dept + " " + raw_ref)

        records.append({
            "slNo": sl_counter,
            "systemTenderNo": raw_id,
            "tenderId": raw_id,
            "refNo": raw_ref,
            "department": raw_dept,
            "district": final_dist,
            "estimatedValue": val_num,
            "estimatedValueFormatted": val_fmt,
            "isEstimatedValueHidden": is_val_hidden,
            "description": raw_desc,
            "detailedDescription": raw_desc,
            "startDate": "Refer Notice",
            "dueDate": raw_date,
            "openDate": "Refer Notice",
            "emdAmount": emd_num,
            "emdAmountFormatted": emd_fmt,
            "emdMode": "As per NIT",
            "tenderFee": 0.0,
            "tenderFeeFormatted": "As per NIT",
            "tenderFeeMode": "As per NIT",
            "processingFee": 0.0,
            "processingFeeFormatted": "As per NIT",
            "processingFeeMode": "As per NIT",
            "termsFiles": [],
            "officialFiles": [],
            "boqFiles": [],
            "hasDownloaded": False,
            "actionExtracted": False
        })
        sl_counter += 1

    return records

def parse_card_elements(soup, source_url):
    cards = soup.find_all(lambda tag: tag.name in ['div', 'li', 'article'] and any(
        cls in ' '.join(tag.get('class', [])).lower() for cls in ['tender', 'bid', 'item', 'card', 'result', 'notice', 'listing']
    ))

    if not cards or len(cards) < 2:
        return []

    records = []
    sl_counter = 1

    for card in cards:
        text = clean_text(card.get_text())
        if len(text) < 20:
            continue

        title_tag = card.find(['h2', 'h3', 'h4', 'h5', 'a', 'strong'])
        title = clean_text(title_tag.get_text()) if title_tag else text[:120]

        id_match = re.search(r'(?:tender\s*id|bid\s*no|nit\s*no|ref)\s*[:#\-]?\s*([A-Za-z0-9\/\-_]+)', text, re.IGNORECASE)
        t_id = id_match.group(1) if id_match else f"TND-{sl_counter:04d}"

        val_match = re.search(r'(?:₹|rs\.?|inr|value|cost|amount)\s*[:\-]?\s*([\d,\.]+(?:\s*(?:cr|crore|lakh|lac))?)', text, re.IGNORECASE)
        val_str = val_match.group(0) if val_match else ""
        val_num, val_fmt, is_hidden = parse_currency_amount(val_str)

        date_match = re.search(r'(?:due|closing|last|end)\s*date\s*[:\-]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}(?:\s+\d{1,2}:\d{2})?)', text, re.IGNORECASE)
        due_date = date_match.group(1) if date_match else "Refer Notice"

        dept = urllib.parse.urlparse(source_url).netloc
        dist = detect_district(title + " " + text)

        records.append({
            "slNo": sl_counter,
            "systemTenderNo": t_id,
            "tenderId": t_id,
            "refNo": f"NIT/{t_id}",
            "department": dept,
            "district": dist,
            "estimatedValue": val_num,
            "estimatedValueFormatted": val_fmt,
            "isEstimatedValueHidden": is_hidden,
            "description": title,
            "detailedDescription": title,
            "startDate": "Refer Notice",
            "dueDate": due_date,
            "openDate": "Refer Notice",
            "emdAmount": None,
            "emdAmountFormatted": "As per NIT",
            "emdMode": "As per NIT",
            "tenderFee": 0.0,
            "tenderFeeFormatted": "As per NIT",
            "tenderFeeMode": "As per NIT",
            "processingFee": 0.0,
            "processingFeeFormatted": "As per NIT",
            "processingFeeMode": "As per NIT",
            "termsFiles": [],
            "officialFiles": [],
            "boqFiles": [],
            "hasDownloaded": False,
            "actionExtracted": False
        })
        sl_counter += 1

    return records

def export_scraped_to_excel(records, source_url, output_filepath):
    wb = openpyxl.Workbook()
    ws_main = wb.active
    ws_main.title = "Main Tenders"

    navy_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    gold_fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    green_font = Font(name="Segoe UI", size=9, bold=True, color="047857")
    gray_font = Font(name="Segoe UI", size=9, italic=True, color="64748B")

    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    headers = [
        "SL. No.", "System Tender No.", "Tender ID", "Reference Number", "Department Name",
        "Estimated Value (INR)", "Detailed Description", "Bid Submission Start Date",
        "Bid Submission Due Date", "Bid Open Date", "EMD Amount (INR)", "EMD Payment Mode",
        "Tender Fee (INR)", "Tender Fee Mode", "Processing Fee (INR)", "Processing Fee Mode",
        "District", "Attachments & BOQ Status"
    ]

    for col_idx, h in enumerate(headers, start=1):
        cell = ws_main.cell(row=1, column=col_idx, value=h)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws_main.row_dimensions[1].height = 28

    for row_idx, t in enumerate(records, start=2):
        status_str = "Files Downloaded" if t.get("hasDownloaded") else ("Action Extracted" if t.get("actionExtracted") else "Available for Action Download")
        row_vals = [
            t.get("slNo", row_idx - 1),
            t.get("systemTenderNo") or t.get("tenderId"),
            t.get("tenderId"),
            t.get("refNo", "N/A"),
            t.get("department", "Government Department"),
            t.get("estimatedValue") if not t.get("isEstimatedValueHidden") else "Hidden by portal",
            t.get("detailedDescription") or t.get("description", "Tender Notice"),
            t.get("startDate", "N/A"),
            t.get("dueDate", "Refer Notice"),
            t.get("openDate", "Refer Notice"),
            t.get("emdAmount") if t.get("emdAmount") else "As per NIT",
            t.get("emdMode", "As per NIT"),
            t.get("tenderFee") if t.get("tenderFee") else "As per NIT",
            t.get("tenderFeeMode", "As per NIT"),
            t.get("processingFee") if t.get("processingFee") else "As per NIT",
            t.get("processingFeeMode", "As per NIT"),
            t.get("district", "State-wide / Headquarter"),
            status_str
        ]

        ws_main.append(row_vals)
        for c_idx in range(1, len(row_vals) + 1):
            cell = ws_main.cell(row=row_idx, column=c_idx)
            cell.font = data_font
            cell.border = thin_border
            if c_idx == 6 and not t.get("isEstimatedValueHidden") and t.get("estimatedValue"):
                cell.number_format = '₹ #,##0.00'
                cell.font = green_font
                cell.alignment = Alignment(horizontal="right")
            if row_idx % 2 == 0:
                cell.fill = zebra_fill

    ws_main.auto_filter.ref = f"A1:R{len(records)+1}"
    col_widths = [8, 18, 14, 22, 28, 22, 45, 20, 20, 20, 18, 28, 16, 26, 16, 26, 18, 24]
    for i, w in enumerate(col_widths, start=1):
        ws_main.column_dimensions[get_column_letter(i)].width = w

    # Shortlist Sheet (sorted descending)
    ws_shortlist = wb.create_sheet(title="Shortlist")
    for col_idx, h in enumerate(headers, start=1):
        cell = ws_shortlist.cell(row=1, column=col_idx, value=h)
        cell.fill = gold_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws_shortlist.row_dimensions[1].height = 28

    sorted_records = sorted(records, key=lambda x: x.get("estimatedValue") if x.get("estimatedValue") is not None else -1, reverse=True)
    for row_idx, t in enumerate(sorted_records, start=2):
        status_str = "Files Downloaded" if t.get("hasDownloaded") else ("Action Extracted" if t.get("actionExtracted") else "Available for Action Download")
        row_vals = [
            row_idx - 1,
            t.get("systemTenderNo") or t.get("tenderId"),
            t.get("tenderId"),
            t.get("refNo", "N/A"),
            t.get("department", "Government Department"),
            t.get("estimatedValue") if not t.get("isEstimatedValueHidden") else "Hidden by portal",
            t.get("detailedDescription") or t.get("description", "Tender Notice"),
            t.get("startDate", "N/A"),
            t.get("dueDate", "Refer Notice"),
            t.get("openDate", "Refer Notice"),
            t.get("emdAmount") if t.get("emdAmount") else "As per NIT",
            t.get("emdMode", "As per NIT"),
            t.get("tenderFee") if t.get("tenderFee") else "As per NIT",
            t.get("tenderFeeMode", "As per NIT"),
            t.get("processingFee") if t.get("processingFee") else "As per NIT",
            t.get("processingFeeMode", "As per NIT"),
            t.get("district", "State-wide / Headquarter"),
            status_str
        ]
        ws_shortlist.append(row_vals)
        for c_idx in range(1, len(row_vals) + 1):
            cell = ws_shortlist.cell(row=row_idx, column=c_idx)
            cell.font = data_font
            cell.border = thin_border
            if c_idx == 6 and not t.get("isEstimatedValueHidden") and t.get("estimatedValue"):
                cell.number_format = '₹ #,##0.00'
                cell.font = green_font
                cell.alignment = Alignment(horizontal="right")
            if row_idx % 2 == 0:
                cell.fill = zebra_fill

    ws_shortlist.auto_filter.ref = f"A1:R{len(sorted_records)+1}"
    for i, w in enumerate(col_widths, start=1):
        ws_shortlist.column_dimensions[get_column_letter(i)].width = w

    # Summary Sheet
    ws_summary = wb.create_sheet(title="Summary")
    title_cell = ws_summary.cell(row=1, column=1, value="WEBSITE SCRAPED TENDER DATA SUMMARY")
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")

    bold_font = Font(name="Segoe UI", size=9, bold=True)
    ws_summary.cell(row=3, column=1, value="Source Website URL:").font = bold_font
    ws_summary.cell(row=3, column=2, value=source_url).font = data_font

    ws_summary.cell(row=4, column=1, value="Extraction Timestamp:").font = bold_font
    ws_summary.cell(row=4, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S")).font = data_font

    ws_summary.cell(row=5, column=1, value="Total Extracted Tenders:").font = bold_font
    ws_summary.cell(row=5, column=2, value=len(records)).font = bold_font

    visible_vals = [t["estimatedValue"] for t in records if t.get("estimatedValue") is not None]
    ws_summary.cell(row=6, column=1, value="Tenders with Visible Value:").font = bold_font
    ws_summary.cell(row=6, column=2, value=len(visible_vals)).font = data_font

    ws_summary.cell(row=7, column=1, value="Tenders with Hidden Value:").font = bold_font
    ws_summary.cell(row=7, column=2, value=len(records) - len(visible_vals)).font = data_font

    total_val = sum(visible_vals)
    ws_summary.cell(row=8, column=1, value="Total Estimated Value (INR):").font = bold_font
    tot_cell = ws_summary.cell(row=8, column=2, value=total_val)
    tot_cell.number_format = '₹ #,##0.00'
    tot_cell.font = green_font

    highest_val = max(visible_vals) if visible_vals else 0
    ws_summary.cell(row=9, column=1, value="Highest Estimated Value (INR):").font = bold_font
    high_cell = ws_summary.cell(row=9, column=2, value=highest_val)
    high_cell.number_format = '₹ #,##0.00'
    high_cell.font = green_font

    ws_summary.column_dimensions['A'].width = 30
    ws_summary.column_dimensions['B'].width = 50

    wb.save(output_filepath)
    return output_filepath

def scrape_any_url(url):
    """
    Universal scraping engine:
    1. Downloads content from given URL
    2. If Bihar eProc portal, fetches ALL live active tenders and merges action details
    3. Normalizes into complete requested schema
    4. Generates Excel file
    5. Returns dictionary with status, records, and Excel filename
    """
    if not url.startswith('http://') and not url.startswith('https://'):
        url = 'https://' + url

    domain = urllib.parse.urlparse(url).netloc.replace(':', '_').replace('.', '_')
    timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"Scraped_{domain}_{timestamp_str}.xlsx"
    filepath = os.path.join(DIRECTORY, filename)

    records = []

    # Special handling for EPSV2Web portals (e.g., Bihar eProc 2.0)
    if 'EPSV2Web' in url or 'eproc2.bihar.gov.in' in url:
        print(f"[Universal Scraper] Connecting live to Bihar eProcurement portal: {url}...")
        try:
            from fetch_eproc_tenders import fetch_live_tenders
            from build_unified_dataset import build_comprehensive_unified_dataset
            # 1. Fetch fresh active tenders from portal
            fetch_live_tenders()
            # 2. Unify with deep details and action cache
            records = build_comprehensive_unified_dataset()
            print(f"[Universal Scraper] Successfully extracted {len(records)} live tenders with full fields!")
        except Exception as e:
            print(f"[Universal Scraper] eProc sync warning: {e}")
            from build_unified_dataset import build_comprehensive_unified_dataset
            records = build_comprehensive_unified_dataset()

    # Generic Scraping for any other URL (HTML or JSON)
    if not records:
        content, content_type = fetch_url_content(url)

        if 'application/json' in content_type or (content.strip().startswith('{') or content.strip().startswith('[')):
            try:
                json_obj = json.loads(content)
                items = []
                if isinstance(json_obj, list):
                    items = json_obj
                elif isinstance(json_obj, dict):
                    for k, v in json_obj.items():
                        if isinstance(v, list) and len(v) > len(items):
                            items = v

                sl = 1
                for item in items:
                    if isinstance(item, dict):
                        t_id = item.get('id') or item.get('tenderId') or item.get('tender_id') or item.get('tenderNo') or f"TND-{sl:04d}"
                        ref = item.get('refNo') or item.get('referenceNumber') or item.get('nit') or f"REF/{t_id}"
                        dept = item.get('department') or item.get('org') or item.get('authority') or domain
                        desc = clean_text(item.get('description') or item.get('title') or item.get('name') or "Tender Item")
                        raw_val = item.get('estimatedValue') or item.get('amount') or item.get('pac') or item.get('value')
                        v_num, v_fmt, is_h = parse_currency_amount(str(raw_val))
                        raw_emd = item.get('emd') or item.get('emdAmount')
                        e_num, e_fmt, e_h = parse_currency_amount(str(raw_emd))
                        due = item.get('dueDate') or item.get('closingDate') or item.get('lastDate') or "Refer Notice"
                        dist = item.get('district') or detect_district(desc + " " + dept)

                        records.append({
                            "slNo": sl,
                            "systemTenderNo": t_id,
                            "tenderId": t_id,
                            "refNo": ref,
                            "department": dept,
                            "district": dist,
                            "estimatedValue": v_num,
                            "estimatedValueFormatted": v_fmt,
                            "isEstimatedValueHidden": is_h,
                            "description": desc,
                            "detailedDescription": desc,
                            "startDate": "Refer Notice",
                            "dueDate": str(due),
                            "openDate": "Refer Notice",
                            "emdAmount": e_num,
                            "emdAmountFormatted": e_fmt if not e_h else "As per NIT",
                            "emdMode": "As per NIT",
                            "tenderFee": 0.0,
                            "tenderFeeFormatted": "As per NIT",
                            "tenderFeeMode": "As per NIT",
                            "processingFee": 0.0,
                            "processingFeeFormatted": "As per NIT",
                            "processingFeeMode": "As per NIT",
                            "termsFiles": [],
                            "officialFiles": [],
                            "boqFiles": [],
                            "hasDownloaded": False,
                            "actionExtracted": False
                        })
                        sl += 1
            except Exception as e:
                pass

        if not records:
            soup = BeautifulSoup(content, 'html.parser')
            records = parse_html_tables(soup, url)
            if not records:
                records = parse_card_elements(soup, url)

    if not records:
        raise ValueError(f"No tabular or tender records could be automatically detected at {url}. Ensure the URL contains an HTML table, tender listing, or valid API endpoint.")

    # Export to Excel
    export_scraped_to_excel(records, url, filepath)

    return {
        "success": True,
        "sourceUrl": url,
        "totalTenders": len(records),
        "excelFilename": filename,
        "excelPath": filepath,
        "tenders": records
    }

if __name__ == "__main__":
    test_url = sys.argv[1] if len(sys.argv) > 1 else "https://eproc2.bihar.gov.in/EPSV2Web/openarea/tenderListingPage.action#latestTenders"
    print(f"Testing Universal Webhook Scraper on: {test_url}")
    try:
        res = scrape_any_url(test_url)
        print(f"Success! Scraped {res['totalTenders']} tenders. Saved to {res['excelFilename']}")
    except Exception as e:
        print(f"Extraction error: {e}")
