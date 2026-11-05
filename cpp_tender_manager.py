"""
CPP / Central Public Procurement Portal (eprocure.gov.in) Tender Manager
Dedicated module for extracting, deduplicating, storing, shortlisting,
downloading, and exporting Central Government e-Procurement tenders.

Zero interference with existing Bihar eProc2 data, files, or workflows.
"""

import os
import re
import sys
import json
import time
import uuid
import ssl
import urllib.request
import urllib.parse
from datetime import datetime
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set unverified SSL context for Indian Govt portal certificates (NIC/CCA)
try:
    ssl._create_default_https_context = ssl._create_unverified_context
except AttributeError:
    pass

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
CPP_VIEW_JSON = os.path.join(DIRECTORY, "cpp_tenders_view.json")
CPP_EXCEL_PATH = os.path.join(DIRECTORY, "CPP_Central_Gov_Tenders_Comprehensive.xlsx")
CPP_DOWNLOADS_DIR = os.path.join(DIRECTORY, "downloads_cpp")

# Ensure CPP downloads directory exists
os.makedirs(CPP_DOWNLOADS_DIR, exist_ok=True)

# In-memory store for active live portal crawl sessions
# session_id -> { cookie: str, form_fields: dict, captcha_image: str, created_at: float, current_url: str }
ACTIVE_CPP_SESSIONS = {}

def clean_text(text):
    if not text:
        return ""
    text = re.sub(r'[\r\n\t]+', ' ', str(text))
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()

def parse_currency_amount(val_str):
    if not val_str:
        return 0.0, "₹ 0.00", True
    
    clean = clean_text(val_str).lower()
    if clean in ['nil', 'zero', '0.00', '0', '-', 'na', 'n/a']:
        return 0.0, "₹ 0.00", False
    if any(k in clean for k in ['hidden', 'not disclosed', 'to be quoted', 'refer nit', 'as per tender']):
        return 0.0, "Hidden by portal", True
    
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
                return val, f"₹ {val:,.2f}", False
        except:
            pass

    return 0.0, "₹ 0.00", False

def compute_budget_slab(val: float) -> str:
    if val <= 0:
        return "Not Specified"
    elif val <= 2500000:
        return "0to25L"
    elif val <= 5000000:
        return "25Lto50L"
    elif val <= 20000000:
        return "50Lto2Cr"
    elif val <= 50000000:
        return "2Crto5Cr"
    else:
        return "Above5Cr"

def extract_state_district_strict(org_chain: str, work_desc: str = "") -> str:
    """
    Extracts State/District/Location strictly from Organisation Chain or Work Description.
    NEVER hallucinate or guess. If not explicitly identified, returns Department/Chain main level.
    """
    combined = f"{org_chain} {work_desc}"
    
    # Common Indian States and Union Territories
    states = [
        "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chhattisgarh", "Goa", "Gujarat",
        "Haryana", "Himachal Pradesh", "Jharkhand", "Karnataka", "Kerala", "Madhya Pradesh",
        "Maharashtra", "Manipur", "Meghalaya", "Mizoram", "Nagaland", "Odisha", "Punjab",
        "Rajasthan", "Sikkim", "Tamil Nadu", "Telangana", "Tripura", "Uttar Pradesh",
        "Uttarakhand", "West Bengal", "Delhi", "New Delhi", "Jammu and Kashmir", "Ladakh",
        "Chandigarh", "Puducherry"
    ]
    
    found_locations = []
    for s in states:
        if re.search(rf'\b{re.escape(s)}\b', combined, re.IGNORECASE):
            found_locations.append(s)
            
    # Major hubs or prominent centers
    prominent = [
        "Roorkee", "Numaligarh", "Trichy", "Tiruchirappalli", "Surat", "Raxaul", "Kolkata", "Patna",
        "Mumbai", "Bengaluru", "Bangalore", "Chennai", "Hyderabad", "Bhubaneswar", "Kanpur", "Lucknow",
        "Varanasi", "Nagpur", "Pune", "Guwahati", "Ahmedabad", "Jaipur", "Kochi", "Cochin"
    ]
    for p in prominent:
        if re.search(rf'\b{re.escape(p)}\b', combined, re.IGNORECASE):
            found_locations.append(p)
            
    if found_locations:
        # Deduplicate preserving order
        seen = set()
        unique = []
        for loc in found_locations:
            l_lower = loc.lower()
            if l_lower not in seen:
                seen.add(l_lower)
                unique.append(loc)
        return " / ".join(unique[:2])
    
    # Fallback to the primary Circle/Office from Org Chain (no hallucination)
    if org_chain:
        parts = [p.strip() for p in org_chain.split('||') if p.strip()]
        if len(parts) >= 2:
            return parts[1]
        elif len(parts) == 1:
            return parts[0]

    return "Central / Pan-India"

# =========================================================================
# 1. DATA PERSISTENCE & STRICT DUPLICATE PREVENTION
# =========================================================================

def load_cpp_tenders_store() -> dict:
    """Loads CPP tenders database from cpp_tenders_view.json."""
    if os.path.exists(CPP_VIEW_JSON):
        try:
            with open(CPP_VIEW_JSON, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict) and "tenders" in data:
                    return data
        except Exception as e:
            print(f"[CPP] Warning reading {CPP_VIEW_JSON}: {e}")
    
    return {
        "portal": "Central Public Procurement Portal (eprocure.gov.in)",
        "updated_at": datetime.now().isoformat(),
        "total_count": 0,
        "tenders": []
    }

def save_cpp_tenders_store(store: dict) -> None:
    """Saves CPP tenders store to cpp_tenders_view.json atomically."""
    store["updated_at"] = datetime.now().isoformat()
    store["total_count"] = len(store.get("tenders", []))
    
    tmp_path = CPP_VIEW_JSON + ".tmp"
    with open(tmp_path, 'w', encoding='utf-8') as f:
        json.dump(store, f, indent=2, ensure_ascii=False)
    
    if os.path.exists(CPP_VIEW_JSON):
        os.replace(tmp_path, CPP_VIEW_JSON)
    else:
        os.rename(tmp_path, CPP_VIEW_JSON)

def get_existing_canonical_keys() -> set:
    """Returns a set of all canonical keys (Tender IDs / normalized identifiers) in CPP store."""
    store = load_cpp_tenders_store()
    keys = set()
    for t in store.get("tenders", []):
        cid = str(t.get("tender_id", "")).strip().lower()
        if cid:
            keys.add(cid)
        ckey = str(t.get("canonical_key", "")).strip().lower()
        if ckey:
            keys.add(ckey)
        ref = str(t.get("tender_ref_no", "")).strip().lower()
        if ref:
            keys.add(ref)
    return keys

def normalize_tender_key(raw_id: str, title_ref_text: str = "") -> str:
    """Produces the canonical unique key for duplicate prevention."""
    clean_id = clean_text(raw_id).strip()
    if clean_id and clean_id.lower() not in ["na", "n/a", "nil", "none"]:
        return clean_id
    
    # Try finding tender ID pattern in title_ref_text
    m = re.search(r'\b\d{4}_[A-Z0-9]+_\d+(?:_\d+)?\b', title_ref_text, re.IGNORECASE)
    if m:
        return m.group(0).strip()
        
    m2 = re.search(r'\b(?:NIT|Tender|Ref)[^/]*[/\-_][A-Za-z0-9/\-_]+\b', title_ref_text, re.IGNORECASE)
    if m2:
        return m2.group(0).strip()

    # Fallback to normalized alphanumeric string
    norm = re.sub(r'[^a-zA-Z0-9]', '', title_ref_text).lower()
    return norm[:80] if norm else str(uuid.uuid4())

def save_cpp_tenders_batch(new_tenders: list, allow_update: bool = False) -> dict:
    """
    Saves a batch of tenders with STRICT DUPLICATE PREVENTION.
    If a tender already exists, it is SKIPPED unless allow_update=True.
    Returns stats: { 'added': int, 'skipped_duplicates': int, 'total': int }
    """
    store = load_cpp_tenders_store()
    existing_tenders = store.get("tenders", [])
    
    # Build fast lookup map
    lookup = {}
    for idx, t in enumerate(existing_tenders):
        cid = str(t.get("tender_id", "")).strip().lower()
        if cid: lookup[cid] = idx
        ckey = str(t.get("canonical_key", "")).strip().lower()
        if ckey: lookup[ckey] = idx

    added_count = 0
    skipped_count = 0

    for item in new_tenders:
        tender_id = str(item.get("tender_id", "")).strip()
        canon_key = str(item.get("canonical_key", "")).strip() or normalize_tender_key(tender_id, item.get("title", ""))
        item["canonical_key"] = canon_key

        key_to_check = tender_id.lower() if tender_id else canon_key.lower()

        if key_to_check in lookup:
            if allow_update:
                idx = lookup[key_to_check]
                # Preserve shortlist status
                item["is_shortlisted"] = existing_tenders[idx].get("is_shortlisted", False)
                existing_tenders[idx] = item
                added_count += 1
            else:
                skipped_count += 1
                continue
        else:
            # New tender
            lookup[key_to_check] = len(existing_tenders)
            existing_tenders.append(item)
            added_count += 1

    store["tenders"] = existing_tenders
    save_cpp_tenders_store(store)
    
    # Regenerate CPP Excel workbook in background
    try:
        generate_cpp_excel_workbook(existing_tenders)
    except Exception as e_xl:
        print(f"[CPP] Excel generation note: {e_xl}")

    return {
        "added": added_count,
        "skipped_duplicates": skipped_count,
        "total": len(existing_tenders)
    }

def toggle_cpp_shortlist(tender_id: str) -> dict:
    """Toggles the shortlist state of a specific CPP tender."""
    store = load_cpp_tenders_store()
    tenders = store.get("tenders", [])
    target = None
    tid_lower = tender_id.strip().lower()

    for t in tenders:
        if str(t.get("tender_id", "")).strip().lower() == tid_lower or str(t.get("canonical_key", "")).strip().lower() == tid_lower:
            curr = t.get("is_shortlisted", False)
            t["is_shortlisted"] = not curr
            target = t
            break

    if target:
        save_cpp_tenders_store(store)
        try:
            generate_cpp_excel_workbook(tenders)
        except Exception:
            pass
        return {"success": True, "tender_id": tender_id, "is_shortlisted": target["is_shortlisted"]}
    
    return {"success": False, "error": f"Tender ID '{tender_id}' not found in CPP database."}

# =========================================================================
# 2. HTML PARSING ENGINES (SEARCH RESULTS & DETAIL PAGES)
# =========================================================================

def parse_cpp_search_results_page(html_content: str, base_url: str = "https://eprocure.gov.in/eprocure/app"):
    """
    Parses the 7 columns from eprocure.gov.in search result page:
    1. S.No
    2. e-Published Date
    3. Bid Submission Closing Date
    4. Tender Opening Date
    5. Title and Ref.No./Tender ID (with detail link)
    6. Organisation Chain
    7. Tender Value in ₹
    Also extracts pagination links.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    results = []
    
    # Find table with id="table" or class containing "list_table"
    table = soup.find('table', id='table')
    if not table:
        for t in soup.find_all('table'):
            txt = t.get_text().lower()
            if 'bid submission' in txt or 'tender opening' in txt or 'organisation chain' in txt:
                table = t
                break
                
    if not table:
        return {"tenders": [], "pagination": {}}

    rows = table.find_all('tr')
    if not rows:
        return {"tenders": [], "pagination": {}}

    # Find header row
    header_idx = -1
    col_map = {
        's_no': -1,
        'pub_date': -1,
        'closing_date': -1,
        'opening_date': -1,
        'title_ref_id': -1,
        'org_chain': -1,
        'tender_value': -1
    }

    for idx, r in enumerate(rows[:5]):
        cells = [clean_text(c.get_text()).lower() for c in r.find_all(['th', 'td'])]
        if any('closing date' in c or 'tender id' in c or 'published date' in c for c in cells):
            header_idx = idx
            for c_idx, h in enumerate(cells):
                if any(k in h for k in ['s.no', 'sl.no', 's no', '#']):
                    col_map['s_no'] = c_idx
                elif 'published date' in h:
                    col_map['pub_date'] = c_idx
                elif 'closing date' in h or 'submission closing' in h:
                    col_map['closing_date'] = c_idx
                elif 'opening date' in h:
                    col_map['opening_date'] = c_idx
                elif 'title' in h or 'ref' in h or 'tender id' in h:
                    col_map['title_ref_id'] = c_idx
                elif 'organisation chain' in h or 'organisation' in h or 'chain' in h:
                    col_map['org_chain'] = c_idx
                elif 'tender value' in h or 'value' in h:
                    col_map['tender_value'] = c_idx
            break

    # If headers not explicitly matched, standard fallback to 7 column layout
    if col_map['title_ref_id'] == -1:
        col_map = {
            's_no': 0,
            'pub_date': 1,
            'closing_date': 2,
            'opening_date': 3,
            'title_ref_id': 4,
            'org_chain': 5,
            'tender_value': 6
        }
        data_rows = rows[1:] if len(rows) > 1 else rows
    else:
        data_rows = rows[header_idx + 1:]

    for r in data_rows:
        cells = r.find_all(['td', 'th'])
        if len(cells) < 4:
            continue

        def get_val(key):
            idx = col_map.get(key, -1)
            if 0 <= idx < len(cells):
                return clean_text(cells[idx].get_text())
            return ""

        raw_sno = get_val('s_no')
        raw_pub = get_val('pub_date')
        raw_close = get_val('closing_date')
        raw_open = get_val('opening_date')
        raw_title_ref = get_val('title_ref_id')
        raw_org = get_val('org_chain')
        raw_val = get_val('tender_value')

        if not raw_title_ref and not raw_org and not raw_close:
            continue

        # Extract link to detail page
        detail_url = ""
        title_cell_idx = col_map.get('title_ref_id', -1)
        if 0 <= title_cell_idx < len(cells):
            a_tag = cells[title_cell_idx].find('a')
            if a_tag and a_tag.get('href'):
                href = a_tag['href']
                detail_url = urllib.parse.urljoin(base_url, href)

        # Parse canonical Tender ID and Reference Number from the cell text
        tender_id = ""
        tender_ref = ""
        m_id = re.search(r'\b(\d{4}_[A-Z0-9]+_\d+(?:_\d+)?)\b', raw_title_ref)
        if m_id:
            tender_id = m_id.group(1).strip()
        else:
            m_id2 = re.search(r'Tender ID\s*[:\-]?\s*([A-Za-z0-9_\-]+)', raw_title_ref, re.I)
            if m_id2:
                tender_id = m_id2.group(1).strip()

        m_ref = re.search(r'(?:Ref\.?\s*No\.?|Tender Reference Number)\s*[:\-]?\s*([^\n\r,]+)', raw_title_ref, re.I)
        if m_ref:
            tender_ref = m_ref.group(1).strip()

        val_float, val_fmt, _ = parse_currency_amount(raw_val)
        canon_key = normalize_tender_key(tender_id, raw_title_ref)

        item = {
            "s_no": raw_sno,
            "canonical_key": canon_key,
            "tender_id": tender_id or canon_key,
            "tender_ref_no": tender_ref,
            "title": raw_title_ref,
            "work_description": raw_title_ref,
            "organisation_chain": raw_org,
            "department": raw_org.split('||')[0].strip() if '||' in raw_org else raw_org,
            "state_district": extract_state_district_strict(raw_org, raw_title_ref),
            "tender_value_inr": val_float,
            "tender_value_formatted": val_fmt,
            "budget_slab": compute_budget_slab(val_float),
            "contract_type": "Tender",
            "published_date": raw_pub,
            "bid_closing_date": raw_close,
            "bid_submission_end_date": raw_close,
            "bid_opening_date": raw_open,
            "detail_url": detail_url,
            "instruments": [],
            "documents": [],
            "zip_download_url": "",
            "is_shortlisted": False,
            "extracted_at": datetime.now().isoformat(),
            "source_portal": "eprocure.gov.in"
        }
        results.append(item)

    # Extract pagination links
    pagination = {
        "current_page": 1,
        "available_pages": [],
        "next_page_url": ""
    }
    
    # Tapestry pagination links: sp=AFrontEndAdvancedSearchResult%2Ctable&sp=2 etc.
    page_links = soup.find_all('a', href=re.compile(r'TablePages\.linkPage|page=FrontEndAdvancedSearchResult'))
    for pl in page_links:
        href = pl.get('href', '')
        txt = clean_text(pl.get_text())
        full_url = urllib.parse.urljoin(base_url, href)
        if txt.isdigit():
            pagination["available_pages"].append({"page": int(txt), "url": full_url})
        elif any(k in txt.lower() for k in ['next', '>', '>>']):
            if not pagination["next_page_url"]:
                pagination["next_page_url"] = full_url

    return {"tenders": results, "pagination": pagination}

def parse_cpp_tender_detail_page(html_or_text: str, base_url: str = "https://eprocure.gov.in/eprocure/app") -> dict:
    """
    Deep dives into FrontEndTenderDetails HTML/text and extracts all required fields:
    - Organisation Chain
    - Tender Reference Number
    - Tender ID
    - Payment Instruments table (S.No, Instrument Type)
    - Tender Fee Details (amount in ₹, fee payable to, fee payable at, exemption allowed)
    - EMD Fee Details (amount in ₹, exemption allowed, fee type, percentage, payable to, payable at)
    - Work Item Details (title, work description, tender value, contract type)
    - Critical Dates (published, bid opening, document download start/end, clarification start/end, submission start/end)
    - Tender Documents (NIT document name, description, size in KB, and download links)
    - 'Download as zip file' link
    """
    data = {}
    is_html = "<html" in html_or_text.lower() or "<table" in html_or_text.lower() or "<div" in html_or_text.lower()
    soup = BeautifulSoup(html_or_text, 'html.parser') if is_html else None
    text = soup.get_text("\n") if soup else str(html_or_text)

    # 1. Header / Basic Info
    m_org = re.search(r'Organisation Chain\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    if not m_org:
        m_org = re.search(r'Organisation Chain\s*[:\-]?\s*([^\n\r]+)', text, re.I)
    data['organisation_chain'] = clean_text(m_org.group(1)) if m_org else ''

    m_ref = re.search(r'Tender Reference Number\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    if not m_ref:
        m_ref = re.search(r'Tender Reference Number\s*[:\-]?\s*([^\n\r]+)', text, re.I)
    data['tender_ref_no'] = clean_text(m_ref.group(1)) if m_ref else ''

    m_id = re.search(r'Tender ID\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    if not m_id:
        m_id = re.search(r'Tender ID\s*[:\-]?\s*([^\n\r]+)', text, re.I)
    data['tender_id'] = clean_text(m_id.group(1)) if m_id else ''

    # 2. Payment Instruments Table
    instruments = []
    if soup:
        # Search for table with header "Instrument Type"
        for t in soup.find_all('table'):
            txt = t.get_text().lower()
            if 'instrument type' in txt:
                for r in t.find_all('tr'):
                    tds = r.find_all(['td', 'th'])
                    if len(tds) >= 2:
                        sno_txt = clean_text(tds[0].get_text())
                        type_txt = clean_text(tds[1].get_text())
                        if sno_txt.isdigit() and type_txt and 'instrument' not in type_txt.lower():
                            instruments.append({"s_no": sno_txt, "instrument_type": type_txt})
                if instruments:
                    break

    # Fallback to regex text parsing for instruments
    if not instruments:
        m_inst_block = re.search(r'Instrument Type\s*[\n\r]+(.*?)(?=Tender Fee|EMD Fee|Work Item Details|Critical Dates)', text, re.DOTALL | re.I)
        if m_inst_block:
            inst_text = m_inst_block.group(1)
            inst_lines = [l.strip() for l in inst_text.splitlines() if l.strip()]
            cur_sno = ""
            for line in inst_lines:
                if line.isdigit():
                    cur_sno = line
                elif cur_sno and line:
                    instruments.append({"s_no": cur_sno, "instrument_type": line})
                    cur_sno = ""
                else:
                    m_pair = re.match(r'^(\d+)[\.\s\-]+([A-Za-z0-9\-\s/]+)$', line)
                    if m_pair:
                        instruments.append({"s_no": m_pair.group(1).strip(), "instrument_type": m_pair.group(2).strip()})

    data['instruments'] = instruments

    # 3. Tender Fee Details
    m_fee = re.search(r'Tender Fee in ₹\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    raw_fee = m_fee.group(1).strip() if m_fee else '0.00'
    fee_float, fee_fmt, _ = parse_currency_amount(raw_fee)
    data['tender_fee_inr'] = fee_float
    data['tender_fee_formatted'] = fee_fmt

    m_fee_to = re.search(r'Fee Payable To\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['fee_payable_to'] = clean_text(m_fee_to.group(1)) if m_fee_to else 'Nil'

    m_fee_at = re.search(r'Fee Payable At\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['fee_payable_at'] = clean_text(m_fee_at.group(1)) if m_fee_at else 'Nil'

    m_fee_ex = re.search(r'Tender Fee Exemption Allowed\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['fee_exemption_allowed'] = clean_text(m_fee_ex.group(1)) if m_fee_ex else 'No'

    # 4. EMD Fee Details
    m_emd = re.search(r'EMD Amount in ₹\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    raw_emd = m_emd.group(1).strip() if m_emd else '0.00'
    emd_float, emd_fmt, _ = parse_currency_amount(raw_emd)
    data['emd_amount_inr'] = emd_float
    data['emd_amount_formatted'] = emd_fmt

    m_emd_ex = re.search(r'EMD Exemption Allowed\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['emd_exemption_allowed'] = clean_text(m_emd_ex.group(1)) if m_emd_ex else 'No'

    m_emd_type = re.search(r'EMD Fee Type\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['emd_fee_type'] = clean_text(m_emd_type.group(1)) if m_emd_type else 'fixed'

    m_emd_pct = re.search(r'EMD Percentage\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['emd_percentage'] = clean_text(m_emd_pct.group(1)) if m_emd_pct else 'NA'

    m_emd_to = re.search(r'EMD Payable To\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['emd_payable_to'] = clean_text(m_emd_to.group(1)) if m_emd_to else 'Nil'

    m_emd_at = re.search(r'EMD Payable At\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['emd_payable_at'] = clean_text(m_emd_at.group(1)) if m_emd_at else 'Nil'

    # 5. Work Item Details
    m_title = re.search(r'Work Item Details.*?Title\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.DOTALL | re.I)
    if not m_title:
        m_title = re.search(r'Title\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['title'] = clean_text(m_title.group(1)) if m_title else ''

    m_desc = re.search(r'Work Description\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['work_description'] = clean_text(m_desc.group(1)) if m_desc else data['title']

    m_val = re.search(r'Work Item Details.*?Tender Value in ₹\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.DOTALL | re.I)
    if not m_val:
        m_val = re.search(r'Tender Value in ₹\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    raw_val = m_val.group(1).strip() if m_val else '0.00'
    val_float, val_fmt, _ = parse_currency_amount(raw_val)
    data['tender_value_inr'] = val_float
    data['tender_value_formatted'] = val_fmt
    data['budget_slab'] = compute_budget_slab(val_float)

    m_ctype = re.search(r'Contract Type\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
    data['contract_type'] = clean_text(m_ctype.group(1)) if m_ctype else 'Tender'

    # 6. Critical Dates
    date_patterns = [
        ('published_date', r'Published Date'),
        ('bid_opening_date', r'Bid Opening Date'),
        ('doc_download_start_date', r'Document Download / Sale Start Date'),
        ('doc_download_end_date', r'Document Download / Sale End Date'),
        ('clarification_start_date', r'Clarification Start Date'),
        ('clarification_end_date', r'Clarification End Date'),
        ('bid_submission_start_date', r'Bid Submission Start Date'),
        ('bid_submission_end_date', r'Bid Submission End Date'),
    ]
    for key, pat in date_patterns:
        m_d = re.search(pat + r'\s*[:\-]?\s*[\n\r]+([^\n\r]+)', text, re.I)
        data[key] = clean_text(m_d.group(1)) if m_d else ''

    # 7. Tender Documents & Download Links
    documents = []
    if soup:
        # Search for documents table
        for t in soup.find_all('table'):
            txt = t.get_text().lower()
            if 'document name' in txt or 'nit document' in txt:
                for r in t.find_all('tr'):
                    tds = r.find_all(['td', 'th'])
                    if len(tds) >= 3:
                        sno_txt = clean_text(tds[0].get_text())
                        name_cell = tds[1]
                        desc_txt = clean_text(tds[2].get_text())
                        size_txt = clean_text(tds[3].get_text()) if len(tds) > 3 else ""
                        
                        a_link = name_cell.find('a')
                        d_name = clean_text(name_cell.get_text())
                        d_url = urllib.parse.urljoin(base_url, a_link['href']) if (a_link and a_link.get('href')) else ""
                        
                        if sno_txt.isdigit() and d_name and 'document name' not in d_name.lower():
                            safe_name = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', d_name)
                            documents.append({
                                "s_no": sno_txt,
                                "name": d_name,
                                "description": desc_txt or "NIT Document",
                                "size_kb": size_txt,
                                "download_url": d_url,
                                "local_path": "",
                                "downloaded": False
                            })
                if documents:
                    break

    # Fallback regex for documents
    if not documents:
        m_doc = re.search(r'NIT Document.*?(\d+)\s+([A-Za-z0-9_\-\.]+\.pdf)\s+([A-Za-z0-9_\-\.]+)\s+([\d\.]+)', text, re.DOTALL | re.I)
        if m_doc:
            documents.append({
                "s_no": m_doc.group(1).strip(),
                "name": m_doc.group(2).strip(),
                "description": m_doc.group(3).strip(),
                "size_kb": m_doc.group(4).strip(),
                "download_url": "",
                "local_path": "",
                "downloaded": False
            })

    data['documents'] = documents

    # 8. Download as zip file link
    zip_url = ""
    if soup:
        zip_a = soup.find('a', string=re.compile(r'Download as zip file', re.I))
        if not zip_a:
            zip_a = soup.find('a', href=re.compile(r'DirectLink.*zip|docDownoad.*zip', re.I))
        if zip_a and zip_a.get('href'):
            zip_url = urllib.parse.urljoin(base_url, zip_a['href'])
    data['zip_download_url'] = zip_url
    data['zip_local_path'] = ""

    # Location
    data['state_district'] = extract_state_district_strict(data['organisation_chain'], data.get('title', ''))
    data['department'] = data['organisation_chain'].split('||')[0].strip() if '||' in data['organisation_chain'] else data['organisation_chain']
    
    # Canonical key
    tid = data.get('tender_id', '')
    data['canonical_key'] = normalize_tender_key(tid, data.get('title', ''))
    if not data.get('tender_id'):
        data['tender_id'] = data['canonical_key']

    return data

def import_cpp_raw_html(html_content: str, base_url: str = "https://eprocure.gov.in/eprocure/app") -> dict:
    """
    Parses user-pasted HTML or uploaded file containing either:
    1. Search results table
    2. Single tender detail view
    Deduplicates against existing records and persists.
    """
    content_lower = html_content.lower()
    is_search_table = (
        ("bid submission closing date" in content_lower or "tender opening date" in content_lower or "e-published date" in content_lower)
        and "work item details" not in content_lower
        and "nit document" not in content_lower
    )
    
    if not is_search_table and ("work item details" in content_lower or "tender reference number" in content_lower or "nit document" in content_lower):
        detail_data = parse_cpp_tender_detail_page(html_content, base_url)
        detail_data["extracted_at"] = datetime.now().isoformat()
        detail_data["is_shortlisted"] = False
        res = save_cpp_tenders_batch([detail_data], allow_update=True)
        return {
            "type": "detail_page",
            "tender_id": detail_data.get("tender_id"),
            "added": res["added"],
            "skipped_duplicates": res["skipped_duplicates"],
            "total": res["total"]
        }
    else:
        parsed = parse_cpp_search_results_page(html_content, base_url)
        tenders = parsed.get("tenders", [])
        if not tenders:
            # Fallback attempt parsing as detail page
            detail_data = parse_cpp_tender_detail_page(html_content, base_url)
            if detail_data.get("tender_id"):
                res = save_cpp_tenders_batch([detail_data], allow_update=True)
                return {
                    "type": "detail_page",
                    "tender_id": detail_data.get("tender_id"),
                    "added": res["added"],
                    "skipped_duplicates": res["skipped_duplicates"],
                    "total": res["total"]
                }
            return {"type": "search_results", "tenders_found": 0, "added": 0, "skipped_duplicates": 0, "total": 0}

        res = save_cpp_tenders_batch(tenders, allow_update=False)
        return {
            "type": "search_results",
            "tenders_found": len(tenders),
            "added": res["added"],
            "skipped_duplicates": res["skipped_duplicates"],
            "total": res["total"],
            "pagination": parsed.get("pagination", {})
        }

# =========================================================================
# 3. LIVE SESSION & USER-CONTROLLED CAPTCHA WORKFLOW
# =========================================================================

def init_cpp_session(mode: str = "active_tenders") -> dict:
    """
    Starts a session with eprocure.gov.in.
    Fetches the live page, extracts JSESSIONID, form fields, and Base64 CAPTCHA image.
    CAPTCHA is returned to frontend for user to solve. NEVER auto-solved.
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8'
    }

    if mode == "advanced_search":
        target_url = 'https://eprocure.gov.in/eprocure/app?page=FrontEndAdvancedSearch&service=page'
    else:
        target_url = 'https://eprocure.gov.in/eprocure/app?page=FrontEndLatestActiveTenders&service=page'

    req = urllib.request.Request(target_url, headers=headers)
    ctx = ssl._create_unverified_context()

    try:
        with urllib.request.urlopen(req, context=ctx, timeout=25) as resp:
            raw_cookies = resp.headers.get_all('Set-Cookie') if hasattr(resp.headers, 'get_all') else [resp.headers.get('Set-Cookie')]
            cookie_parts = []
            for c in raw_cookies:
                if c:
                    cookie_parts.append(c.split(';')[0].strip())
            cookie_header = '; '.join(cookie_parts)

            html = resp.read().decode('utf-8', errors='ignore')
            soup = BeautifulSoup(html, 'html.parser')
            form = soup.find('form')
            if not form:
                return {"success": False, "error": "No form found on eprocure.gov.in page"}

            form_action = form.get('action', '/eprocure/app')
            if not form_action.startswith('http'):
                form_action = urllib.parse.urljoin('https://eprocure.gov.in', form_action)

            # Collect all hidden and default input fields
            form_fields = {}
            for inp in form.find_all('input'):
                n = inp.get('name')
                if n:
                    form_fields[n] = inp.get('value', '')

            # Extract live Base64 Captcha image
            captcha_image = ""
            for img in form.find_all('img'):
                src = img.get('src', '')
                if 'data:image' in src or 'captcha' in src.lower():
                    captcha_image = src
                    break

            session_id = str(uuid.uuid4())
            ACTIVE_CPP_SESSIONS[session_id] = {
                "cookie": cookie_header,
                "form_fields": form_fields,
                "form_action": form_action,
                "target_url": target_url,
                "mode": mode,
                "created_at": time.time(),
                "captcha_image": captcha_image
            }

            return {
                "success": True,
                "session_id": session_id,
                "captcha_image": captcha_image,
                "mode": mode,
                "message": "Session initialized. Please solve the CAPTCHA image and submit."
            }

    except Exception as e:
        return {"success": False, "error": f"Failed to connect to eprocure.gov.in: {str(e)}"}

def submit_cpp_search_and_crawl(session_id: str, captcha_text: str, tender_id: str = "", tender_title: str = "", max_detail_fetch: int = 15) -> dict:
    """
    Submits user's manually entered CAPTCHA to eprocure.gov.in.
    If CAPTCHA is invalid, returns new CAPTCHA for retry.
    If valid, crawls search results and deep-dives into tender details politely.
    """
    session = ACTIVE_CPP_SESSIONS.get(session_id)
    if not session:
        return {"success": False, "error": "Session expired or invalid. Please refresh CAPTCHA."}

    cookie = session["cookie"]
    form_fields = session["form_fields"].copy()
    form_action = session["form_action"]

    # Fill user inputs
    form_fields['captchaText'] = captcha_text.strip()
    if tender_id:
        form_fields['TenderId'] = tender_id.strip()
    if tender_title:
        form_fields['TenderTitle'] = tender_title.strip()
    form_fields['Submit'] = 'Search'

    post_data = urllib.parse.urlencode(form_fields).encode('utf-8')
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': cookie,
        'Referer': session["target_url"]
    }

    ctx = ssl._create_unverified_context()
    req = urllib.request.Request(form_action, data=post_data, headers=headers)

    try:
        with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
            # Update cookie if modified
            new_cookies = resp.headers.get_all('Set-Cookie') if hasattr(resp.headers, 'get_all') else [resp.headers.get('Set-Cookie')]
            for c in new_cookies:
                if c and 'JSESSIONID' in c:
                    cookie = c.split(';')[0].strip()
                    session["cookie"] = cookie

            html = resp.read().decode('utf-8', errors='ignore')

            # Check if captcha failed
            if "Invalid Captcha" in html or "Please Enter Correct Captcha" in html:
                # Extract new captcha image
                soup = BeautifulSoup(html, 'html.parser')
                form = soup.find('form')
                new_captcha = ""
                if form:
                    # Refresh form hidden fields
                    for inp in form.find_all('input'):
                        n = inp.get('name')
                        if n:
                            session["form_fields"][n] = inp.get('value', '')
                    for img in form.find_all('img'):
                        src = img.get('src', '')
                        if 'data:image' in src or 'captcha' in src.lower():
                            new_captcha = src
                            break

                return {
                    "success": False,
                    "captcha_error": True,
                    "error": "Invalid Captcha! Please enter the fresh letters shown in the image.",
                    "captcha_image": new_captcha or session.get("captcha_image")
                }

            # Parse search results
            parsed_res = parse_cpp_search_results_page(html, base_url='https://eprocure.gov.in/eprocure/app')
            tenders = parsed_res.get("tenders", [])
            pagination = parsed_res.get("pagination", {})

            # Deep-dive each tender detail if URL present
            existing_keys = get_existing_canonical_keys()
            processed_tenders = []
            duplicates_skipped = 0

            for t in tenders:
                canon = t.get("canonical_key", "")
                tid = t.get("tender_id", "")
                if (tid and tid.lower() in existing_keys) or (canon and canon.lower() in existing_keys):
                    duplicates_skipped += 1
                    continue

                detail_url = t.get("detail_url")
                if detail_url and len(processed_tenders) < max_detail_fetch:
                    try:
                        time.sleep(1.2)  # Polite pacing delay
                        d_req = urllib.request.Request(detail_url, headers={'User-Agent': headers['User-Agent'], 'Cookie': cookie})
                        with urllib.request.urlopen(d_req, context=ctx, timeout=20) as d_resp:
                            d_html = d_resp.read().decode('utf-8', errors='ignore')
                            details = parse_cpp_tender_detail_page(d_html, base_url='https://eprocure.gov.in/eprocure/app')
                            
                            # Merge details
                            t.update(details)
                    except Exception as e_det:
                        print(f"[CPP] Note: Detail page fetch for {tid} skipped: {e_det}")

                processed_tenders.append(t)

            # Save batch with duplicate rejection
            save_res = save_cpp_tenders_batch(processed_tenders, allow_update=False)

            # Store last pagination in session for traversal
            session["last_pagination"] = pagination
            session["last_results_html"] = html

            return {
                "success": True,
                "tenders_found": len(tenders),
                "added": save_res["added"],
                "skipped_duplicates": duplicates_skipped + save_res["skipped_duplicates"],
                "total_in_system": save_res["total"],
                "pagination": pagination,
                "message": f"Successfully imported {save_res['added']} new tenders. {duplicates_skipped + save_res['skipped_duplicates']} duplicates skipped."
            }

    except Exception as e:
        return {"success": False, "error": f"Error during search & crawl: {str(e)}"}

def crawl_next_cpp_page(session_id: str, page_url: str) -> dict:
    """Traverses to the next page of results using existing session."""
    session = ACTIVE_CPP_SESSIONS.get(session_id)
    if not session:
        return {"success": False, "error": "Crawl session expired. Please start a new search."}

    cookie = session["cookie"]
    ctx = ssl._create_unverified_context()
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Cookie': cookie,
        'Referer': session["target_url"]
    }

    try:
        time.sleep(1.5)  # Courteous pacing
        req = urllib.request.Request(page_url, headers=headers)
        with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
            parsed_res = parse_cpp_search_results_page(html, base_url='https://eprocure.gov.in/eprocure/app')
            tenders = parsed_res.get("tenders", [])
            pagination = parsed_res.get("pagination", {})

            existing_keys = get_existing_canonical_keys()
            processed = []
            duplicates_skipped = 0

            for t in tenders:
                canon = t.get("canonical_key", "")
                tid = t.get("tender_id", "")
                if (tid and tid.lower() in existing_keys) or (canon and canon.lower() in existing_keys):
                    duplicates_skipped += 1
                    continue

                detail_url = t.get("detail_url")
                if detail_url and len(processed) < 15:
                    try:
                        time.sleep(1.2)
                        d_req = urllib.request.Request(detail_url, headers={'User-Agent': headers['User-Agent'], 'Cookie': cookie})
                        with urllib.request.urlopen(d_req, context=ctx, timeout=20) as d_resp:
                            d_html = d_resp.read().decode('utf-8', errors='ignore')
                            details = parse_cpp_tender_detail_page(d_html, base_url='https://eprocure.gov.in/eprocure/app')
                            t.update(details)
                    except Exception:
                        pass

                processed.append(t)

            save_res = save_cpp_tenders_batch(processed, allow_update=False)
            session["last_pagination"] = pagination

            return {
                "success": True,
                "tenders_found": len(tenders),
                "added": save_res["added"],
                "skipped_duplicates": duplicates_skipped + save_res["skipped_duplicates"],
                "total_in_system": save_res["total"],
                "pagination": pagination
            }

    except Exception as e:
        return {"success": False, "error": f"Failed to fetch page: {str(e)}"}

# =========================================================================
# 4. DOCUMENT DOWNLOADER
# =========================================================================

def download_cpp_document(tender_id: str, doc_type: str = "nit", doc_index: int = 0, session_id: str = "") -> dict:
    """
    Downloads NIT Document or Zip File to downloads_cpp/tender_<tenderId>/<safe_filename>.
    Updates record in cpp_tenders_view.json.
    """
    store = load_cpp_tenders_store()
    tenders = store.get("tenders", [])
    target = None
    tid_lower = tender_id.strip().lower()

    for t in tenders:
        if str(t.get("tender_id", "")).strip().lower() == tid_lower or str(t.get("canonical_key", "")).strip().lower() == tid_lower:
            target = t
            break

    if not target:
        return {"success": False, "error": f"Tender ID '{tender_id}' not found"}

    clean_tid = re.sub(r'[^a-zA-Z0-9_\-]', '_', target.get("tender_id", tender_id))
    tender_folder = os.path.join(CPP_DOWNLOADS_DIR, f"tender_{clean_tid}")
    os.makedirs(tender_folder, exist_ok=True)

    session = ACTIVE_CPP_SESSIONS.get(session_id)
    cookie = session["cookie"] if session else ""

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    }
    if cookie:
        headers['Cookie'] = cookie

    ctx = ssl._create_unverified_context()

    if doc_type == "zip":
        url = target.get("zip_download_url")
        if not url:
            return {"success": False, "error": "No ZIP download link available for this tender"}
        filename = f"{clean_tid}_all_documents.zip"
        dest_path = os.path.join(tender_folder, filename)
    else:
        docs = target.get("documents", [])
        if not docs or doc_index >= len(docs):
            return {"success": False, "error": "Document not found in tender records"}
        doc_item = docs[doc_index]
        url = doc_item.get("download_url")
        filename = doc_item.get("name") or f"Tendernotice_{clean_tid}.pdf"
        safe_fn = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', filename)
        dest_path = os.path.join(tender_folder, safe_fn)

    if not url:
        return {"success": False, "error": "Download URL missing from tender record"}

    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, context=ctx, timeout=40) as resp:
            content = resp.read()
            with open(dest_path, 'wb') as f:
                f.write(content)

        # Update record
        rel_path = os.path.relpath(dest_path, DIRECTORY).replace('\\', '/')
        if doc_type == "zip":
            target["zip_local_path"] = rel_path
        else:
            docs[doc_index]["local_path"] = rel_path
            docs[doc_index]["downloaded"] = True

        save_cpp_tenders_store(store)
        return {
            "success": True,
            "filename": filename,
            "local_path": rel_path,
            "size_bytes": len(content),
            "message": f"Successfully downloaded to {rel_path}"
        }

    except Exception as e:
        return {"success": False, "error": f"Download failed: {str(e)}"}

# =========================================================================
# 5. DEDICATED EXCEL EXPORT WORKBOOK GENERATOR
# =========================================================================

def generate_cpp_excel_workbook(tenders: list = None) -> str:
    """
    Generates CPP_Central_Gov_Tenders_Comprehensive.xlsx with 4 styled sheets:
    1. CPP Main Tenders
    2. Payment Instruments
    3. Tender Documents
    4. Shortlisted Tenders
    """
    if tenders is None:
        store = load_cpp_tenders_store()
        tenders = store.get("tenders", [])

    wb = openpyxl.Workbook()
    
    # Styles
    navy_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    gold_fill = PatternFill(start_color="D69E2E", end_color="D69E2E", fill_type="solid")
    stripe_fill = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    shortlist_fill = PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid")
    
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    bold_data_font = Font(name="Segoe UI", size=10, bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E0'),
        right=Side(style='thin', color='CBD5E0'),
        top=Side(style='thin', color='CBD5E0'),
        bottom=Side(style='thin', color='CBD5E0')
    )

    # ----------------- Sheet 1: CPP Main Tenders -----------------
    ws_main = wb.active
    ws_main.title = "CPP Main Tenders"
    ws_main.views.sheetView[0].showGridLines = True

    main_headers = [
        "S.No", "Tender ID", "Tender Reference No", "Title / Work Description",
        "Organisation Chain", "Department", "State / Location", "Tender Value (₹)",
        "Value Slab", "EMD Amount (₹)", "Tender Fee (₹)", "Contract Type",
        "Published Date", "Closing Date", "Opening Date", "Shortlisted?"
    ]
    
    ws_main.append(main_headers)
    for col_num in range(1, len(main_headers) + 1):
        cell = ws_main.cell(row=1, column=col_num)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for idx, t in enumerate(tenders, start=1):
        row = [
            idx,
            t.get("tender_id", ""),
            t.get("tender_ref_no", ""),
            t.get("title", ""),
            t.get("organisation_chain", ""),
            t.get("department", ""),
            t.get("state_district", ""),
            float(t.get("tender_value_inr") or 0.0),
            t.get("budget_slab", ""),
            float(t.get("emd_amount_inr") or 0.0),
            float(t.get("tender_fee_inr") or 0.0),
            t.get("contract_type", "Tender"),
            t.get("published_date", ""),
            t.get("bid_closing_date", t.get("bid_submission_end_date", "")),
            t.get("bid_opening_date", ""),
            "★ SHORTLISTED" if t.get("is_shortlisted") else "No"
        ]
        ws_main.append(row)
        curr_row = ws_main.max_row
        is_sl = t.get("is_shortlisted")

        for c_idx in range(1, len(row) + 1):
            cell = ws_main.cell(row=curr_row, column=c_idx)
            cell.font = data_font
            cell.border = thin_border
            if is_sl:
                cell.fill = shortlist_fill
            elif idx % 2 == 0:
                cell.fill = stripe_fill

            # Format Currency
            if c_idx in [8, 10, 11]:
                cell.number_format = '"₹ "#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            elif c_idx in [1, 12, 13, 14, 15, 16]:
                cell.alignment = Alignment(horizontal="center")

    # ----------------- Sheet 2: Payment Instruments -----------------
    ws_inst = wb.create_sheet(title="Payment Instruments")
    ws_inst.views.sheetView[0].showGridLines = True
    inst_headers = ["Tender ID", "Tender Ref No", "S.No", "Instrument Type", "EMD Amount (₹)", "EMD Payable To", "EMD Payable At"]
    ws_inst.append(inst_headers)
    for col_num in range(1, len(inst_headers) + 1):
        cell = ws_inst.cell(row=1, column=col_num)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    inst_row_cnt = 0
    for t in tenders:
        tid = t.get("tender_id", "")
        ref = t.get("tender_ref_no", "")
        emd = float(t.get("emd_amount_inr") or 0.0)
        emd_to = t.get("emd_payable_to", "")
        emd_at = t.get("emd_payable_at", "")
        
        for inst in t.get("instruments", []):
            inst_row_cnt += 1
            ws_inst.append([
                tid, ref, inst.get("s_no", ""), inst.get("instrument_type", ""), emd, emd_to, emd_at
            ])
            r_num = ws_inst.max_row
            for c_idx in range(1, len(inst_headers) + 1):
                cell = ws_inst.cell(row=r_num, column=c_idx)
                cell.font = data_font
                cell.border = thin_border
                if c_idx == 5:
                    cell.number_format = '"₹ "#,##0.00'
                    cell.alignment = Alignment(horizontal="right")

    # ----------------- Sheet 3: Tender Documents -----------------
    ws_doc = wb.create_sheet(title="Tender Documents")
    ws_doc.views.sheetView[0].showGridLines = True
    doc_headers = ["Tender ID", "Document Name", "Description", "Size (KB)", "Download Link", "Local Path", "Downloaded?"]
    ws_doc.append(doc_headers)
    for col_num in range(1, len(doc_headers) + 1):
        cell = ws_doc.cell(row=1, column=col_num)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for t in tenders:
        tid = t.get("tender_id", "")
        for d in t.get("documents", []):
            ws_doc.append([
                tid,
                d.get("name", ""),
                d.get("description", ""),
                d.get("size_kb", ""),
                d.get("download_url", ""),
                d.get("local_path", ""),
                "YES" if d.get("downloaded") else "NO"
            ])
            r_num = ws_doc.max_row
            for c_idx in range(1, len(doc_headers) + 1):
                cell = ws_doc.cell(row=r_num, column=c_idx)
                cell.font = data_font
                cell.border = thin_border

    # ----------------- Sheet 4: Shortlisted Tenders -----------------
    ws_sl = wb.create_sheet(title="Shortlisted Tenders")
    ws_sl.views.sheetView[0].showGridLines = True
    ws_sl.append(main_headers)
    for col_num in range(1, len(main_headers) + 1):
        cell = ws_sl.cell(row=1, column=col_num)
        cell.fill = gold_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sl_tenders = [t for t in tenders if t.get("is_shortlisted")]
    for idx, t in enumerate(sl_tenders, start=1):
        row = [
            idx,
            t.get("tender_id", ""),
            t.get("tender_ref_no", ""),
            t.get("title", ""),
            t.get("organisation_chain", ""),
            t.get("department", ""),
            t.get("state_district", ""),
            float(t.get("tender_value_inr") or 0.0),
            t.get("budget_slab", ""),
            float(t.get("emd_amount_inr") or 0.0),
            float(t.get("tender_fee_inr") or 0.0),
            t.get("contract_type", "Tender"),
            t.get("published_date", ""),
            t.get("bid_closing_date", t.get("bid_submission_end_date", "")),
            t.get("bid_opening_date", ""),
            "★ SHORTLISTED"
        ]
        ws_sl.append(row)
        curr_row = ws_sl.max_row
        for c_idx in range(1, len(row) + 1):
            cell = ws_sl.cell(row=curr_row, column=c_idx)
            cell.font = bold_data_font
            cell.border = thin_border
            cell.fill = shortlist_fill
            if c_idx in [8, 10, 11]:
                cell.number_format = '"₹ "#,##0.00'
                cell.alignment = Alignment(horizontal="right")

    # Auto-adjust column widths across all sheets
    for ws in [ws_main, ws_inst, ws_doc, ws_sl]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = min(len(val_str), 50)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(CPP_EXCEL_PATH)
    return CPP_EXCEL_PATH

# Self-test block
if __name__ == "__main__":
    print("[CPP Manager] Module loaded successfully.")
