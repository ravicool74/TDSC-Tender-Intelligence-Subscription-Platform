"""
Batch Action Details & Attachment Enricher for Bihar e-Procurement Portal
Enriches ALL tenders from https://eproc2.bihar.gov.in/EPSV2Web/openarea/tenderListingPage.action#latestTenders
with full Action Deep Dive details:
- System Tender No.
- Reference No.
- Estimated Value in INR (reveals PAC even when hidden in list)
- EMD Amount & Payment Mode
- Tender Fee & Payment Mode
- Processing Fee & Payment Mode
- Bid Submission Start Date, Due Date, Open Date
- Detailed Description
- District
- Terms & Conditions files metadata
- Official NIT attachments metadata
- BOQ (Excel / PDF) files metadata
"""

import os
import sys
import json
import time
import re
import urllib.request
import concurrent.futures
from datetime import datetime
from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
from action_details_manager import load_cache, save_cache, CACHE_FILE
from build_unified_dataset import build_comprehensive_unified_dataset

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
TENDERS_DATA_PATH = os.path.join(DIRECTORY, "tenders_data.json")

BIHAR_DISTRICT_ALIASES = {
    # East Champaran
    "motihari": "East Champaran", "sugaon": "East Champaran", "sugauli": "East Champaran",
    "raxaul": "East Champaran", "chakia": "East Champaran", "dhaka": "East Champaran",
    "areraj": "East Champaran", "pakridayal": "East Champaran", "mehsi": "East Champaran",
    "kesaria": "East Champaran", "kalyanpur": "East Champaran",
    # West Champaran
    "bettiah": "West Champaran", "bagaha": "West Champaran", "narkatiaganj": "West Champaran",
    "ramnagar": "West Champaran", "lauriya": "West Champaran", "chanpatia": "West Champaran",
    # Nalanda
    "biharsharif": "Nalanda", "bihar sharif": "Nalanda", "rajgir": "Nalanda", "hilsa": "Nalanda", "islampur": "Nalanda",
    # Bhojpur
    "ara": "Bhojpur", "arrah": "Bhojpur", "jagdishpur": "Bhojpur", "piro": "Bhojpur",
    # Saran
    "chhapra": "Saran", "chapra": "Saran", "marhaura": "Saran", "sonepur": "Saran",
    # Rohtas
    "sasaram": "Rohtas", "dehri": "Rohtas", "bikramganj": "Rohtas", "nokha": "Rohtas",
    # Kaimur
    "bhabua": "Kaimur", "mohania": "Kaimur", "kudra": "Kaimur",
    # Vaishali
    "hajipur": "Vaishali", "mahua": "Vaishali", "lalganj": "Vaishali",
    # Purnia
    "purnea": "Purnia", "banmankhi": "Purnia", "kasba": "Purnia", "dhamdaha": "Purnia",
    # Munger
    "monghyr": "Munger", "jamalpur": "Munger", "tarapur": "Munger",
    # Patna
    "danapur": "Patna", "barh": "Patna", "mokama": "Patna", "fatuha": "Patna",
    "masaurhi": "Patna", "paliganj": "Patna", "bikram": "Patna", "phulwari": "Patna",
    # Gaya
    "bodhgaya": "Gaya", "sherghati": "Gaya", "tekari": "Gaya", "barachatti": "Gaya",
    # Samastipur
    "rosera": "Samastipur", "dalsinghsarai": "Samastipur", "patori": "Samastipur",
    # Madhubani
    "jhanjharpur": "Madhubani", "benipatti": "Madhubani", "phulparas": "Madhubani",
    # Darbhanga
    "benipur": "Darbhanga", "baheri": "Darbhanga", "biraul": "Darbhanga",
    # Bhagalpur
    "kahalgaon": "Bhagalpur", "sultanganj": "Bhagalpur", "naugachia": "Bhagalpur",
    # Siwan
    "maharajganj": "Siwan", "mairwa": "Siwan",
    # Gopalganj
    "hathua": "Gopalganj", "barauli": "Gopalganj",
    # Araria
    "forbesganj": "Araria", "jogbani": "Araria",
    # Katihar
    "manihari": "Katihar", "barsoi": "Katihar",
    # Begusarai
    "teghra": "Begusarai", "bakhri": "Begusarai", "balia": "Begusarai",
    # Muzaffarpur
    "motipur": "Muzaffarpur", "kanti": "Muzaffarpur", "marwan": "Muzaffarpur", "sakra": "Muzaffarpur",
    # Sitamarhi
    "pupri": "Sitamarhi", "belsand": "Sitamarhi", "dumra": "Sitamarhi"
}

STANDARD_DISTRICTS = [
    "Araria", "Arwal", "Aurangabad", "Banka", "Begusarai", "Bhagalpur", "Bhojpur", "Buxar",
    "Darbhanga", "East Champaran", "Gaya", "Gopalganj", "Jamui", "Jehanabad", "Kaimur",
    "Katihar", "Khagaria", "Kishanganj", "Lakhisarai", "Madhepura", "Madhubani", "Munger",
    "Muzaffarpur", "Nalanda", "Nawada", "Patna", "Purnia", "Rohtas", "Saharsa",
    "Samastipur", "Saran", "Sheikhpura", "Sheohar", "Sitamarhi", "Siwan", "Supaul",
    "Vaishali", "West Champaran"
]

def detect_district_enhanced(text, dept=""):
    combined = f"{text} {dept}".lower()
    for alias, standard in BIHAR_DISTRICT_ALIASES.items():
        if re.search(rf'\b{re.escape(alias)}\b', combined):
            return standard
    for d in STANDARD_DISTRICTS:
        if re.search(rf'\b{re.escape(d.lower())}\b', combined):
            return d
    return "State-wide / Headquarter"

def format_ts_date(ts):
    if not ts: return "N/A"
    try:
        ts = int(ts)
        if ts <= 0: return "N/A"
        return datetime.fromtimestamp(ts / 1000.0).strftime("%d-%m-%Y %I:%M %p")
    except:
        return "N/A"

def fetch_single_action_preview(tender_id, opener, auth_token):
    """Fetches full Action button preview JSON from the portal."""
    url = f"{BASE_URL}/rest/tender/previewTenderByTenderId?tenderId={tender_id}"
    api_headers = {
        'User-Agent': HEADERS['User-Agent'],
        'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
        'Content-Type': 'application/json',
        'Authorization': auth_token
    }
    req = urllib.request.Request(url, data=b"", headers=api_headers)
    with opener.open(req, timeout=18) as resp:
        raw = resp.read().decode('utf-8', errors='replace')
        data = json.loads(raw)

    sys_tender = data.get('orgtenderid') or data.get('tenderid') or tender_id
    ref_no = data.get('tenderrefno') or "N/A"
    desc = (data.get('description') or "").strip()
    
    # PAC
    pac_amt = data.get('pacamt')
    pac_val = None
    is_hidden = True
    pac_formatted = "Hidden by portal"
    if pac_amt is not None:
        try:
            pac_val = float(pac_amt)
            if pac_val > 0:
                is_hidden = False
                pac_formatted = f"₹ {pac_val:,.2f}"
        except:
            pass

    # Dates
    start_date = "N/A"
    due_date = "Refer Notice"
    open_date = "Refer Notice"
    doc_sub_date = "N/A"

    # Payments
    emd_amount = 0.0
    emd_mode = "Internet Payment Gateway(IPG),Challan|Bank Guarantee"
    tender_fee = 0.0
    tender_fee_mode = "Internet Payment Gateway(IPG),Challan|"
    processing_fee = 0.0
    processing_fee_mode = "Internet Payment Gateway(IPG),Challan|"

    terms_files = []
    official_files = []
    boq_files = []

    for tmpl in data.get('templates', []):
        short = tmpl.get('shortName', '')
        grp_id = tmpl.get('templategroupId')
        fdict = {f.get('code'): f.get('value') for f in tmpl.get('templateFieldList', []) if f.get('code')}

        if short == 'rfqdate':
            start_date = format_ts_date(fdict.get('bid_start_date'))
            due_date = format_ts_date(fdict.get('bid_end_date'))
            open_date = format_ts_date(fdict.get('bid_open_date'))
            doc_sub_date = format_ts_date(fdict.get('doc_sub_date'))

        elif short == 'rfqpayment':
            ptype = str(fdict.get('payment_type') or '').lower()
            amt_str = fdict.get('amount') or "0"
            mode = fdict.get('payment_mode') or "Internet Payment Gateway(IPG),Challan|"
            amt_val = 0.0
            try: amt_val = float(amt_str)
            except: pass

            if 'emd' in ptype:
                emd_amount = amt_val
                emd_mode = mode
            elif 'processing' in ptype:
                processing_fee = amt_val
                processing_fee_mode = mode
            elif 'fee' in ptype:
                tender_fee = amt_val
                tender_fee_mode = mode

        elif short == 'br_rfq_technical':
            unf2 = fdict.get('unf2') or "Terms and Conditions"
            unf3 = fdict.get('unf3') or ""
            if unf3:
                fname = unf3.split('|')[-1] if '|' in unf3 else "Terms_Conditions.pdf"
                terms_files.append({
                    "title": unf2,
                    "attachFile": unf3,
                    "templategroupId": grp_id,
                    "fileName": fname
                })

        elif short == 'rfqattachment':
            lbl = fdict.get('label') or "Attachment"
            fname = fdict.get('file_name') or "Document.pdf"
            att = fdict.get('attach_file') or ""
            if att:
                entry = {
                    "label": lbl,
                    "fileName": fname,
                    "attachFile": att,
                    "templategroupId": grp_id
                }
                if 'boq' in lbl.lower() or 'boq' in fname.lower():
                    boq_files.append(entry)
                else:
                    official_files.append(entry)

        elif short in ['br_rfq_item_rc', 'br_rfq_item', 'auction_rfq_item']:
            item_name = fdict.get('item_name') or fdict.get('item_code') or "BOQ Item"
            unf3 = fdict.get('unf3') or ""
            if unf3:
                fname = unf3.split('|')[-1] if '|' in unf3 else "BOQ.pdf"
                boq_files.append({
                    "label": f"BOQ ({item_name})",
                    "fileName": fname,
                    "attachFile": unf3,
                    "templategroupId": grp_id
                })

    folder_name = f"tender_{sys_tender}_{tender_id}"
    folder_path = os.path.join(DIRECTORY, "downloads", folder_name)

    # Attach webUrls if files already exist locally
    for f in terms_files:
        p = os.path.join(folder_path, "terms_and_conditions", re.sub(r'[\\/*?:"<>|]', '_', f["fileName"]))
        if os.path.exists(p):
            f["localPath"] = p
            f["fileSize"] = os.path.getsize(p)
            f["webUrl"] = f"/downloads/{folder_name}/terms_and_conditions/{urllib.parse.quote(os.path.basename(p))}"

    for f in official_files:
        p = os.path.join(folder_path, "attachments", re.sub(r'[\\/*?:"<>|]', '_', f["fileName"]))
        if os.path.exists(p):
            f["localPath"] = p
            f["fileSize"] = os.path.getsize(p)
            f["webUrl"] = f"/downloads/{folder_name}/attachments/{urllib.parse.quote(os.path.basename(p))}"

    for f in boq_files:
        p = os.path.join(folder_path, "boq", re.sub(r'[\\/*?:"<>|]', '_', f["fileName"]))
        if os.path.exists(p):
            f["localPath"] = p
            f["fileSize"] = os.path.getsize(p)
            f["webUrl"] = f"/downloads/{folder_name}/boq/{urllib.parse.quote(os.path.basename(p))}"

    return {
        "tenderId": tender_id,
        "systemTenderNo": sys_tender,
        "tenderRefNo": ref_no,
        "estimatedValue": pac_val,
        "estimatedValueFormatted": pac_formatted,
        "isEstimatedValueHidden": is_hidden,
        "detailedDescription": desc,
        "bidSubmissionStartDate": start_date,
        "bidSubmissionDueDate": due_date,
        "bidOpenDate": open_date,
        "physicalDocSubmissionDate": doc_sub_date,
        "emd": {
            "amount": emd_amount,
            "amountFormatted": f"₹ {emd_amount:,.2f}" if emd_amount > 0 else "As per NIT",
            "paymentMode": emd_mode
        },
        "tenderFee": {
            "amount": tender_fee,
            "amountFormatted": f"₹ {tender_fee:,.2f}" if tender_fee > 0 else "As per NIT",
            "paymentMode": tender_fee_mode
        },
        "tenderProcessingFee": {
            "amount": processing_fee,
            "amountFormatted": f"₹ {processing_fee:,.2f}" if processing_fee > 0 else "As per NIT",
            "paymentMode": processing_fee_mode
        },
        "tenderFolder": folder_path,
        "termsAndConditionsFiles": terms_files,
        "attachmentsFiles": official_files,
        "boqFiles": boq_files,
        "totalAttachmentsCount": len(terms_files) + len(official_files) + len(boq_files),
        "district": detect_district_enhanced(desc)
    }

def enrich_all_tenders(max_count=None, workers=10):
    """
    Crawls and enriches all active tenders from portal using concurrent threads.
    Updates action_details_cache.json and rebuilds unified datasets.
    """
    print(f"\n[Enricher] Loading target tenders for Action Preview enrichment...")
    cache = load_cache()

    target_tids = []
    # 1. Active tenders from tenders_data.json
    if os.path.exists(TENDERS_DATA_PATH):
        with open(TENDERS_DATA_PATH, "r", encoding="utf-8") as f:
            td = json.load(f)
        for t in td.get("tenders", []):
            tid = t.get("id")
            if tid and tid not in target_tids:
                target_tids.append(tid)

    if max_count:
        target_tids = target_tids[:max_count]

    # Filter out already enriched tenders unless force
    needed_tids = [tid for tid in target_tids if str(tid) not in cache or not cache[str(tid)].get('bidOpenDate')]
    print(f"[Enricher] Total target: {len(target_tids)}, Already enriched: {len(target_tids) - len(needed_tids)}, Need to enrich: {len(needed_tids)}")

    if not needed_tids:
        print("[Enricher] All tenders already enriched in cache! Rebuilding unified dataset...")
        build_comprehensive_unified_dataset()
        return

    print(f"[Enricher] Connecting to Bihar eProcurement portal...")
    opener, auth_token = get_auth_session()
    print(f"[Enricher] Session established. Launching {workers} concurrent workers...")

    start_time = time.time()
    success_count = 0
    fail_count = 0

    def worker_task(tid):
        try:
            return fetch_single_action_preview(tid, opener, auth_token)
        except Exception as e:
            # Retry once
            try:
                time.sleep(0.5)
                return fetch_single_action_preview(tid, opener, auth_token)
            except Exception as e2:
                return {"tenderId": tid, "error": str(e2)}

    # Process in chunks of 50 to periodically save cache
    chunk_size = 50
    for i in range(0, len(needed_tids), chunk_size):
        chunk = needed_tids[i:i + chunk_size]
        print(f"[Enricher] Processing batch {i+1} to {min(i+len(chunk), len(needed_tids))} of {len(needed_tids)}...")

        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            results = list(executor.map(worker_task, chunk))

        for res in results:
            tid_str = str(res.get("tenderId"))
            if "error" not in res:
                cache[tid_str] = res
                success_count += 1
            else:
                fail_count += 1

        # Periodically persist cache
        save_cache(cache)
        print(f"      [Progress] {success_count} enriched, {fail_count} skipped. Elapsed: {time.time()-start_time:.1f}s")

    total_time = time.time() - start_time
    print(f"\n[Enricher] Completed enrichment in {total_time:.2f}s! Successfully enriched {success_count} tenders.")

    # Rebuild unified dataset and Excel workbook
    print("[Enricher] Updating unified datasets (excel_tenders_view.json & Bihar_eProc_Tenders_Comprehensive.xlsx)...")
    records = build_comprehensive_unified_dataset()
    print(f"[Enricher] Done! All {len(records)} tenders now have full Action details, EMD, Fees, 3 Dates & Attachments!")

if __name__ == "__main__":
    count_arg = int(sys.argv[1]) if len(sys.argv) > 1 else None
    enrich_all_tenders(max_count=count_arg, workers=12)
