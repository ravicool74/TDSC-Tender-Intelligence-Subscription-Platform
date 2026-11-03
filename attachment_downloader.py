r"""
Bihar e-Procurement Attachment & BOQ Downloader Module
Extracts deep tender details and downloads:
1. Terms and Conditions attachments
2. Official Attachments (NIT, etc.)
3. BOQ (Bill of Quantities) Excel / PDF files
Organizes them into local folders:
c:\e-tender\downloads\tender_<systemTenderNo>_<tenderId>\
"""

import os
import re
import json
import urllib.request
import urllib.parse
import ssl
from datetime import datetime
from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS, format_currency_inr

try:
    ssl._create_default_https_context = ssl._create_unverified_context
except AttributeError:
    pass

DOWNLOADS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")

def format_timestamp_date(ts):
    if not ts:
        return "N/A"
    try:
        ts = int(ts)
        if ts <= 0: return "N/A"
        dt = datetime.fromtimestamp(ts / 1000.0)
        return dt.strftime("%d-%m-%Y %I:%M %p")
    except:
        return "N/A"

def fetch_and_download_tender_data(tender_id, download_files=True):
    """
    Simulates clicking the 'Action' button on eproc2.bihar.gov.in:
    - Extracts all exact fields requested by user.
    - If download_files=True, downloads:
        1. Terms and Conditions attachment -> downloads/<tender_dir>/terms_and_conditions/
        2. Official Attachments (NIT, etc.) -> downloads/<tender_dir>/attachments/
        3. BOQ (Bill of Quantities) file -> downloads/<tender_dir>/boq/
    - Returns structured dictionary with details and local file paths.
    """
    opener, auth_token = get_auth_session()
    
    # Fetch deep preview
    preview_url = f"{BASE_URL}/rest/tender/previewTenderByTenderId?tenderId={tender_id}"
    api_headers = {
        'User-Agent': HEADERS['User-Agent'],
        'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
        'Content-Type': 'application/json',
        'Authorization': auth_token
    }
    req = urllib.request.Request(preview_url, data=b"", headers=api_headers)
    with opener.open(req, timeout=25) as resp:
        data = json.loads(resp.read().decode('utf-8', errors='replace'))

    system_tender_no = data.get('orgtenderid') or data.get('tenderid')
    tender_ref_no = data.get('tenderrefno') or "N/A"
    description = (data.get('description') or "").strip()
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
    bid_start_date = "N/A"
    bid_due_date = "N/A"
    bid_open_date = "N/A"
    doc_sub_date = "N/A"

    # Payments
    emd_amount = 0.0
    emd_mode = "Internet Payment Gateway(IPG),Challan|Bank Guarantee"
    tender_fee = 0.0
    tender_fee_mode = "Internet Payment Gateway(IPG),Challan|"
    processing_fee = 0.0
    processing_fee_mode = "Internet Payment Gateway(IPG),Challan|"

    # Attachments to collect
    terms_files = []
    official_files = []
    boq_files = []

    templates = data.get('templates', [])
    for tmpl in templates:
        short = tmpl.get('shortName', '')
        grp_id = tmpl.get('templategroupId')
        fdict = {f.get('code'): f.get('value') for f in tmpl.get('templateFieldList', []) if f.get('code')}

        if short == 'rfqdate':
            bid_start_date = format_timestamp_date(fdict.get('bid_start_date'))
            bid_due_date = format_timestamp_date(fdict.get('bid_end_date'))
            bid_open_date = format_timestamp_date(fdict.get('bid_open_date'))
            doc_sub_date = format_timestamp_date(fdict.get('doc_sub_date'))

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
            # Terms and Conditions
            unf2 = fdict.get('unf2') or "Terms and Conditions"
            unf3 = fdict.get('unf3') or ""
            if unf3:
                terms_files.append({
                    "title": unf2,
                    "attachFile": unf3,
                    "templategroupId": grp_id,
                    "fileName": unf3.split('|')[-1] if '|' in unf3 else "Terms_Conditions.pdf"
                })

        elif short == 'rfqattachment':
            # Official Attachments
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
            # BOQ items
            item_name = fdict.get('item_name') or fdict.get('item_code') or "BOQ Item"
            unf3 = fdict.get('unf3') or ""
            if unf3:
                boq_files.append({
                    "label": f"BOQ ({item_name})",
                    "fileName": unf3.split('|')[-1] if '|' in unf3 else "BOQ.pdf",
                    "attachFile": unf3,
                    "templategroupId": grp_id
                })

    # Prepare tender folder
    tender_folder_name = f"tender_{system_tender_no}_{tender_id}"
    tender_folder_path = os.path.join(DOWNLOADS_DIR, tender_folder_name)
    terms_dir = os.path.join(tender_folder_path, "terms_and_conditions")
    attach_dir = os.path.join(tender_folder_path, "attachments")
    boq_dir = os.path.join(tender_folder_path, "boq")

    if download_files:
        os.makedirs(terms_dir, exist_ok=True)
        os.makedirs(attach_dir, exist_ok=True)
        os.makedirs(boq_dir, exist_ok=True)

    def download_single_file(file_info, dest_folder):
        raw_attach = file_info.get("attachFile", "")
        grp = file_info.get("templategroupId")
        if not raw_attach:
            return None

        # Determine templategroupId from path if not explicitly provided
        if not grp and '/' in raw_attach:
            try:
                grp = int(raw_attach.split('/')[0])
            except:
                pass

        # Target file name
        orig_name = file_info.get("fileName") or raw_attach.split('|')[-1]
        clean_name = re.sub(r'[\\/*?:"<>|]', '_', orig_name)
        dest_file_path = os.path.join(dest_folder, clean_name)

        if os.path.exists(dest_file_path) and os.path.getsize(dest_file_path) > 1000:
            file_info["localPath"] = dest_file_path
            file_info["fileSize"] = os.path.getsize(dest_file_path)
            file_info["downloaded"] = True
            file_info["webUrl"] = f"/downloads/{tender_folder_name}/{os.path.basename(dest_folder)}/{urllib.parse.quote(clean_name)}"
            return dest_file_path

        download_action_url = f"{BASE_URL}/rest/openarea/downloadattachment.action"
        form_data = {
            'Authorization': auth_token,
            'tempGrpId': str(grp),
            'tenderId': str(tender_id),
            'fileName': raw_attach,
            'orgId': '538'
        }
        encoded_data = urllib.parse.urlencode(form_data).encode('utf-8')
        d_headers = {
            'User-Agent': HEADERS['User-Agent'],
            'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        try:
            d_req = urllib.request.Request(download_action_url, data=encoded_data, headers=d_headers)
            with opener.open(d_req, timeout=20) as d_resp:
                content = d_resp.read()
                if len(content) > 500 and not content.strip().startswith(b'<!DOCTYPE html>'):
                    with open(dest_file_path, "wb") as f_out:
                        f_out.write(content)
                    file_info["localPath"] = dest_file_path
                    file_info["fileSize"] = len(content)
                    file_info["downloaded"] = True
                    file_info["webUrl"] = f"/downloads/{tender_folder_name}/{os.path.basename(dest_folder)}/{urllib.parse.quote(clean_name)}"
                    return dest_file_path
                else:
                    file_info["downloaded"] = False
                    file_info["error"] = "Invalid file payload or auth expired"
        except Exception as e:
            file_info["downloaded"] = False
            file_info["error"] = str(e)
        return None

    if download_files:
        # Download Terms and Conditions
        for t_file in terms_files:
            download_single_file(t_file, terms_dir)

        # Download Official Attachments (NIT, etc.)
        for o_file in official_files:
            download_single_file(o_file, attach_dir)

        # Download BOQ files
        for b_file in boq_files:
            download_single_file(b_file, boq_dir)

    result = {
        "tenderId": tender_id,
        "systemTenderNo": system_tender_no,
        "tenderRefNo": tender_ref_no,
        "estimatedValue": pac_val,
        "estimatedValueFormatted": pac_formatted,
        "isEstimatedValueHidden": is_hidden,
        "detailedDescription": description,
        "bidSubmissionStartDate": bid_start_date,
        "bidSubmissionDueDate": bid_due_date,
        "bidOpenDate": bid_open_date,
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
        "tenderFolder": tender_folder_path,
        "termsAndConditionsFiles": terms_files,
        "attachmentsFiles": official_files,
        "boqFiles": boq_files,
        "totalAttachmentsCount": len(terms_files) + len(official_files) + len(boq_files)
    }

    return result

if __name__ == "__main__":
    import sys
    test_id = 142360
    if len(sys.argv) > 1:
        test_id = int(sys.argv[1])
    print(f"Testing deep action extraction & download for Tender ID: {test_id}...")
    res = fetch_and_download_tender_data(test_id, download_files=True)
    print("\nExtraction Summary:")
    print(f"System Tender No: {res['systemTenderNo']}")
    print(f"Ref No: {res['tenderRefNo']}")
    print(f"Estimated Value: {res['estimatedValue']}")
    print(f"EMD: {res['emd']['amount']} ({res['emd']['paymentMode']})")
    print(f"Tender Fee: {res['tenderFee']['amount']} ({res['tenderFee']['paymentMode']})")
    print(f"Processing Fee: {res['tenderProcessingFee']['amount']} ({res['tenderProcessingFee']['paymentMode']})")
    print(f"Start Date: {res['bidSubmissionStartDate']}")
    print(f"Due Date: {res['bidSubmissionDueDate']}")
    print(f"Open Date: {res['bidOpenDate']}")
    print(f"Terms Files: {len(res['termsAndConditionsFiles'])}")
    print(f"Attachment Files: {len(res['attachmentsFiles'])}")
    print(f"BOQ Files: {len(res['boqFiles'])}")
    print(f"Local Folder: {res['tenderFolder']}")
