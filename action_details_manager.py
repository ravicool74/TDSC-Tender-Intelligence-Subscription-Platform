"""
Bihar e-Procurement Action Details & Attachment Cache Manager
Provides on-demand and batch extraction of deep Action details and attachments:
- System Tender No.
- Reference No.
- Estimated Value in INR
- Detailed Description
- Bid Submission Start Date, Due Date, Open Date
- EMD (Amount & Payment Mode)
- Tender Fee (Amount & Payment Mode)
- Tender Processing Fee (Amount & Payment Mode)
- Terms and Conditions attachments (downloads to downloads/tender_.../terms_and_conditions/)
- Official Attachments (downloads to downloads/tender_.../attachments/)
- BOQ (downloads to downloads/tender_.../boq/)
"""

import os
import json
import threading
import time
from attachment_downloader import fetch_and_download_tender_data, DOWNLOADS_DIR

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(DIRECTORY, "action_details_cache.json")

_cache_lock = threading.Lock()

def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_cache(cache):
    with _cache_lock:
        tmp_file = CACHE_FILE + ".tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, ensure_ascii=False)
        try:
            os.replace(tmp_file, CACHE_FILE)
        except Exception:
            pass

def get_action_details(tender_id, download_files=True, force_refresh=False):
    cache = load_cache()
    tid_str = str(tender_id)
    
    if not force_refresh and tid_str in cache:
        cached = cache[tid_str]
        # If download was requested, check if files are on disk
        if download_files:
            folder = cached.get('tenderFolder')
            if folder and os.path.exists(folder):
                # Check if at least one file exists in folder
                has_files = any(len(files) > 0 for _, _, files in os.walk(folder))
                if has_files:
                    return cached

    # Fetch live from portal
    data = fetch_and_download_tender_data(int(tender_id), download_files=download_files)
    cache[tid_str] = data
    save_cache(cache)
    return data

def sync_single_tender_live(tender_id, download_files=False):
    """
    Connects directly to eproc2.bihar.gov.in for a SINGLE tender.
    Extracts latest live due date, PAC amount, EMD, fee, corrigendum and attachments.
    Updates action_details_cache.json, excel_tenders_view.json, and returns
    detailed change comparison without reloading or disturbing other tenders.
    """
    import datetime
    tid_int = int(tender_id)
    tid_str = str(tender_id)

    # 1. Fetch live from eproc2 portal
    live_data = fetch_and_download_tender_data(tid_int, download_files=download_files)

    # 2. Update action details cache
    cache = load_cache()
    cache[tid_str] = live_data
    save_cache(cache)

    # 3. Update in excel_tenders_view.json
    view_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_tenders_view.json")
    updated_tender = None
    changes_detected = []

    if os.path.exists(view_path):
        try:
            with open(view_path, "r", encoding="utf-8") as f:
                view_data = json.load(f)
            tenders = view_data.get("tenders", [])

            # Match by tenderId or systemTenderNo
            target_t = next((t for t in tenders if str(t.get("tenderId")) == tid_str or str(t.get("systemTenderNo")) == tid_str), None)

            if target_t:
                # Compare Due Date
                old_due = target_t.get("dueDate", "N/A")
                new_due = live_data.get("bidSubmissionDueDate")
                if new_due and new_due != "N/A" and new_due != old_due:
                    changes_detected.append(f"Due Date: {old_due} ➔ {new_due}")
                    target_t["dueDate"] = new_due

                # Compare PAC Value
                old_pac = target_t.get("estimatedValue")
                new_pac = live_data.get("estimatedValue")
                if new_pac is not None and new_pac > 0 and new_pac != old_pac:
                    changes_detected.append(f"PAC: {target_t.get('estimatedValueFormatted', 'Hidden')} ➔ {live_data.get('estimatedValueFormatted')}")
                    target_t["estimatedValue"] = new_pac
                    target_t["estimatedValueFormatted"] = live_data.get("estimatedValueFormatted")
                    target_t["isEstimatedValueHidden"] = False

                # Compare EMD
                old_emd = target_t.get("emdAmount")
                new_emd = live_data.get("emd", {}).get("amount")
                if new_emd and new_emd > 0 and new_emd != old_emd:
                    changes_detected.append(f"EMD: {target_t.get('emdAmountFormatted')} ➔ {live_data.get('emd', {}).get('amountFormatted')}")
                    target_t["emdAmount"] = new_emd
                    target_t["emdAmountFormatted"] = live_data.get("emd", {}).get("amountFormatted")
                    if live_data.get("emd", {}).get("paymentMode"):
                        target_t["emdMode"] = live_data.get("emd", {}).get("paymentMode")

                # Update Start & Open Dates if present
                if live_data.get("bidSubmissionStartDate") and live_data.get("bidSubmissionStartDate") != "N/A":
                    target_t["startDate"] = live_data.get("bidSubmissionStartDate")
                if live_data.get("bidOpenDate") and live_data.get("bidOpenDate") != "N/A":
                    target_t["openDate"] = live_data.get("bidOpenDate")

                # Update Fee
                if live_data.get("tenderFee", {}).get("amount"):
                    target_t["tenderFee"] = live_data.get("tenderFee", {}).get("amount")
                    target_t["tenderFeeFormatted"] = live_data.get("tenderFee", {}).get("amountFormatted")
                if live_data.get("tenderProcessingFee", {}).get("amount"):
                    target_t["processingFee"] = live_data.get("tenderProcessingFee", {}).get("amount")
                    target_t["processingFeeFormatted"] = live_data.get("tenderProcessingFee", {}).get("amountFormatted")

                # Update Attachments & BOQ lists
                target_t["termsFiles"] = live_data.get("termsAndConditionsFiles", [])
                target_t["officialFiles"] = live_data.get("attachmentsFiles", [])
                target_t["boqFiles"] = live_data.get("boqFiles", [])
                target_t["actionExtracted"] = True

                folder = target_t.get("tenderFolder") or live_data.get("tenderFolder")
                if folder and os.path.exists(folder):
                    target_t["hasDownloaded"] = any(len(files) > 0 for _, _, files in os.walk(folder))
                target_t["tenderFolder"] = folder

                target_t["lastLiveChecked"] = datetime.datetime.now().strftime("%d-%m-%Y %I:%M %p")
                updated_tender = target_t

                # Save updated JSON
                with open(view_path, "w", encoding="utf-8") as f:
                    json.dump(view_data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Notice: Failed to update view json during single tender refresh: {e}")

    return {
        "success": True,
        "tenderId": tid_int,
        "updatedTender": updated_tender or live_data,
        "liveDetails": live_data,
        "changesDetected": changes_detected,
        "message": f"Tender #{tid_int} successfully verified with eProc2 live server!" + (f" ({', '.join(changes_detected)})" if changes_detected else " (Data is current, no changes on portal)")
    }

def pre_populate_known_tenders():
    """Pre-populates sample tenders like 142360 into cache."""
    try:
        get_action_details(142360, download_files=True)
    except Exception as e:
        print(f"Notice: pre-population error for 142360: {e}")

if __name__ == "__main__":
    print("Testing action_details_manager on 142360...")
    res = get_action_details(142360, download_files=True)
    print("Cached & Downloaded successfully:")
    print("System Tender No:", res['systemTenderNo'])
    print("Estimated Value:", res['estimatedValue'])
    print("Total Attachments:", res['totalAttachmentsCount'])
