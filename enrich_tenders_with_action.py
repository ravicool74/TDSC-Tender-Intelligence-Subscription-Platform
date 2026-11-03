"""
Enrich tenders dataset with Action details, system tender numbers,
dates, payments, and attachment links.
"""

import json
import os
from action_details_manager import load_cache, get_action_details

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
TENDERS_DATA_PATH = os.path.join(DIRECTORY, "tenders_data.json")
EXCEL_VIEW_PATH = os.path.join(DIRECTORY, "excel_tenders_view.json")

def enrich_dataset():
    cache = load_cache()
    
    # Pre-load known sample tender 142360 if not present
    if "142360" not in cache:
        try:
            get_action_details(142360, download_files=True)
            cache = load_cache()
        except Exception as e:
            print("Notice:", e)

    with open(TENDERS_DATA_PATH, "r", encoding="utf-8") as f:
        tdata = json.load(f)
    raw_tenders = tdata.get("tenders", [])

    enriched = []
    for idx, t in enumerate(raw_tenders, start=1):
        tid = t.get("id")
        tid_str = str(tid)
        rfq_id = t.get("rfqId") or tid
        
        deep = cache.get(tid_str, {})

        # System Tender No
        sys_tender_no = deep.get("systemTenderNo") or rfq_id or tid

        # Ref No
        ref_no = deep.get("tenderRefNo") or t.get("refNo") or "N/A"

        # Description
        desc = deep.get("detailedDescription") or t.get("title") or t.get("description") or "Tender Notice"

        # Dates
        start_date = deep.get("bidSubmissionStartDate") or t.get("startDate") or "N/A"
        due_date = deep.get("bidSubmissionDueDate") or t.get("closingDate") or t.get("dueDate") or "Refer Notice"
        open_date = deep.get("bidOpenDate") or t.get("openingDate") or "Refer Notice"

        # Estimated Value
        est_val = deep.get("estimatedValue") if deep.get("estimatedValue") is not None else t.get("estimatedValue")
        if est_val is not None and est_val > 0:
            est_val_formatted = f"₹ {est_val:,.2f}"
            is_hidden = False
        else:
            est_val = None
            est_val_formatted = "Hidden by portal"
            is_hidden = True

        # Payments
        deep_emd = deep.get("emd", {})
        emd_amt = deep_emd.get("amount") if deep_emd.get("amount") else t.get("emdAmount")
        emd_mode = deep_emd.get("paymentMode") or "Internet Payment Gateway(IPG),Challan|Bank Guarantee"
        emd_formatted = f"₹ {emd_amt:,.2f}" if (emd_amt and emd_amt > 0) else "As per NIT"

        deep_fee = deep.get("tenderFee", {})
        tender_fee_amt = deep_fee.get("amount") or 0.0
        tender_fee_mode = deep_fee.get("paymentMode") or "Internet Payment Gateway(IPG),Challan|"
        tender_fee_formatted = f"₹ {tender_fee_amt:,.2f}" if tender_fee_amt > 0 else "As per NIT"

        deep_proc = deep.get("tenderProcessingFee", {})
        proc_fee_amt = deep_proc.get("amount") or 0.0
        proc_fee_mode = deep_proc.get("paymentMode") or "Internet Payment Gateway(IPG),Challan|"
        proc_fee_formatted = f"₹ {proc_fee_amt:,.2f}" if proc_fee_amt > 0 else "As per NIT"

        # Files
        terms_files = deep.get("termsAndConditionsFiles", [])
        official_files = deep.get("attachmentsFiles", [])
        boq_files = deep.get("boqFiles", [])
        has_downloaded = deep.get("totalAttachmentsCount", 0) > 0 or (len(terms_files) + len(official_files) + len(boq_files) > 0)

        folder = deep.get("tenderFolder") or os.path.join(DIRECTORY, "downloads", f"tender_{sys_tender_no}_{tid}")

        enriched.append({
            "slNo": idx,
            "systemTenderNo": sys_tender_no,
            "tenderId": tid,
            "refNo": ref_no,
            "department": t.get("department", "Government Department"),
            "district": t.get("district", "State-wide / Headquarter"),
            "estimatedValue": est_val,
            "estimatedValueFormatted": est_val_formatted,
            "isEstimatedValueHidden": is_hidden,
            "description": desc,
            "detailedDescription": desc,
            "startDate": start_date,
            "dueDate": due_date,
            "openDate": open_date,
            "emdAmount": emd_amt,
            "emdAmountFormatted": emd_formatted,
            "emdMode": emd_mode,
            "tenderFee": tender_fee_amt,
            "tenderFeeFormatted": tender_fee_formatted,
            "tenderFeeMode": tender_fee_mode,
            "processingFee": proc_fee_amt,
            "processingFeeFormatted": proc_fee_formatted,
            "processingFeeMode": proc_fee_mode,
            "termsFiles": terms_files,
            "officialFiles": official_files,
            "boqFiles": boq_files,
            "tenderFolder": folder,
            "hasDownloaded": has_downloaded,
            "actionExtracted": bool(deep)
        })

    with open(EXCEL_VIEW_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "totalTenders": len(enriched),
            "tenders": enriched
        }, f, indent=2, ensure_ascii=False)

    print(f"Enriched {len(enriched)} tenders in {EXCEL_VIEW_PATH}")
    return enriched

if __name__ == "__main__":
    enrich_dataset()
