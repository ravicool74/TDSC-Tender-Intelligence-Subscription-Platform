"""
Comprehensive Dataset Unifier & Excel Updater
Merges tenders from:
1. action_details_cache.json (Deep action details, real downloaded files)
2. tenders_data.json (Active tenders enriched with rfqId, category, etc.)
3. Bihar_eProc_Tenders_Comprehensive.xlsx (All tenders)
Ensures Tender 142360 (System Tender No 142165) and all 1450+ tenders are included with all requested fields.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
import os
import re
from action_details_manager import load_cache, get_action_details

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(DIRECTORY, "Bihar_eProc_Tenders_Comprehensive.xlsx")
VIEW_JSON_PATH = os.path.join(DIRECTORY, "excel_tenders_view.json")
TENDERS_DATA_PATH = os.path.join(DIRECTORY, "tenders_data.json")

def format_inr(val):
    if val is None or val == "" or str(val).lower() == "none":
        return "Hidden by portal"
    try:
        fval = float(val)
        return f"₹ {fval:,.2f}"
    except:
        return str(val)

def build_comprehensive_unified_dataset(build_excel=False):
    cache = load_cache()
    
    # Ensure sample tender 142360 is in cache and downloaded
    if "142360" not in cache:
        try:
            get_action_details(142360, download_files=True)
            cache = load_cache()
        except Exception as e:
            print("Notice:", e)

    tenders_map = {}

    # 1. Load from tenders_data.json
    if os.path.exists(TENDERS_DATA_PATH):
        with open(TENDERS_DATA_PATH, "r", encoding="utf-8") as f:
            td = json.load(f)
        for t in td.get("tenders", []):
            tid = t.get("id")
            if not tid: continue
            rfq_id = t.get("rfqId") or tid
            tenders_map[tid] = {
                "tenderId": tid,
                "systemTenderNo": rfq_id,
                "refNo": t.get("refNo") or "N/A",
                "department": t.get("department") or "Government Department",
                "district": t.get("district") or "State-wide / Headquarter",
                "estimatedValue": t.get("estimatedValue") if (t.get("estimatedValue") and t.get("estimatedValue") > 0) else None,
                "description": t.get("title") or t.get("description") or "Tender Notice",
                "startDate": t.get("startDate") or "N/A",
                "dueDate": t.get("closingDate") or t.get("dueDate") or "Refer Notice",
                "openDate": t.get("openingDate") or "Refer Notice",
                "emdAmount": t.get("emdAmount") if (t.get("emdAmount") and t.get("emdAmount") > 0) else None,
                "emdMode": "Internet Payment Gateway(IPG),Challan|Bank Guarantee",
                "tenderFee": 0.0,
                "tenderFeeMode": "Internet Payment Gateway(IPG),Challan|",
                "processingFee": 0.0,
                "processingFeeMode": "Internet Payment Gateway(IPG),Challan|"
            }

    # 2. Merge from Bihar_eProc_Tenders_Comprehensive.xlsx
    if os.path.exists(EXCEL_PATH):
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb['Main Tenders']
        headers = [cell for cell in next(ws.iter_rows(values_only=True))]
        h_map = {h: i for i, h in enumerate(headers)}

        for row in ws.iter_rows(min_row=2, values_only=True):
            tid = row[h_map.get("Tender ID", 1)]
            if not tid: continue
            try: tid = int(tid)
            except: pass

            ref_no = str(row[h_map.get("Reference Number", 2)] or "N/A").strip()
            desc = str(row[h_map.get("Tender Description", 3)] or "").strip()
            dept = str(row[h_map.get("Department", 4)] or "").strip()
            cell_pac = row[h_map.get("Estimated Value (INR)", 7)]
            cell_emd = row[h_map.get("EMD Amount (INR)", 9)]
            cell_fee = row[h_map.get("Tender Fee (INR)", 10)]
            cell_proc = row[h_map.get("Processing Fee (INR)", 11)]
            start_d = str(row[h_map.get("Bid Submission Start", 12)] or "N/A").strip()
            due_d = str(row[h_map.get("Bid Submission Due Date", 13)] or "Refer Notice").strip()
            open_d = str(row[h_map.get("Bid Open Date", 14)] or "Refer Notice").strip()
            dist = str(row[h_map.get("District", 19)] or "State-wide / Headquarter").strip()

            num_pac = None
            if cell_pac is not None:
                try: num_pac = float(cell_pac)
                except: pass

            num_emd = None
            if cell_emd is not None:
                try: num_emd = float(cell_emd)
                except: pass

            num_fee = 0.0
            if cell_fee is not None:
                try:
                    num_fee = float(cell_fee)
                except:
                    num_fee = 0.0

            num_proc = 0.0
            if cell_proc is not None:
                try:
                    num_proc = float(cell_proc)
                except:
                    num_proc = 0.0

            if tid not in tenders_map:
                tenders_map[tid] = {
                    "tenderId": tid,
                    "systemTenderNo": tid,
                    "refNo": ref_no,
                    "department": dept,
                    "district": dist,
                    "estimatedValue": num_pac,
                    "description": desc,
                    "startDate": start_d,
                    "dueDate": due_d,
                    "openDate": open_d,
                    "emdAmount": num_emd,
                    "emdMode": "Internet Payment Gateway(IPG),Challan|Bank Guarantee",
                    "tenderFee": num_fee,
                    "tenderFeeMode": "Internet Payment Gateway(IPG),Challan|",
                    "processingFee": num_proc,
                    "processingFeeMode": "Internet Payment Gateway(IPG),Challan|"
                }
            else:
                # Fill missing details
                if num_pac and not tenders_map[tid]["estimatedValue"]:
                    tenders_map[tid]["estimatedValue"] = num_pac
                if num_emd and not tenders_map[tid]["emdAmount"]:
                    tenders_map[tid]["emdAmount"] = num_emd
                if num_fee: tenders_map[tid]["tenderFee"] = num_fee
                if num_proc: tenders_map[tid]["processingFee"] = num_proc
                if start_d != "N/A": tenders_map[tid]["startDate"] = start_d
                if open_d != "Refer Notice": tenders_map[tid]["openDate"] = open_d

    # 3. Apply Deep Action Cache details (like for 142360)
    for tid_str, deep in cache.items():
        try: tid = int(tid_str)
        except: tid = tid_str

        if tid in tenders_map:
            t = tenders_map[tid]
        else:
            t = {"tenderId": tid}
            tenders_map[tid] = t

        t["systemTenderNo"] = deep.get("systemTenderNo") or t.get("systemTenderNo") or tid
        t["refNo"] = deep.get("tenderRefNo") or t.get("refNo") or "N/A"
        if deep.get("detailedDescription"):
            t["description"] = deep.get("detailedDescription")
        if deep.get("estimatedValue") is not None and deep.get("estimatedValue") > 0:
            t["estimatedValue"] = deep.get("estimatedValue")
        if deep.get("bidSubmissionStartDate") and deep.get("bidSubmissionStartDate") != "N/A":
            t["startDate"] = deep.get("bidSubmissionStartDate")
        if deep.get("bidSubmissionDueDate") and deep.get("bidSubmissionDueDate") != "Refer Notice":
            t["dueDate"] = deep.get("bidSubmissionDueDate")
        if deep.get("bidOpenDate") and deep.get("bidOpenDate") != "Refer Notice":
            t["openDate"] = deep.get("bidOpenDate")

        deep_emd = deep.get("emd", {})
        if deep_emd.get("amount") is not None and deep_emd.get("amount") > 0:
            t["emdAmount"] = deep_emd.get("amount")
            t["emdMode"] = deep_emd.get("paymentMode") or t.get("emdMode")

        deep_fee = deep.get("tenderFee", {})
        if deep_fee.get("amount") is not None and deep_fee.get("amount") > 0:
            t["tenderFee"] = deep_fee.get("amount")
            t["tenderFeeMode"] = deep_fee.get("paymentMode") or t.get("tenderFeeMode")

        deep_proc = deep.get("tenderProcessingFee", {})
        if deep_proc.get("amount") is not None and deep_proc.get("amount") > 0:
            t["processingFee"] = deep_proc.get("amount")
            t["processingFeeMode"] = deep_proc.get("paymentMode") or t.get("processingFeeMode")

        if deep.get("district") and deep.get("district") not in ["State-wide / Headquarter", "Other / State-wide"]:
            t["district"] = deep.get("district")
        elif "sugaon" in (t.get("description", "") + " " + deep.get("detailedDescription", "")).lower():
            t["district"] = "East Champaran"

        t["termsFiles"] = deep.get("termsAndConditionsFiles", [])
        t["officialFiles"] = deep.get("attachmentsFiles", [])
        t["boqFiles"] = deep.get("boqFiles", [])
        t["tenderFolder"] = deep.get("tenderFolder") or os.path.join(DIRECTORY, "downloads", f"tender_{t['systemTenderNo']}_{tid}")
        t["actionExtracted"] = True

    # Build final list
    # Put 142360 first so it's row 1!
    ordered_tids = []
    if 142360 in tenders_map:
        ordered_tids.append(142360)
    for tid in tenders_map.keys():
        if tid != 142360:
            ordered_tids.append(tid)

    final_tenders = []
    for idx, tid in enumerate(ordered_tids, start=1):
        t = tenders_map[tid]
        est_val = t.get("estimatedValue")
        is_hidden = (est_val is None or est_val <= 0)
        est_formatted = format_inr(est_val) if not is_hidden else "Hidden by portal"

        emd_val = t.get("emdAmount")
        emd_formatted = format_inr(emd_val) if (emd_val and emd_val > 0) else "As per NIT"

        fee_raw = t.get("tenderFee", 0.0)
        try: fee_val = float(fee_raw) if fee_raw is not None else 0.0
        except: fee_val = 0.0
        fee_formatted = format_inr(fee_val) if fee_val > 0 else "As per NIT"

        proc_raw = t.get("processingFee", 0.0)
        try: proc_val = float(proc_raw) if proc_raw is not None else 0.0
        except: proc_val = 0.0
        proc_formatted = format_inr(proc_val) if proc_val > 0 else "As per NIT"

        sys_no = t.get("systemTenderNo") or tid
        folder = t.get("tenderFolder") or os.path.join(DIRECTORY, "downloads", f"tender_{sys_no}_{tid}")
        has_down = os.path.exists(folder) and any(len(files) > 0 for _, _, files in os.walk(folder))

        terms_f = t.get("termsFiles", [])
        official_f = t.get("officialFiles", [])
        boq_f = t.get("boqFiles", [])

        final_tenders.append({
            "slNo": idx,
            "systemTenderNo": sys_no,
            "tenderId": tid,
            "refNo": t.get("refNo", "N/A"),
            "department": t.get("department", "Government Department"),
            "district": t.get("district", "State-wide / Headquarter"),
            "estimatedValue": est_val if not is_hidden else None,
            "estimatedValueFormatted": est_formatted,
            "isEstimatedValueHidden": is_hidden,
            "description": t.get("description", "Tender Notice"),
            "detailedDescription": t.get("description", "Tender Notice"),
            "startDate": t.get("startDate", "N/A"),
            "dueDate": t.get("dueDate", "Refer Notice"),
            "openDate": t.get("openDate", "Refer Notice"),
            "emdAmount": emd_val,
            "emdAmountFormatted": emd_formatted,
            "emdMode": t.get("emdMode", "Internet Payment Gateway(IPG),Challan|Bank Guarantee"),
            "tenderFee": fee_val,
            "tenderFeeFormatted": fee_formatted,
            "tenderFeeMode": t.get("tenderFeeMode", "Internet Payment Gateway(IPG),Challan|"),
            "processingFee": proc_val,
            "processingFeeFormatted": proc_formatted,
            "processingFeeMode": t.get("processingFeeMode", "Internet Payment Gateway(IPG),Challan|"),
            "termsFiles": terms_f,
            "officialFiles": official_f,
            "boqFiles": boq_f,
            "tenderFolder": folder,
            "hasDownloaded": has_down,
            "actionExtracted": t.get("actionExtracted", False)
        })

    # Save to excel_tenders_view.json
    with open(VIEW_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "totalTenders": len(final_tenders),
            "tenders": final_tenders
        }, f, indent=2, ensure_ascii=False)

    print(f"Generated {len(final_tenders)} fully unified tenders in {VIEW_JSON_PATH}!")
    
    # Also update Bihar_eProc_Tenders_Comprehensive.xlsx only if build_excel is requested
    if build_excel:
        update_excel_file(final_tenders)
    return final_tenders

def update_excel_file(final_tenders):
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    gold_fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    bold_font = Font(name="Segoe UI", size=9, bold=True)
    green_font = Font(name="Segoe UI", size=9, bold=True, color="047857")
    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    ws_main = wb.active
    ws_main.title = "Main Tenders"

    headers = [
        "SL. No.", "System Tender No.", "Tender ID", "Reference Number", "Department Name",
        "Estimated Value (INR)", "Detailed Description", "Bid Submission Start Date",
        "Bid Submission Due Date", "Bid Open Date", "EMD Amount (INR)", "EMD Payment Mode",
        "Tender Fee (INR)", "Tender Fee Mode", "Processing Fee (INR)", "Processing Fee Mode",
        "District", "Attachments & BOQ Status"
    ]
    ws_main.append(headers)
    ws_main.row_dimensions[1].height = 28

    for c_idx in range(1, len(headers) + 1):
        cell = ws_main.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx, t in enumerate(final_tenders, start=2):
        status_str = "Files Downloaded in Folder" if t["hasDownloaded"] else ("Action Preview Extracted" if t["actionExtracted"] else "Available for Action Download")
        row = [
            t["slNo"],
            t["systemTenderNo"],
            t["tenderId"],
            t["refNo"],
            t["department"],
            t["estimatedValue"] if not t["isEstimatedValueHidden"] else "Hidden by portal",
            t["description"],
            t["startDate"],
            t["dueDate"],
            t["openDate"],
            t["emdAmount"] if t["emdAmount"] else "As per NIT",
            t["emdMode"],
            t["tenderFee"] if (isinstance(t.get("tenderFee"), (int, float)) and t["tenderFee"] > 0) else "As per NIT",
            t["tenderFeeMode"],
            t["processingFee"] if (isinstance(t.get("processingFee"), (int, float)) and t["processingFee"] > 0) else "As per NIT",
            t["processingFeeMode"],
            t["district"],
            status_str
        ]
        ws_main.append(row)
        for c_idx in range(1, len(row) + 1):
            cell = ws_main.cell(row=r_idx, column=c_idx)
            cell.font = data_font
            cell.border = thin_border
            if c_idx == 6 and not t["isEstimatedValueHidden"]:
                cell.number_format = '₹ #,##0.00'
                cell.font = green_font
                cell.alignment = Alignment(horizontal="right")

    ws_main.auto_filter.ref = f"A1:R{len(final_tenders)+1}"
    col_widths = [8, 18, 14, 22, 28, 22, 45, 20, 20, 20, 18, 28, 16, 26, 16, 26, 18, 24]
    for i, w in enumerate(col_widths, start=1):
        ws_main.column_dimensions[get_column_letter(i)].width = w

    # Shortlist sheet (descending estimated value)
    ws_short = wb.create_sheet(title="Shortlist")
    ws_short.append(headers)
    ws_short.row_dimensions[1].height = 28
    for c_idx in range(1, len(headers) + 1):
        cell = ws_short.cell(row=1, column=c_idx)
        cell.fill = gold_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    sorted_tenders = sorted(final_tenders, key=lambda x: x["estimatedValue"] if x["estimatedValue"] is not None else -1, reverse=True)
    for r_idx, t in enumerate(sorted_tenders, start=2):
        status_str = "Files Downloaded in Folder" if t["hasDownloaded"] else ("Action Preview Extracted" if t["actionExtracted"] else "Available for Action Download")
        row = [
            r_idx - 1,
            t["systemTenderNo"],
            t["tenderId"],
            t["refNo"],
            t["department"],
            t["estimatedValue"] if not t["isEstimatedValueHidden"] else "Hidden by portal",
            t["description"],
            t["startDate"],
            t["dueDate"],
            t["openDate"],
            t["emdAmount"] if t["emdAmount"] else "As per NIT",
            t["emdMode"],
            t["tenderFee"] if (isinstance(t.get("tenderFee"), (int, float)) and t["tenderFee"] > 0) else "As per NIT",
            t["tenderFeeMode"],
            t["processingFee"] if (isinstance(t.get("processingFee"), (int, float)) and t["processingFee"] > 0) else "As per NIT",
            t["processingFeeMode"],
            t["district"],
            status_str
        ]
        ws_short.append(row)
        for c_idx in range(1, len(row) + 1):
            cell = ws_short.cell(row=r_idx, column=c_idx)
            cell.font = data_font
            cell.border = thin_border
            if c_idx == 6 and not t["isEstimatedValueHidden"]:
                cell.number_format = '₹ #,##0.00'
                cell.font = green_font
                cell.alignment = Alignment(horizontal="right")

    ws_short.auto_filter.ref = f"A1:R{len(sorted_tenders)+1}"
    for i, w in enumerate(col_widths, start=1):
        ws_short.column_dimensions[get_column_letter(i)].width = w

    wb.save(EXCEL_PATH)
    print(f"Updated Excel file: {EXCEL_PATH} with {len(final_tenders)} rows across all requested fields!")

if __name__ == "__main__":
    build_comprehensive_unified_dataset()
