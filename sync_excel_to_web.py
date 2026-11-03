"""
Sync data from Bihar_eProc_Tenders_Comprehensive.xlsx to web-ready JSON format.
Extracts the exact 9 fields requested:
1. SL. No.
2. Tender ID
3. Reference Number
4. Department Name
5. Estimated Value (INR)
6. EMD Amount (INR)
7. Bid Submission Due Date
8. Tender Description
9. District (searches description, venue, ref no, dept)
"""

import openpyxl
import json
import re
import os

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

    # Also check Hindi district names
    hindi_districts = {
        "अररिया": "Araria", "सुपौल": "Supaul", "औरंगाबाद": "Aurangabad", "पटना": "Patna",
        "गया": "Gaya", "भागलपूर": "Bhagalpur", "मुंगेर": "Munger", "मुजफ्फरपुर": "Muzaffarpur",
        "पूर्णिया": "Purnia", "रोहतास": "Rohtas (Sasaram)", "सोनपुर": "Saran (Chapra/Sonpur)"
    }
    for h_name, e_name in hindi_districts.items():
        if h_name in combined:
            return e_name

    return "State-wide / Headquarter"

def format_inr(val):
    if val is None or val == "" or val == "Hidden by portal":
        return "Hidden by portal"
    try:
        fval = float(val)
        if fval >= 10000000:
            return f"₹ {fval / 10000000:.2f} Cr"
        elif fval >= 100000:
            return f"₹ {fval / 100000:.2f} Lakh"
        else:
            return f"₹ {fval:,.0f}"
    except:
        return str(val)

def extract_excel_data():
    excel_path = os.path.join(os.path.dirname(__file__), "Bihar_eProc_Tenders_Comprehensive.xlsx")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb['Main Tenders']

    tenders = []
    sl_no = 1

    for row in range(2, ws.max_row + 1):
        tender_id = ws.cell(row=row, column=1).value
        if not tender_id:
            continue

        ref_no = str(ws.cell(row=row, column=2).value or "").strip()
        desc = str(ws.cell(row=row, column=3).value or "").strip()
        dept = str(ws.cell(row=row, column=4).value or "").strip()
        est_val = ws.cell(row=row, column=7).value
        visibility = str(ws.cell(row=row, column=8).value or "")
        emd_val = ws.cell(row=row, column=9).value
        due_date = str(ws.cell(row=row, column=13).value or "").strip()
        prebid_venue = str(ws.cell(row=row, column=18).value or "").strip()

        # Find District
        district = find_district(desc, prebid_venue, ref_no, dept)

        is_hidden = (est_val is None or "hidden" in visibility.lower())
        num_est_val = float(est_val) if not is_hidden else None
        num_emd = float(emd_val) if (emd_val and str(emd_val).replace('.','',1).isdigit()) else None

        tenders.append({
            "slNo": sl_no,
            "tenderId": tender_id,
            "refNo": ref_no,
            "department": dept,
            "estimatedValue": num_est_val,
            "estimatedValueFormatted": format_inr(num_est_val) if num_est_val else "Hidden by portal",
            "isEstimatedValueHidden": is_hidden,
            "emdAmount": num_emd,
            "emdAmountFormatted": format_inr(num_emd) if num_emd else "As per NIT",
            "dueDate": due_date,
            "description": desc,
            "district": district,
            "prebidVenue": prebid_venue
        })
        sl_no += 1

    output_json = os.path.join(os.path.dirname(__file__), "excel_tenders_view.json")
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump({
            "totalTenders": len(tenders),
            "tenders": tenders
        }, f, indent=2, ensure_ascii=False)

    print(f"Successfully extracted {len(tenders)} records into {output_json}!")
    return tenders

if __name__ == "__main__":
    extract_excel_data()
