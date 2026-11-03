"""
Bihar eProc2 Tender Data Fetcher & Detail Processor
Fetches active tenders, enriches with Category, Amount, Dates, District, Place,
and provides on-demand Action (Deep-Dive) inspection for each tender.
"""

import urllib.request
import http.cookiejar
import json
import re
import datetime
import os
import sys
import ssl

# Globally disable SSL certificate verification for NIC / eProcurement India portals
try:
    ssl._create_default_https_context = ssl._create_unverified_context
except AttributeError:
    pass

BASE_URL = "https://eproc2.bihar.gov.in/EPSV2Web"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5'
}

BIHAR_DISTRICTS = [
    "Araria", "Arwal", "Aurangabad", "Banka", "Begusarai", "Bhagalpur", "Bhojpur", "Buxar",
    "Darbhanga", "East Champaran", "Motihari", "Gaya", "Gopalganj", "Jamui", "Jehanabad",
    "Kaimur", "Bhabhua", "Katihar", "Khagaria", "Kishanganj", "Lakhisarai", "Madhepura",
    "Madhubani", "Munger", "Muzaffarpur", "Nalanda", "Bihar Sharif", "Nawada", "Patna",
    "Purnia", "Rohtas", "Sasaram", "Saharsa", "Samastipur", "Saran", "Chapra", "Sheikhpura",
    "Sheohar", "Sitamarhi", "Siwan", "Supaul", "Vaishali", "Hajipur", "West Champaran", "Bettiah"
]

def format_currency_inr(val):
    if val is None or val == "" or val <= 0:
        return "As per NIT"
    try:
        val = float(val)
    except:
        return str(val)
    if val >= 10000000:
        return f"₹ {val / 10000000:.2f} Cr"
    elif val >= 100000:
        return f"₹ {val / 100000:.2f} Lakh"
    else:
        return f"₹ {val:,.0f}"

def format_timestamp(ts):
    if not ts:
        return "N/A", 999, "00:00:00"
    try:
        ts = int(ts)
        if ts <= 0:
            return "N/A", 999, "00:00:00"
        dt = datetime.datetime.fromtimestamp(ts / 1000.0)
        now = datetime.datetime.now()
        diff = dt - now
        total_seconds = int(diff.total_seconds())
        days_left = round(total_seconds / 86400.0, 1)

        if total_seconds > 0:
            d = total_seconds // 86400
            rem = total_seconds % 86400
            h = rem // 3600
            m = (rem % 3600) // 60
            time_left_str = f"{d:02d}:{h:02d}:{m:02d}"
        else:
            time_left_str = "00:00:00 (Expired)"

        return dt.strftime("%d-%m-%Y %I:%M %p"), days_left, time_left_str
    except Exception:
        return "N/A", 999, "00:00:00"

def extract_district_and_place(text, dept_name=""):
    combined = f"{text} {dept_name}"
    
    # 1. District matching
    matched_district = "Other / State-wide"
    for d in BIHAR_DISTRICTS:
        if re.search(r'\b' + re.escape(d) + r'\b', combined, re.I):
            matched_district = d.title()
            if matched_district == "Motihari": matched_district = "East Champaran"
            elif matched_district == "Bettiah": matched_district = "West Champaran"
            elif matched_district == "Sasaram": matched_district = "Rohtas"
            elif matched_district == "Chapra": matched_district = "Saran"
            elif matched_district == "Bihar Sharif": matched_district = "Nalanda"
            elif matched_district == "Bhabhua": matched_district = "Kaimur"
            elif matched_district == "Hajipur": matched_district = "Vaishali"
            break

    # 2. Extract Place / Location keywords
    place = "Bihar"
    place_match = re.search(r'(?:at|near|in|block|dist[ -]|district|ward no[\.\- ]*)\s*([A-Za-z0-9\u0900-\u097F\s,\-]{3,40}?)(?:,|\.|\(|\)|dist|\-|\;|\n|$)', text, re.I)
    if place_match:
        cand = place_match.group(1).strip()
        if len(cand) > 3 and not cand.lower().startswith('the'):
            place = cand
    elif matched_district != "Other / State-wide":
        place = matched_district

    return matched_district, place

def classify_category(title, ref_no):
    txt = f"{title} {ref_no}".lower()
    if any(k in txt for k in ['supply', 'equipment', 'instrument', 'purchase', 'procurement of', 'medicines', 'drugs', 'uniform', 'meter', 'pipes', 'laboratory']):
        return "Goods & Supply"
    elif any(k in txt for k in ['outsourcing', 'security', 'housekeeping', 'sweeper', 'manpower', 'hiring of', 'maintenance', 'operation and maintenance', 'consultancy', 'survey', 'o&m']):
        return "Services & Manpower"
    elif any(k in txt for k in ['electrical', 'substation', 'gss', 'transmission', 'transformer', 'wiring', 'cable', 'lighting', 'solar', 'pole', 'feeder']):
        return "Electrical Works"
    elif any(k in txt for k in ['mechanical', 'pump', 'boring', 'motor', 'pipeline', 'tubewell', 'machinery']):
        return "Mechanical Works"
    else:
        return "Civil Works"

def get_auth_session():
    cj = http.cookiejar.CookieJar()
    ssl_ctx = ssl._create_unverified_context()
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(cj),
        urllib.request.HTTPSHandler(context=ssl_ctx)
    )

    auth_token = ""
    # 1. Primary Method: Official eProc2 REST Token Object API (Fast & Reliable)
    try:
        token_req = urllib.request.Request(
            f"{BASE_URL}/rest/login/provideTokenObject",
            data=b"",
            headers={
                'User-Agent': HEADERS['User-Agent'],
                'Content-Type': 'application/json',
                'Referer': f"{BASE_URL}/openarea/tenderListingPage.action"
            }
        )
        with opener.open(token_req, timeout=12) as token_resp:
            token_json = json.loads(token_resp.read().decode('utf-8', errors='replace'))
            auth_token = token_json.get('jwt', '')
    except Exception as e_tok:
        pass

    # 2. Fallback: Parse HTML from listing page if token API failed
    if not auth_token:
        try:
            req_page = urllib.request.Request(f"{BASE_URL}/openarea/tenderListingPage.action", headers=HEADERS)
            with opener.open(req_page, timeout=20) as resp:
                page_html = resp.read().decode('utf-8', errors='ignore')

            token_match = re.search(r'id=["\']Authorization["\'][^>]*value=["\']([^"\']+)["\']', page_html)
            if not token_match or not token_match.group(1):
                token_match = re.search(r'value=["\']([^"\']+)["\'][^>]*id=["\']Authorization["\']', page_html)
            if token_match and token_match.group(1):
                auth_token = token_match.group(1)
        except Exception:
            pass

    if not auth_token:
        raise RuntimeError("eProc2 Bihar portal se active session token prapt nahi ho saka. Kripya thoda intezar karein.")
    return opener, auth_token

def fetch_live_tenders():
    print("[1/4] Establishing session with eproc2.bihar.gov.in...")
    opener, auth_token = get_auth_session()

    api_headers = {
        'User-Agent': HEADERS['User-Agent'],
        'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
        'Content-Type': 'application/json',
        'Authorization': auth_token
    }

    print("[2/4] Fetching department master list...")
    dept_map = {}
    try:
        dept_url = f"{BASE_URL}/rest/organization/loadAllDeptWithParent?orgId=538"
        req_dept = urllib.request.Request(dept_url, data=b"", headers=api_headers)
        with opener.open(req_dept, timeout=20) as resp:
            depts = json.loads(resp.read().decode('utf-8', errors='replace'))
            for d in depts:
                dept_map[d.get('organizationId')] = {
                    'name': d.get('organizationName'),
                    'code': d.get('organizationCode'),
                    'address': d.get('address')
                }
    except Exception as e:
        print(f"      Warning loading departments: {e}")

    print("[3/4] Fetching active tenders list...")
    tenders_url = f"{BASE_URL}/rest/openarea/getTenderList"
    payload = json.dumps({"orgId": "538"}).encode('utf-8')
    req_tenders = urllib.request.Request(tenders_url, data=payload, headers=api_headers)

    try:
        with opener.open(req_tenders, timeout=30) as resp:
            raw_tenders = json.loads(resp.read().decode('utf-8', errors='replace'))
    except Exception as e:
        if hasattr(e, 'partial'):
            raw_tenders = json.loads(e.partial.decode('utf-8', errors='replace'))
        else:
            raise
    print(f"      Successfully retrieved {len(raw_tenders)} raw active tenders.")

    print("[4/4] Smart Delta Sync (System Tender No Se Milan - Preserving Existing Tenders)...")
    view_path = os.path.join(os.path.dirname(__file__), "excel_tenders_view.json")
    existing_list = []
    if os.path.exists(view_path):
        try:
            with open(view_path, "r", encoding="utf-8") as f:
                existing_list = json.load(f).get("tenders", [])
        except Exception as e:
            print(f"      Notice reading existing view: {e}")

    existing_by_sysno = {str(t.get('systemTenderNo')): t for t in existing_list if t.get('systemTenderNo')}
    existing_by_tid = {str(t.get('tenderId')): t for t in existing_list if t.get('tenderId')}

    new_tenders = []
    new_count = 0
    updated_count = 0
    skipped_count = 0

    for idx, t in enumerate(raw_tenders):
        tender_id = t.get('currenttenderid')
        tid_str = str(tender_id)
        rfq_id = t.get('currentOrgTenderId') or tender_id
        sys_no = str(rfq_id)

        matched = existing_by_sysno.get(sys_no) or existing_by_tid.get(tid_str)

        # 1. Existing Tender: Check only for Date Extension or Cost revisions (Skip full re-parse)
        if matched:
            close_date_str, days_left, time_left_str = format_timestamp(t.get('currentbidEndDate'))
            if close_date_str != 'N/A' and close_date_str != matched.get('dueDate'):
                matched['dueDate'] = close_date_str
                matched['daysLeft'] = days_left
                matched['timeLeft'] = time_left_str
                matched['isClosingSoon'] = (0 <= days_left <= 3)
                matched['isClosingThisWeek'] = (0 <= days_left <= 7)
                updated_count += 1

            pac_val = t.get('currentpacamt')
            if pac_val is not None:
                try: p_num = float(pac_val)
                except: p_num = 0.0
                if p_num > 0 and p_num != matched.get('estimatedValue'):
                    matched['estimatedValue'] = p_num
                    matched['estimatedValueFormatted'] = format_currency_inr(p_num)
                    matched['isEstimatedValueHidden'] = False
                    updated_count += 1

            skipped_count += 1
            continue

        # 2. Brand New Tender: Process fully and prepend to master list
        dept_id = t.get('currentdeptid')
        dept_info = dept_map.get(dept_id, {})
        dept_name = dept_info.get('name') or f"Department #{dept_id}"

        pac_val = t.get('currentpacamt')
        if pac_val is not None:
            try: pac_val = float(pac_val)
            except: pac_val = 0.0
        else:
            pac_val = 0.0

        close_date_str, days_left, time_left_str = format_timestamp(t.get('currentbidEndDate'))
        start_date_str, _, _ = format_timestamp(t.get('currentbidStartDate'))
        open_date_str, _, _ = format_timestamp(t.get('currentbidOpenDate'))

        ref_no = t.get('currenttenderrefno') or "N/A"
        desc = (t.get('currentdescription') or "No description provided.").strip()

        district, place = extract_district_and_place(desc, dept_name)
        work_category = classify_category(desc, ref_no)
        approx_emd = round(pac_val * 0.02) if pac_val > 0 else 0

        new_item = {
            "slNo": len(new_tenders) + 1,
            "systemTenderNo": sys_no,
            "tenderId": int(tid_str) if tid_str.isdigit() else tid_str,
            "refNo": ref_no,
            "department": dept_name,
            "district": district,
            "estimatedValue": pac_val if pac_val > 0 else None,
            "estimatedValueFormatted": format_currency_inr(pac_val) if pac_val > 0 else "Hidden by portal",
            "isEstimatedValueHidden": (pac_val <= 0),
            "description": desc,
            "detailedDescription": desc,
            "startDate": start_date_str,
            "dueDate": close_date_str,
            "openDate": open_date_str,
            "emdAmount": approx_emd if approx_emd > 0 else None,
            "emdAmountFormatted": format_currency_inr(approx_emd) if approx_emd > 0 else "As per NIT",
            "emdMode": "Internet Payment Gateway(IPG),Challan|Bank Guarantee",
            "tenderFee": 0.0,
            "tenderFeeFormatted": "As per NIT",
            "tenderFeeMode": "Internet Payment Gateway(IPG),Challan|",
            "processingFee": 0.0,
            "processingFeeFormatted": "As per NIT",
            "processingFeeMode": "Internet Payment Gateway(IPG),Challan|",
            "termsFiles": [],
            "officialFiles": [],
            "boqFiles": [],
            "tenderFolder": os.path.join(os.path.dirname(__file__), "downloads", f"tender_{sys_no}_{tid_str}"),
            "hasDownloaded": False,
            "actionExtracted": False,
            "isNew": True
        }
        new_tenders.append(new_item)
        new_count += 1

    final_tenders = new_tenders + existing_list
    for idx, t_obj in enumerate(final_tenders, 1):
        t_obj['slNo'] = idx

    with open(view_path, "w", encoding="utf-8") as f:
        json.dump({
            "totalTenders": len(final_tenders),
            "tenders": final_tenders,
            "lastSyncStats": {
                "newAdded": new_count,
                "updatedCount": updated_count,
                "skippedCount": skipped_count
            }
        }, f, indent=2, ensure_ascii=False)

    print(f"Done! Smart Delta Sync completed: {new_count} new tenders added, {updated_count} tenders updated with latest dates/PAC, {skipped_count} existing tenders preserved without reload.")
    return final_tenders

def fetch_tender_deep_dive(tender_id):
    """
    Simulates clicking the 'Action' (👁️) button on eproc2.bihar.gov.in.
    Parses EVERY section shown in the official 7-page PDF report:
    1. General Information (System Tender No, Ref No, Category, Hierarchy, Authorities)
    2. Date Schedule (Start, Due Date, Open Date, Physical Doc Date)
    3. Pre-bid Discussion (Type, Dates, Venue, Remarks)
    4. Payment Details (EMD, Tender Fee, Processing Fee, Payment Modes, Exemption)
    5. Required Attachments from Bidder (Eligibility Checklist with 18+ items)
    6. BOQ Items (Item Code, Description, UOM, Quantity, Mandatory)
    7. Official Attachments (PDF / DOCX links)
    """
    opener, auth_token = get_auth_session()
    api_headers = {
        'User-Agent': HEADERS['User-Agent'],
        'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
        'Content-Type': 'application/json',
        'Authorization': auth_token
    }

    url = f"{BASE_URL}/rest/tender/previewTenderByTenderId?tenderId={tender_id}"
    req = urllib.request.Request(url, data=b"", headers=api_headers)
    with opener.open(req, timeout=20) as resp:
        data = json.loads(resp.read().decode('utf-8', errors='replace'))

    tpm = data.get('tenderPreviewMap', {})
    creator_detail = tpm.get('creatorDetail', {})
    org_hierarchy = creator_detail.get('OrgHirarchy', {})
    hierarchy_list = []
    if isinstance(org_hierarchy, dict):
        hierarchy_list = [v for k, v in sorted(org_hierarchy.items(), key=lambda x: str(x[0]))]

    # 1. Date Schedule parsing
    date_schedule = {}
    pre_bid = {
        "type": "N/A",
        "startDate": "N/A",
        "endDate": "N/A",
        "venue": "N/A",
        "remarks": ""
    }
    payments = []
    total_payment_required = 0.0
    emd_amount = 0.0
    tender_fee = 0.0
    processing_fee = 0.0

    required_bidder_attachments = []
    boq_items = []
    official_attachments = []
    general_particulars = ""

    templates = data.get('templates', [])

    for tmpl in templates:
        short = tmpl.get('shortName', '')
        field_dict = {}
        for f in tmpl.get('templateFieldList', []):
            code = f.get('code')
            val = f.get('value')
            if code:
                field_dict[code] = val

        if short == 'rfqdate':
            s_date, _, _ = format_timestamp(field_dict.get('bid_start_date'))
            e_date, days_left, time_left = format_timestamp(field_dict.get('bid_end_date'))
            o_date, _, _ = format_timestamp(field_dict.get('bid_open_date'))
            p_date, _, _ = format_timestamp(field_dict.get('doc_sub_date'))
            date_schedule = {
                "bidStartDate": s_date,
                "bidDueDate": e_date,
                "bidOpenDate": o_date,
                "physicalDocDate": p_date,
                "daysLeft": days_left,
                "timeLeft": time_left
            }

        elif short == 'prebiddiscussion':
            pb_start, _, _ = format_timestamp(field_dict.get('meeting_start_date'))
            pb_end, _, _ = format_timestamp(field_dict.get('meeting_end_date'))
            pre_bid = {
                "type": field_dict.get('discussion_type') or "OFFLINE",
                "startDate": pb_start,
                "endDate": pb_end,
                "venue": field_dict.get('venue') or "Not Specified",
                "remarks": field_dict.get('remarks') or ""
            }

        elif short == 'rfqpayment':
            p_type = field_dict.get('payment_type') or "Payment"
            amt_str = field_dict.get('amount') or "0"
            mode = field_dict.get('payment_mode') or "Internet Payment Gateway(IPG), Challan|Bank Guarantee"
            curr = field_dict.get('payment_currency') or "INR"
            exempt = field_dict.get('exemption_flag') or "N"
            exempt_reason = field_dict.get('supporting_doc') or ""

            amt_val = 0.0
            try:
                amt_val = float(amt_str)
            except:
                pass

            if amt_val > 0:
                total_payment_required += amt_val

            if 'emd' in p_type.lower():
                emd_amount = amt_val
            elif 'processing' in p_type.lower():
                processing_fee = amt_val
            elif 'fee' in p_type.lower():
                tender_fee = amt_val

            payments.append({
                "paymentType": p_type,
                "amount": amt_val,
                "amountFormatted": format_currency_inr(amt_val),
                "paymentMode": mode,
                "currency": curr,
                "exemptionAllowed": exempt,
                "exemptionReason": exempt_reason
            })

        elif short == 'requiredbidderattachment':
            doc_name = field_dict.get('supporting_doc') or ""
            if doc_name:
                required_bidder_attachments.append({
                    "supportingDocument": doc_name.strip(),
                    "isMandatory": (field_dict.get('mandatory_flag') == 'Y'),
                    "allowExemption": field_dict.get('exemption_flag') or 'N',
                    "attachmentGroup": field_dict.get('attachment_group') or 'Others',
                    "evaluationType": field_dict.get('evaluation_type') or 'BidPart1'
                })

        elif short in ['br_rfq_item', 'br_rfq_item_rc'] or 'item' in short:
            item_code = field_dict.get('item_code') or str(len(boq_items) + 1).zfill(2)
            item_desc = field_dict.get('item_name') or field_dict.get('item_description') or "Item"
            uom = field_dict.get('uom') or "No"
            qty = field_dict.get('item_qty') or field_dict.get('quantity') or "1"
            mandatory = field_dict.get('mandatory_item') or "Yes"
            boq_items.append({
                "itemCode": item_code,
                "description": item_desc.strip(),
                "uom": uom,
                "quantity": qty,
                "isMandatory": mandatory
            })

        elif short == 'rfqattachment':
            fname = field_dict.get('file_name') or ""
            attach = field_dict.get('attach_file') or ""
            lbl = field_dict.get('label') or "Attachment"
            if fname:
                official_attachments.append({
                    "label": lbl,
                    "fileName": fname,
                    "attachFile": attach
                })

        elif short == 'general_particulars_rh':
            general_particulars = field_dict.get('unf1') or "Name of bidder, Firm Registration, Contact Details, PAN, GST"

    pac = data.get('pacamt') or 0.0
    try: pac = float(pac)
    except: pac = 0.0

    # Key eligibility highlights extracted
    eligibility_summary = {
        "turnoverRequired": "Not Specified",
        "experienceYears": "Not Specified",
        "bankSolvency": False,
        "warrantyYears": "Not Specified",
        "panGst": True
    }
    for req_att in required_bidder_attachments:
        doc_lower = req_att["supportingDocument"].lower()
        if "turnover" in doc_lower:
            m = re.search(r'turnover\s*(?:at\s*least)?\s*([^\,\;]+)', req_att["supportingDocument"], re.I)
            if m: eligibility_summary["turnoverRequired"] = m.group(1).strip()
        if "experience" in doc_lower:
            m = re.search(r'(\d+)\s*years?\s*of\s*experience', req_att["supportingDocument"], re.I)
            if m: eligibility_summary["experienceYears"] = f"{m.group(1)} Years"
        if "solvency" in doc_lower:
            eligibility_summary["bankSolvency"] = True
        if "warranty" in doc_lower:
            m = re.search(r'(\w+)\s*years?\s*(?:onsite)?\s*warranty', req_att["supportingDocument"], re.I)
            if m: eligibility_summary["warrantyYears"] = f"{m.group(1)} Years"

    return {
        "tenderId": data.get('tenderid'),
        "systemTenderNo": data.get('tenderid') or data.get('orgtenderid'),
        "tenderRefNo": data.get('tenderrefno'),
        "detailedDescription": data.get('description'),
        "nitTitle": data.get('nit') or data.get('description'),
        "templateGroup": tpm.get('templateGroupId'),
        "procurementCategory": tpm.get('procCat') or "GENERAL",
        "categoryDescription": tpm.get('catagory', {}).get('categoryDescription'),
        "tenderType": tpm.get('tenderType'),
        "rankingSequence": data.get('rankingsequence') or "L1 Ranking",
        "offerValidityDays": data.get('offerValidity') or 180,
        "minBids": data.get('minbidno') or 2,
        "bidParts": data.get('bidPartNo') or 2,
        "tenderCurrency": data.get('tendercurrency') or "INR",
        "estimatedPacAmount": pac,
        "estimatedPacFormatted": format_currency_inr(pac),
        "totalPaymentRequired": total_payment_required,
        "totalPaymentFormatted": format_currency_inr(total_payment_required),
        "emdAmount": emd_amount,
        "emdFormatted": format_currency_inr(emd_amount),
        "tenderFee": tender_fee,
        "tenderFeeFormatted": format_currency_inr(tender_fee),
        "processingFee": processing_fee,
        "processingFeeFormatted": format_currency_inr(processing_fee),
        "hierarchy": hierarchy_list,
        "issuingAuthority": str(data.get('tenderIssuingAuthority') or creator_detail.get('DealingOfficer', {}).get('officerName') or ""),
        "dateSchedule": date_schedule,
        "preBidDiscussion": pre_bid,
        "payments": payments,
        "requiredBidderAttachments": required_bidder_attachments,
        "totalRequiredDocuments": len(required_bidder_attachments),
        "eligibilitySummary": eligibility_summary,
        "boqItems": boq_items,
        "totalBoqItems": len(boq_items),
        "attachments": official_attachments,
        "generalParticulars": general_particulars
    }

if __name__ == "__main__":
    fetch_live_tenders()
