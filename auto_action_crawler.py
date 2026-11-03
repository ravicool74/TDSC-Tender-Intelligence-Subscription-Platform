import urllib.request
import http.cookiejar
import json
import re
import concurrent.futures
import time
import os

BASE_URL = "https://eproc2.bihar.gov.in/EPSV2Web"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

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
        return "N/A"
    try:
        ts = int(ts)
        if ts <= 0: return "N/A"
        dt = datetime.datetime.fromtimestamp(ts / 1000.0)
        return dt.strftime("%d-%m-%Y %I:%M %p")
    except:
        return "N/A"

def run_auto_action_crawler(max_tenders=100):
    print(f"[Crawler] Starting automated Action-button click extraction for top {max_tenders} tenders...")
    
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    req_page = urllib.request.Request(f"{BASE_URL}/openarea/tenderListingPage.action", data=b"", headers=HEADERS)
    with opener.open(req_page, timeout=15) as r:
        page_html = r.read().decode('utf-8')

    token_match = re.search(r'id=["\']Authorization["\'][^>]*value=["\']([^"\']+)["\']', page_html)
    auth_token = token_match.group(1) if token_match else ""

    api_headers = {
        'User-Agent': HEADERS['User-Agent'],
        'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
        'Content-Type': 'application/json',
        'Authorization': auth_token
    }

    data_file = "tenders_data.json"
    with open(data_file, "r", encoding="utf-8") as f:
        stored = json.load(f)
    tenders = stored["tenders"]

    # Target first N tenders
    target_tenders = tenders[:max_tenders]

    def fetch_action_page(t):
        tid = t['id']
        url = f"{BASE_URL}/rest/tender/previewTenderByTenderId?tenderId={tid}"
        try:
            req = urllib.request.Request(url, data=b"", headers=api_headers)
            with opener.open(req, timeout=15) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                
                emd_val = 0.0
                tender_fee = 0.0
                processing_fee = 0.0
                prebid_venue = "N/A"
                prebid_type = "N/A"
                turnover_req = "As per NIT"
                exp_req = "As per NIT"
                solvency_req = False
                boq_count = 0
                boq_sample = []
                attachments = []
                hierarchy = []

                tpm = data.get('tenderPreviewMap', {})
                creator = tpm.get('creatorDetail', {})
                org_h = creator.get('OrgHirarchy', {})
                if isinstance(org_h, dict):
                    hierarchy = [v for k, v in sorted(org_h.items(), key=lambda x: str(x[0]))]

                for tmpl in data.get('templates', []):
                    short = tmpl.get('shortName', '')
                    fdict = {f.get('code'): f.get('value') for f in tmpl.get('templateFieldList', []) if f.get('code')}

                    if short == 'rfqpayment':
                        ptype = str(fdict.get('payment_type') or '').lower()
                        amt = fdict.get('amount')
                        if 'emd' in ptype and amt:
                            try: emd_val = float(amt)
                            except: pass
                        elif 'processing' in ptype and amt:
                            try: processing_fee = float(amt)
                            except: pass
                        elif 'fee' in ptype and amt:
                            try: tender_fee = float(amt)
                            except: pass

                    elif short == 'prebiddiscussion':
                        prebid_venue = fdict.get('venue') or "N/A"
                        prebid_type = fdict.get('discussion_type') or "N/A"

                    elif short == 'requiredbidderattachment':
                        sdoc = str(fdict.get('supporting_doc') or '')
                        sdoc_low = sdoc.lower()
                        if 'turnover' in sdoc_low and turnover_req == "As per NIT":
                            turnover_req = sdoc.strip()
                        if 'experience' in sdoc_low and exp_req == "As per NIT":
                            exp_req = sdoc.strip()
                        if 'solvency' in sdoc_low:
                            solvency_req = True

                    elif 'item' in short:
                        boq_count += 1
                        iname = fdict.get('item_name') or fdict.get('item_description')
                        if iname and len(boq_sample) < 5:
                            boq_sample.append(iname.strip())

                    elif short == 'rfqattachment':
                        fn = fdict.get('file_name')
                        if fn: attachments.append(fn)

                total_outflow = emd_val + tender_fee + processing_fee
                if emd_val > 0:
                    t['emdAmount'] = emd_val
                    t['emdAmountFormatted'] = format_currency_inr(emd_val)
                if tender_fee > 0:
                    t['tenderFee'] = tender_fee
                    t['tenderFeeFormatted'] = format_currency_inr(tender_fee)
                if total_outflow > 0:
                    t['totalOutflow'] = total_outflow
                    t['totalOutflowFormatted'] = format_currency_inr(total_outflow)

                t['prebidVenue'] = prebid_venue
                t['prebidType'] = prebid_type
                t['turnoverRequirement'] = turnover_req
                t['experienceRequirement'] = exp_req
                t['bankSolvencyRequired'] = solvency_req
                t['boqItemCount'] = boq_count
                t['boqSampleItems'] = boq_sample
                t['attachmentsList'] = attachments
                if hierarchy:
                    t['orgHierarchy'] = " -> ".join(hierarchy)
                t['isActionCrawled'] = True
        except Exception as e:
            t['isActionCrawled'] = False

    t0 = time.time()
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        list(executor.map(fetch_action_page, target_tenders))
    t1 = time.time()
    print(f"[Crawler] Completed automated Action-clicks on {len(target_tenders)} tenders in {t1 - t0:.2f}s!")

    stored["lastActionCrawl"] = time.strftime("%d-%b-%Y %I:%M %p")
    with open(data_file, "w", encoding="utf-8") as f:
        json.dump(stored, f, indent=2, ensure_ascii=False)
    print(f"[Crawler] Saved enriched data back to {data_file}")

if __name__ == "__main__":
    run_auto_action_crawler(100)
