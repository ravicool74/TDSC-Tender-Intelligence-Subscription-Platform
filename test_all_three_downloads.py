from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import urllib.parse

opener, auth = get_auth_session()
download_url = f"{BASE_URL}/rest/openarea/downloadattachment.action"

tests = [
    ('33396', '33396/1792838901230_BOQ.pdf|94344|BOQ.pdf', 'BOQ_from_rfqattachment.pdf'),
    ('33414', '33414/1792839131765_BOQ.pdf|BOQ.pdf', 'BOQ_from_item.pdf'),
    ('33399', '33399/1792838816856_NIT DOCUMENT.pdf|NIT DOCUMENT.pdf', 'Terms_and_Conditions.pdf')
]

for grp_id, fn, out_name in tests:
    form_data = {
        'Authorization': auth,
        'tempGrpId': grp_id,
        'tenderId': '142360',
        'fileName': fn,
        'orgId': '538'
    }
    encoded_data = urllib.parse.urlencode(form_data).encode('utf-8')
    headers = {
        'User-Agent': HEADERS['User-Agent'],
        'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    req = urllib.request.Request(download_url, data=encoded_data, headers=headers)
    with opener.open(req, timeout=15) as resp:
        data = resp.read()
        print(f"Downloaded {out_name}: {len(data)} bytes, starts with: {data[:10]}")
        with open(f"c:\\e-tender\\{out_name}", "wb") as f:
            f.write(data)
print("All 3 types downloaded successfully!")
