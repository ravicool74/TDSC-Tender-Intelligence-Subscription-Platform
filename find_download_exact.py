from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import re

opener, auth = get_auth_session()

# Search in openareaTenderListController.js
url = f"{BASE_URL}/js/tenderModule/openareaTenderListController.js?version=15.1.241"
req = urllib.request.Request(url, headers=HEADERS)
with opener.open(req) as resp:
    js = resp.read().decode('utf-8', errors='ignore')

matches = [m.start() for m in re.finditer(r'downloadattachment', js, re.I)]
for m in matches:
    print("Match at", m)
    print(js[max(0, m-200):min(len(js), m+300)])
    print("="*60)
