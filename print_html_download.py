from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import re

opener, auth = get_auth_session()
url = f"{BASE_URL}/openarea/tenderListingPage.action"
req = urllib.request.Request(url, headers=HEADERS)
with opener.open(req) as resp:
    html = resp.read().decode('utf-8', errors='ignore')

matches = [m.start() for m in re.finditer(r'downloadFile\(', html)]
for m in matches:
    print("HTML around downloadFile:")
    print(html[max(0, m-250):min(len(html), m+350)])
    print("="*60)
