from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import re

opener, auth = get_auth_session()
req = urllib.request.Request(f"{BASE_URL}/openarea/tenderListingPage.action", headers=HEADERS)
with opener.open(req) as resp:
    html = resp.read().decode('utf-8', errors='ignore')

scripts = re.findall(r'<script[^>]+src=[\'"]([^\'"]+)[\'"]', html)
for s in scripts:
    print(s)
