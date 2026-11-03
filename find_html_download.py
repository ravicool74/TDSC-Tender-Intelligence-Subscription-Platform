from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import re

opener, auth = get_auth_session()

url = f"{BASE_URL}/openarea/tenderListingPage.action"
req = urllib.request.Request(url, headers=HEADERS)
with opener.open(req) as resp:
    html = resp.read().decode('utf-8', errors='ignore')

matches = re.findall(r'downloadFile\([^)]+\)', html)
print("Matches in listing page:", matches)

# Now check all html/template files referenced in JS
url_js = f"{BASE_URL}/js/tenderModule/openareaTenderListController.js?version=15.1.241"
with opener.open(urllib.request.Request(url_js, headers=HEADERS)) as resp:
    js = resp.read().decode('utf-8', errors='ignore')

matches_js = re.findall(r'downloadFile\([^)]+\)', js)
print("Matches in JS:", matches_js)

# Also check template html files
tmpl_urls = re.findall(r'[\'"]([^\'"]+\.html)[\'"]', js)
print("HTML templates:", set(tmpl_urls))
