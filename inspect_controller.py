from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import re

opener, auth = get_auth_session()

url = f"{BASE_URL}/js/tenderModule/openareaTenderListController.js?version=15.1.241"
req = urllib.request.Request(url, headers=HEADERS)
with opener.open(req) as resp:
    js = resp.read().decode('utf-8', errors='ignore')

# find downloadFile implementation
pos = js.find('downloadFile = function')
if pos != -1:
    print("=== downloadFile in openareaTenderListController ===")
    print(js[pos:pos+1500])

# Also check commonService or other services
pos2 = js.find('downloadTenderAttachment')
if pos2 != -1:
    print("=== downloadTenderAttachment ===")
    print(js[pos2:pos2+1000])
