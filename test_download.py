from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import re
import json

opener, auth = get_auth_session()
api_headers = {
    'User-Agent': HEADERS['User-Agent'],
    'Referer': f'{BASE_URL}/openarea/tenderListingPage.action',
    'Authorization': auth
}

req = urllib.request.Request(f'{BASE_URL}/openarea/tenderListingPage.action', headers=HEADERS)
with opener.open(req) as resp:
    html = resp.read().decode('utf-8', errors='ignore')

js_files = re.findall(r'src=["\']([^"\']+\.js[^"\']*)["\']', html)
print('JS files:', len(js_files))

# Search inside JS files for download functions
for js in js_files:
    if not js.startswith('http'):
        if js.startswith('/'):
            js_url = f"https://eproc2.bihar.gov.in{js}"
        else:
            js_url = f"{BASE_URL}/{js}"
    else:
        js_url = js
    
    try:
        j_req = urllib.request.Request(js_url, headers=HEADERS)
        with opener.open(j_req, timeout=10) as r:
            js_text = r.read().decode('utf-8', errors='ignore')
            matches = re.findall(r'(\b\w*download\w*\b|\b\w*fileDownload\w*\b|\b\w*docDownload\w*\b)', js_text, re.I)
            if matches:
                endpoints = re.findall(r'["\'](/EPSV2Web/[^"\']+)["\']', js_text)
                ep_down = [ep for ep in endpoints if 'download' in ep.lower() or 'file' in ep.lower() or 'doc' in ep.lower()]
                if ep_down:
                    print(f"File {js}: download endpoints:", set(ep_down))
    except Exception as e:
        pass
