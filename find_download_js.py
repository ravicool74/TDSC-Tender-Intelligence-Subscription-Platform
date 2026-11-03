from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import re

opener, auth = get_auth_session()
api_headers = {'User-Agent': HEADERS['User-Agent'], 'Authorization': auth}

# Let's inspect the page HTML of tenderListingPage.action and preview pages
req = urllib.request.Request(f"{BASE_URL}/openarea/tenderListingPage.action", headers=HEADERS)
with opener.open(req) as resp:
    html = resp.read().decode('utf-8', errors='ignore')

scripts = re.findall(r'src=["\']([^"\']+)["\']', html)
print("Total scripts in listing page:", len(scripts))

for s in scripts:
    if not s.startswith('http'):
        s_url = f"{BASE_URL}/{s}" if not s.startswith('/') else f"https://eproc2.bihar.gov.in{s}"
    else:
        s_url = s
    try:
        r = opener.open(urllib.request.Request(s_url, headers=HEADERS), timeout=10)
        js = r.read().decode('utf-8', errors='ignore')
        if 'download' in js.lower() or 'attach_file' in js.lower() or 'rfqattachment' in js.lower():
            # Look for functions
            matches = re.findall(r'(\b\w*download\w*\s*[:=]\s*function[^{]*\{[^}]{0,200})', js, re.I)
            if matches:
                print(f"Match in {s}:")
                for m in matches[:5]:
                    print("  ", m.strip().replace('\n', ' '))
            urls = re.findall(r'[\'"]([^\'"]*download[^\'"]*)[\'"]', js, re.I)
            if urls:
                print(f"Urls in {s}:", set(urls))
    except Exception as e:
        pass
