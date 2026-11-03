from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import urllib.parse
import os

opener, auth = get_auth_session()
download_url = f"{BASE_URL}/rest/openarea/downloadattachment.action"

form_data = {
    'Authorization': auth,
    'tempGrpId': '33396',
    'tenderId': '142360',
    'fileName': '33396/1792838907769_NIT DOCUMENT.pdf|146953|NIT DOCUMENT.pdf',
    'orgId': '538'
}
encoded_data = urllib.parse.urlencode(form_data).encode('utf-8')
headers = {
    'User-Agent': HEADERS['User-Agent'],
    'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
    'Content-Type': 'application/x-www-form-urlencoded'
}

req = urllib.request.Request(download_url, data=encoded_data, headers=headers)
try:
    with opener.open(req, timeout=15) as resp:
        print(f"Status: {resp.status}")
        print("Headers:", resp.headers)
        data = resp.read()
        print(f"Downloaded bytes: {len(data)}")
        if len(data) > 0:
            print("First 20 bytes:", data[:20])
            save_path = r"c:\e-tender\test_downloaded_nit.pdf"
            with open(save_path, "wb") as f:
                f.write(data)
            print(f"File saved to {save_path}!")
except urllib.error.HTTPError as e:
    print(f"HTTP Error {e.code}: {e.read().decode('utf-8', errors='ignore')}")
except Exception as e:
    print(f"Error: {e}")
