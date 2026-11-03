from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import urllib.parse
import json

opener, auth = get_auth_session()
api_headers = {
    'User-Agent': HEADERS['User-Agent'],
    'Referer': f'{BASE_URL}/openarea/tenderListingPage.action',
    'Authorization': auth
}

file_arg = "33396/1792838907769_NIT DOCUMENT.pdf"
encoded_file = urllib.parse.quote(file_arg)

candidates = [
    f"{BASE_URL}/rest/common/downloadFile?filePath={encoded_file}&fileName=NIT_DOCUMENT.pdf",
    f"{BASE_URL}/rest/common/downloadDocument?filePath={encoded_file}",
    f"{BASE_URL}/rest/tender/downloadAttachment?filePath={encoded_file}",
    f"{BASE_URL}/openarea/downloadDocument.action?filePath={encoded_file}",
    f"{BASE_URL}/openarea/downloadAttachment.action?filePath={encoded_file}",
    f"{BASE_URL}/rest/openarea/downloadFile?filePath={encoded_file}",
    f"{BASE_URL}/rest/file/download?fileName={encoded_file}",
    f"{BASE_URL}/downloadAttachment.action?filePath={encoded_file}",
    f"{BASE_URL}/common/downloadDocument.action?fileName={encoded_file}"
]

for url in candidates:
    try:
        req = urllib.request.Request(url, headers=api_headers)
        with opener.open(req, timeout=5) as resp:
            content_type = resp.headers.get('content-type', '')
            content_disp = resp.headers.get('content-disposition', '')
            body = resp.read()
            print(f"[FOUND!] URL: {url} -> Status {resp.status}, Content-Type: {content_type}, Length: {len(body)}")
            if len(body) > 100:
                print("First 50 bytes:", body[:50])
                break
    except urllib.error.HTTPError as e:
        print(f"HTTP {e.code} for {url.split('?')[0]}")
    except Exception as e:
        print(f"Error {e} for {url.split('?')[0]}")
