from fetch_eproc_tenders import get_auth_session, BASE_URL, HEADERS
import urllib.request
import urllib.parse
import json

opener, auth = get_auth_session()

# First get preview to get templateGroupId and attachments
preview_url = f"{BASE_URL}/rest/tender/previewTenderByTenderId?tenderId=142360"
api_headers = {
    'User-Agent': HEADERS['User-Agent'],
    'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
    'Content-Type': 'application/json',
    'Authorization': auth
}
req = urllib.request.Request(preview_url, data=b"", headers=api_headers)
with opener.open(req) as resp:
    data = json.loads(resp.read().decode('utf-8'))

template_group_id = data.get('tenderPreviewMap', {}).get('templateGroupId')
print(f"TemplateGroupId: {template_group_id}")

attachments = []
for t in data.get('templates', []):
    fields = {f.get('code'): f.get('value') for f in t.get('templateFieldList', [])}
    if t.get('shortName') == 'rfqattachment':
        attachments.append(fields)
    elif t.get('shortName') == 'br_rfq_item_rc' and fields.get('unf3'):
        attachments.append({'label': 'BOQ Item File', 'attach_file': fields.get('unf3'), 'file_name': 'BOQ.pdf'})
    elif t.get('shortName') == 'br_rfq_technical' and fields.get('unf3'):
        attachments.append({'label': 'Technical Terms', 'attach_file': fields.get('unf3'), 'file_name': 'NIT DOCUMENT.pdf'})

print(f"Found {len(attachments)} attachments:")
for a in attachments:
    print(a)

# Now test download attachment
download_url = f"{BASE_URL}/rest/openarea/downloadattachment.action"
for a in attachments:
    fn = a.get('attach_file') or a.get('file_name')
    # Try different variants of fileName parameter
    for fn_candidate in [fn, a.get('file_name'), fn.split('|')[0] if '|' in fn else fn]:
        form_data = {
            'Authorization': auth,
            'tempGrpId': str(template_group_id),
            'tenderId': '142360',
            'fileName': fn_candidate,
            'orgId': '538'
        }
        encoded_data = urllib.parse.urlencode(form_data).encode('utf-8')
        d_headers = {
            'User-Agent': HEADERS['User-Agent'],
            'Referer': f"{BASE_URL}/openarea/tenderListingPage.action",
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        try:
            d_req = urllib.request.Request(download_url, data=encoded_data, headers=d_headers)
            with opener.open(d_req, timeout=10) as d_resp:
                ct = d_resp.headers.get('content-type', '')
                cd = d_resp.headers.get('content-disposition', '')
                content = d_resp.read()
                print(f"Candidate '{fn_candidate[:30]}' -> Status {d_resp.status}, CT: {ct}, CD: {cd}, Bytes: {len(content)}")
                if len(content) > 2000 and ('pdf' in ct.lower() or 'octet-stream' in ct.lower() or content.startswith(b'%PDF')):
                    print(f"SUCCESS! Downloaded valid PDF ({len(content)} bytes) using candidate '{fn_candidate}'")
                    break
        except Exception as e:
            print(f"Failed candidate '{fn_candidate}': {e}")
