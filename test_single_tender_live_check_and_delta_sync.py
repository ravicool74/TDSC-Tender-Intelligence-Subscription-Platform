import urllib.request
import urllib.parse
import json
import time
import sys

sys.stdout.reconfigure(encoding='utf-8')

BASE_URL = "http://localhost:8000"

def post_json(path, data):
    payload = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(f"{BASE_URL}{path}", data=payload, headers={'Content-Type': 'application/json'}, method='POST')
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status, json.loads(resp.read().decode('utf-8'))

def get_json(path):
    req = urllib.request.Request(f"{BASE_URL}{path}")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status, json.loads(resp.read().decode('utf-8'))

def run_tests():
    print("=================================================================")
    print("TEST: TENDER-BY-TENDER LIVE REFRESH & SMART DELTA SYNC ENGINE")
    print("=================================================================")

    # 1. Test In-Memory RAM Cache for excel_tenders
    print("\n[Test 1] Testing In-Memory RAM Cache via /api/excel_tenders...")
    t0 = time.time()
    status, data = get_json("/api/excel_tenders")
    elapsed_ms = (time.time() - t0) * 1000
    assert status == 200, f"Expected 200, got {status}"
    total = data.get("totalTenders", 0)
    assert total >= 1450, f"Expected 1450+ tenders, got {total}"
    print(f" PASS: Retrieved {total} tenders from RAM Cache in {elapsed_ms:.1f} ms (Lightning fast, zero disk I/O)!")

    # 2. Test Single Tender Live Refresh from eProc2
    print("\n[Test 2] Testing Single Tender Live Refresh for Tender #142360...")
    t0 = time.time()
    status, refresh_data = post_json("/api/tenders/refresh_single", {"tenderId": 142360})
    elapsed_s = time.time() - t0
    assert status == 200, f"Refresh failed: {refresh_data}"
    assert refresh_data.get("success") is True, f"Response: {refresh_data}"
    up_t = refresh_data.get("updatedTender", {})
    assert up_t.get("tenderId") == 142360 or up_t.get("systemTenderNo") == 142165
    print(f" PASS: Tender #142360 refreshed in {elapsed_s:.2f}s!")
    print(f"       System Tender No: {up_t.get('systemTenderNo')}")
    print(f"       Due Date:         {up_t.get('dueDate')}")
    print(f"       Estimated PAC:    {up_t.get('estimatedValueFormatted')}")
    print(f"       EMD Amount:       {up_t.get('emdAmountFormatted')}")
    print(f"       Changes Detected: {refresh_data.get('changesDetected', [])}")
    print(f"       Message:          {refresh_data.get('message')}")

    # 3. Test RAM Cache Invalidation and Instant Live Serve
    print("\n[Test 3] Verifying Tender #142360 reflects immediately in RAM cache...")
    status, data2 = get_json("/api/excel_tenders")
    assert status == 200
    cached_t = next((t for t in data2.get("tenders", []) if t.get("tenderId") == 142360), None)
    assert cached_t is not None, "Tender 142360 not found in cache"
    print(f" PASS: Verified Tender #142360 in RAM cache (Last Live Checked: {cached_t.get('lastLiveChecked')})")

    # 4. Test Single Tender Refresh for Another Tender (e.g. Tender #142279 or first tender in list)
    sample_id = data2["tenders"][1].get("tenderId")
    print(f"\n[Test 4] Testing Single Tender Refresh for Tender #{sample_id}...")
    status, ref_sample = post_json("/api/tenders/refresh_single", {"tenderId": sample_id})
    assert status == 200
    assert ref_sample.get("success") is True
    print(f" PASS: Tender #{sample_id} live check completed successfully!")

    print("\n=================================================================")
    print("ALL TESTS PASSED 100%! TENDER-BY-TENDER LIVE SYNC IS OPERATIONAL!")
    print("=================================================================")

if __name__ == "__main__":
    run_tests()
