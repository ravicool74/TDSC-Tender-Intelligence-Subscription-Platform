import urllib.request
import urllib.parse
import http.cookiejar
import json
import time
import sys

BASE_URL = "http://localhost:8000"

class TestClient:
    def __init__(self):
        self.cj = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cj))

    def get(self, path):
        req = urllib.request.Request(f"{BASE_URL}{path}")
        try:
            with self.opener.open(req) as resp:
                data = resp.read().decode('utf-8')
                return resp.status, json.loads(data)
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read().decode('utf-8'))

    def post(self, path, payload):
        data_bytes = json.dumps(payload).encode('utf-8')
        req = urllib.request.Request(
            f"{BASE_URL}{path}",
            data=data_bytes,
            headers={'Content-Type': 'application/json'}
        )
        try:
            with self.opener.open(req) as resp:
                data = resp.read().decode('utf-8')
                return resp.status, json.loads(data)
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read().decode('utf-8'))

def run_tests():
    print("==================================================")
    print("TDSC EMPLOYEE LOGIN & CONTRACTOR CRM TEST SUITE")
    print("==================================================")

    client = TestClient()

    # 1. Employee Login
    print("\n[Test 1] Logging in as Employee (employee@tdsc.in)...")
    status, data = client.post("/api/auth/login", {
        "email": "employee@tdsc.in",
        "password": "Employee@Tdsc2026"
    })
    assert status == 200, f"Login failed (status {status}): {data}"
    assert data.get("success") is True, f"Login response unsuccessful: {data}"
    user = data.get("user")
    assert user.get("role") == "employee", f"Expected role 'employee', got: {user.get('role')}"
    print(f" PASS: Logged in successfully as {user.get('name')} (Role: {user.get('role')})")

    # 2. Check Session & Staff Entitlements
    print("\n[Test 2] Verifying Employee Entitlements via /api/auth/me...")
    status, me_data = client.get("/api/auth/me")
    assert status == 200
    assert me_data.get("authenticated") is True
    sub = me_data["user"].get("subscription", {})
    assert sub.get("isPaidActive") is True, "Employee should have active staff entitlements"
    print(f" PASS: Employee has unrestricted staff entitlements (Plan: {sub.get('planName')})")

    # 3. Employee Tender Search & Full Access
    print("\n[Test 3] Testing Employee Full Tender Access via /api/excel_tenders...")
    status, tenders_data = client.get("/api/excel_tenders")
    assert status == 200, f"Failed with status {status}: {tenders_data}"
    tenders = tenders_data.get("tenders", [])
    assert len(tenders) > 0, "No tenders returned"
    for t in tenders[:10]:
        assert t.get("isLocked") is False, "Staff should have no locked tenders"
    print(f" PASS: Employee can access full unmasked tender data ({len(tenders)} total tenders in portal)")

    # 4. Contractor CRM Directory
    print("\n[Test 4] Fetching Contractors Directory via /api/contractors...")
    status, c_data = client.get("/api/contractors")
    assert status == 200
    assert c_data.get("success") is True
    contractors = c_data.get("contractors", [])
    assert len(contractors) >= 10, f"Expected at least 10 contractors, got {len(contractors)}"
    print(f" PASS: Retrieved {len(contractors)} contractors across categories: {c_data.get('categories')}")

    # 5. Category-wise Filtering
    print("\n[Test 5] Testing Category-wise Filtering (Roads & Highways)...")
    status, roads_data = client.get("/api/contractors?category=Roads+%26+Highways")
    assert status == 200
    roads_contractors = roads_data.get("contractors", [])
    assert len(roads_contractors) > 0, "No road contractors found"
    for rc in roads_contractors:
        assert rc["category"] == "Roads & Highways"
    print(f" PASS: Category filter works accurately ({len(roads_contractors)} Roads & Highways contractors)")

    # 6. TDSC Office Management System Sync
    print("\n[Test 6] Syncing with TDSC Office Management System via /api/contractors/sync_tdsc_office...")
    status, sync_data = client.post("/api/contractors/sync_tdsc_office", {})
    assert status == 200
    assert sync_data.get("success") is True
    print(f" PASS: {sync_data.get('message')}")

    # 7. Add Custom Contractor
    print("\n[Test 7] Adding a New Contractor via /api/contractors/create...")
    new_con_payload = {
        "companyName": "Mithila Smart Infra Ltd",
        "contactPerson": "Suresh Jha",
        "mobile": "9835123456",
        "email": "suresh@mithilainfra.in",
        "district": "Darbhanga",
        "category": "Building Construction",
        "workClass": "Class 1",
        "budgetSlab": "2Crto5Cr",
        "notes": "Experienced in institutional & residential schools building."
    }
    status, created_c = client.post("/api/contractors/create", new_con_payload)
    assert status == 200
    assert created_c.get("success") is True
    created_id = created_c.get("contractorId")
    print(f" PASS: Contractor created with ID {created_id} ({created_c.get('message')})")

    # 8. Tender Matcher Engine (Mode A: Tenders for Contractor)
    print("\n[Test 8] Running Smart Matcher Engine for Contractor #1...")
    status, match_data = client.get("/api/contractors/match_tenders?contractorId=1")
    assert status == 200
    assert match_data.get("success") is True
    matched_tenders = match_data.get("matchedTenders", [])
    print(f" PASS: Matched {len(matched_tenders)} high-affinity tenders for contractor '{match_data['contractor']['company_name']}'")

    # 9. Tender Matcher Engine (Mode B: Contractors for Tender)
    print("\n[Test 9] Running Smart Matcher Engine for Tender...")
    sample_tender = tenders[0]
    sample_tid = sample_tender.get("systemTenderNo") or sample_tender.get("tenderId")
    status, t_match_data = client.get(f"/api/contractors/match_tenders?tenderId={sample_tid}")
    assert status == 200
    assert t_match_data.get("success") is True
    matched_c = t_match_data.get("matchedContractors", [])
    print(f" PASS: Matched {len(matched_c)} specialized contractors for Tender #{sample_tid}")

    # 10. Tender Suggestion & WhatsApp Dispatch
    print("\n[Test 10] Dispatching Tender Suggestion via /api/contractors/suggest_tender...")
    status, sugg_data = client.post("/api/contractors/suggest_tender", {
        "contractorId": 1,
        "tenderId": sample_tid,
        "channel": "whatsapp",
        "notes": "Urgent: Bid due date is approaching."
    })
    assert status == 200
    assert sugg_data.get("success") is True
    assert "whatsappUrl" in sugg_data
    assert "api.whatsapp.com/send" in sugg_data["whatsappUrl"]
    print(f" PASS: Tender suggested! Generated WhatsApp Link: {sugg_data['whatsappUrl'][:75]}...")

    # 11. Suggestions Log Verification
    print("\n[Test 11] Checking Suggestions History Log via /api/contractors/suggestions_log...")
    status, log_data = client.get("/api/contractors/suggestions_log")
    assert status == 200
    assert log_data.get("success") is True
    assert len(log_data.get("logs", [])) > 0
    print(f" PASS: Suggestions log verified ({len(log_data['logs'])} records found)")

    # 12. Admin Adds New Employee
    print("\n[Test 12] Admin creates new Employee via /api/admin/employees/create...")
    admin_client = TestClient()
    adm_status, login_adm = admin_client.post("/api/auth/login", {
        "email": "admin@bihar-etender.in",
        "password": "Admin@Eproc2026"
    })
    assert adm_status == 200
    emp_email = f"priya_{int(time.time())}@tdsc.in"
    status, adm_emp_data = admin_client.post("/api/admin/employees/create", {
        "name": "Priya Sharma",
        "email": emp_email,
        "password": "Priya@Tdsc2026",
        "mobile": "9876543211",
        "department": "Tender Desk Executive"
    })
    assert status == 200, f"Expected 200, got {status}: {adm_emp_data}"
    assert adm_emp_data.get("success") is True
    print(f" PASS: Admin created employee '{adm_emp_data.get('message')}'")

    # 13. New Employee Login
    print("\n[Test 13] Verifying newly created employee login...")
    new_emp_client = TestClient()
    status, new_res = new_emp_client.post("/api/auth/login", {
        "email": emp_email,
        "password": "Priya@Tdsc2026"
    })
    assert status == 200
    assert new_res.get("user", {}).get("role") == "employee"
    print(" PASS: Newly created employee logged in successfully!")

    print("\n==================================================")
    print("ALL 13 TESTS PASSED SUCCESSFULLY! 100% VERIFIED!")
    print("==================================================")

if __name__ == "__main__":
    run_tests()
