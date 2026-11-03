"""
Verification Test for Plan-Level "Hide/Unhide" Switch Feature:
Tests:
1. Admin authentication.
2. Toggle Hide/Unhide for a plan (1st click -> Hides plan, is_active=0).
3. Confirm that Hidden plan is excluded from public customer-facing /api/plans.
4. Toggle Hide/Unhide again (2nd click -> Unhides plan, is_active=1).
5. Confirm that Unhidden plan is restored and visible to customers in /api/plans.
6. Verify security access control (non-admin blocked with HTTP 403).
"""

import urllib.request
import urllib.parse
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

BASE_URL = "http://localhost:8000"

def request_json(url, method="GET", data=None, token=None):
    headers = {'Content-Type': 'application/json'}
    if token:
        headers['Authorization'] = f"Bearer {token}"
    body = json.dumps(data).encode('utf-8') if data is not None else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as resp:
            content = resp.read().decode('utf-8', errors='replace')
            try: return resp.status, json.loads(content)
            except: return resp.status, content
    except urllib.error.HTTPError as e:
        content = e.read().decode('utf-8', errors='replace')
        try: return e.code, json.loads(content)
        except: return e.code, content

def run_plan_hide_unhide_tests():
    print("=" * 75)
    print("TESTING EACH-PLAN 'HIDE/UNHIDE' SWITCH (SUBSCRIPTION PLANS & DURATION MANAGER)")
    print("=" * 75)

    # 1. Admin Login
    print("\n[Step 1] Admin Authentication...")
    code, login = request_json(f"{BASE_URL}/api/auth/login", "POST", {
        "email": "admin@bihar-etender.in",
        "password": "Admin@Eproc2026"
    })
    assert code == 200 and login.get('token'), f"Admin login failed: {login}"
    admin_token = login['token']
    print(" [OK] Admin authenticated successfully.")

    # 2. Get Admin Plans
    code, admin_plans_res = request_json(f"{BASE_URL}/api/admin/plans", "GET", token=admin_token)
    assert code == 200 and admin_plans_res.get('plans'), f"Failed to fetch admin plans: {admin_plans_res}"
    plans = admin_plans_res['plans']
    starter_plan = next((p for p in plans if p['code'] == 'starter'), plans[1])
    target_id = starter_plan['id']
    target_name = starter_plan['name']
    print(f" [OK] Selected Target Plan: '{target_name}' (ID #{target_id}, Current Active: {starter_plan.get('is_active')})")

    # 3. 1st Click on Switch -> HIDE PLAN (is_active becomes 0)
    print("\n[Step 2] 1st Switch Click: Hiding Plan from Customers (POST /api/admin/plans/toggle_visibility)...")
    code, hide_res = request_json(f"{BASE_URL}/api/admin/plans/toggle_visibility", "POST", {"planId": target_id}, token=admin_token)
    assert code == 200 and hide_res.get('success'), f"Hide failed: {hide_res}"
    assert hide_res['isActive'] == 0, f"Expected isActive=0, got {hide_res['isActive']}"
    print(f" [OK] Plan successfully HIDDEN: {hide_res['message']}")

    # 4. Verify Customer Pricing Grid /api/plans Excludes the Hidden Plan
    print("\n[Step 3] Confirm Customer View (/api/plans) No Longer Shows the Hidden Plan...")
    code, pub_plans_res = request_json(f"{BASE_URL}/api/plans", "GET")
    assert code == 200, "Failed to load public plans"
    pub_plan_ids = [p['id'] for p in pub_plans_res['plans']]
    assert target_id not in pub_plan_ids, f"Plan #{target_id} should be HIDDEN from customer /api/plans, but found in: {pub_plan_ids}"
    print(f" [OK] Confirmed: Plan #{target_id} is completely hidden from customers on public pricing page.")

    # 5. 2nd Click on Switch -> UNHIDE PLAN (is_active becomes 1)
    print("\n[Step 4] 2nd Switch Click: Unhiding Plan for Customers (POST /api/admin/plans/toggle_visibility)...")
    code, unhide_res = request_json(f"{BASE_URL}/api/admin/plans/toggle_visibility", "POST", {"planId": target_id}, token=admin_token)
    assert code == 200 and unhide_res.get('success'), f"Unhide failed: {unhide_res}"
    assert unhide_res['isActive'] == 1, f"Expected isActive=1, got {unhide_res['isActive']}"
    print(f" [OK] Plan successfully UNHIDDEN: {unhide_res['message']}")

    # 6. Verify Customer Pricing Grid /api/plans Shows the Plan Again
    print("\n[Step 5] Confirm Customer View (/api/plans) Now Restores the Unhidden Plan...")
    code, pub_plans_res2 = request_json(f"{BASE_URL}/api/plans", "GET")
    assert code == 200
    pub_plan_ids2 = [p['id'] for p in pub_plans_res2['plans']]
    assert target_id in pub_plan_ids2, f"Plan #{target_id} should now be VISIBLE in /api/plans"
    print(f" [OK] Confirmed: Plan #{target_id} is live and visible again to all customers.")

    # 7. Verify Non-Admin Access Control (HTTP 403 Security Check)
    print("\n[Step 6] Security Check: Verify Unauthorized User Cannot Toggle Plan Visibility...")
    code, reg_user = request_json(f"{BASE_URL}/api/auth/register", "POST", {
        "email": "contractor_sec_test@patna.in",
        "password": "Password@123",
        "name": "Sec Contractor",
        "mobile": "9811122233"
    })
    _, cust_login = request_json(f"{BASE_URL}/api/auth/login", "POST", {"email": "contractor_sec_test@patna.in", "password": "Password@123"})
    cust_token = cust_login['token']

    code, unauth_toggle = request_json(f"{BASE_URL}/api/admin/plans/toggle_visibility", "POST", {"planId": target_id}, token=cust_token)
    assert code == 403, f"Expected HTTP 403 Forbidden for non-admin, got {code}: {unauth_toggle}"
    print(f" [OK] Security verified: Non-admin correctly rejected with HTTP 403 Forbidden ({unauth_toggle.get('error')}).")

    print("\n" + "=" * 75)
    print("ALL PLAN-LEVEL 'HIDE/UNHIDE' SWITCH TESTS PASSED 100%!")
    print("=" * 75)

if __name__ == "__main__":
    run_plan_hide_unhide_tests()
