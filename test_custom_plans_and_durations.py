"""
Automated Verification Suite for Custom Plan Creation & Billing Durations:
- Monthly (30 Days)
- Quarterly (90 Days / 3 Months)
- Half-Yearly (180 Days / 6 Months)
- Yearly / Annual (365 Days / 1 Year)
Tests:
1. Admin Plan Editing (Price, Name, Badge, Billing Period)
2. Admin Custom Quarterly Plan Creation
3. Admin Custom Half-Yearly Plan Creation
4. Admin Custom Annual Plan Creation
5. Customer Subscription & Verification of 90-day, 180-day, and 365-day expiry calculations
6. Plan Deletion / Deactivation
"""

import urllib.request
import urllib.parse
import json
import time
from datetime import datetime, timedelta
import sys
import hmac
import hashlib
from database import get_setting

sys.stdout.reconfigure(encoding='utf-8')

BASE_URL = "http://localhost:8000"

def request_json(url, method="GET", data=None, token=None, headers=None):
    req_headers = {'Content-Type': 'application/json'}
    if token:
        req_headers['Authorization'] = f"Bearer {token}"
    if headers:
        req_headers.update(headers)

    body = json.dumps(data).encode('utf-8') if data is not None else None
    req = urllib.request.Request(url, data=body, headers=req_headers, method=method)
    try:
        with urllib.request.urlopen(req) as resp:
            content = resp.read().decode('utf-8', errors='replace')
            try:
                return resp.status, json.loads(content)
            except Exception:
                return resp.status, content
    except urllib.error.HTTPError as e:
        content = e.read().decode('utf-8', errors='replace')
        try:
            return e.code, json.loads(content)
        except Exception:
            return e.code, content

def run_custom_plan_tests():
    print("=" * 70)
    print("TESTING CUSTOM PLAN CREATION & DURATIONS (MONTHLY, QUARTERLY, HALF-YEARLY, ANNUAL)")
    print("=" * 70)

    # 1. Admin Login
    print("\n[Step 1] Admin Authentication...")
    code, admin_login = request_json(f"{BASE_URL}/api/auth/login", "POST", {
        "email": "admin@bihar-etender.in",
        "password": "Admin@Eproc2026"
    })
    assert code == 200 and admin_login.get('token'), f"Admin login failed: {admin_login}"
    admin_token = admin_login['token']
    print(" [OK] Admin authenticated successfully.")

    # 2. Edit Existing Plan (e.g., Plan #2 Starter)
    print("\n[Step 2] Admin Edits Existing Starter Plan...")
    code, edit_res = request_json(f"{BASE_URL}/api/admin/plans/edit", "POST", {
        "planId": 2,
        "name": "Starter Pro Edition (Updated)",
        "badgeText": "Hot Deal",
        "priceInr": 1299.0,
        "billingPeriod": "monthly",
        "features": ["Unlimited tender search", "PAC revealed", "5 BOQ downloads", "Email alerts"],
        "allowBoqDownload": True,
        "allowExcelExport": False,
        "allowWhatsappAlert": False,
        "isActive": True
    }, token=admin_token)
    assert code == 200 and edit_res.get('success'), f"Plan edit failed: {edit_res}"
    print(f" [OK] Existing Plan updated: {edit_res['message']}")

    # 3. Create Custom Quarterly Plan (3 Months / 90 Days)
    print("\n[Step 3] Create Custom Quarterly Plan (3 Months)...")
    q_code = f"quarterly_pro_{int(time.time())}"
    code, create_q = request_json(f"{BASE_URL}/api/admin/plans/create", "POST", {
        "code": q_code,
        "name": "Professional (Quarterly - 3 Months)",
        "badgeText": "Save 15%",
        "priceInr": 5999.0,
        "billingPeriod": "quarterly",
        "features": ["All Pro features for 3 Months", "Unlimited BOQ downloads", "Excel exports", "WhatsApp alerts"],
        "allowBoqDownload": True,
        "allowExcelExport": True,
        "allowWhatsappAlert": True,
        "allowEmailAlert": True,
        "allowSavedSearches": True,
        "maxSavedSearches": 20,
        "sortOrder": 5
    }, token=admin_token)
    assert code == 200 and create_q.get('success'), f"Quarterly plan creation failed: {create_q}"
    q_plan_id = create_q['planId']
    print(f" [OK] Created Quarterly Plan ID #{q_plan_id}: {create_q['message']}")

    # 4. Create Custom Half-Yearly Plan (6 Months / 180 Days)
    print("\n[Step 4] Create Custom Half-Yearly Plan (6 Months)...")
    hy_code = f"halfyearly_gold_{int(time.time())}"
    code, create_hy = request_json(f"{BASE_URL}/api/admin/plans/create", "POST", {
        "code": hy_code,
        "name": "Gold Contractor (Half-Yearly - 6 Months)",
        "badgeText": "Save 25%",
        "priceInr": 11999.0,
        "billingPeriod": "half_yearly",
        "features": ["Full Access for 6 Months", "Unlimited BOQ & NIT downloads", "Full Excel Export", "Priority Support"],
        "allowBoqDownload": True,
        "allowExcelExport": True,
        "allowWhatsappAlert": True,
        "allowEmailAlert": True,
        "allowSavedSearches": True,
        "maxSavedSearches": 30,
        "sortOrder": 6
    }, token=admin_token)
    assert code == 200 and create_hy.get('success'), f"Half-yearly plan creation failed: {create_hy}"
    hy_plan_id = create_hy['planId']
    print(f" [OK] Created Half-Yearly Plan ID #{hy_plan_id}: {create_hy['message']}")

    # 5. Create Custom Yearly Plan (1 Year / 365 Days)
    print("\n[Step 5] Create Custom Annual / Yearly Plan (1 Year)...")
    yr_code = f"annual_vip_{int(time.time())}"
    code, create_yr = request_json(f"{BASE_URL}/api/admin/plans/create", "POST", {
        "code": yr_code,
        "name": "Enterprise VIP (Yearly - 1 Year)",
        "badgeText": "Best Value / 2 Months Free",
        "priceInr": 21999.0,
        "billingPeriod": "annual",
        "features": ["365 Days Unlimited Access", "Consortium / Multi-user", "Unlimited Downloads", "Dedicated Support"],
        "allowBoqDownload": True,
        "allowExcelExport": True,
        "allowWhatsappAlert": True,
        "allowEmailAlert": True,
        "allowSavedSearches": True,
        "maxSavedSearches": 999,
        "sortOrder": 7
    }, token=admin_token)
    assert code == 200 and create_yr.get('success'), f"Yearly plan creation failed: {create_yr}"
    yr_plan_id = create_yr['planId']
    print(f" [OK] Created Yearly Plan ID #{yr_plan_id}: {create_yr['message']}")

    # 6. Verify Customer Plans List contains all periods
    print("\n[Step 6] Public Plans Verification...")
    code, public_plans = request_json(f"{BASE_URL}/api/plans", "GET")
    assert code == 200 and public_plans.get('plans'), "Public plans failed"
    periods_found = set(p['billing_period'] for p in public_plans['plans'])
    print(f" [OK] Active billing periods available to customers: {sorted(list(periods_found))}")
    assert 'monthly' in periods_found
    assert 'quarterly' in periods_found
    assert 'half_yearly' in periods_found
    assert 'annual' in periods_found

    # 7. Customer Subscribes to Quarterly Plan (Test 90-day expiry calculation)
    print("\n[Step 7] Test Customer Subscription on Quarterly Plan (90 Days Duration)...")
    q_cust_email = f"contractor_q_{int(time.time())}@patna.in"
    request_json(f"{BASE_URL}/api/auth/register", "POST", {
        "email": q_cust_email,
        "password": "Password@123",
        "name": "Quarterly Contractor",
        "mobile": "9812345678"
    })
    _, q_login = request_json(f"{BASE_URL}/api/auth/login", "POST", {"email": q_cust_email, "password": "Password@123"})
    q_token = q_login['token']

    key_sec = get_setting("razorpay_key_secret", "test_sec_bihar_eproc_demo_2026")

    _, q_order = request_json(f"{BASE_URL}/api/payments/create_order", "POST", {"planId": q_plan_id}, token=q_token)
    q_pay_id = f"pay_q_{int(time.time())}"
    q_order_id = q_order['orderId']
    q_sig = hmac.new(key_sec.encode('utf-8'), f"{q_order_id}|{q_pay_id}".encode('utf-8'), hashlib.sha256).hexdigest()

    code, q_verify = request_json(f"{BASE_URL}/api/payments/verify", "POST", {
        "orderId": q_order_id,
        "paymentId": q_pay_id,
        "signature": q_sig,
        "method": "razorpay_checkout"
    }, token=q_token)
    assert code == 200 and q_verify.get('success'), f"Quarterly verification failed: {q_verify}"

    _, q_me = request_json(f"{BASE_URL}/api/auth/me", "GET", token=q_token)
    q_sub = q_me['user']['subscription']
    start_dt = datetime.fromisoformat(q_sub['startDate'])
    end_dt = datetime.fromisoformat(q_sub['endDate'])
    duration_days = (end_dt - start_dt).days
    print(f" [OK] Quarterly Subscription Activated! Days: {duration_days} days (Expected 90). End Date: {end_dt.strftime('%d-%b-%Y')}")
    assert duration_days == 90, f"Expected 90 days for quarterly, got {duration_days}"

    # 8. Customer Subscribes to Half-Yearly Plan (Test 180-day expiry calculation)
    print("\n[Step 8] Test Customer Subscription on Half-Yearly Plan (180 Days Duration)...")
    hy_cust_email = f"contractor_hy_{int(time.time())}@patna.in"
    request_json(f"{BASE_URL}/api/auth/register", "POST", {
        "email": hy_cust_email,
        "password": "Password@123",
        "name": "Half-Yearly Contractor",
        "mobile": "9823456789"
    })
    _, hy_login = request_json(f"{BASE_URL}/api/auth/login", "POST", {"email": hy_cust_email, "password": "Password@123"})
    hy_token = hy_login['token']

    _, hy_order = request_json(f"{BASE_URL}/api/payments/create_order", "POST", {"planId": hy_plan_id}, token=hy_token)
    hy_pay_id = f"pay_hy_{int(time.time())}"
    hy_order_id = hy_order['orderId']
    hy_sig = hmac.new(key_sec.encode('utf-8'), f"{hy_order_id}|{hy_pay_id}".encode('utf-8'), hashlib.sha256).hexdigest()

    code, hy_verify = request_json(f"{BASE_URL}/api/payments/verify", "POST", {
        "orderId": hy_order_id,
        "paymentId": hy_pay_id,
        "signature": hy_sig,
        "method": "razorpay_checkout"
    }, token=hy_token)
    assert code == 200 and hy_verify.get('success'), f"Half-Yearly verification failed: {hy_verify}"

    _, hy_me = request_json(f"{BASE_URL}/api/auth/me", "GET", token=hy_token)
    hy_sub = hy_me['user']['subscription']
    hy_start = datetime.fromisoformat(hy_sub['startDate'])
    hy_end = datetime.fromisoformat(hy_sub['endDate'])
    hy_days = (hy_end - hy_start).days
    print(f" [OK] Half-Yearly Subscription Activated! Days: {hy_days} days (Expected 180). End Date: {hy_end.strftime('%d-%b-%Y')}")
    assert hy_days == 180, f"Expected 180 days for half-yearly, got {hy_days}"

    # 9. Customer Subscribes to Annual Plan (Test 365-day expiry calculation)
    print("\n[Step 9] Test Customer Subscription on Yearly / Annual Plan (365 Days Duration)...")
    yr_cust_email = f"contractor_yr_{int(time.time())}@patna.in"
    request_json(f"{BASE_URL}/api/auth/register", "POST", {
        "email": yr_cust_email,
        "password": "Password@123",
        "name": "Annual VIP Contractor",
        "mobile": "9834567890"
    })
    _, yr_login = request_json(f"{BASE_URL}/api/auth/login", "POST", {"email": yr_cust_email, "password": "Password@123"})
    yr_token = yr_login['token']

    _, yr_order = request_json(f"{BASE_URL}/api/payments/create_order", "POST", {"planId": yr_plan_id}, token=yr_token)
    yr_pay_id = f"pay_yr_{int(time.time())}"
    yr_order_id = yr_order['orderId']
    yr_sig = hmac.new(key_sec.encode('utf-8'), f"{yr_order_id}|{yr_pay_id}".encode('utf-8'), hashlib.sha256).hexdigest()

    code, yr_verify = request_json(f"{BASE_URL}/api/payments/verify", "POST", {
        "orderId": yr_order_id,
        "paymentId": yr_pay_id,
        "signature": yr_sig,
        "method": "razorpay_checkout"
    }, token=yr_token)
    assert code == 200 and yr_verify.get('success'), f"Yearly verification failed: {yr_verify}"

    _, yr_me = request_json(f"{BASE_URL}/api/auth/me", "GET", token=yr_token)
    yr_sub = yr_me['user']['subscription']
    yr_start = datetime.fromisoformat(yr_sub['startDate'])
    yr_end = datetime.fromisoformat(yr_sub['endDate'])
    yr_days = (yr_end - yr_start).days
    print(f" [OK] Annual Subscription Activated! Days: {yr_days} days (Expected 365). End Date: {yr_end.strftime('%d-%b-%Y')}")
    assert yr_days == 365, f"Expected 365 days for annual, got {yr_days}"

    # 10. Delete / Deactivate a Plan Test
    print("\n[Step 10] Admin Deactivates / Deletes Custom Plan...")
    # Plan with active subscriber should be deactivated (is_active = 0)
    code, del_active = request_json(f"{BASE_URL}/api/admin/plans/delete", "POST", {"planId": q_plan_id}, token=admin_token)
    assert code == 200 and del_active.get('success'), f"Plan deactivation failed: {del_active}"
    print(f" [OK] Active subscriber plan safe deactivation: {del_active['message']}")

    # Plan without subscribers should be permanently deleted
    code, create_temp = request_json(f"{BASE_URL}/api/admin/plans/create", "POST", {
        "code": f"temp_plan_{int(time.time())}",
        "name": "Temporary Test Plan",
        "priceInr": 100.0,
        "billingPeriod": "monthly"
    }, token=admin_token)
    temp_id = create_temp['planId']
    code, del_temp = request_json(f"{BASE_URL}/api/admin/plans/delete", "POST", {"planId": temp_id}, token=admin_token)
    assert code == 200 and del_temp.get('success'), f"Temp plan delete failed: {del_temp}"
    print(f" [OK] Unused plan permanent deletion: {del_temp['message']}")

    print("\n" + "=" * 70)
    print("ALL CUSTOM PLAN & DURATION TESTS (MONTHLY, QUARTERLY, HALF-YEARLY, ANNUAL) PASSED 100%!")
    print("=" * 70)

if __name__ == "__main__":
    run_custom_plan_tests()
