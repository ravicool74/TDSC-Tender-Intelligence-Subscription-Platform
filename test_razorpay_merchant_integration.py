"""
Automated Verification Suite for Razorpay Merchant Integration
Validates all native Razorpay Merchant capabilities:
1. Admin Razorpay Settings retrieval
2. Admin Razorpay Connectivity Health Check
3. Admin Razorpay Settings persistence (DB + .env)
4. Customer Razorpay Order Creation (Paise conversion + Prefill)
5. Cryptographic HMAC-SHA256 Payment Verification
6. Invalid Signature Rejection & Audit Logging
7. Signed Idempotent Webhook Processing (X-Razorpay-Signature)
8. Refund Webhook Subscription Revocation
"""

import urllib.request
import urllib.parse
import json
import time
import hmac
import hashlib
import sys

sys.stdout.reconfigure(encoding='utf-8')

BASE_URL = "http://localhost:8000"

def request_json(url, method="GET", data=None, token=None, headers=None):
    req_headers = {'Content-Type': 'application/json'}
    if token:
        req_headers['Authorization'] = f"Bearer {token}"
    if headers:
        req_headers.update(headers)

    if isinstance(data, bytes):
        body = data
    elif data is not None:
        body = json.dumps(data).encode('utf-8')
    else:
        body = None

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

def run_razorpay_merchant_tests():
    print("=" * 70)
    print("RAZORPAY MERCHANT INTEGRATION - AUTOMATED VERIFICATION SUITE")
    print("=" * 70)

    # 1. Admin Login
    print("\n[Step 1] Admin Login...")
    code, admin_login = request_json(f"{BASE_URL}/api/auth/login", "POST", {
        "email": "admin@bihar-etender.in",
        "password": "Admin@Eproc2026"
    })
    assert code == 200 and admin_login.get('token'), f"Admin login failed: {admin_login}"
    admin_token = admin_login['token']
    print(f" [OK] Admin authenticated. Token: {admin_token[:15]}...")

    # 2. Reset and Get Default Simulation Settings
    print("\n[Step 2] Admin Razorpay Settings Retrieval & Reset to Test Simulation...")
    code, update_sim = request_json(f"{BASE_URL}/api/admin/razorpay_settings", "POST", {
        "keyId": "rzp_test_bihar_eproc_demo",
        "keySecret": "test_sec_bihar_eproc_demo_2026",
        "webhookSecret": "whsec_bihar_eproc_demo_test",
        "mode": "test",
        "companyName": "Bihar e-Procurement Portal",
        "themeColor": "#1e3a8a"
    }, token=admin_token)
    assert code == 200 and update_sim.get('success'), f"Settings reset failed: {update_sim}"

    code, rzp_settings = request_json(f"{BASE_URL}/api/admin/razorpay_settings", "GET", token=admin_token)
    assert code == 200 and rzp_settings.get('success'), f"Settings retrieval failed: {rzp_settings}"
    curr_settings = rzp_settings['settings']
    print(f" [OK] Retrieved settings: Key ID: {curr_settings['keyId']}, Mode: {curr_settings['mode']}, Secret Masked: {curr_settings['keySecretMasked']}")

    # 3. Test Razorpay Connectivity Health Check (Simulator Mode & Live Ping Mode)
    print("\n[Step 3] Test Razorpay Connectivity Health Check (POST /api/admin/test_razorpay)...")
    # A. In Simulation Mode
    code, test_res = request_json(f"{BASE_URL}/api/admin/test_razorpay", "POST", {}, token=admin_token)
    assert code == 200 and test_res.get('connected'), f"Connectivity test failed: {test_res}"
    print(f" [OK] Simulation Mode Verified: {test_res['message']} (Mode: {test_res.get('mode')})")

    # B. Test Real API Auth Check with arbitrary unverified keys (verifying API intercepts HTTP 401 safely)
    code, ping_fail = request_json(f"{BASE_URL}/api/admin/test_razorpay", "POST", {
        "keyId": "rzp_test_invalid_sample_key",
        "keySecret": "fake_secret_123456"
    }, token=admin_token)
    assert code == 200 and not ping_fail.get('connected') and ping_fail.get('httpStatus') == 401, f"Expected 401 auth rejection from Razorpay API: {ping_fail}"
    print(f" [OK] Live Razorpay REST API ping verified: Correctly detected unauthenticated keys ({ping_fail.get('error')})")

    # 4. Admin Updates Razorpay Merchant Settings
    print("\n[Step 4] Admin Updates Razorpay Merchant Settings (POST /api/admin/razorpay_settings)...")
    test_key_id = "rzp_test_bihar_merchant_889"
    test_key_sec = "secret_merchant_bihar_key_9911"
    test_wh_sec = "whsec_merchant_bihar_wh_5544"
    code, update_res = request_json(f"{BASE_URL}/api/admin/razorpay_settings", "POST", {
        "keyId": test_key_id,
        "keySecret": test_key_sec,
        "webhookSecret": test_wh_sec,
        "mode": "test",
        "companyName": "Bihar e-Procurement Tender System",
        "themeColor": "#0284c7"
    }, token=admin_token)
    assert code == 200 and update_res.get('success'), f"Update failed: {update_res}"
    new_summary = update_res['settings']
    assert new_summary['keyId'] == test_key_id, "Key ID mismatch after update"
    assert new_summary['companyName'] == "Bihar e-Procurement Tender System"
    print(f" [OK] Credentials saved successfully! New Key: {new_summary['keyId']}, Masked: {new_summary['keySecretMasked']}")

    # 5. Customer Registration & Checkout Order Creation
    print("\n[Step 5] Customer Order Creation with Razorpay Prefill & Paise Calculation...")
    cust_email = f"merchant_test_user_{int(time.time())}@example.com"
    code, reg = request_json(f"{BASE_URL}/api/auth/register", "POST", {
        "email": cust_email,
        "password": "Password@1234",
        "name": "Ravi Shankar Contractor",
        "mobile": "9876543210",
        "company": "Bihar Infra Buildcon"
    })
    assert code == 200 and reg.get('success'), f"Registration failed: {reg}"

    code, login_res = request_json(f"{BASE_URL}/api/auth/login", "POST", {
        "email": cust_email,
        "password": "Password@1234"
    })
    assert code == 200 and login_res.get('token'), f"Login failed: {login_res}"
    cust_token = login_res['token']

    # Create order for Professional plan (id=3, price=2499 INR = 249900 paise)
    code, order = request_json(f"{BASE_URL}/api/payments/create_order", "POST", {"planId": 3}, token=cust_token)
    assert code == 200 and order.get('orderId'), f"Order creation failed: {order}"
    assert order['amountPaise'] == 249900, f"Expected 249900 paise, got {order['amountPaise']}"
    assert order['currency'] == "INR"
    assert order['keyId'] == test_key_id
    assert order['user']['email'] == cust_email
    assert order['user']['mobile'] == "9876543210"
    print(f" [OK] Order Created: {order['orderId']} | Amount: {order['amountPaise']} paise | Prefill User: {order['user']['name']}")

    # 6. Cryptographic HMAC-SHA256 Signature Calculation & Verification
    print("\n[Step 6] Cryptographic HMAC-SHA256 Payment Signature Verification...")
    payment_id = f"pay_{int(time.time())}_test"
    order_id = order['orderId']
    # Authentic Razorpay signature generation: HMAC-SHA256(order_id + "|" + payment_id, secret)
    raw_sig_msg = f"{order_id}|{payment_id}".encode('utf-8')
    valid_signature = hmac.new(test_key_sec.encode('utf-8'), raw_sig_msg, hashlib.sha256).hexdigest()

    code, verify_res = request_json(f"{BASE_URL}/api/payments/verify", "POST", {
        "orderId": order_id,
        "paymentId": payment_id,
        "signature": valid_signature,
        "method": "razorpay_checkout"
    }, token=cust_token)
    assert code == 200 and verify_res.get('success'), f"Payment verification failed: {verify_res}"
    print(f" [OK] Payment verified with HMAC-SHA256: {verify_res['message']} (Invoice: {verify_res.get('invoiceNumber')})")

    # 7. Invalid Signature Tampering Rejection
    print("\n[Step 7] Tampered Signature Rejection Test...")
    tampered_sig = "invalid_tampered_signature_hex_1234567890"
    code, fake_order = request_json(f"{BASE_URL}/api/payments/create_order", "POST", {"planId": 2}, token=cust_token)
    code, reject_res = request_json(f"{BASE_URL}/api/payments/verify", "POST", {
        "orderId": fake_order['orderId'],
        "paymentId": "pay_fake_999",
        "signature": tampered_sig,
        "method": "razorpay_checkout"
    }, token=cust_token)
    assert code == 400 and not reject_res.get('success'), f"Expected 400 failure, got: {code} {reject_res}"
    print(f" [OK] Cryptographic tamper rejected with HTTP 400: {reject_res.get('error')}")

    # 8. Razorpay Webhook HMAC-SHA256 Signed Event Processing
    print("\n[Step 8] Razorpay Signed Webhook Event (payment.captured)...")
    wh_cust_email = f"wh_user_{int(time.time())}@example.com"
    code, wh_reg = request_json(f"{BASE_URL}/api/auth/register", "POST", {
        "email": wh_cust_email,
        "password": "Password@1234",
        "name": "Webhook Contractor",
        "mobile": "9811223344"
    })
    assert code == 200 and wh_reg.get('success'), f"Registration failed: {wh_reg}"

    code, wh_login = request_json(f"{BASE_URL}/api/auth/login", "POST", {
        "email": wh_cust_email,
        "password": "Password@1234"
    })
    assert code == 200 and wh_login.get('token'), f"Login failed: {wh_login}"
    wh_token = wh_login['token']
    code, wh_order = request_json(f"{BASE_URL}/api/payments/create_order", "POST", {"planId": 2}, token=wh_token)
    wh_order_id = wh_order['orderId']
    wh_payment_id = f"pay_wh_{int(time.time())}"

    wh_payload_dict = {
        "entity": "event",
        "id": f"evt_test_{int(time.time())}",
        "event": "payment.captured",
        "payload": {
            "payment": {
                "entity": {
                    "id": wh_payment_id,
                    "order_id": wh_order_id,
                    "amount": 99900,
                    "currency": "INR",
                    "status": "captured",
                    "method": "upi"
                }
            }
        }
    }
    raw_wh_bytes = json.dumps(wh_payload_dict).encode('utf-8')
    wh_signature = hmac.new(test_wh_sec.encode('utf-8'), raw_wh_bytes, hashlib.sha256).hexdigest()

    code, wh_res = request_json(
        f"{BASE_URL}/api/payments/webhook",
        "POST",
        raw_wh_bytes,
        headers={"X-Razorpay-Signature": wh_signature}
    )
    assert code == 200 and wh_res.get('success'), f"Webhook processing failed: {wh_res}"
    print(f" [OK] Webhook captured successfully: {wh_res.get('result', {}).get('message')}")

    # 9. Webhook Idempotency Test
    print("\n[Step 9] Webhook Idempotency Protection (Duplicate Event Ignored)...")
    code, wh_dup_res = request_json(
        f"{BASE_URL}/api/payments/webhook",
        "POST",
        raw_wh_bytes,
        headers={"X-Razorpay-Signature": wh_signature}
    )
    assert code == 200 and "idempotent" in wh_dup_res.get('message', '').lower(), f"Idempotency failed: {wh_dup_res}"
    print(f" [OK] Duplicate webhook safely ignored: {wh_dup_res.get('message')}")

    # 10. Webhook Refund Processing (Subscription Revocation)
    print("\n[Step 10] Refund Webhook Processing (Subscription Revocation)...")
    refund_payload_dict = {
        "entity": "event",
        "id": f"evt_rfnd_{int(time.time())}",
        "event": "refund.processed",
        "payload": {
            "refund": {
                "entity": {
                    "id": f"rfnd_{int(time.time())}",
                    "payment_id": wh_payment_id,
                    "amount": 99900,
                    "status": "processed"
                }
            }
        }
    }
    raw_refund_bytes = json.dumps(refund_payload_dict).encode('utf-8')
    refund_sig = hmac.new(test_wh_sec.encode('utf-8'), raw_refund_bytes, hashlib.sha256).hexdigest()

    code, refund_res = request_json(
        f"{BASE_URL}/api/payments/webhook",
        "POST",
        raw_refund_bytes,
        headers={"X-Razorpay-Signature": refund_sig}
    )
    assert code == 200 and refund_res.get('success'), f"Refund processing failed: {refund_res}"
    print(f" [OK] Refund processed and subscription revoked: {refund_res.get('message')}")

    # Check that webhook customer plan is now revoked
    code, wh_me = request_json(f"{BASE_URL}/api/auth/me", "GET", token=wh_token)
    sub = wh_me['user']['subscription']
    assert sub['status'] == 'cancelled', f"Expected cancelled status, got: {sub['status']}"
    assert not sub['isPaidActive'], "Expected isPaidActive to be False after refund"
    print(" [OK] Confirmed: Contractor access successfully revoked back to inactive/cancelled on refund.")

    print("\n" + "=" * 70)
    print("ALL 10 RAZORPAY MERCHANT INTEGRATION TESTS PASSED (100% SUCCESS)!")
    print("=" * 70)

if __name__ == "__main__":
    run_razorpay_merchant_tests()
