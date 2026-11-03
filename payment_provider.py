"""
Bihar e-Procurement Subscription Platform - Payment Provider Module
Implements Native Razorpay Merchant Gateway Integration:
- Official Razorpay Orders API (POST https://api.razorpay.com/v1/orders)
- Official Razorpay Checkout Modal SDK support (https://checkout.razorpay.com/v1/checkout.js)
- Cryptographic HMAC-SHA256 signature verification for payments and webhooks
- Direct Razorpay API Connectivity & Merchant Credential Health Check
- Dynamic Admin Configuration stored in SQLite database system_settings & .env
- Webhook idempotency protection
- Payment reconciliation, Invoice generation, and Subscription Lifecycle management
- Zero-storage of sensitive card/UPI data (Hosted Checkout modal only)
"""

import os
import hmac
import hashlib
import json
import time
import secrets
import urllib.request
import urllib.error
import base64
from datetime import datetime, timedelta
from database import get_connection, get_setting, set_setting
from auth import log_audit

DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class BasePaymentAdapter:
    def create_order(self, user_id: int, plan_dict: dict, notes: dict = None) -> dict:
        raise NotImplementedError
    def verify_payment_signature(self, order_id: str, payment_id: str, signature: str) -> bool:
        raise NotImplementedError
    def verify_webhook_signature(self, raw_body: bytes, signature_header: str) -> bool:
        raise NotImplementedError


class RazorpayMerchantAdapter(BasePaymentAdapter):
    """
    Native Razorpay Merchant Adapter supporting both Test & Live production modes.
    Directly interfaces with Razorpay REST APIs and verifies HMAC-SHA256 signatures.
    """
    def __init__(self):
        self.reload_config()

    def reload_config(self):
        """Loads or refreshes credentials from SQLite system_settings table, falling back to os.environ."""
        self.key_id = get_setting("razorpay_key_id", os.environ.get("RAZORPAY_KEY_ID", "rzp_test_bihar_eproc_demo"))
        self.key_secret = get_setting("razorpay_key_secret", os.environ.get("RAZORPAY_KEY_SECRET", "test_sec_bihar_eproc_demo_2026"))
        self.webhook_secret = get_setting("razorpay_webhook_secret", os.environ.get("RAZORPAY_WEBHOOK_SECRET", "whsec_bihar_eproc_demo_test"))
        self.mode = get_setting("razorpay_mode", os.environ.get("RAZORPAY_MODE", "test")).lower()
        self.company_name = get_setting("razorpay_company_name", "Bihar e-Procurement Portal")
        self.theme_color = get_setting("razorpay_theme_color", "#1e3a8a")

    def is_live_mode(self) -> bool:
        """Determines if the gateway is running in Live production mode."""
        return self.mode == "live" or (self.key_id and self.key_id.startswith("rzp_live_"))

    def is_configured(self) -> bool:
        """Checks if real merchant keys are supplied (beyond local mock placeholders)."""
        if not self.key_id or not self.key_secret:
            return False
        return not (self.key_id == "rzp_test_bihar_eproc_demo" and self.key_secret == "test_sec_bihar_eproc_demo_2026")

    def test_connectivity(self, key_id: str = None, key_secret: str = None) -> dict:
        """
        Pings Razorpay Merchant API (GET https://api.razorpay.com/v1/payments?count=1)
        using HTTP Basic Auth to verify merchant credentials and API health.
        """
        k_id = key_id or self.key_id
        k_sec = key_secret or self.key_secret

        if not k_id or not k_sec:
            return {
                "success": False,
                "connected": False,
                "error": "Razorpay Key ID aur Key Secret dono required hain."
            }

        # If using demo placeholder keys without real internet call
        if k_id == "rzp_test_bihar_eproc_demo" and k_sec == "test_sec_bihar_eproc_demo_2026":
            return {
                "success": True,
                "connected": True,
                "isMock": True,
                "mode": "test_simulator",
                "message": "Local Test Mode credentials active (Simulation Ready). Real live payments require valid rzp_live / rzp_test keys."
            }

        auth_str = f"{k_id}:{k_sec}"
        b64_auth = base64.b64encode(auth_str.encode('utf-8')).decode('ascii')

        req = urllib.request.Request(
            "https://api.razorpay.com/v1/payments?count=1",
            headers={
                "Authorization": f"Basic {b64_auth}",
                "User-Agent": "Bihar-eProc-Tender-Subscription/1.0"
            },
            method="GET"
        )

        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                mode_str = "Live Mode" if k_id.startswith("rzp_live_") else "Test Mode"
                return {
                    "success": True,
                    "connected": True,
                    "isMock": False,
                    "mode": mode_str,
                    "message": f"Razorpay API connected successfully! Merchant account verified in {mode_str}."
                }
        except urllib.error.HTTPError as e:
            try:
                err_body = json.loads(e.read().decode('utf-8'))
                err_desc = err_body.get("error", {}).get("description", str(e))
            except Exception:
                err_desc = str(e)
            return {
                "success": False,
                "connected": False,
                "httpStatus": e.code,
                "error": f"Razorpay API Error ({e.code}): {err_desc}"
            }
        except Exception as e:
            return {
                "success": False,
                "connected": False,
                "error": f"Network / Connection Error: {str(e)}"
            }

    def create_order(self, user_id: int, plan_dict: dict, notes: dict = None) -> dict:
        """
        Creates a payment order on Razorpay servers (or server fallback if offline/mock),
        and records pending transaction in the SQLite payments table.
        """
        price_inr = float(plan_dict.get("price_inr", 0.0))
        amount_paise = int(round(price_inr * 100))
        now = datetime.now().isoformat()
        notes_dict = notes or {}

        order_id = None
        is_real_razorpay = False

        # Attempt to call Razorpay Orders API if real credentials configured
        if self.is_configured():
            try:
                auth_str = f"{self.key_id}:{self.key_secret}"
                b64_auth = base64.b64encode(auth_str.encode('utf-8')).decode('ascii')

                payload = {
                    "amount": amount_paise,
                    "currency": "INR",
                    "receipt": f"rcpt_{int(time.time())}_{secrets.token_hex(3)}",
                    "notes": {
                        "userId": str(user_id),
                        "planId": str(plan_dict.get("id")),
                        "planCode": plan_dict.get("code", "")
                    }
                }

                req = urllib.request.Request(
                    "https://api.razorpay.com/v1/orders",
                    data=json.dumps(payload).encode('utf-8'),
                    headers={
                        "Authorization": f"Basic {b64_auth}",
                        "Content-Type": "application/json",
                        "User-Agent": "Bihar-eProc-Tender-Subscription/1.0"
                    },
                    method="POST"
                )

                with urllib.request.urlopen(req, timeout=12) as resp:
                    rzp_res = json.loads(resp.read().decode('utf-8'))
                    if rzp_res.get("id"):
                        order_id = rzp_res.get("id")
                        is_real_razorpay = True
            except Exception as e:
                # Log error and fallback to server-generated order ID so checkout is not blocked
                print(f"[Razorpay] API call failed ({e}). Falling back to local order ID.")

        if not order_id:
            order_id = f"order_{int(time.time())}_{secrets.token_hex(4)}"

        # Record pending transaction in SQLite database
        conn = get_connection()
        with conn:
            cursor = conn.cursor()
            cursor.execute("""
            INSERT INTO payments (
                user_id, order_id, provider, amount_inr, currency, status, raw_payload, created_at, updated_at
            ) VALUES (?, ?, 'razorpay', ?, 'INR', 'pending', ?, ?, ?);
            """, (
                user_id,
                order_id,
                price_inr,
                json.dumps({
                    "planId": plan_dict["id"],
                    "planCode": plan_dict["code"],
                    "isRealRazorpay": is_real_razorpay,
                    "notes": notes_dict
                }),
                now, now
            ))

        return {
            "success": True,
            "orderId": order_id,
            "razorpayOrderId": order_id,
            "amountPaise": amount_paise,
            "amountInr": price_inr,
            "currency": "INR",
            "keyId": self.key_id,
            "planName": plan_dict["name"],
            "planCode": plan_dict["code"],
            "companyName": self.company_name,
            "themeColor": self.theme_color,
            "mode": self.mode,
            "isRealRazorpay": is_real_razorpay
        }

    def verify_payment_signature(self, order_id: str, payment_id: str, signature: str) -> bool:
        """
        Cryptographically verifies Razorpay Checkout signature:
        HMAC-SHA256(order_id + "|" + payment_id, key_secret) == signature
        """
        if not order_id or not payment_id or not signature:
            return False

        # In Test mode or mock environment, allow demo simulator signature
        if signature.startswith("demo_test_sig_") and not self.is_live_mode():
            return True

        message = f"{order_id}|{payment_id}".encode('utf-8')
        expected_sig = hmac.new(self.key_secret.encode('utf-8'), message, hashlib.sha256).hexdigest()
        return hmac.compare_digest(expected_sig, signature)

    def verify_webhook_signature(self, raw_body: bytes, signature_header: str) -> bool:
        """
        Verifies Razorpay Webhook header 'X-Razorpay-Signature':
        HMAC-SHA256(raw_body, webhook_secret) == signature_header
        """
        if not signature_header or not self.webhook_secret:
            return False
        expected = hmac.new(self.webhook_secret.encode('utf-8'), raw_body, hashlib.sha256).hexdigest()
        return hmac.compare_digest(expected, signature_header)

    def update_credentials(self, key_id: str, key_secret: str, webhook_secret: str, mode: str, company_name: str = None, theme_color: str = None):
        """
        Saves updated Razorpay Merchant credentials to database system_settings table,
        updates active in-memory instance, and updates .env file.
        """
        if key_id:
            set_setting("razorpay_key_id", key_id.strip())
        if key_secret:
            set_setting("razorpay_key_secret", key_secret.strip())
        if webhook_secret:
            set_setting("razorpay_webhook_secret", webhook_secret.strip())
        if mode:
            set_setting("razorpay_mode", mode.strip().lower())
        if company_name:
            set_setting("razorpay_company_name", company_name.strip())
        if theme_color:
            set_setting("razorpay_theme_color", theme_color.strip())

        # Reload in-memory config
        self.reload_config()

        # Update or create .env file for environment backup
        env_path = os.path.join(DIRECTORY, ".env")
        try:
            existing_lines = []
            if os.path.exists(env_path):
                with open(env_path, "r", encoding="utf-8") as f:
                    existing_lines = f.readlines()
            
            settings_map = {
                "RAZORPAY_KEY_ID": self.key_id,
                "RAZORPAY_KEY_SECRET": self.key_secret,
                "RAZORPAY_WEBHOOK_SECRET": self.webhook_secret,
                "RAZORPAY_MODE": self.mode,
                "RAZORPAY_COMPANY_NAME": self.company_name,
                "RAZORPAY_THEME_COLOR": self.theme_color
            }

            written_keys = set()
            new_lines = []
            for line in existing_lines:
                matched = False
                for k, v in settings_map.items():
                    if line.strip().startswith(f"{k}="):
                        new_lines.append(f"{k}={v}\n")
                        written_keys.add(k)
                        matched = True
                        break
                if not matched:
                    new_lines.append(line)

            for k, v in settings_map.items():
                if k not in written_keys:
                    new_lines.append(f"{k}={v}\n")

            with open(env_path, "w", encoding="utf-8") as f:
                f.writelines(new_lines)
        except Exception as e:
            print(f"[Razorpay] Could not write .env file: {e}")

    def get_admin_summary(self) -> dict:
        """Returns safe summary of current Razorpay Merchant configuration (secret masked)."""
        masked_secret = ""
        if self.key_secret:
            if len(self.key_secret) > 8:
                masked_secret = self.key_secret[:4] + "••••••••" + self.key_secret[-4:]
            else:
                masked_secret = "••••••••"

        masked_wh = ""
        if self.webhook_secret:
            if len(self.webhook_secret) > 8:
                masked_wh = self.webhook_secret[:4] + "••••••••" + self.webhook_secret[-4:]
            else:
                masked_wh = "••••••••"

        return {
            "keyId": self.key_id,
            "keySecretMasked": masked_secret,
            "webhookSecretMasked": masked_wh,
            "mode": self.mode,
            "companyName": self.company_name,
            "themeColor": self.theme_color,
            "isLive": self.is_live_mode(),
            "isConfigured": self.is_configured()
        }


# Global adapter instance and alias
RazorpayTestAdapter = RazorpayMerchantAdapter
payment_gateway = RazorpayMerchantAdapter()


def complete_successful_payment(order_id: str, payment_id: str, method: str = "hosted_checkout", signature: str = "", ip_address: str = ""):
    """
    Reconciles order, marks payment successful, generates invoice,
    and activates paid subscription.
    """
    conn = get_connection()
    with conn:
        cursor = conn.cursor()
        cursor.execute("""
        SELECT id, user_id, amount_inr, status, raw_payload
        FROM payments WHERE order_id = ?;
        """, (order_id,))
        payment = cursor.fetchone()
        if not payment:
            return {"success": False, "error": f"Order {order_id} not found."}

        if payment['status'] == 'successful':
            return {"success": True, "message": "Payment already processed.", "alreadyProcessed": True}

        user_id = payment['user_id']
        amount_inr = payment['amount_inr']

        payload = {}
        try: payload = json.loads(payment['raw_payload'] or '{}')
        except: pass
        plan_id = payload.get('planId')

        cursor.execute("SELECT id, code, name, price_inr, billing_period FROM plans WHERE id = ?;", (plan_id,))
        plan = cursor.fetchone()
        if not plan:
            return {"success": False, "error": "Associated plan not found."}

        # Validate amount reconciliation
        if abs(plan['price_inr'] - amount_inr) > 0.01:
            log_audit("PAYMENT_RECONCILIATION_MISMATCH", actor_user_id=user_id,
                      details={"orderId": order_id, "expected": plan['price_inr'], "received": amount_inr},
                      ip_address=ip_address)
            return {"success": False, "error": "Payment amount does not match plan price."}

        now = datetime.now()
        now_iso = now.isoformat()

        # Calculate subscription duration based on billing period
        plan_billing_period = plan['billing_period'] if ('billing_period' in plan.keys() and plan['billing_period']) else 'monthly'
        period = str(plan_billing_period).lower().strip()
        if period in ['quarterly', 'quarter', '3_months']:
            duration_days = 90
        elif period in ['half_yearly', 'halfyearly', 'semi_annual', '6_months']:
            duration_days = 180
        elif period in ['annual', 'yearly', '1_year', 'annual_plan']:
            duration_days = 365
        else:
            duration_days = 30
        end_iso = (now + timedelta(days=duration_days)).isoformat()

        # 1. Update or Create Subscription
        cursor.execute("""
        INSERT INTO subscriptions (user_id, plan_id, status, start_date, end_date, auto_renew, created_at, updated_at)
        VALUES (?, ?, 'active', ?, ?, 1, ?, ?);
        """, (user_id, plan['id'], now_iso, end_iso, now_iso, now_iso))
        sub_id = cursor.lastrowid

        # 2. Update Payment
        cursor.execute("""
        UPDATE payments
        SET subscription_id = ?, transaction_id = ?, status = 'successful', payment_method = ?, signature = ?, updated_at = ?
        WHERE order_id = ?;
        """, (sub_id, payment_id, method, signature, now_iso, order_id))
        payment_record_id = payment['id']

        # 3. Generate Official Invoice
        invoice_num = f"INV-{now.strftime('%Y%m')}-{secrets.token_hex(3).upper()}"
        cursor.execute("""
        INSERT INTO invoices (invoice_number, user_id, payment_id, subscription_id, amount_inr, tax_inr, total_inr, status, invoice_date, created_at)
        VALUES (?, ?, ?, ?, ?, 0.0, ?, 'paid', ?, ?);
        """, (invoice_num, user_id, payment_record_id, sub_id, amount_inr, amount_inr, now.strftime('%d-%m-%Y'), now_iso))

    log_audit("SUBSCRIPTION_ACTIVATED", actor_user_id=user_id, target_user_id=user_id,
              details={"plan": plan['code'], "orderId": order_id, "txn": payment_id, "amount": amount_inr},
              ip_address=ip_address)

    return {
        "success": True,
        "message": f"Payment verified! Your {plan['name']} is now ACTIVE until {now + timedelta(days=duration_days):%d-%b-%Y}.",
        "subscriptionId": sub_id,
        "invoiceNumber": invoice_num
    }

def record_failed_payment(order_id: str, error_code: str, error_desc: str, ip_address: str = ""):
    conn = get_connection()
    with conn:
        cursor = conn.cursor()
        cursor.execute("""
        UPDATE payments
        SET status = 'failed', error_code = ?, error_description = ?, updated_at = ?
        WHERE order_id = ?;
        """, (error_code, error_desc, datetime.now().isoformat(), order_id))
        cursor.execute("SELECT user_id, amount_inr FROM payments WHERE order_id = ?;", (order_id,))
        row = cursor.fetchone()
        user_id = row['user_id'] if row else None

    log_audit("PAYMENT_FAILED", actor_user_id=user_id,
              details={"orderId": order_id, "code": error_code, "reason": error_desc},
              ip_address=ip_address)
    return {"success": True, "message": "Failed payment recorded."}

def process_webhook_payload(raw_body: bytes, signature_header: str, ip_address: str = ""):
    """
    Idempotent signed webhook receiver:
    1. Verifies HMAC-SHA256 signature
    2. Checks idempotency cache (processed_webhooks)
    3. Handles: payment.captured, payment.failed, refund.created, subscription.cancelled
    """
    if not payment_gateway.verify_webhook_signature(raw_body, signature_header):
        return {"success": False, "error": "Invalid webhook cryptographic signature", "status_code": 400}

    try:
        data = json.loads(raw_body.decode('utf-8'))
    except Exception as e:
        return {"success": False, "error": f"Invalid JSON payload: {e}", "status_code": 400}

    event_id = data.get("id") or f"evt_{hashlib.md5(raw_body).hexdigest()}"
    event_type = data.get("event", "payment.captured")

    conn = get_connection()
    # Check Idempotency
    with conn:
        cursor = conn.cursor()
        cursor.execute("SELECT event_id FROM processed_webhooks WHERE event_id = ?;", (event_id,))
        if cursor.fetchone():
            return {"success": True, "message": "Event already processed (idempotent skipped)", "status_code": 200}
        
        cursor.execute("""
        INSERT INTO processed_webhooks (event_id, provider, event_type, processed_at)
        VALUES (?, 'razorpay', ?, ?);
        """, (event_id, event_type, datetime.now().isoformat()))

    # Process by Event Type
    payload = data.get("payload", {})
    payment_entity = payload.get("payment", {}).get("entity", {})
    order_id = payment_entity.get("order_id")
    payment_id = payment_entity.get("id")
    method = payment_entity.get("method", "webhook")

    if event_type in ["payment.captured", "order.paid"]:
        if order_id and payment_id:
            res = complete_successful_payment(order_id, payment_id, method=method, signature=signature_header, ip_address=ip_address)
            return {"success": True, "result": res, "status_code": 200}

    elif event_type in ["payment.failed"]:
        err_code = payment_entity.get("error_code", "PAYMENT_FAILED")
        err_desc = payment_entity.get("error_description", "Payment failed at gateway")
        if order_id:
            record_failed_payment(order_id, err_code, err_desc, ip_address=ip_address)
        return {"success": True, "message": "Recorded failure", "status_code": 200}

    elif event_type in ["refund.created", "refund.processed"]:
        # Handle refund: revoke subscription
        refund_entity = payload.get("refund", {}).get("entity", {})
        p_id = refund_entity.get("payment_id")
        user_id_to_log = None
        with conn:
            cursor = conn.cursor()
            cursor.execute("SELECT subscription_id, user_id FROM payments WHERE transaction_id = ?;", (p_id,))
            p_row = cursor.fetchone()
            if p_row and p_row['subscription_id']:
                user_id_to_log = p_row['user_id']
                cursor.execute("UPDATE subscriptions SET status = 'cancelled', cancelled_at = ? WHERE id = ?;",
                               (datetime.now().isoformat(), p_row['subscription_id']))

        if user_id_to_log:
            log_audit("SUBSCRIPTION_CANCELLED_REFUND", actor_user_id=user_id_to_log,
                      details={"paymentId": p_id, "refundId": refund_entity.get("id")}, ip_address=ip_address)

        return {"success": True, "message": "Refund processed and subscription cancelled.", "status_code": 200}

    return {"success": True, "message": f"Webhook event '{event_type}' handled.", "status_code": 200}
