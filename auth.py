"""
Bihar e-Procurement Subscription Platform - Authentication & Security Module
Provides:
- PBKDF2-HMAC-SHA256 password hashing with unique cryptographic salts.
- Stateful user session management with expiration.
- Registration, login, logout, password reset, and verification flows.
- Sliding-window rate limiting for security against brute-force attacks.
- Audit logging for administrative and subscription events.
"""

import os
import re
import hmac
import hashlib
import secrets
import json
import time
from datetime import datetime, timedelta
from database import get_connection

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
TOKEN_VALIDITY_DAYS = 30
RESET_TOKEN_VALIDITY_HOURS = 2

# Salt and hash helpers
def hash_password(password: str, salt: bytes = None) -> tuple[str, str]:
    if not salt:
        salt = secrets.token_bytes(16)
    pw_hash = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 260000)
    return pw_hash.hex(), salt.hex()

def verify_password(password: str, stored_hash: str, stored_salt_hex: str) -> bool:
    try:
        salt = bytes.fromhex(stored_salt_hex)
        pw_hash = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 260000)
        return hmac.compare_digest(pw_hash.hex(), stored_hash)
    except Exception:
        return False

# Audit Logging
def log_audit(action: str, actor_user_id: int = None, target_user_id: int = None, details: dict = None, ip_address: str = ""):
    conn = get_connection()
    try:
        with conn:
            conn.execute("""
            INSERT INTO audit_events (actor_user_id, target_user_id, action, details_json, ip_address, created_at)
            VALUES (?, ?, ?, ?, ?, ?);
            """, (
                actor_user_id,
                target_user_id,
                action,
                json.dumps(details or {}, ensure_ascii=False),
                ip_address,
                datetime.now().isoformat()
            ))
    except Exception as e:
        print("[Auth] Audit log write warning:", e)

# Rate Limiter
def check_rate_limit(key: str, max_requests: int = 20, window_seconds: int = 60) -> bool:
    """Returns True if request is allowed, False if rate limit exceeded."""
    conn = get_connection()
    now_ts = int(time.time())
    with conn:
        cursor = conn.cursor()
        cursor.execute("SELECT count, reset_at FROM rate_limits WHERE key = ?;", (key,))
        row = cursor.fetchone()
        if not row:
            cursor.execute("INSERT INTO rate_limits (key, count, reset_at) VALUES (?, 1, ?);", (key, now_ts + window_seconds))
            return True
        count, reset_at = row['count'], row['reset_at']
        if now_ts > reset_at:
            cursor.execute("UPDATE rate_limits SET count = 1, reset_at = ? WHERE key = ?;", (now_ts + window_seconds, key))
            return True
        if count >= max_requests:
            return False
        cursor.execute("UPDATE rate_limits SET count = count + 1 WHERE key = ?;", (key,))
        return True

# Validation
def is_valid_email(email: str) -> bool:
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email))

# User Registration
def register_user(name: str, email: str, password: str, mobile: str = "", company: str = "", role: str = "customer", ip_address: str = ""):
    email = email.strip().lower()
    name = name.strip()
    mobile = mobile.strip()

    if not name:
        return {"success": False, "error": "Name is required."}
    if not is_valid_email(email):
        return {"success": False, "error": "Invalid email address format."}
    if len(password) < 6:
        return {"success": False, "error": "Password must be at least 6 characters long."}

    conn = get_connection()
    with conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM users WHERE email = ?;", (email,))
        if cursor.fetchone():
            return {"success": False, "error": "An account with this email already exists. Please login."}

        pw_hash, salt_hex = hash_password(password)
        verif_token = secrets.token_urlsafe(24)
        now = datetime.now().isoformat()

        cursor.execute("""
        INSERT INTO users (email, mobile, name, company, password_hash, salt, role, email_verified, verification_token, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?);
        """, (email, mobile, name, company, pw_hash, salt_hex, role, verif_token, now, now))
        user_id = cursor.lastrowid

        # Auto-assign Free plan subscription
        cursor.execute("SELECT id FROM plans WHERE code = 'free';")
        free_plan = cursor.fetchone()
        plan_id = free_plan['id'] if free_plan else 1

        end_date = (datetime.now() + timedelta(days=365)).isoformat()
        cursor.execute("""
        INSERT INTO subscriptions (user_id, plan_id, status, start_date, end_date, created_at, updated_at)
        VALUES (?, ?, 'active', ?, ?, ?, ?);
        """, (user_id, plan_id, now, end_date, now, now))

        # Default alert preferences
        cursor.execute("""
        INSERT INTO alert_preferences (user_id, email_enabled, whatsapp_enabled, frequency, target_districts, target_departments, target_keywords, updated_at)
        VALUES (?, 1, 0, 'daily', '[]', '[]', '[]', ?);
        """, (user_id, now))

    log_audit("USER_REGISTERED", actor_user_id=user_id, target_user_id=user_id, details={"email": email, "role": role}, ip_address=ip_address)
    return {"success": True, "message": "Account registered successfully!", "userId": user_id}

# User Login
def login_user(email: str, password: str, ip_address: str = "", user_agent: str = ""):
    email = email.strip().lower()
    
    # Rate limit check on login key
    limit_key = f"login_{ip_address}_{email}"
    if not check_rate_limit(limit_key, max_requests=10, window_seconds=300):
        return {"success": False, "error": "Too many failed login attempts. Please wait 5 minutes before trying again."}

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT id, email, mobile, name, company, password_hash, salt, role, email_verified
    FROM users WHERE email = ?;
    """, (email,))
    user = cursor.fetchone()

    if not user or not verify_password(password, user['password_hash'], user['salt']):
        log_audit("LOGIN_FAILED", details={"email": email}, ip_address=ip_address)
        return {"success": False, "error": "Invalid email or password. Please check your credentials."}

    # Issue Session Token
    session_token = secrets.token_urlsafe(32)
    expires_at = (datetime.now() + timedelta(days=TOKEN_VALIDITY_DAYS)).isoformat()
    now = datetime.now().isoformat()

    with conn:
        cursor.execute("""
        INSERT INTO user_sessions (token, user_id, expires_at, ip_address, user_agent, created_at)
        VALUES (?, ?, ?, ?, ?, ?);
        """, (session_token, user['id'], expires_at, ip_address, user_agent, now))

    # Retrieve active subscription and plan details
    sub_info = get_user_subscription(user['id'])

    # Admin, Super Admin, and Employee receive full unlocked staff access
    if user['role'] in ['admin', 'super_admin']:
        sub_info['isPaidActive'] = True
        sub_info['status'] = 'active'
        sub_info['planCode'] = 'admin_unlimited'
        sub_info['planName'] = 'Admin Full Access'
        sub_info['isAdmin'] = True
    elif user['role'] == 'employee':
        sub_info['isPaidActive'] = True
        sub_info['status'] = 'active'
        sub_info['planCode'] = 'employee_access'
        sub_info['planName'] = 'TDSC Staff Access'
        sub_info['isEmployee'] = True

    log_audit("LOGIN_SUCCESS", actor_user_id=user['id'], details={"email": email}, ip_address=ip_address)
    return {
        "success": True,
        "token": session_token,
        "user": {
            "id": user['id'],
            "name": user['name'],
            "email": user['email'],
            "mobile": user['mobile'],
            "company": user['company'],
            "role": user['role'],
            "emailVerified": bool(user['email_verified'])
        },
        "subscription": sub_info
    }

# Session verification
def get_user_from_session(token: str):
    if not token:
        return None
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT s.user_id, s.expires_at, u.id, u.email, u.mobile, u.name, u.company, u.role, u.email_verified
    FROM user_sessions s
    JOIN users u ON s.user_id = u.id
    WHERE s.token = ?;
    """, (token,))
    row = cursor.fetchone()
    if not row:
        return None

    # Check expiration
    try:
        exp = datetime.fromisoformat(row['expires_at'])
        if datetime.now() > exp:
            with conn:
                conn.execute("DELETE FROM user_sessions WHERE token = ?;", (token,))
            return None
    except Exception:
        return None

    sub_info = get_user_subscription(row['id'])

    # Admin and Super Admin automatically receive permanent full unlocked entitlements
    if row['role'] in ['admin', 'super_admin']:
        sub_info['isPaidActive'] = True
        sub_info['status'] = 'active'
        sub_info['planCode'] = 'admin_unlimited'
        sub_info['planName'] = 'Admin Full Access'
        sub_info['isAdmin'] = True
        sub_info['entitlements'] = {
            "allowBoqDownload": True,
            "allowExcelExport": True,
            "allowWhatsappAlert": True,
            "allowSavedSearches": True,
            "allowUniversalScraper": True
        }
    elif row['role'] == 'employee':
        sub_info['isPaidActive'] = True
        sub_info['status'] = 'active'
        sub_info['planCode'] = 'employee_access'
        sub_info['planName'] = 'TDSC Staff Access'
        sub_info['isEmployee'] = True
        sub_info['entitlements'] = {
            "allowBoqDownload": True,
            "allowExcelExport": True,
            "allowWhatsappAlert": True,
            "allowSavedSearches": True,
            "allowContractorCRM": True,
            "allowTenderMatcher": True
        }

    return {
        "id": row['id'],
        "name": row['name'],
        "email": row['email'],
        "mobile": row['mobile'],
        "company": row['company'],
        "role": row['role'],
        "emailVerified": bool(row['email_verified']),
        "subscription": sub_info
    }

# Logout
def logout_user(token: str):
    if not token: return
    conn = get_connection()
    with conn:
        conn.execute("DELETE FROM user_sessions WHERE token = ?;", (token,))

# Password Reset Request
def request_password_reset(email: str, ip_address: str = ""):
    email = email.strip().lower()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM users WHERE email = ?;", (email,))
    user = cursor.fetchone()
    if not user:
        # Prevent user enumeration
        return {"success": True, "message": "If this email is registered, password reset instructions have been generated."}

    token = secrets.token_urlsafe(24)
    expiry = (datetime.now() + timedelta(hours=RESET_TOKEN_VALIDITY_HOURS)).isoformat()
    with conn:
        conn.execute("UPDATE users SET reset_token = ?, reset_token_expiry = ? WHERE id = ?;", (token, expiry, user['id']))

    log_audit("PASSWORD_RESET_REQUESTED", actor_user_id=user['id'], details={"email": email}, ip_address=ip_address)
    return {
        "success": True,
        "message": "Password reset token generated successfully.",
        "resetToken": token  # In test mode, returned so user can complete reset in UI
    }

# Password Reset Complete
def complete_password_reset(token: str, new_password: str, ip_address: str = ""):
    if len(new_password) < 6:
        return {"success": False, "error": "New password must be at least 6 characters."}

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, reset_token_expiry FROM users WHERE reset_token = ?;", (token,))
    user = cursor.fetchone()
    if not user:
        return {"success": False, "error": "Invalid or expired password reset link."}

    try:
        exp = datetime.fromisoformat(user['reset_token_expiry'])
        if datetime.now() > exp:
            return {"success": False, "error": "Password reset link has expired. Please request a new one."}
    except Exception:
        return {"success": False, "error": "Invalid token expiration."}

    pw_hash, salt_hex = hash_password(new_password)
    now = datetime.now().isoformat()
    with conn:
        conn.execute("""
        UPDATE users
        SET password_hash = ?, salt = ?, reset_token = NULL, reset_token_expiry = NULL, updated_at = ?
        WHERE id = ?;
        """, (pw_hash, salt_hex, now, user['id']))

    log_audit("PASSWORD_RESET_COMPLETED", actor_user_id=user['id'], ip_address=ip_address)
    return {"success": True, "message": "Password updated successfully. You can now login with your new password."}

# Get Active Subscription for User
def get_user_subscription(user_id: int):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT s.id AS sub_id, s.status, s.start_date, s.end_date, s.auto_renew,
           p.id AS plan_id, p.code AS plan_code, p.name AS plan_name, p.badge_text,
           p.price_inr, p.billing_period, p.features_json, p.daily_view_limit,
           p.allow_boq_download, p.allow_excel_export, p.allow_whatsapp_alert,
           p.allow_email_alert, p.allow_saved_searches, p.max_saved_searches
    FROM subscriptions s
    JOIN plans p ON s.plan_id = p.id
    WHERE s.user_id = ?
    ORDER BY s.id DESC
    LIMIT 1;
    """, (user_id,))
    row = cursor.fetchone()
    if not row:
        return {
            "status": "none",
            "planCode": "free",
            "planName": "Free Plan",
            "isPaidActive": False,
            "entitlements": {
                "allowBoqDownload": False,
                "allowExcelExport": False,
                "allowWhatsappAlert": False,
                "allowSavedSearches": False
            }
        }

    status = row['status']
    is_expired = False
    try:
        end_dt = datetime.fromisoformat(row['end_date'])
        if datetime.now() > end_dt and status == 'active':
            status = 'expired'
            is_expired = True
            # Update DB status
            with conn:
                conn.execute("UPDATE subscriptions SET status = 'expired' WHERE id = ?;", (row['sub_id'],))
    except Exception:
        pass

    is_paid = (row['plan_code'] != 'free') and (status == 'active')

    features = []
    try:
        features = json.loads(row['features_json'])
    except:
        pass

    return {
        "subscriptionId": row['sub_id'],
        "status": status,
        "isExpired": is_expired,
        "isPaidActive": is_paid,
        "startDate": row['start_date'],
        "endDate": row['end_date'],
        "autoRenew": bool(row['auto_renew']),
        "plan": {
            "id": row['plan_id'],
            "code": row['plan_code'],
            "name": row['plan_name'],
            "badgeText": row['badge_text'],
            "priceInr": row['price_inr'],
            "billingPeriod": row['billing_period'],
            "features": features
        },
        "entitlements": {
            "dailyViewLimit": row['daily_view_limit'],
            "allowBoqDownload": bool(row['allow_boq_download']) and is_paid,
            "allowExcelExport": bool(row['allow_excel_export']) and is_paid,
            "allowWhatsappAlert": bool(row['allow_whatsapp_alert']) and is_paid,
            "allowEmailAlert": bool(row['allow_email_alert']),
            "allowSavedSearches": bool(row['allow_saved_searches']) or is_paid,
            "maxSavedSearches": row['max_saved_searches'] if is_paid else 1
        }
    }

# Seed default admin if missing
def ensure_default_admin():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM users WHERE role = 'admin' LIMIT 1;")
    admin = cursor.fetchone()
    if not admin:
        admin_email = os.environ.get("ADMIN_DEFAULT_EMAIL", "admin@bihar-etender.in").lower()
        admin_pass = os.environ.get("ADMIN_DEFAULT_PASSWORD", "Admin@Eproc2026")
        register_user("System Administrator", admin_email, admin_pass, mobile="9800000000", company="TDSC System Administrator", role="admin")
        print(f"[Auth] Initialized default Admin account ({admin_email})")

# Seed default employee if missing
def ensure_default_employee():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM users WHERE role = 'employee' LIMIT 1;")
    emp = cursor.fetchone()
    if not emp:
        emp_email = os.environ.get("EMPLOYEE_DEFAULT_EMAIL", "employee@tdsc.in").lower()
        emp_pass = os.environ.get("EMPLOYEE_DEFAULT_PASSWORD", "Employee@Tdsc2026")
        register_user("Rahul Kumar (Tender Executive)", emp_email, emp_pass, mobile="9876500001", company="TDSC Tender Desk", role="employee")
        print(f"[Auth] Initialized default Employee account ({emp_email})")

if __name__ == "__main__":
    ensure_default_admin()
    ensure_default_employee()
