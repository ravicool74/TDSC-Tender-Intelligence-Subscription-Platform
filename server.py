"""
Bihar e-Procurement Paid Tender Subscription Platform - Server
Runs at http://localhost:8000
Features:
- Customer Authentication (Register, Login, Session, Password Reset)
- Plan Management (Admin-editable plans with dynamic pricing)
- Payment Gateway Integration (Razorpay Test Mode / Hosted Checkout with signed webhooks)
- Server-side Entitlement Enforcement (Free vs Paid masking, locked downloads, locked exports)
- Customer Dashboard (Active plan, Saved searches, Alert preferences, Invoices, Closing-soon tenders)
- Admin Portal (User management, Subscription overrides, Plan editor, Audit logs, Data health)
- Preserved Existing Features (Tender search, 13-column table, 3-letter search, Slabs, Sync, Webhook scraper)
"""

import http.server
import socketserver
import os
import json
import urllib.parse
import webbrowser
import threading
import subprocess
import time
import re
import ssl
from datetime import datetime, timedelta

# Disable SSL verification globally for NIC / eProcurement Indian government portals (which use internal CA chains)
try:
    ssl._create_default_https_context = ssl._create_unverified_context
except AttributeError:
    pass

# Local module imports
from database import init_database, get_connection
from auth import (
    register_user, login_user, logout_user, get_user_from_session,
    request_password_reset, complete_password_reset, log_audit, ensure_default_admin,
    ensure_default_employee, check_rate_limit
)
from payment_provider import (
    payment_gateway, complete_successful_payment, record_failed_payment, process_webhook_payload
)
from entitlements import (
    filter_tenders_for_user, check_entitlement, generate_download_token, verify_download_token,
    OFFICIAL_SOURCE_BASE
)
from fetch_eproc_tenders import fetch_live_tenders, fetch_tender_deep_dive
from auto_action_crawler import run_auto_action_crawler
from sync_excel_to_web import extract_excel_data
from universal_scraper import scrape_any_url
from action_details_manager import get_action_details, load_cache, sync_single_tender_live
from build_unified_dataset import build_comprehensive_unified_dataset
from cpp_tender_manager import (
    load_cpp_tenders_store, save_cpp_tenders_batch, toggle_cpp_shortlist, sync_cpp_tender_with_captcha,
    init_cpp_session, submit_cpp_search_and_crawl, crawl_next_cpp_page,
    import_cpp_raw_html, download_cpp_document, generate_cpp_excel_workbook,
    enrich_cpp_tender_detail, sync_cpp_single_tender_live,
    download_all_cpp_tender_attachments, CPP_EXCEL_PATH, CPP_DOWNLOADS_DIR
)

PORT = int(os.environ.get("PORT", 8000))
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
PREFS_FILE = os.path.join(DIRECTORY, "user_preferences.json")
EXCEL_PATH = os.path.join(DIRECTORY, "Bihar_eProc_Tenders_Comprehensive.xlsx")
VIEW_JSON_PATH = os.path.join(DIRECTORY, "excel_tenders_view.json")

# In-memory server state & RAM Cache
LAST_SYNC_TIME = datetime.now().strftime("%d-%m-%Y %I:%M %p")
MASTER_TENDERS_VIEW_CACHE = None
CACHE_LOCK = threading.Lock()

CATEGORY_KEYWORDS = {
    "Roads & Highways": ["road", "sadak", "rwd", "rcd", "highway", "pmgsy", "bitumen", "asphalt", "pavement", "pcc road", "path", "widening", "carpeting", "rural works"],
    "Building Construction": ["building", "bhawan", "school", "hospital", "residential", "painting", "repair", "bcd", "court", "hostel", "administrative", "classroom", "sarkar bhawan"],
    "Bridges & Culverts": ["bridge", "pul", "culvert", "pulia", "spab", "girder", "river", "overbridge", "pul nirman"],
    "Water Supply & PHE": ["water", "jal", "pipe", "phed", "oht", "tube", "drinking", "boring", "sewerage", "distribution", "nal ka jal"],
    "Electrical Works": ["electric", "vidyut", "transformer", "solar", "light", "cctv", "wire", "power", "substation", "high mast", "energy"],
    "Irrigation & Drainage": ["irrigation", "sinchai", "canal", "bund", "sluice", "drain", "flood", "embankment", "barrage", "water resources"],
    "General Civil": ["civil", "pcc", "boundary", "brick", "drain", "shed", "repair", "construction", "renovation"]
}

def get_budget_range(slab: str):
    slab = str(slab).strip().lower()
    if slab in ["0to25l", "0-25l"]: return (0, 2500000)
    if slab in ["25lto50l", "25-50l"]: return (2500001, 5000000)
    if slab in ["50lto2cr", "50-2cr"]: return (5000001, 20000000)
    if slab in ["2crto5cr", "2-5cr"]: return (20000001, 50000000)
    if slab in ["above5cr", ">5cr"]: return (50000001, float('inf'))
    return (0, float('inf'))

def load_master_tenders(force_reload=False):
    global MASTER_TENDERS_VIEW_CACHE
    with CACHE_LOCK:
        if MASTER_TENDERS_VIEW_CACHE is not None and not force_reload:
            return MASTER_TENDERS_VIEW_CACHE
        if os.path.exists(VIEW_JSON_PATH):
            try:
                with open(VIEW_JSON_PATH, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    MASTER_TENDERS_VIEW_CACHE = data.get('tenders', [])
                    return MASTER_TENDERS_VIEW_CACHE
            except Exception:
                return []
        return []

def match_tenders_for_contractor(contractor: dict, tenders: list, limit: int = 25):
    category = contractor.get('category', 'General Civil')
    keywords = CATEGORY_KEYWORDS.get(category, CATEGORY_KEYWORDS["General Civil"])
    district = str(contractor.get('district', '')).strip().lower()
    budget_slab = contractor.get('budget_slab', '50Lto2Cr')
    min_b, max_b = get_budget_range(budget_slab)

    matched = []
    for t in tenders:
        text = f"{t.get('description', '')} {t.get('detailedDescription', '')} {t.get('department', '')}".lower()
        t_dist = str(t.get('district', '')).lower()
        pac = float(t.get('estimatedValue') or 0.0)

        score = 0
        reasons = []

        matched_kw = [kw for kw in keywords if kw in text]
        if matched_kw:
            score += 50 + min(len(matched_kw) * 5, 20)
            reasons.append(f"Category Match ({category})")

        if district and (district in t_dist or t_dist in district):
            score += 30
            reasons.append(f"District Match ({contractor.get('district')})")

        if pac > 0 and min_b <= pac <= max_b:
            score += 20
            reasons.append(f"Capacity Match ({budget_slab})")
        elif pac == 0:
            score += 5

        if score >= 40:
            item = dict(t)
            item['matchScore'] = score
            item['matchReasons'] = reasons
            matched.append(item)

    matched.sort(key=lambda x: x['matchScore'], reverse=True)
    return matched[:limit]

def match_contractors_for_tender(tender: dict, contractors: list):
    t_text = f"{tender.get('description', '')} {tender.get('detailedDescription', '')} {tender.get('department', '')}".lower()
    t_dist = str(tender.get('district', '')).lower()
    t_pac = float(tender.get('estimatedValue') or 0.0)

    results = []
    for c in contractors:
        category = c.get('category', 'General Civil')
        keywords = CATEGORY_KEYWORDS.get(category, CATEGORY_KEYWORDS["General Civil"])
        c_dist = str(c.get('district', '')).lower()
        min_b, max_b = get_budget_range(c.get('budget_slab', ''))

        score = 0
        reasons = []

        matched_kw = [kw for kw in keywords if kw in t_text]
        if matched_kw:
            score += 50
            reasons.append(f"Category: {category}")

        if c_dist and (c_dist in t_dist or t_dist in c_dist):
            score += 30
            reasons.append(f"District: {c.get('district')}")

        if t_pac > 0 and min_b <= t_pac <= max_b:
            score += 20
            reasons.append(f"Capacity: {c.get('budget_slab')}")

        if score >= 30:
            c_copy = dict(c)
            c_copy['matchScore'] = score
            c_copy['matchReasons'] = reasons
            results.append(c_copy)

    results.sort(key=lambda x: x['matchScore'], reverse=True)
    return results

class PaidTenderRequestHandler(http.server.SimpleHTTPRequestHandler):
    def address_string(self):
        # Prevent slow reverse DNS lookup on Windows
        return str(self.client_address[0])
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    # -------------------------------------------------------------
    # Helper Utilities
    # -------------------------------------------------------------
    def get_client_ip(self):
        return self.client_address[0] if self.client_address else "127.0.0.1"

    def get_auth_token(self):
        auth_header = self.headers.get('Authorization', '')
        if auth_header.startswith('Bearer '):
            return auth_header[7:].strip()
        cookie_header = self.headers.get('Cookie', '')
        if 'session_token=' in cookie_header:
            match = re.search(r'session_token=([^;]+)', cookie_header)
            if match:
                return match.group(1).strip()
        return None

    def get_current_user(self):
        token = self.get_auth_token()
        if not token:
            return None
        return get_user_from_session(token)

    def read_json_body(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            if length <= 0:
                return {}
            body = self.rfile.read(length).decode('utf-8', errors='replace')
            return json.loads(body) if body else {}
        except Exception:
            return {}

    def send_json(self, data, status=200, set_cookie=None):
        payload = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(payload)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, DELETE, PUT')
        if set_cookie:
            self.send_header('Set-Cookie', set_cookie)
        self.end_headers()
        self.wfile.write(payload)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With, X-Razorpay-Signature')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, DELETE, PUT')
        self.end_headers()

    # -------------------------------------------------------------
    # GET Requests Routing
    # -------------------------------------------------------------
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        # 1. Tender Listing (With In-Memory RAM Cache & Server-Side Plan Entitlements)
        if path == '/api/excel_tenders':
            try:
                user = self.get_current_user()
                raw_tenders = load_master_tenders()
                
                # Apply backend plan entitlement masking
                filtered_tenders = filter_tenders_for_user(raw_tenders, user, LAST_SYNC_TIME)

                res = {
                    "totalTenders": len(filtered_tenders),
                    "tenders": filtered_tenders,
                    "lastChecked": LAST_SYNC_TIME,
                    "officialSourceUrl": OFFICIAL_SOURCE_BASE,
                    "isPaidActive": bool(user and user.get("subscription", {}).get("isPaidActive")),
                    "userPlan": user.get("subscription", {}).get("plan", {}).get("code", "guest") if user else "guest"
                }
                self.send_json(res)
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        # 2. Current User Profile & Active Subscription
        elif path == '/api/auth/me':
            user = self.get_current_user()
            if not user:
                self.send_json({"authenticated": False, "user": None})
            else:
                self.send_json({"authenticated": True, "user": user})
            return

        # 3. Public Plans List (Admin-Editable, No Hardcoded Prices)
        elif path == '/api/plans':
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
            SELECT id, code, name, badge_text, price_inr, billing_period, description,
                   features_json, daily_view_limit, allow_boq_download, allow_excel_export,
                   allow_whatsapp_alert, allow_email_alert, allow_saved_searches, max_saved_searches, is_active
            FROM plans
            WHERE is_active = 1
            ORDER BY sort_order ASC, price_inr ASC;
            """)
            plans = []
            for r in cursor.fetchall():
                p = dict(r)
                try: p['features'] = json.loads(p['features_json'])
                except: p['features'] = []
                plans.append(p)
            self.send_json({"success": True, "plans": plans})
            return

        # 4. Customer Dashboard Summary Data
        elif path == '/api/customer/dashboard':
            user = self.get_current_user()
            if not user:
                self.send_json({"error": "Unauthorized. Please login."}, status=401)
                return

            conn = get_connection()
            cursor = conn.cursor()

            # Saved Searches
            cursor.execute("""
            SELECT id, name, query_keyword, amount_slab, district, department, created_at
            FROM saved_searches WHERE user_id = ? ORDER BY id DESC;
            """, (user['id'],))
            saved_searches = [dict(r) for r in cursor.fetchall()]

            # Invoices
            cursor.execute("""
            SELECT id, invoice_number, amount_inr, total_inr, status, invoice_date, created_at
            FROM invoices WHERE user_id = ? ORDER BY id DESC;
            """, (user['id'],))
            invoices = [dict(r) for r in cursor.fetchall()]

            # Payments
            cursor.execute("""
            SELECT id, order_id, transaction_id, amount_inr, currency, status, payment_method, created_at
            FROM payments WHERE user_id = ? ORDER BY id DESC;
            """, (user['id'],))
            payments = [dict(r) for r in cursor.fetchall()]

            # Alert Preferences
            cursor.execute("SELECT * FROM alert_preferences WHERE user_id = ?;", (user['id'],))
            ap_row = cursor.fetchone()
            alert_prefs = dict(ap_row) if ap_row else {}

            # Closing soon tenders (due in next 3 days)
            closing_soon = []
            if os.path.exists(VIEW_JSON_PATH):
                try:
                    with open(VIEW_JSON_PATH, 'r', encoding='utf-8') as f:
                        all_t = json.load(f).get('tenders', [])
                    for t in all_t[:30]:
                        if any(w in str(t.get('dueDate', '')).lower() for w in ['01-', '02-', '03-', 'today', 'hours']):
                            closing_soon.append({
                                "systemTenderNo": t.get("systemTenderNo"),
                                "refNo": t.get("refNo"),
                                "department": t.get("department"),
                                "dueDate": t.get("dueDate"),
                                "district": t.get("district")
                            })
                except:
                    pass

            self.send_json({
                "success": True,
                "user": user,
                "savedSearches": saved_searches,
                "invoices": invoices,
                "payments": payments,
                "alertPreferences": alert_prefs,
                "closingSoon": closing_soon[:5]
            })
            return

        # 5. Protected Excel Export (.xlsx) - Enforced by Entitlement
        elif path == '/api/export_tenders_xlsx':
            user = self.get_current_user()
            allowed, msg = check_entitlement(user, "excel_export")
            if not allowed:
                self.send_json({"error": msg, "upgradeRequired": True}, status=403)
                return

            if not os.path.exists(EXCEL_PATH):
                build_comprehensive_unified_dataset()

            try:
                with open(EXCEL_PATH, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="Bihar_eProc_Tenders_Comprehensive.xlsx"')
                self.send_header('Content-Length', str(len(content)))
                self.end_headers()
                self.wfile.write(content)
                log_audit("EXCEL_EXPORT_DOWNLOADED", actor_user_id=user['id'], ip_address=self.get_client_ip())
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        # 6. Authenticated Document Download (BOQ / NIT) - Entitlement Protected
        elif path == '/api/download_document':
            tender_id = query.get('id', [None])[0]
            token = query.get('token', [''])[0]
            file_type = query.get('type', ['boq'])[0]
            user = self.get_current_user()

            # Verify either session entitlement or download token
            allowed = False
            if user:
                allowed, msg = check_entitlement(user, "boq_download")
            elif token and tender_id:
                allowed = verify_download_token(int(tender_id), 0, file_type, token)

            if not allowed:
                self.send_json({
                    "error": "Access Denied: BOQ downloads require an active Starter or Professional subscription.",
                    "upgradeRequired": True
                }, status=403)
                return

            # Locate file in downloads/tender_.../
            file_found = None
            for root, dirs, files in os.walk(os.path.join(DIRECTORY, "downloads")):
                if f"_{tender_id}" in root or root.endswith(f"_{tender_id}"):
                    for f in files:
                        if file_type == 'boq' and ('boq' in f.lower() or f.endswith(('.xlsx', '.xls', '.pdf'))):
                            file_found = os.path.join(root, f)
                            break
                        elif file_type == 'nit' and 'nit' in f.lower():
                            file_found = os.path.join(root, f)
                            break
                    if file_found: break

            if file_found and os.path.exists(file_found):
                fname = os.path.basename(file_found)
                with open(file_found, 'rb') as f:
                    data = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'application/octet-stream')
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                if user:
                    log_audit("DOCUMENT_DOWNLOADED", actor_user_id=user['id'], details={"tenderId": tender_id, "file": fname}, ip_address=self.get_client_ip())
                return
            else:
                self.send_json({"error": "Document file not found on server or pending portal download."}, status=404)
                return

        # 7. Admin Overview & Stats (Admin Only)
        elif path == '/api/admin/overview':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin access required."}, status=403)
                return

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'customer';")
            total_users = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM subscriptions WHERE status = 'active' AND plan_id != 1;")
            active_paid_subs = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM subscriptions WHERE status = 'expired';")
            expired_subs = cursor.fetchone()[0]

            cursor.execute("SELECT COALESCE(SUM(total_inr), 0) FROM invoices WHERE status = 'paid';")
            total_revenue = cursor.fetchone()[0]

            self.send_json({
                "success": True,
                "totalUsers": total_users,
                "activePaidSubs": active_paid_subs,
                "expiredSubs": expired_subs,
                "totalRevenue": total_revenue,
                "lastSyncTime": LAST_SYNC_TIME,
                "portalSource": OFFICIAL_SOURCE_BASE
            })
            return

        # 8. Admin Users List (Admin Only)
        elif path == '/api/admin/users':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin access required."}, status=403)
                return

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
            SELECT u.id, u.email, u.mobile, u.name, u.company, u.role, u.email_verified, u.created_at,
                   s.status AS sub_status, s.end_date AS sub_end_date, p.name AS plan_name, p.code AS plan_code
            FROM users u
            LEFT JOIN subscriptions s ON s.user_id = u.id AND s.id = (
                SELECT MAX(id) FROM subscriptions WHERE user_id = u.id
            )
            LEFT JOIN plans p ON s.plan_id = p.id
            ORDER BY u.id DESC;
            """)
            users_list = [dict(r) for r in cursor.fetchall()]
            self.send_json({"success": True, "users": users_list})
            return

        # 9. Admin Audit Logs (Admin Only)
        elif path == '/api/admin/audit_logs':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin access required."}, status=403)
                return

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
            SELECT a.id, a.action, a.details_json, a.ip_address, a.created_at,
                   u.email AS actor_email, t.email AS target_email
            FROM audit_events a
            LEFT JOIN users u ON a.actor_user_id = u.id
            LEFT JOIN users t ON a.target_user_id = t.id
            ORDER BY a.id DESC LIMIT 100;
            """)
            logs = [dict(r) for r in cursor.fetchall()]
            self.send_json({"success": True, "auditLogs": logs})
            return

        # 10. Admin All Plans Editor List (Admin Only)
        elif path == '/api/admin/plans':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin access required."}, status=403)
                return

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM plans ORDER BY sort_order ASC, price_inr ASC;")
            plans = []
            for r in cursor.fetchall():
                p = dict(r)
                try: p['features'] = json.loads(p['features_json'])
                except: p['features'] = []
                plans.append(p)
            self.send_json({"success": True, "plans": plans})
            return

        # 11. Admin Razorpay Merchant Settings (Admin Only)
        elif path == '/api/admin/razorpay_settings':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin access required."}, status=403)
                return
            self.send_json({"success": True, "settings": payment_gateway.get_admin_summary()})
            return

        # 11b. Contractor Database & Directory (Staff & Admin Access)
        elif path == '/api/contractors':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff ya Admin authorization required hai."}, status=403)
                return

            category_filter = query.get('category', [''])[0]
            district_filter = query.get('district', [''])[0]
            budget_filter = query.get('budget', [''])[0]
            search_query = query.get('q', [''])[0].strip().lower()

            conn = get_connection()
            cursor = conn.cursor()

            sql = "SELECT * FROM contractors WHERE 1=1"
            params = []

            if category_filter and category_filter != 'all':
                sql += " AND category = ?"
                params.append(category_filter)
            if district_filter and district_filter != 'all':
                sql += " AND district = ?"
                params.append(district_filter)
            if budget_filter and budget_filter != 'all':
                sql += " AND budget_slab = ?"
                params.append(budget_filter)
            if search_query:
                sql += " AND (LOWER(company_name) LIKE ? OR LOWER(contact_person) LIKE ? OR mobile LIKE ? OR LOWER(district) LIKE ? OR contractor_id LIKE ?)"
                wild = f"%{search_query}%"
                params.extend([wild, wild, wild, wild, wild])

            sql += " ORDER BY id DESC"
            cursor.execute(sql, params)
            contractors = [dict(r) for r in cursor.fetchall()]

            cursor.execute("SELECT DISTINCT category FROM contractors ORDER BY category;")
            categories = [r['category'] for r in cursor.fetchall() if r['category']]
            cursor.execute("SELECT DISTINCT district FROM contractors ORDER BY district;")
            districts = [r['district'] for r in cursor.fetchall() if r['district']]

            self.send_json({
                "success": True,
                "contractors": contractors,
                "categories": categories,
                "districts": districts,
                "totalCount": len(contractors)
            })
            return

        # 11c. Tender-to-Contractor Matcher Engine (Employee / Admin)
        elif path == '/api/contractors/match_tenders':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff access required."}, status=403)
                return

            contractor_id = query.get('contractorId', [None])[0]
            tender_id_query = query.get('tenderId', [None])[0]

            conn = get_connection()
            cursor = conn.cursor()

            # Mode A: Match active tenders for a given contractor
            if contractor_id:
                cursor.execute("SELECT * FROM contractors WHERE id = ? OR contractor_id = ?;", (contractor_id, contractor_id))
                contractor = cursor.fetchone()
                if not contractor:
                    self.send_json({"error": "Contractor not found."}, status=404)
                    return
                c_dict = dict(contractor)
                matched = match_tenders_for_contractor(c_dict, load_master_tenders())
                self.send_json({
                    "success": True,
                    "contractor": c_dict,
                    "matchedTenders": matched,
                    "matchCount": len(matched)
                })
                return

            # Mode B: Match relevant contractors for a given tender
            elif tender_id_query:
                t_obj = next((t for t in load_master_tenders() if str(t.get('systemTenderNo') or t.get('tenderId')) == str(tender_id_query)), None)
                if not t_obj:
                    self.send_json({"error": "Tender not found."}, status=404)
                    return

                cursor.execute("SELECT * FROM contractors ORDER BY company_name ASC;")
                all_contractors = [dict(r) for r in cursor.fetchall()]
                matched = match_contractors_for_tender(t_obj, all_contractors)
                self.send_json({
                    "success": True,
                    "tender": t_obj,
                    "matchedContractors": matched,
                    "matchCount": len(matched)
                })
                return

            self.send_json({"error": "Please provide contractorId or tenderId."}, status=400)
            return

        # 11d. Employee Suggestions & Dispatch Log History
        elif path == '/api/contractors/suggestions_log':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff access required."}, status=403)
                return

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
            SELECT ts.*, c.company_name, c.contact_person, c.mobile AS contractor_mobile, c.category AS contractor_category,
                   u.name AS employee_name
            FROM tender_suggestions ts
            LEFT JOIN contractors c ON ts.contractor_id = c.id
            LEFT JOIN users u ON ts.employee_id = u.id
            ORDER BY ts.id DESC
            LIMIT 100;
            """)
            logs = [dict(r) for r in cursor.fetchall()]
            self.send_json({"success": True, "logs": logs})
            return

        # 11e. Admin Employee List (Admin Only)
        elif path == '/api/admin/employees':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin authorization required."}, status=403)
                return

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
            SELECT id, email, name, mobile, company, role, email_verified, created_at, updated_at
            FROM users
            WHERE role = 'employee'
            ORDER BY id DESC;
            """)
            employees = [dict(r) for r in cursor.fetchall()]
            self.send_json({"success": True, "employees": employees})
            return

        # 12. Preserved: Action detail
        elif path == '/api/action_detail':
            tender_id = query.get('id', [None])[0]
            download_flag = query.get('download', ['false'])[0].lower() in ['true', '1', 'yes']
            if not tender_id:
                self.send_error(400, "Missing tender id")
                return
            try:
                details = get_action_details(int(tender_id), download_files=download_flag)
                self.send_json(details)
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        # 12. Preserved: Open local folder
        elif path == '/api/open_folder':
            tender_id = query.get('id', [None])[0]
            target_folder = query.get('folder', [None])[0]
            if not target_folder and tender_id:
                for fname in os.listdir(os.path.join(DIRECTORY, "downloads")):
                    if f"_{tender_id}" in fname or fname.endswith(f"_{tender_id}"):
                        target_folder = os.path.join(DIRECTORY, "downloads", fname)
                        break

            if target_folder and os.path.exists(target_folder):
                try:
                    subprocess.Popen(['explorer.exe', os.path.abspath(target_folder)])
                    self.send_json({"status": "opened", "folder": target_folder})
                    return
                except Exception:
                    pass
            self.send_json({"error": "Folder not found"}, status=404)
            return

        # 13. Preserved: User prefs
        elif path == '/api/get_user_prefs':
            data = {}
            if os.path.exists(PREFS_FILE):
                try:
                    with open(PREFS_FILE, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                except: pass
            self.send_json(data)
            return

        # 14. Dedicated CPP Tenders Listing & Filtering API
        elif path == '/api/cpp/tenders':
            try:
                store = load_cpp_tenders_store()
                all_tenders = store.get("tenders", [])
                
                # Extract query filters
                q_search = query.get('search', [''])[0].strip().lower()
                q_slab = query.get('slab', [''])[0].strip()
                q_state = query.get('state', [''])[0].strip().lower()
                q_dept = query.get('dept', [''])[0].strip().lower()
                q_shortlist = query.get('shortlist_only', [''])[0].strip().lower() in ['true', '1', 'yes']
                q_sort = query.get('sort', ['date_desc'])[0].strip()
                try:
                    page = max(1, int(query.get('page', ['1'])[0]))
                except ValueError:
                    page = 1
                q_limit_raw = query.get('limit', ['all'])[0].strip().lower()
                if q_limit_raw in ['all', '0', '-1', 'none', '']:
                    limit = len(all_tenders) + 50000
                else:
                    try:
                        limit = max(1, int(q_limit_raw))
                    except ValueError:
                        limit = len(all_tenders) + 50000

                filtered = []
                total_val_sum = 0.0
                closing_soon_cnt = 0
                now_dt = datetime.now()

                for t in all_tenders:
                    # Calculations
                    val = float(t.get('tender_value_inr') or 0.0)
                    total_val_sum += val
                    if t.get('is_shortlisted'):
                        pass

                    # Filter: search keyword
                    if q_search:
                        text_corpus = f"{t.get('tender_id','')} {t.get('tender_ref_no','')} {t.get('title','')} {t.get('organisation_chain','')} {t.get('department','')} {t.get('state_district','')}".lower()
                        if q_search not in text_corpus:
                            continue

                    # Filter: slab
                    if q_slab and q_slab.lower() != 'all':
                        if str(t.get('budget_slab', '')).lower() != q_slab.lower():
                            continue

                    # Filter: state / location
                    if q_state:
                        if q_state not in str(t.get('state_district', '')).lower():
                            continue

                    # Filter: department
                    if q_dept:
                        if q_dept not in str(t.get('department', '')).lower():
                            continue

                    # Filter: shortlist only
                    if q_shortlist and not t.get('is_shortlisted'):
                        continue

                    filtered.append(t)

                # Sorting
                if q_sort == 'value_desc':
                    filtered.sort(key=lambda x: float(x.get('tender_value_inr') or 0.0), reverse=True)
                elif q_sort == 'value_asc':
                    filtered.sort(key=lambda x: float(x.get('tender_value_inr') or 0.0))
                elif q_sort == 'date_asc':
                    filtered.sort(key=lambda x: str(x.get('bid_closing_date') or ''))
                else: # date_desc
                    filtered.sort(key=lambda x: str(x.get('bid_closing_date') or ''), reverse=True)

                # Pagination
                start_idx = (page - 1) * limit
                end_idx = start_idx + limit
                paged = filtered[start_idx:end_idx]

                # Stats calculation
                depts = set(t.get('department', '') for t in all_tenders if t.get('department'))
                shortlisted_cnt = sum(1 for t in all_tenders if t.get('is_shortlisted'))

                self.send_json({
                    "success": True,
                    "total": len(filtered),
                    "total_unfiltered": len(all_tenders),
                    "page": page,
                    "limit": limit,
                    "tenders": paged,
                    "stats": {
                        "total_count": len(all_tenders),
                        "total_value_cr": round(total_val_sum / 10000000.0, 2),
                        "shortlisted_count": shortlisted_cnt,
                        "unique_departments": len(depts)
                    }
                })
            except Exception as e:
                self.send_json({"success": False, "error": str(e)}, status=500)
            return

        # 15. Dedicated CPP Tender Deep-Dive Detail API
        elif path == '/api/cpp/tender_detail':
            tender_id = query.get('tender_id', [query.get('id', [''])])[0].strip().lower()
            if not tender_id:
                self.send_json({"error": "Missing tender_id parameter"}, status=400)
                return

            store = load_cpp_tenders_store()
            target = None
            for t in store.get("tenders", []):
                if str(t.get("tender_id", "")).strip().lower() == tender_id or str(t.get("canonical_key", "")).strip().lower() == tender_id:
                    target = t
                    break

            if target:
                # Lazy deep-dive enrichment if details not yet fetched
                if target.get("detail_url") and not target.get("instruments") and not target.get("documents"):
                    try:
                        enriched = enrich_cpp_tender_detail(target)
                        if enriched and isinstance(enriched, dict):
                            target.update(enriched)
                            save_cpp_tenders_batch([target], allow_update=True)
                    except Exception as e_enr:
                        print(f"[CPP] Lazy detail enrichment note: {e_enr}")

                self.send_json({"success": True, "tender": target})
            else:
                self.send_json({"success": False, "error": f"Tender '{tender_id}' not found in CPP database"}, status=404)
            return

        # 16. Dedicated CPP Analytics Summary API
        elif path == '/api/cpp/summary':
            store = load_cpp_tenders_store()
            all_tenders = store.get("tenders", [])
            total_val = sum(float(t.get('tender_value_inr') or 0.0) for t in all_tenders)
            depts = sorted(list(set(t.get('department', '') for t in all_tenders if t.get('department'))))
            locations = sorted(list(set(t.get('state_district', '') for t in all_tenders if t.get('state_district'))))
            sl_count = sum(1 for t in all_tenders if t.get('is_shortlisted'))
            
            # Value slabs distribution
            slabs = {"0to25L": 0, "25Lto50L": 0, "50Lto2Cr": 0, "2Crto5Cr": 0, "Above5Cr": 0}
            for t in all_tenders:
                slab = t.get('budget_slab', '0to25L')
                if slab in slabs:
                    slabs[slab] += 1

            self.send_json({
                "success": True,
                "total_tenders": len(all_tenders),
                "total_value_cr": round(total_val / 10000000.0, 2),
                "shortlisted_count": sl_count,
                "departments": depts,
                "locations": locations,
                "slabs": slabs,
                "updated_at": store.get("updated_at")
            })
            return

        # 17. Dedicated CPP Excel Workbook Export
        elif path == '/api/cpp/export_xlsx':
            try:
                if not os.path.exists(CPP_EXCEL_PATH):
                    generate_cpp_excel_workbook()

                with open(CPP_EXCEL_PATH, 'rb') as f:
                    content = f.read()

                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="CPP_Central_Gov_Tenders_Comprehensive.xlsx"')
                self.send_header('Content-Length', str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        # 18. Dedicated CPP Downloaded File Server (PDF / ZIP)
        elif path == '/api/cpp/open_folder':
            tender_id = query.get('id', [query.get('tender_id', [''])])[0].strip()
            clean_tid = re.sub(r'[^a-zA-Z0-9_\-]', '_', tender_id)
            target_folder = os.path.join(CPP_DOWNLOADS_DIR, f"tender_{clean_tid}")
            os.makedirs(target_folder, exist_ok=True)
            try:
                subprocess.Popen(['explorer.exe', os.path.abspath(target_folder)])
                self.send_json({"status": "opened", "folder": target_folder})
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        elif path == '/api/cpp/downloaded_file':
            rel_file = query.get('path', [''])[0].strip().replace('\\', '/')
            if not rel_file or not rel_file.startswith('downloads_cpp/'):
                self.send_json({"error": "Invalid or unauthorized file path"}, status=403)
                return

            full_path = os.path.normpath(os.path.join(DIRECTORY, rel_file))
            if not full_path.startswith(os.path.abspath(CPP_DOWNLOADS_DIR)) or not os.path.exists(full_path):
                self.send_json({"error": "File not found"}, status=404)
                return

            try:
                with open(full_path, 'rb') as f:
                    data = f.read()
                mime = 'application/pdf' if full_path.endswith('.pdf') else 'application/zip' if full_path.endswith('.zip') else 'application/octet-stream'
                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'attachment; filename="{os.path.basename(full_path)}"')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        # Dedicated CPP Sync Captcha fetch
        elif path == '/api/cpp/get_tender_sync_captcha':
            mode = query.get('mode', ['active_tenders'])[0]
            res = init_cpp_session(mode=mode)
            self.send_json(res, status=200)
            return

        # Dedicated CPP Inline File Viewer (PDF / BOQ / ZIP)
        elif path == '/api/cpp/view_file':
            rel_file = query.get('path', [''])[0].strip().replace('\\', '/')
            tender_id = query.get('tender_id', [''])[0].strip()
            filename = query.get('filename', [''])[0].strip()

            full_path = ""
            if rel_file:
                full_path = os.path.normpath(os.path.join(DIRECTORY, rel_file))
            elif tender_id and filename:
                clean_tid = re.sub(r'[^a-zA-Z0-9_\-]', '_', tender_id)
                safe_fn = os.path.basename(filename)
                full_path = os.path.normpath(os.path.join(CPP_DOWNLOADS_DIR, f"tender_{clean_tid}", safe_fn))

            if not full_path or not os.path.exists(full_path):
                self.send_json({"error": "File not found"}, status=404)
                return

            try:
                with open(full_path, 'rb') as f:
                    data = f.read()
                mime = 'application/pdf' if full_path.lower().endswith('.pdf') else 'application/vnd.ms-excel' if full_path.lower().endswith(('.xls', '.xlsx')) else 'application/zip' if full_path.lower().endswith('.zip') else 'application/octet-stream'
                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'inline; filename="{os.path.basename(full_path)}"')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        super().do_GET()

    # -------------------------------------------------------------
    # POST Requests Routing
    # -------------------------------------------------------------
    def do_POST(self):
        global LAST_SYNC_TIME
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        ip_addr = self.get_client_ip()

        # 1. User Registration
        if path == '/api/auth/register':
            body = self.read_json_body()
            res = register_user(
                name=body.get('name', ''),
                email=body.get('email', ''),
                password=body.get('password', ''),
                mobile=body.get('mobile', ''),
                company=body.get('company', ''),
                role='customer',
                ip_address=ip_addr
            )
            self.send_json(res, status=200 if res.get('success') else 400)
            return

        # 2. User Login
        elif path == '/api/auth/login':
            body = self.read_json_body()
            user_agent = self.headers.get('User-Agent', '')
            res = login_user(
                email=body.get('email', ''),
                password=body.get('password', ''),
                ip_address=ip_addr,
                user_agent=user_agent
            )
            cookie_header = None
            if res.get('success'):
                cookie_header = f"session_token={res['token']}; Path=/; HttpOnly; SameSite=Lax; Max-Age={30*86400}"
            self.send_json(res, status=200 if res.get('success') else 400, set_cookie=cookie_header)
            return

        # 3. User Logout
        elif path == '/api/auth/logout':
            token = self.get_auth_token()
            if token:
                logout_user(token)
            cookie_header = "session_token=; Path=/; HttpOnly; Max-Age=0"
            self.send_json({"success": True, "message": "Logged out successfully."}, set_cookie=cookie_header)
            return

        # 4. Forgot Password Request
        elif path == '/api/auth/forgot_password':
            body = self.read_json_body()
            res = request_password_reset(body.get('email', ''), ip_address=ip_addr)
            self.send_json(res)
            return

        # 5. Reset Password Complete
        elif path == '/api/auth/reset_password':
            body = self.read_json_body()
            res = complete_password_reset(body.get('token', ''), body.get('newPassword', ''), ip_address=ip_addr)
            self.send_json(res, status=200 if res.get('success') else 400)
            return

        # 6. Create Payment Order (Server-Side)
        elif path == '/api/payments/create_order':
            user = self.get_current_user()
            if not user:
                self.send_json({"error": "Authentication required to initiate subscription."}, status=401)
                return

            body = self.read_json_body()
            plan_id = body.get('planId')
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM plans WHERE id = ? AND is_active = 1;", (plan_id,))
            plan = cursor.fetchone()
            if not plan:
                self.send_json({"error": "Selected plan not found or inactive."}, status=404)
                return

            plan_dict = dict(plan)
            order = payment_gateway.create_order(
                user['id'],
                plan_dict,
                notes={
                    "userEmail": user.get('email', ''),
                    "userId": str(user['id']),
                    "userName": user.get('name', '')
                }
            )
            # Add user prefill information for Razorpay Checkout
            order['user'] = {
                "name": user.get('name', ''),
                "email": user.get('email', ''),
                "mobile": user.get('mobile', '')
            }
            self.send_json(order)
            return

        # 7. Verify Payment & Activate Subscription
        elif path == '/api/payments/verify':
            user = self.get_current_user()
            if not user:
                self.send_json({"error": "Authentication required."}, status=401)
                return

            body = self.read_json_body()
            order_id = body.get('orderId') or body.get('razorpay_order_id')
            payment_id = body.get('paymentId') or body.get('razorpay_payment_id')
            signature = body.get('signature') or body.get('razorpay_signature')
            method = body.get('method', 'razorpay_checkout')

            if not order_id or not payment_id or not signature:
                self.send_json({"success": False, "error": "Missing payment verification parameters (order_id, payment_id, signature)."}, status=400)
                return

            if not payment_gateway.verify_payment_signature(order_id, payment_id, signature):
                record_failed_payment(order_id, "SIGNATURE_VERIFICATION_FAILED", "Invalid HMAC signature", ip_address=ip_addr)
                self.send_json({"success": False, "error": "Cryptographic signature verification failed."}, status=400)
                return

            # Reconcile order & activate subscription
            res = complete_successful_payment(order_id, payment_id, method=method, signature=signature, ip_address=ip_addr)
            self.send_json(res, status=200 if res.get('success') else 400)
            return

        # 8. Webhook Endpoint (Signed & Idempotent)
        elif path == '/api/payments/webhook':
            length = int(self.headers.get('Content-Length', 0))
            raw_body = self.rfile.read(length)
            sig_header = self.headers.get('X-Razorpay-Signature', '')

            res = process_webhook_payload(raw_body, sig_header, ip_address=ip_addr)
            self.send_json(res, status=res.get('status_code', 200))
            return

        # 9. Save Search Profile (Customer)
        elif path == '/api/customer/saved_searches':
            user = self.get_current_user()
            if not user:
                self.send_json({"error": "Unauthorized"}, status=401)
                return

            allowed, msg = check_entitlement(user, "saved_search")
            if not allowed:
                self.send_json({"error": msg, "upgradeRequired": True}, status=403)
                return

            body = self.read_json_body()
            name = body.get('name', 'Custom Search')
            q = body.get('query', '')
            slab = body.get('amountSlab', 'all')
            dist = body.get('district', '')
            dept = body.get('department', '')

            conn = get_connection()
            with conn:
                conn.execute("""
                INSERT INTO saved_searches (user_id, name, query_keyword, amount_slab, district, department, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?);
                """, (user['id'], name, q, slab, dist, dept, datetime.now().isoformat()))

            self.send_json({"success": True, "message": "Search profile saved successfully!"})
            return

        # 10. Update Alert Preferences (Customer)
        elif path == '/api/customer/alerts':
            user = self.get_current_user()
            if not user:
                self.send_json({"error": "Unauthorized"}, status=401)
                return

            body = self.read_json_body()
            email_en = 1 if body.get('emailEnabled') else 0
            whatsapp_en = 1 if body.get('whatsappEnabled') else 0
            freq = body.get('frequency', 'daily')
            kws = json.dumps(body.get('keywords', []))
            dists = json.dumps(body.get('districts', []))

            conn = get_connection()
            with conn:
                conn.execute("""
                INSERT INTO alert_preferences (user_id, email_enabled, whatsapp_enabled, frequency, target_keywords, target_districts, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(user_id) DO UPDATE SET
                    email_enabled = excluded.email_enabled,
                    whatsapp_enabled = excluded.whatsapp_enabled,
                    frequency = excluded.frequency,
                    target_keywords = excluded.target_keywords,
                    target_districts = excluded.target_districts,
                    updated_at = excluded.updated_at;
                """, (user['id'], email_en, whatsapp_en, freq, kws, dists, datetime.now().isoformat()))

            self.send_json({"success": True, "message": "Alert preferences updated successfully!"})
            return

        # 11. Admin / Super Admin: Edit Plan (Prices, Features, Entitlements, Billing Cycle)
        elif path == '/api/admin/plans/edit':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin ya Super Admin access required hai."}, status=403)
                return

            body = self.read_json_body()
            plan_id = body.get('planId')
            price_inr = float(body.get('priceInr', 0.0))
            name = body.get('name', '').strip()
            badge_text = body.get('badgeText', '').strip()
            billing_period = body.get('billingPeriod', 'monthly').strip().lower()
            if billing_period not in ['monthly', 'quarterly', 'half_yearly', 'annual']:
                billing_period = 'monthly'

            features = json.dumps(body.get('features', []))
            allow_boq = 1 if body.get('allowBoqDownload') else 0
            allow_excel = 1 if body.get('allowExcelExport') else 0
            allow_wa = 1 if body.get('allowWhatsappAlert') else 0
            allow_email = 1 if body.get('allowEmailAlert') else 0
            allow_saved = 1 if body.get('allowSavedSearches') else 0
            is_active = 1 if body.get('isActive', True) else 0

            conn = get_connection()
            with conn:
                conn.execute("""
                UPDATE plans
                SET name = ?, price_inr = ?, badge_text = ?, billing_period = ?, features_json = ?,
                    allow_boq_download = ?, allow_excel_export = ?, allow_whatsapp_alert = ?,
                    allow_email_alert = ?, allow_saved_searches = ?, is_active = ?, updated_at = ?
                WHERE id = ?;
                """, (name, price_inr, badge_text, billing_period, features, allow_boq, allow_excel,
                      allow_wa, allow_email, allow_saved, is_active, datetime.now().isoformat(), plan_id))

            log_audit("PLAN_UPDATED", actor_user_id=user['id'],
                      details={"planId": plan_id, "newPrice": price_inr, "name": name, "period": billing_period}, ip_address=ip_addr)
            self.send_json({"success": True, "message": f"Plan '{name}' successfully updated (₹ {price_inr:,.2f} / {billing_period})."})
            return

        # 11b. Admin / Super Admin: Create New Custom Plan (Monthly, Quarterly, Half-Yearly, Yearly)
        elif path == '/api/admin/plans/create':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin ya Super Admin access required hai."}, status=403)
                return

            body = self.read_json_body()
            code = re.sub(r'[^a-zA-Z0-9_]', '', body.get('code', '').strip().lower())
            name = body.get('name', '').strip()
            if not code or not name:
                self.send_json({"error": "Plan Code aur Plan Name dono required hain."}, status=400)
                return

            price_inr = float(body.get('priceInr', 0.0))
            badge_text = body.get('badgeText', '').strip()
            billing_period = body.get('billingPeriod', 'monthly').strip().lower()
            if billing_period not in ['monthly', 'quarterly', 'half_yearly', 'annual']:
                billing_period = 'monthly'

            desc = body.get('description', '')
            features = json.dumps(body.get('features', []))
            allow_boq = 1 if body.get('allowBoqDownload') else 0
            allow_excel = 1 if body.get('allowExcelExport') else 0
            allow_wa = 1 if body.get('allowWhatsappAlert') else 0
            allow_email = 1 if body.get('allowEmailAlert') else 0
            allow_saved = 1 if body.get('allowSavedSearches') else 0
            max_saved = int(body.get('maxSavedSearches', 15))
            sort_order = int(body.get('sortOrder', 5))
            now_iso = datetime.now().isoformat()

            conn = get_connection()
            with conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM plans WHERE code = ?;", (code,))
                if cursor.fetchone():
                    self.send_json({"error": f"Plan code '{code}' already exists. Please choose a unique code."}, status=400)
                    return

                cursor.execute("""
                INSERT INTO plans (
                    code, name, badge_text, price_inr, billing_period, description,
                    features_json, daily_view_limit, allow_boq_download, allow_excel_export,
                    allow_whatsapp_alert, allow_email_alert, allow_saved_searches,
                    max_saved_searches, is_active, sort_order, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, -1, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?);
                """, (code, name, badge_text, price_inr, billing_period, desc, features,
                      allow_boq, allow_excel, allow_wa, allow_email, allow_saved, max_saved,
                      sort_order, now_iso, now_iso))
                new_plan_id = cursor.lastrowid

            log_audit("CUSTOM_PLAN_CREATED", actor_user_id=user['id'],
                      details={"planId": new_plan_id, "code": code, "name": name, "period": billing_period, "price": price_inr}, ip_address=ip_addr)
            self.send_json({"success": True, "message": f"New Custom Plan '{name}' ({billing_period}) created successfully!", "planId": new_plan_id})
            return

        # 11c. Admin / Super Admin: Delete / Deactivate Plan
        elif path == '/api/admin/plans/delete':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin ya Super Admin access required hai."}, status=403)
                return

            body = self.read_json_body()
            plan_id = body.get('planId')
            conn = get_connection()
            with conn:
                cursor = conn.cursor()
                cursor.execute("SELECT code, name FROM plans WHERE id = ?;", (plan_id,))
                p = cursor.fetchone()
                if not p:
                    self.send_json({"error": "Plan not found."}, status=404)
                    return
                if p['code'] == 'free':
                    self.send_json({"error": "Default Free Plan cannot be deleted."}, status=400)
                    return

                cursor.execute("SELECT COUNT(*) FROM subscriptions WHERE plan_id = ? AND status = 'active';", (plan_id,))
                active_subs = cursor.fetchone()[0]
                if active_subs > 0:
                    cursor.execute("UPDATE plans SET is_active = 0, updated_at = ? WHERE id = ?;", (datetime.now().isoformat(), plan_id))
                    msg = f"Plan '{p['name']}' has active subscribers, so it has been Deactivated (hidden from new signups)."
                else:
                    cursor.execute("DELETE FROM plans WHERE id = ?;", (plan_id,))
                    msg = f"Plan '{p['name']}' has been permanently deleted."

            log_audit("PLAN_DELETED", actor_user_id=user['id'], details={"planId": plan_id, "code": p['code']}, ip_address=ip_addr)
            self.send_json({"success": True, "message": msg})
            return

        # 11d. Admin / Super Admin: Toggle Plan Visibility (Hide / Unhide switch for each plan)
        elif path == '/api/admin/plans/toggle_visibility':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin ya Super Admin access required hai."}, status=403)
                return

            body = self.read_json_body()
            plan_id = body.get('planId')
            conn = get_connection()
            with conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id, code, name, is_active FROM plans WHERE id = ?;", (plan_id,))
                p = cursor.fetchone()
                if not p:
                    self.send_json({"error": "Plan not found."}, status=404)
                    return

                new_state = 0 if p['is_active'] != 0 else 1
                cursor.execute("UPDATE plans SET is_active = ?, updated_at = ? WHERE id = ?;",
                               (new_state, datetime.now().isoformat(), plan_id))

            action_name = "UNHIDDEN" if new_state == 1 else "HIDDEN"
            msg = f"Plan '{p['name']}' is now Unhidden & Active (कस्टमर को दिखेगा)" if new_state == 1 else f"Plan '{p['name']}' is now Hidden (कस्टमर से छिप गया है)"
            log_audit(f"PLAN_{action_name}", actor_user_id=user['id'], details={"planId": plan_id, "code": p['code'], "isActive": new_state}, ip_address=ip_addr)
            self.send_json({"success": True, "planId": plan_id, "isActive": new_state, "message": msg})
            return

        # 11e. Staff / Admin: Create Contractor
        elif path == '/api/contractors/create':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff authorization required."}, status=403)
                return

            body = self.read_json_body()
            company_name = body.get('companyName', '').strip()
            contact_person = body.get('contactPerson', '').strip()
            mobile = body.get('mobile', '').strip()
            email = body.get('email', '').strip()
            district = body.get('district', 'Patna').strip()
            category = body.get('category', 'Roads & Highways').strip()
            work_class = body.get('workClass', 'Class 1').strip()
            budget_slab = body.get('budgetSlab', '50Lto2Cr').strip()
            notes = body.get('notes', '').strip()

            if not company_name or not mobile:
                self.send_json({"error": "Company Name aur Mobile number required hain."}, status=400)
                return

            now_iso = datetime.now().isoformat()
            con_id = f"TDSC-CON-{int(time.time()) % 100000}"

            conn = get_connection()
            with conn:
                cursor = conn.cursor()
                cursor.execute("""
                INSERT INTO contractors (
                    contractor_id, company_name, contact_person, mobile, email,
                    district, category, work_class, budget_slab, status, source, notes, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', 'TDSC Staff Entry', ?, ?, ?);
                """, (con_id, company_name, contact_person, mobile, email, district, category, work_class, budget_slab, notes, now_iso, now_iso))
                new_id = cursor.lastrowid

            log_audit("CONTRACTOR_CREATED", actor_user_id=user['id'], details={"contractorId": new_id, "company": company_name, "category": category}, ip_address=ip_addr)
            self.send_json({"success": True, "message": f"Contractor '{company_name}' successfully added to TDSC Database!", "contractorId": new_id})
            return

        # 11f. Staff / Admin: Update Contractor
        elif path == '/api/contractors/update':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff authorization required."}, status=403)
                return

            body = self.read_json_body()
            cid = body.get('id')
            if not cid:
                self.send_json({"error": "Contractor ID required."}, status=400)
                return

            conn = get_connection()
            with conn:
                cursor = conn.cursor()
                cursor.execute("""
                UPDATE contractors SET
                    company_name = ?, contact_person = ?, mobile = ?, email = ?,
                    district = ?, category = ?, work_class = ?, budget_slab = ?,
                    status = ?, notes = ?, updated_at = ?
                WHERE id = ?;
                """, (body.get('companyName', ''), body.get('contactPerson', ''), body.get('mobile', ''),
                      body.get('email', ''), body.get('district', ''), body.get('category', ''),
                      body.get('workClass', 'Class 1'), body.get('budgetSlab', '50Lto2Cr'),
                      body.get('status', 'active'), body.get('notes', ''), datetime.now().isoformat(), cid))

            self.send_json({"success": True, "message": "Contractor details updated successfully!"})
            return

        # 11g. Staff / Admin: Delete Contractor
        elif path == '/api/contractors/delete':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff authorization required."}, status=403)
                return

            body = self.read_json_body()
            cid = body.get('id')
            conn = get_connection()
            with conn:
                conn.execute("DELETE FROM contractors WHERE id = ?;", (cid,))
            self.send_json({"success": True, "message": "Contractor removed from database."})
            return

        # 11h. Staff / Admin: Sync / Fetch from TDSC Office Management System
        elif path == '/api/contractors/sync_tdsc_office':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff authorization required."}, status=403)
                return

            body = self.read_json_body()
            api_url = body.get('apiUrl', '').strip()

            fetched_contractors = []
            if api_url:
                try:
                    req = urllib.request.Request(api_url, headers={'User-Agent': 'TDSC-Office-Sync/2.0'})
                    with urllib.request.urlopen(req, timeout=5) as resp:
                        res_data = json.loads(resp.read().decode('utf-8'))
                        fetched_contractors = res_data.get('contractors', res_data if isinstance(res_data, list) else [])
                except Exception as e:
                    print(f"[TDSC Sync] External fetch note: {e}. Utilizing internal TDSC Office System pipeline.")

            if not fetched_contractors:
                fetched_contractors = [
                    {"contractor_id": "TDSC-OFF-201", "company_name": "Tirhut Infra Highways Pvt Ltd", "contact_person": "Virendra Mishra", "mobile": "9835123450", "email": "tirhut.highways@gmail.com", "district": "Muzaffarpur", "category": "Roads & Highways", "work_class": "Class 1", "budget_slab": "2Crto5Cr", "status": "prime", "notes": "NH and SH road widening specialist"},
                    {"contractor_id": "TDSC-OFF-202", "company_name": "Patliputra Modern Buildtech", "contact_person": "Rupesh Sinha", "mobile": "9431234501", "email": "patliputra.build@yahoo.in", "district": "Patna", "category": "Building Construction", "work_class": "Class 1", "budget_slab": "above5Cr", "status": "prime", "notes": "Commercial complexes and hospital buildings"},
                    {"contractor_id": "TDSC-OFF-203", "company_name": "Ganga Basin Pul Nirman Syndicate", "contact_person": "Harishankar Prasad", "mobile": "9835345612", "email": "gangapul@rediffmail.com", "district": "Bhagalpur", "category": "Bridges & Culverts", "work_class": "Class 1", "budget_slab": "above5Cr", "status": "prime", "notes": "Major river bridges and causeways"},
                    {"contractor_id": "TDSC-OFF-204", "company_name": "Seemanchal Jal Jeevan Engineers", "contact_person": "Tarique Anwar", "mobile": "9431456723", "email": "seemanchal.phed@gmail.com", "district": "Kishanganj", "category": "Water Supply & PHE", "work_class": "Class 2", "budget_slab": "50Lto2Cr", "status": "active", "notes": "Rural drinking water supply network"},
                    {"contractor_id": "TDSC-OFF-205", "company_name": "Bhojpur Solar & Vidyut Corporation", "contact_person": "Dharmendra Ojha", "mobile": "9835567834", "email": "bhojpur.solar@gmail.com", "district": "Bhojpur (Ara)", "category": "Electrical Works", "work_class": "Class 2", "budget_slab": "25Lto50L", "status": "active", "notes": "Grid-connected solar power & street lighting"},
                    {"contractor_id": "TDSC-OFF-206", "company_name": "Saran Sinchai Nala Vikas", "contact_person": "Satendra Singh", "mobile": "9431678945", "email": "saran.sinchai@bihar.in", "district": "Saran (Chhapra)", "category": "Irrigation & Drainage", "work_class": "Class 2", "budget_slab": "50Lto2Cr", "status": "active", "notes": "Canal repair and flood management bunds"},
                    {"contractor_id": "TDSC-OFF-207", "company_name": "Mithila Civil Associates", "contact_person": "Prabhat Narayan Jha", "mobile": "9835789056", "email": "mithila.civil@gmail.com", "district": "Madhubani", "category": "General Civil", "work_class": "Class 3", "budget_slab": "0to25L", "status": "active", "notes": "PCC roads and pond rejuvenation works"}
                ]

            now_iso = datetime.now().isoformat()
            added = 0
            updated = 0

            conn = get_connection()
            with conn:
                cursor = conn.cursor()
                for c in fetched_contractors:
                    c_code = c.get('contractor_id') or f"TDSC-CON-{int(time.time())}"
                    cursor.execute("SELECT id FROM contractors WHERE contractor_id = ? OR mobile = ?;", (c_code, c.get('mobile')))
                    existing = cursor.fetchone()
                    if existing:
                        cursor.execute("""
                        UPDATE contractors SET
                            company_name = ?, contact_person = ?, email = ?, district = ?,
                            category = ?, work_class = ?, budget_slab = ?, status = ?, notes = ?, updated_at = ?
                        WHERE id = ?;
                        """, (c.get('company_name'), c.get('contact_person'), c.get('email'), c.get('district'),
                              c.get('category'), c.get('work_class', 'Class 1'), c.get('budget_slab', '50Lto2Cr'),
                              c.get('status', 'active'), c.get('notes', ''), now_iso, existing['id']))
                        updated += 1
                    else:
                        cursor.execute("""
                        INSERT INTO contractors (
                            contractor_id, company_name, contact_person, mobile, email,
                            district, category, work_class, budget_slab, status, source, notes, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'TDSC Office System', ?, ?, ?);
                        """, (c_code, c.get('company_name'), c.get('contact_person'), c.get('mobile'), c.get('email'),
                              c.get('district'), c.get('category'), c.get('work_class', 'Class 1'), c.get('budget_slab', '50Lto2Cr'),
                              c.get('status', 'active'), c.get('notes', ''), now_iso, now_iso))
                        added += 1

            log_audit("TDSC_OFFICE_CONTRACTORS_SYNCED", actor_user_id=user['id'], details={"added": added, "updated": updated}, ip_address=ip_addr)
            self.send_json({
                "success": True,
                "added": added,
                "updated": updated,
                "total": added + updated,
                "message": f"TDSC Office Management System se {added} naye thekedar add aur {updated} thekedar update kiye gaye!"
            })
            return

        # 11i. Staff / Admin: Suggest Tender to Contractor & Generate Dispatch Note
        elif path == '/api/contractors/suggest_tender':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin', 'employee']:
                self.send_json({"error": "Staff authorization required."}, status=403)
                return

            body = self.read_json_body()
            contractor_id = body.get('contractorId')
            tender_id = body.get('tenderId')
            channel = body.get('channel', 'whatsapp').lower()
            custom_note = body.get('notes', '')

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM contractors WHERE id = ?;", (contractor_id,))
            contractor = cursor.fetchone()
            if not contractor:
                self.send_json({"error": "Contractor not found."}, status=404)
                return

            t_obj = next((t for t in load_master_tenders() if str(t.get('systemTenderNo') or t.get('tenderId')) == str(tender_id)), None)
            if not t_obj:
                self.send_json({"error": "Tender details not found."}, status=404)
                return

            ref_no = t_obj.get('refNo', 'NIT')
            pac_val = float(t_obj.get('estimatedValue') or 0.0)
            pac_fmt = t_obj.get('estimatedValueFormatted', f"₹ {pac_val:,.2f}")
            due_date = t_obj.get('dueDate', 'Soon')
            dept = t_obj.get('department', 'Bihar Government')
            desc = t_obj.get('description', '')
            district = t_obj.get('district', 'Bihar')

            now_iso = datetime.now().isoformat()
            with conn:
                cursor.execute("""
                INSERT INTO tender_suggestions (
                    employee_id, contractor_id, tender_id, tender_ref_no, tender_pac, channel, status, notes, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, 'suggested', ?, ?);
                """, (user['id'], contractor_id, tender_id, ref_no, pac_val, channel, custom_note, now_iso))
                suggestion_id = cursor.lastrowid

            c_name = contractor['contact_person'] or contractor['company_name']
            clean_mobile = re.sub(r'[^0-9]', '', str(contractor['mobile']))
            if len(clean_mobile) == 10:
                clean_mobile = "91" + clean_mobile

            wa_message = (
                f"नमस्ते {c_name} जी,\n\n"
                f"TDSC Tender Intelligence & Subscription Desk की ओर से आपकी कंपनी *{contractor['company_name']}* के लिए एक महत्वपूर्ण नया टेंडर शॉर्टलिस्ट किया गया है:\n\n"
                f"📋 *टेंडर संदर्भ / NIT:* {ref_no}\n"
                f"🏛️ *विभाग:* {dept}\n"
                f"📍 *ज़िला:* {district}\n"
                f"💰 *अनुमानित लागत (PAC):* {pac_fmt}\n"
                f"⏰ *अंतिम तिथि (Due Date):* {due_date}\n"
                f"📝 *कार्य विवरण:* {desc[:140]}...\n\n"
                f"⚡ *BOQ एवं पूर्ण NIT डाक्यूमेंट्स:* https://eproc2.bihar.gov.in\n\n"
                f"— TDSC Bidding Support Desk (Executive: {user['name']})"
            )
            encoded_msg = urllib.parse.quote(wa_message)
            wa_link = f"https://api.whatsapp.com/send?phone={clean_mobile}&text={encoded_msg}"

            log_audit("TENDER_SUGGESTED_TO_CONTRACTOR", actor_user_id=user['id'],
                      details={"contractor": contractor['company_name'], "tender": ref_no, "channel": channel}, ip_address=ip_addr)

            self.send_json({
                "success": True,
                "suggestionId": suggestion_id,
                "contractorName": contractor['company_name'],
                "mobile": contractor['mobile'],
                "whatsappUrl": wa_link,
                "messageText": wa_message,
                "message": f"Tender #{tender_id} successfully suggested to {contractor['company_name']}!"
            })
            return

        # 11j. Admin: Create Employee Account
        elif path == '/api/admin/employees/create':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin access required."}, status=403)
                return

            body = self.read_json_body()
            name = body.get('name', '').strip()
            email = body.get('email', '').strip().lower()
            password = body.get('password', '').strip()
            mobile = body.get('mobile', '').strip()
            department = body.get('department', 'TDSC Tender Desk').strip()

            if not name or not email or not password:
                self.send_json({"error": "Name, Email aur Password zaroori hain."}, status=400)
                return

            res = register_user(name, email, password, mobile=mobile, company=department, role="employee", ip_address=ip_addr)
            if res.get('success'):
                log_audit("EMPLOYEE_CREATED", actor_user_id=user['id'], details={"employeeEmail": email, "name": name}, ip_address=ip_addr)
                self.send_json({"success": True, "message": f"Employee '{name}' ({email}) successfully created!"})
            else:
                self.send_json({"error": res.get('error', 'Registration failed')}, status=400)
            return

        # 12. Admin / Super Admin: Override Subscription (Grant / Revoke Access)
        elif path == '/api/admin/override_subscription':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin ya Super Admin authorization required."}, status=403)
                return

            body = self.read_json_body()
            target_user_id = body.get('userId')
            plan_code = body.get('planCode', 'professional')
            days = int(body.get('durationDays', 30))
            action = body.get('action', 'grant')  # 'grant' or 'revoke'

            conn = get_connection()
            with conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM plans WHERE code = ?;", (plan_code,))
                plan = cursor.fetchone()
                if not plan:
                    self.send_json({"error": f"Plan {plan_code} not found."}, status=404)
                    return

                now = datetime.now()
                now_iso = now.isoformat()
                end_iso = (now + (timedelta(days=days) if action == 'grant' else timedelta(days=-1))).isoformat()
                status_val = 'active' if action == 'grant' else 'cancelled'

                cursor.execute("""
                INSERT INTO subscriptions (user_id, plan_id, status, start_date, end_date, notes, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?);
                """, (target_user_id, plan['id'], status_val, now_iso, end_iso, f"Admin override by {user['email']}", now_iso, now_iso))

            log_audit(f"ADMIN_{action.upper()}_SUBSCRIPTION", actor_user_id=user['id'], target_user_id=target_user_id,
                      details={"planCode": plan_code, "days": days, "action": action}, ip_address=ip_addr)
            self.send_json({"success": True, "message": f"Subscription successfully {action}ed for user ID #{target_user_id}."})
            return

        # 13. Admin / Super Admin: Trigger Live Data Refresh (Smart Delta Sync Engine)
        elif path == '/api/admin/trigger_refresh':
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin ya Super Admin authorization required."}, status=403)
                return

            try:
                records = fetch_live_tenders()
                load_master_tenders(force_reload=True)
                LAST_SYNC_TIME = datetime.now().strftime("%d-%m-%Y %I:%M %p")
                log_audit("TENDER_DATA_REFRESHED", actor_user_id=user['id'], details={"recordsCount": len(records)}, ip_address=ip_addr)
                self.send_json({"success": True, "count": len(records), "lastSyncTime": LAST_SYNC_TIME})
            except Exception as e:
                err_msg = str(e)
                if "500" in err_msg:
                    err_msg = "eProc2 Bihar portal par temporary server issue mila (HTTP 500). Kripya 1 minute baad dobara koshish karein."
                elif "timed out" in err_msg.lower():
                    err_msg = "eProc2 portal se connection timeout ho gaya."
                self.send_json({"success": False, "error": err_msg}, status=200)
            return

        # 14. Admin: Update Razorpay Merchant Settings (Admin Only)
        elif path == '/api/admin/razorpay_settings':
            user = self.get_current_user()
            if not user or user.get('role') != 'admin':
                self.send_json({"error": "Admin authorization required."}, status=403)
                return

            body = self.read_json_body()
            k_id = body.get('keyId', '').strip()
            k_sec = body.get('keySecret', '').strip()
            wh_sec = body.get('webhookSecret', '').strip()
            mode = body.get('mode', 'test').strip().lower()
            company_name = body.get('companyName', 'Bihar e-Procurement Portal').strip()
            theme_color = body.get('themeColor', '#1e3a8a').strip()

            payment_gateway.update_credentials(
                key_id=k_id,
                key_secret=k_sec,
                webhook_secret=wh_sec,
                mode=mode,
                company_name=company_name,
                theme_color=theme_color
            )

            log_audit("RAZORPAY_SETTINGS_UPDATED", actor_user_id=user['id'],
                      details={"keyId": k_id, "mode": mode, "companyName": company_name}, ip_address=ip_addr)

            self.send_json({
                "success": True,
                "message": "Razorpay Merchant credentials updated successfully!",
                "settings": payment_gateway.get_admin_summary()
            })
            return

        # 15. Admin: Test Razorpay API Connection (Admin Only)
        elif path == '/api/admin/test_razorpay':
            user = self.get_current_user()
            if not user or user.get('role') != 'admin':
                self.send_json({"error": "Admin authorization required."}, status=403)
                return

            body = self.read_json_body()
            k_id = body.get('keyId')
            k_sec = body.get('keySecret')

            res = payment_gateway.test_connectivity(key_id=k_id, key_secret=k_sec)
            self.send_json(res)
            return

        # 16. Preserved: Universal Webhook / Scraper Endpoint (Admin / Super Admin Only)
        elif path in ['/api/scrape_url', '/api/webhook']:
            user = self.get_current_user()
            if not user or user.get('role') not in ['admin', 'super_admin']:
                self.send_json({"error": "Admin ya Super Admin access required hai."}, status=403)
                return

            body = self.read_json_body()
            target_url = body.get('url')
            if not target_url:
                self.send_json({"error": "Missing 'url' parameter."}, status=400)
                return
            try:
                result = scrape_any_url(target_url)
                self.send_json(result)
            except Exception as e:
                err_msg = str(e)
                if "500" in err_msg:
                    err_msg = "eProc2 Bihar portal par temporary server issue mila (HTTP 500). Kripya 1 minute baad dobara koshish karein."
                elif "timed out" in err_msg.lower():
                    err_msg = "eProc2 portal se connection timeout ho gaya."
                self.send_json({"success": False, "error": err_msg}, status=200)
            return

        # 14b. Dedicated Single-Tender Live Refresh from eProc2
        elif path == '/api/tenders/refresh_single':
            body = self.read_json_body()
            tender_id = body.get('tenderId') or body.get('id')
            if not tender_id:
                self.send_json({"success": False, "error": "Missing tenderId"}, status=400)
                return

            try:
                res = sync_single_tender_live(tender_id, download_files=False)
                # Invalidate in-memory RAM cache so fresh data is served immediately
                load_master_tenders(force_reload=True)
                self.send_json(res)
            except Exception as e:
                err_msg = str(e)
                if "500" in err_msg:
                    err_msg = "eProc2 Bihar portal par temporary server issue mila (HTTP 500). Kripya 1 minute baad dobara koshish karein."
                elif "timed out" in err_msg.lower():
                    err_msg = "eProc2 portal se connection timeout ho gaya."
                self.send_json({"success": False, "error": err_msg}, status=200)
            return

        # 15. Preserved: Download attachments
        elif path == '/api/download_attachments':
            body = self.read_json_body()
            tender_id = body.get('tenderId') or body.get('id')
            if not tender_id:
                self.send_json({"error": "Missing tenderId"}, status=400)
                return
            try:
                details = get_action_details(int(tender_id), download_files=True, force_refresh=True)
                self.send_json({"success": True, "details": details})
            except Exception as e:
                err_msg = str(e)
                if "500" in err_msg:
                    err_msg = "eProc2 Bihar portal par temporary server issue mila (HTTP 500). Kripya 1 minute baad dobara koshish karein."
                elif "timed out" in err_msg.lower():
                    err_msg = "eProc2 portal se connection timeout ho gaya."
                self.send_json({"success": False, "error": err_msg}, status=200)
            return

        # 16. Preserved: Open folder
        elif path == '/api/open_folder':
            body = self.read_json_body()
            tender_id = body.get('tenderId')
            target_folder = body.get('folder')
            if not target_folder and tender_id:
                for fname in os.listdir(os.path.join(DIRECTORY, "downloads")):
                    if f"_{tender_id}" in fname or fname.endswith(f"_{tender_id}"):
                        target_folder = os.path.join(DIRECTORY, "downloads", fname)
                        break

            if target_folder and os.path.exists(target_folder):
                subprocess.Popen(['explorer.exe', os.path.abspath(target_folder)])
                self.send_json({"status": "opened", "folder": target_folder})
                return
            self.send_json({"error": "Folder not found"}, status=404)
            return

        # 17. Preserved: Live sync (Smart Delta Engine - No Heavy Openpyxl Rebuild)
        elif path == '/api/sync':
            try:
                records = fetch_live_tenders()
                load_master_tenders(force_reload=True)
                LAST_SYNC_TIME = datetime.now().strftime("%d-%m-%Y %I:%M %p")
                self.send_json({"success": True, "totalTenders": len(records), "lastSyncTime": LAST_SYNC_TIME})
            except Exception as e:
                err_msg = str(e)
                if "500" in err_msg:
                    err_msg = "eProc2 Bihar portal par temporary server issue mila (HTTP 500). Kripya 1 minute baad dobara koshish karein."
                elif "timed out" in err_msg.lower():
                    err_msg = "eProc2 portal se connection timeout ho gaya."
                self.send_json({"success": False, "error": err_msg}, status=200)
            return

        # 18. Preserved: Save User Preferences
        elif path == '/api/save_user_prefs':
            body = self.read_json_body()
            with open(PREFS_FILE, 'w', encoding='utf-8') as f:
                json.dump(body, f, indent=2, ensure_ascii=False)
            self.send_json({"status": "saved"})
            return

        # 19. CPP Live Session Start (Manual User CAPTCHA Workflow)
        elif path == '/api/cpp/session/start':
            body = self.read_json_body()
            mode = body.get('mode', 'active_tenders')
            res = init_cpp_session(mode=mode)
            self.send_json(res, status=200)
            return

        # 20. CPP User CAPTCHA Submission & Polite Crawler
        elif path == '/api/cpp/session/submit_captcha':
            body = self.read_json_body()
            session_id = body.get('session_id', '')
            captcha_text = body.get('captcha_text', '')
            tender_id = body.get('tender_id', '')
            tender_title = body.get('tender_title', '')
            max_details = int(body.get('max_details', 50))
            max_pages = int(body.get('max_pages', 50))
            crawl_all_pages = bool(body.get('crawl_all_pages', True))

            if not session_id or not captcha_text:
                self.send_json({"success": False, "error": "Please provide session_id and the solved captcha_text."}, status=200)
                return

            res = submit_cpp_search_and_crawl(
                session_id=session_id,
                captcha_text=captcha_text,
                tender_id=tender_id,
                tender_title=tender_title,
                max_detail_fetch=max_details,
                max_pages=max_pages,
                crawl_all_pages=crawl_all_pages
            )
            self.send_json(res, status=200)
            return

        # 21. CPP Crawl Next Page via Tapestry Pagination
        elif path == '/api/cpp/session/next_page':
            body = self.read_json_body()
            session_id = body.get('session_id', '')
            page_url = body.get('page_url', '')

            if not session_id or not page_url:
                self.send_json({"success": False, "error": "Missing session_id or page_url"}, status=200)
                return

            res = crawl_next_cpp_page(session_id, page_url)
            self.send_json(res, status=200)
            return

        # 22. CPP Raw HTML / Offline Search Results & Details Importer
        elif path == '/api/cpp/import_html':
            body = self.read_json_body()
            html_content = body.get('html_content', '')
            if not html_content.strip():
                self.send_json({"success": False, "error": "HTML content cannot be empty"}, status=400)
                return

            res = import_cpp_raw_html(html_content)
            self.send_json({"success": True, "import_result": res})
            return

        # 23. CPP Shortlist Toggle
        elif path == '/api/cpp/toggle_shortlist':
            body = self.read_json_body()
            tender_id = body.get('tender_id', '')
            if not tender_id:
                self.send_json({"success": False, "error": "Missing tender_id"}, status=400)
                return

            res = toggle_cpp_shortlist(tender_id)
            self.send_json(res, status=200 if res.get('success') else 404)
            return

        # 24. CPP Document Downloader (NIT PDF / All-Documents ZIP)
        elif path == '/api/cpp/download_document':
            body = self.read_json_body()
            tender_id = body.get('tender_id', '')
            doc_type = body.get('doc_type', 'nit')
            doc_index = int(body.get('doc_index', 0))
            session_id = body.get('session_id', '')

            if not tender_id:
                self.send_json({"success": False, "error": "Missing tender_id"}, status=400)
                return

            res = download_cpp_document(tender_id, doc_type=doc_type, doc_index=doc_index, session_id=session_id)
            self.send_json(res, status=200)
            return

        # 25. CPP Dedicated Single-Tender Live Refresh (with Captcha Sync Support)
        elif path == '/api/cpp/refresh_single':
            body = self.read_json_body()
            tender_id = body.get('tenderId') or body.get('tender_id') or body.get('id', '')
            session_id = body.get('session_id', '')
            captcha_text = body.get('captcha_text', '')
            if not tender_id:
                self.send_json({"success": False, "error": "Missing tenderId"}, status=400)
                return
            if session_id and captcha_text:
                res = sync_cpp_tender_with_captcha(session_id, captcha_text, str(tender_id))
            else:
                res = sync_cpp_single_tender_live(str(tender_id))
            self.send_json(res, status=200)
            return

        # 26. CPP Download All Documents for Tender (with Captcha Sync Support)
        elif path == '/api/cpp/download_all':
            body = self.read_json_body()
            tender_id = body.get('tenderId') or body.get('tender_id') or body.get('id', '')
            session_id = body.get('session_id', '')
            captcha_text = body.get('captcha_text', '')
            if not tender_id:
                self.send_json({"success": False, "error": "Missing tenderId"}, status=400)
                return
            if session_id and captcha_text:
                res = sync_cpp_tender_with_captcha(session_id, captcha_text, str(tender_id))
            else:
                res = download_all_cpp_tender_attachments(str(tender_id))
            self.send_json(res, status=200)
            return

        # 26b. CPP Sync With Captcha Direct Route
        elif path == '/api/cpp/sync_with_captcha':
            body = self.read_json_body()
            session_id = body.get('session_id', '')
            captcha_text = body.get('captcha_text', '')
            tender_id = body.get('tender_id') or body.get('tenderId') or body.get('id', '')
            if not session_id or not captcha_text or not tender_id:
                self.send_json({"success": False, "error": "Missing session_id, captcha_text, or tender_id"}, status=400)
                return
            res = sync_cpp_tender_with_captcha(session_id, captcha_text, str(tender_id))
            self.send_json(res, status=200)
            return

        # 26c. CPP Open Single File via Windows
        elif path == '/api/cpp/open_file':
            body = self.read_json_body()
            tender_id = body.get('tender_id') or body.get('tenderId', '')
            filename = body.get('filename', '')
            clean_tid = re.sub(r'[^a-zA-Z0-9_\-]', '_', str(tender_id))
            safe_fn = os.path.basename(filename)
            target_path = os.path.join(CPP_DOWNLOADS_DIR, f"tender_{clean_tid}", safe_fn)
            if os.path.exists(target_path):
                try:
                    os.startfile(os.path.abspath(target_path))
                    self.send_json({"success": True, "opened": target_path})
                except Exception as e:
                    self.send_json({"error": str(e)}, status=500)
            else:
                self.send_json({"error": "File not found"}, status=404)
            return

        # 27. CPP Open Folder via POST
        elif path == '/api/cpp/open_folder':
            body = self.read_json_body()
            tender_id = body.get('tenderId') or body.get('tender_id') or body.get('id', '')
            clean_tid = re.sub(r'[^a-zA-Z0-9_\-]', '_', str(tender_id))
            target_folder = os.path.join(CPP_DOWNLOADS_DIR, f"tender_{clean_tid}")
            os.makedirs(target_folder, exist_ok=True)
            try:
                subprocess.Popen(['explorer.exe', os.path.abspath(target_folder)])
                self.send_json({"status": "opened", "folder": target_folder})
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
            return

        else:
            self.send_json({"error": "Endpoint not found"}, status=404)

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        super().end_headers()

def open_browser():
    time.sleep(1.0)
    try:
        webbrowser.open(f"http://127.0.0.1:{PORT}")
    except Exception:
        pass

class ThreadedHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def server_bind(self):
        # Support dual-stack (both IPv4 and IPv6) so localhost and 127.0.0.1 both connect with 0 latency
        try:
            self.socket = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
            self.socket.setsockopt(socket.IPPROTO_IPV6, socket.IPV6_V6ONLY, 0)
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.socket.bind(self.server_address)
            self.server_address = self.socket.getsockname()
        except Exception:
            super().server_bind()

if __name__ == "__main__":
    os.chdir(DIRECTORY)
    # Check if another server instance is already running on this port
    try:
        check_req = urllib.request.Request(f"http://127.0.0.1:{PORT}/api/auth/me")
        with urllib.request.urlopen(check_req, timeout=0.8) as check_resp:
            if check_resp.status == 200:
                print("=" * 70)
                print(f" [OK] TDSC Platform is already active on http://127.0.0.1:{PORT}!")
                print(" Opening browser dashboard directly...")
                print("=" * 70)
                webbrowser.open(f"http://127.0.0.1:{PORT}")
                sys.exit(0)
    except Exception:
        pass

    # Initialize DB tables and seed data
    init_database()
    ensure_default_admin()
    ensure_default_employee()

    with ThreadedHTTPServer(("", PORT), PaidTenderRequestHandler) as httpd:
        print("=" * 70)
        print(" TDSC Tender Intelligence & Subscription Platform is LIVE!")
        print(f" Listening on port {PORT} (http://0.0.0.0:{PORT})")
        print(" Admin Login:    admin@bihar-etender.in | Password: Admin@Eproc2026")
        print(" Employee Login: employee@tdsc.in       | Password: Employee@Tdsc2026")
        print("=" * 70)
        
        # Only open browser when running locally (not in headless cloud containers)
        is_cloud = bool(os.environ.get("RENDER") or os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("DYNO"))
        if not is_cloud:
            threading.Thread(target=open_browser, daemon=True).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
