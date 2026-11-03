"""
Bihar e-Procurement Subscription Platform - Database Module
SQLite database with migrations, schema initialization, and transactional helpers.
Zero external dependencies - standard library sqlite3.
"""

import sqlite3
import os
import json
import logging
from datetime import datetime

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(DIRECTORY, "etender_subscription.db")

logger = logging.getLogger("etender_db")

def get_connection():
    """Returns a SQLite connection with row_factory set to sqlite3.Row."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")  # High concurrency write-ahead logging
    return conn

def init_database():
    """Initializes tables, indexes, and default admin/plans if they do not exist."""
    conn = get_connection()
    with conn:
        cursor = conn.cursor()

        # 1. Users Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL COLLATE NOCASE,
            mobile TEXT,
            name TEXT NOT NULL,
            company TEXT,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'customer', -- 'customer' or 'admin'
            email_verified INTEGER NOT NULL DEFAULT 0,
            mobile_verified INTEGER NOT NULL DEFAULT 0,
            verification_token TEXT,
            reset_token TEXT,
            reset_token_expiry TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_email ON users(email);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_role ON users(role);")

        # 2. Plans Table (Admin-editable, prices not hardcoded)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL, -- 'free', 'starter', 'professional', 'business'
            name TEXT NOT NULL,
            badge_text TEXT,
            price_inr REAL NOT NULL DEFAULT 0.0,
            billing_period TEXT NOT NULL DEFAULT 'monthly', -- 'monthly', 'quarterly', 'annual'
            description TEXT,
            features_json TEXT NOT NULL DEFAULT '[]',
            daily_view_limit INTEGER NOT NULL DEFAULT -1, -- -1 = unlimited
            allow_boq_download INTEGER NOT NULL DEFAULT 0,
            allow_excel_export INTEGER NOT NULL DEFAULT 0,
            allow_whatsapp_alert INTEGER NOT NULL DEFAULT 0,
            allow_email_alert INTEGER NOT NULL DEFAULT 0,
            allow_saved_searches INTEGER NOT NULL DEFAULT 0,
            max_saved_searches INTEGER NOT NULL DEFAULT 0,
            is_active INTEGER NOT NULL DEFAULT 1,
            sort_order INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_plans_code ON plans(code);")

        # 3. Subscriptions Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS subscriptions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            plan_id INTEGER NOT NULL REFERENCES plans(id),
            status TEXT NOT NULL DEFAULT 'active', -- 'active', 'expired', 'cancelled', 'pending_payment', 'renewal_due'
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            cancelled_at TEXT,
            auto_renew INTEGER NOT NULL DEFAULT 1,
            notes TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_subs_user_id ON subscriptions(user_id);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_subs_status ON subscriptions(status);")

        # 4. Payments Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            subscription_id INTEGER REFERENCES subscriptions(id),
            order_id TEXT UNIQUE NOT NULL,
            transaction_id TEXT UNIQUE,
            provider TEXT NOT NULL DEFAULT 'razorpay',
            amount_inr REAL NOT NULL,
            currency TEXT NOT NULL DEFAULT 'INR',
            status TEXT NOT NULL DEFAULT 'pending', -- 'pending', 'successful', 'failed', 'refunded'
            payment_method TEXT,
            signature TEXT,
            error_code TEXT,
            error_description TEXT,
            raw_payload TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_payments_order ON payments(order_id);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_payments_user ON payments(user_id);")

        # 5. Invoices Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_number TEXT UNIQUE NOT NULL,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            payment_id INTEGER REFERENCES payments(id),
            subscription_id INTEGER REFERENCES subscriptions(id),
            amount_inr REAL NOT NULL,
            tax_inr REAL NOT NULL DEFAULT 0.0,
            total_inr REAL NOT NULL,
            status TEXT NOT NULL DEFAULT 'paid', -- 'paid', 'unpaid', 'void'
            invoice_date TEXT NOT NULL,
            download_url TEXT,
            created_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_invoices_user ON invoices(user_id);")

        # 6. Saved Searches Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS saved_searches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            query_keyword TEXT,
            amount_slab TEXT,
            district TEXT,
            department TEXT,
            created_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_saved_searches_user ON saved_searches(user_id);")

        # 7. Alert Preferences Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS alert_preferences (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER UNIQUE NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            email_enabled INTEGER NOT NULL DEFAULT 1,
            whatsapp_enabled INTEGER NOT NULL DEFAULT 0,
            frequency TEXT NOT NULL DEFAULT 'daily', -- 'instant', 'daily', 'weekly'
            target_districts TEXT NOT NULL DEFAULT '[]',
            target_departments TEXT NOT NULL DEFAULT '[]',
            target_keywords TEXT NOT NULL DEFAULT '[]',
            min_value REAL DEFAULT 0,
            updated_at TEXT NOT NULL
        );
        """)

        # 8. Audit Events Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actor_user_id INTEGER REFERENCES users(id),
            target_user_id INTEGER REFERENCES users(id),
            action TEXT NOT NULL,
            details_json TEXT,
            ip_address TEXT,
            created_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_events(created_at);")

        # 9. Webhook Idempotency Log
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS processed_webhooks (
            event_id TEXT PRIMARY KEY,
            provider TEXT NOT NULL,
            event_type TEXT NOT NULL,
            processed_at TEXT NOT NULL
        );
        """)

        # 10. User Sessions Table (Stateful secure bearer tokens)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS user_sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            expires_at TEXT NOT NULL,
            ip_address TEXT,
            user_agent TEXT,
            created_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sessions_user ON user_sessions(user_id);")

        # 11. Rate Limiting Table (Sliding window counts)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS rate_limits (
            key TEXT PRIMARY KEY,
            count INTEGER NOT NULL DEFAULT 0,
            reset_at INTEGER NOT NULL
        );
        """)

        # 12. System Settings Table (Razorpay Merchant Credentials & Configurations)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS system_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        """)

        # 13. Contractors Table (TDSC Contractor Database & CRM)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS contractors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            contractor_id TEXT UNIQUE NOT NULL,
            company_name TEXT NOT NULL,
            contact_person TEXT NOT NULL,
            mobile TEXT NOT NULL,
            email TEXT,
            district TEXT NOT NULL,
            category TEXT NOT NULL,
            work_class TEXT DEFAULT 'Class 1',
            budget_slab TEXT DEFAULT '50Lto2Cr',
            status TEXT DEFAULT 'active',
            source TEXT DEFAULT 'TDSC Office System',
            notes TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_contractors_category ON contractors(category);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_contractors_district ON contractors(district);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_contractors_budget ON contractors(budget_slab);")

        # 14. Tender Suggestions Log Table (Employee Suggestion & Dispatch Records)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS tender_suggestions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER REFERENCES users(id),
            contractor_id INTEGER REFERENCES contractors(id),
            tender_id INTEGER NOT NULL,
            tender_ref_no TEXT,
            tender_pac REAL DEFAULT 0,
            channel TEXT DEFAULT 'whatsapp',
            status TEXT DEFAULT 'suggested',
            notes TEXT,
            created_at TEXT NOT NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_suggestions_contractor ON tender_suggestions(contractor_id);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_suggestions_employee ON tender_suggestions(employee_id);")

        # Seed Default Plans if not present
        cursor.execute("SELECT COUNT(*) FROM plans;")
        if cursor.fetchone()[0] == 0:
            now = datetime.now().isoformat()
            default_plans = [
                (
                    "free",
                    "Free Plan (मुफ़्त)",
                    "Basic Access",
                    0.0,
                    "monthly",
                    "Explore public Bihar tenders with limited information preview.",
                    json.dumps([
                        "Tender listing view (limited 5 full previews)",
                        "Basic search & district filter",
                        "Summary stats view",
                        "❌ BOQ Excel/PDF downloads locked",
                        "❌ Bulk Excel export locked",
                        "❌ WhatsApp alerts locked",
                        "❌ Saved searches locked"
                    ], ensure_ascii=False),
                    5, 0, 0, 0, 0, 0, 0, 1, 1, now, now
                ),
                (
                    "starter",
                    "Starter Plan (शुरुआती)",
                    "Best for Small Contractors",
                    999.0,
                    "monthly",
                    "For independent contractors bidding up to 25 Lakhs.",
                    json.dumps([
                        "Unlimited tender search & view",
                        "Exact Estimated Value (PAC) revealed",
                        "EMD & Tender Fee payment modes",
                        "5 BOQ & NIT downloads per day",
                        "3 Saved search profiles",
                        "Email notifications for new tenders",
                        "❌ Bulk Excel export locked",
                        "❌ WhatsApp alerts locked"
                    ], ensure_ascii=False),
                    -1, 1, 0, 0, 1, 1, 3, 1, 2, now, now
                ),
                (
                    "professional",
                    "Professional Plan (पेशेवर)",
                    "Most Popular / सर्वाधिक लोकप्रिय",
                    2499.0,
                    "monthly",
                    "For established contractors & suppliers bidding across Bihar.",
                    json.dumps([
                        "Everything in Starter Plan",
                        "Unlimited BOQ, NIT & Terms downloads",
                        "Full Excel Workbook (.xlsx) export",
                        "15 Saved searches with instant filters",
                        "Instant 1-Click WhatsApp & Email sharing",
                        "Daily email tender alert digest",
                        "Closing-soon urgency warnings"
                    ], ensure_ascii=False),
                    -1, 1, 1, 1, 1, 1, 15, 1, 3, now, now
                ),
                (
                    "business",
                    "Business / Team Plan (बिज़नेस टीम)",
                    "Enterprise & Multi-User",
                    4999.0,
                    "monthly",
                    "For large construction firms, consortiums, and bidding agencies.",
                    json.dumps([
                        "Everything in Professional Plan",
                        "Multi-user team access (up to 5 team members)",
                        "Unlimited Saved Searches & instant alerts",
                        "Priority WhatsApp tender alerts directly to phone",
                        "Automated Central CPPP & Bihar cross-portal alerts",
                        "Dedicated phone & WhatsApp bidding support",
                        "Custom billing & GST invoices"
                    ], ensure_ascii=False),
                    -1, 1, 1, 1, 1, 1, 999, 1, 4, now, now
                )
            ]
            cursor.executemany("""
            INSERT INTO plans (
                code, name, badge_text, price_inr, billing_period, description,
                features_json, daily_view_limit, allow_boq_download, allow_excel_export,
                allow_whatsapp_alert, allow_email_alert, allow_saved_searches,
                max_saved_searches, is_active, sort_order, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, default_plans)
            print("[Database] Seeded 4 default subscription plans (Free, Starter, Professional, Business).")

        # Seed initial sample contractors from TDSC Office Management System
        cursor.execute("SELECT COUNT(*) FROM contractors;")
        if cursor.fetchone()[0] == 0:
            now = datetime.now().isoformat()
            sample_contractors = [
                ("TDSC-CON-101", "Patna Roadlines & Infra Projects", "Ajay Kumar Sharma", "9835012345", "ajay.patna@infra.in", "Patna", "Roads & Highways", "Class 1", "2Crto5Cr", "prime", "TDSC Office System", "Specialist in RWD PMGSY road construction and bitumen carpeting.", now, now),
                ("TDSC-CON-102", "Maa Sharda Civil Construction Ltd", "Rajesh Kumar Singh", "9431023456", "sharda.build@bihar.in", "Muzaffarpur", "Building Construction", "Class 2", "50Lto2Cr", "active", "TDSC Office System", "School, hospital and administrative building repairs.", now, now),
                ("TDSC-CON-103", "Magadh Bridge & Structural Works", "Sanjay Gupta", "9835234567", "magadh.bridge@gmail.com", "Gaya", "Bridges & Culverts", "Class 1", "above5Cr", "prime", "TDSC Office System", "Bridge and culvert RCC works under Pul Nirman Nigam.", now, now),
                ("TDSC-CON-104", "Anga Jal Shakti & PHE Engineers", "Vinod Prasad Sah", "9431456789", "anga.phed@yahoo.co.in", "Bhagalpur", "Water Supply & PHE", "Class 2", "25Lto50L", "active", "TDSC Office System", "PHE tube wells, piped drinking water and OHT projects.", now, now),
                ("TDSC-CON-105", "Kosi Vidyut & Power Services", "Manoj Kumar Jha", "9835678901", "kosi.electric@bihar.com", "Purnia", "Electrical Works", "Class 3", "0to25L", "active", "TDSC Office System", "Solar installations, sub-station maintenance and building wiring.", now, now),
                ("TDSC-CON-106", "Mithila Sinchai & Drainage Corp", "Ramesh Choudhary", "9431890123", "mithila.irrigation@gmail.com", "Darbhanga", "Irrigation & Drainage", "Class 1", "50Lto2Cr", "active", "TDSC Office System", "Canal lining, flood protection bunds and sluice gates.", now, now),
                ("TDSC-CON-107", "Vaishali Civil Buildcon Group", "Vikash Kumar Yadav", "9835901234", "vaishali.build@gmail.com", "Vaishali", "General Civil", "Class 2", "50Lto2Cr", "active", "TDSC Office System", "PCC roads, boundary walls, community halls.", now, now),
                ("TDSC-CON-108", "Champaran High-Tech Infra", "Sunil Kumar Tiwary", "9431123456", "champaran.infra@outlook.com", "Motihari", "Roads & Highways", "Class 1", "2Crto5Cr", "prime", "TDSC Office System", "State Highway, SH and MDR road widenings.", now, now),
                ("TDSC-CON-109", "Shahabad Building Developers", "Anand Prakash Rai", "9835345678", "shahabad.dev@gmail.com", "Bhojpur (Ara)", "Building Construction", "Class 2", "25Lto50L", "active", "TDSC Office System", "Panchayat sarkar bhawan and health center construction.", now, now),
                ("TDSC-CON-110", "Nalanda Pioneer Projects", "Santosh Kumar Verma", "9431567890", "nalanda.pipe@rediffmail.com", "Nalanda", "Water Supply & PHE", "Class 1", "50Lto2Cr", "active", "TDSC Office System", "Har Ghar Nal Ka Jal pipeline distribution networks.", now, now)
            ]
            cursor.executemany("""
            INSERT INTO contractors (
                contractor_id, company_name, contact_person, mobile, email,
                district, category, work_class, budget_slab, status, source, notes, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, sample_contractors)
            print("[Database] Seeded 10 verified contractors from TDSC Office Management System.")

        conn.commit()

    print("[Database] SQLite database initialized successfully at:", DB_PATH)

def get_setting(key: str, default: str = None) -> str:
    """Retrieve a setting by key from system_settings, fallback to default or os.environ."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM system_settings WHERE key = ?;", (key,))
        row = cursor.fetchone()
        if row and row['value'] is not None:
            return row['value']
    except Exception:
        pass
    env_key = key.upper()
    return os.environ.get(env_key, default)

def set_setting(key: str, value: str):
    """Store or update a setting in system_settings."""
    conn = get_connection()
    now = datetime.now().isoformat()
    with conn:
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO system_settings (key, value, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at;
        """, (key, value, now))

def get_all_settings() -> dict:
    """Retrieve all settings as a dictionary."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT key, value FROM system_settings;")
        return {row['key']: row['value'] for row in cursor.fetchall()}
    except Exception:
        return {}

if __name__ == "__main__":
    init_database()
