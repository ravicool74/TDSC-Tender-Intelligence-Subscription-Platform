# TDSC Tender Intelligence & Paid Subscription Platform

Complete SaaS platform transforming the **eproc2.bihar.gov.in** tender management tool into a full **commercial paid tender subscription service**.

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/ravicool74/TDSC-Tender-Intelligence-Subscription-Platform)

---

## ☁️ 1-Click Free Cloud Deployment (24/7 Live Online)

Pura platform (Python backend + SQLite Database + Live eProc2 Delta Sync + Contractor CRM + Razorpay) ko 24/7 live internet par chalane ke liye:

### Vikalp A: Render.com (Recommended Free Cloud Hosting)
1. Upar diye gaye **[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/ravicool74/TDSC-Tender-Intelligence-Subscription-Platform)** button par click karein.
2. Render.com par **Sign in with GitHub** karein.
3. Render repository aur `render.yaml` configuration automatically load kar lega.
4. **"Apply"** par click karein.
5. 1 se 2 minute me aapka full-stack platform live ho jayega (e.g. `https://tdsc-tender-platform-xxxx.onrender.com`).

### Vikalp B: Railway.app
1. [Railway.app](https://railway.app) par jayein aur GitHub se login karein.
2. **"New Project"** -> **"Deploy from GitHub repo"** -> select **`ravicool74/TDSC-Tender-Intelligence-Subscription-Platform`**.
3. Railway automatically `Procfile` aur `railway.json` detect karke live deploy kar dega.

---

## 🚀 Quick Start (Kaise Start Karein)

### Method 1: 1-Click Launch (Recommended)
1. Double-click **`start_dashboard.bat`** in `c:\e-tender`.
2. The server initializes the SQLite database, seeds default plans and default admin, and automatically launches your browser at **`http://localhost:8000`**.

### Method 2: Command Line
```powershell
cd c:\e-tender
python server.py
```

### Default Credentials
- **Admin Email:** `admin@bihar-etender.in`
- **Admin Password:** `Admin@Eproc2026`
- **Customer Registration:** Open [http://localhost:8000](http://localhost:8000), click **"🔑 Login / Register"** -> **"Register (नया खाता)"**.

---

## 💎 Product Architecture & Subscription Plans

The platform provides 4 subscription tiers with backend entitlement checks:

| Plan | Price (INR) | Billing Cycle | Tender Browsing | PAC Value | BOQ / NIT Downloads | Excel Export (.xlsx) | WhatsApp / Email Alerts |
|---|---|---|---|---|---|---|---|
| **Free Plan** | ₹ 0 | Monthly | 5 Full Previews | 🔒 Masked after 5th | ❌ Locked | ❌ Locked | ❌ None |
| **Starter Plan** | ₹ 1,199 | Monthly | Unlimited | ✅ Revealed | 5 / day | ❌ Locked | Email Digest |
| **Professional** | ₹ 2,499 | Monthly | Unlimited | ✅ Revealed | Unlimited | ✅ Full (.xlsx) | Instant WhatsApp & Email |
| **Business / Team**| ₹ 4,999 | Monthly | Unlimited | ✅ Revealed | Unlimited | ✅ Full (.xlsx) | Priority WhatsApp Alerts |

> **Note:** Plan prices and features are **admin-editable** in the database (`plans` table) and via the **Admin Portal** UI. Prices are never hardcoded.

---

## 🔒 Security & Backend Access Control

- **Server-Side Entitlement Checks:** Free tier limits and paid entitlements are enforced on the backend (`entitlements.py`). Tender PAC and EMD values beyond the free quota are masked at the server level, not merely hidden in CSS/HTML.
- **Protected File Downloads:** Document downloads (`/api/download_document`) and bulk Excel exports (`/api/export_tenders_xlsx`) verify active plan entitlements and enforce HTTP 403 Forbidden on expired/unauthorized requests.
- **Password Security:** Cryptographic password hashing using `PBKDF2-HMAC-SHA256` with 260,000 iterations and unique 16-byte random salts.
- **Rate Limiting:** Sliding-window rate limiting on login, registration, and sensitive endpoints to protect against brute-force attacks.
- **Audit Logging:** Every administrative override (grant/revoke access) and plan update is permanently recorded in the `audit_events` table with actor, target user, IP address, and timestamp.

---

## 💳 Payment Gateway Integration (Test Mode)

The system integrates an adapter-based payment engine (`payment_provider.py`):

- **Hosted Checkout Simulator:** Customers select UPI, Debit/Credit Card, Netbanking, or Wallet. No sensitive card/UPI data is stored on our server or database.
- **Order Creation on Server:** Orders are generated via `POST /api/payments/create_order` before payment begins.
- **Cryptographic Signature Verification:** `POST /api/payments/verify` validates HMAC-SHA256 signatures (`order_id|payment_id`).
- **Signed Webhook Support:** `POST /api/payments/webhook` verifies `X-Razorpay-Signature`.
- **Webhook Idempotency:** Duplicate webhook events are tracked in `processed_webhooks` and safely skipped without duplicate subscription extensions or invoices.
- **Automated Invoices:** Every successful payment generates an official tax invoice (`INV-YYYYMM-XXXX`) viewable in the customer dashboard.

### Payment Environment Setup (.env)
Copy `.env.example` to `.env` to configure gateway credentials:
```ini
RAZORPAY_KEY_ID=rzp_test_placeholder_key_id
RAZORPAY_KEY_SECRET=placeholder_test_key_secret_do_not_commit
RAZORPAY_WEBHOOK_SECRET=placeholder_webhook_secret_for_signature_verification
```

---

## 👤 Customer Features

1. **Dashboard (`#viewDashboard`):**
   - Active plan card with validity dates and 1-click renewal.
   - Saved Searches with 1-click execution.
   - Email & WhatsApp Alert Preferences with keyword and frequency settings.
   - Invoice and payment receipt history.
   - Closing-Soon urgency alerts (tenders closing in $\le 72$ hours).
2. **Preserved Tender Search (`#viewSearch`):**
   - 13 comprehensive columns (System Tender No., Tender ID, Ref No., Department, Estimated PAC, EMD with payment mode, Fees, 3 Dates, Description, District, Action & Attachments, Share options).
   - 3-letter keyword search across all fields.
   - Exact amount slabs (`₹0 - 25L`, `₹25 - 50L`, `₹50L - 2Cr`, `₹2 - 5Cr`, `> ₹5Cr`, `Hidden PAC`).
   - Quick range pills and 1-click filter reset.
   - 1-Click WhatsApp, Email, and Clipboard Copy sharing.
   - Universal Webhook / URL Data Extractor for external websites.

---

## ⚙️ Admin Portal (`#viewAdmin`)

Accessible only to accounts with `role = 'admin'`:
1. **Plan & Price Editor:** Edit plan names, prices in INR, and toggle feature entitlements live without code modifications.
2. **Customers & Subscriptions:** View all registered contractors, current plan status, and manually grant or revoke subscriptions with custom duration and audit reasons.
3. **💳 Razorpay Merchant Settings:** Configure Razorpay API Key ID, Key Secret, Webhook Secret, Operating Mode (`test` vs `live`), Brand Display Name, and checkout theme color. Includes a **"⚡ Test Razorpay Connection"** button to verify live credentials against the official Razorpay REST API in 1 click.
4. **Audit Trail:** Complete chronological audit log of all administrative actions.
5. **Data Health & Sync:** View last checked timestamp and trigger live background sync from `eproc2.bihar.gov.in`.

---

## 💳 Razorpay Merchant Integration (Inbuilt)

The platform includes **native, direct Razorpay Merchant integration**:
- **Standard Checkout Modal SDK:** Uses official Razorpay checkout (`https://checkout.razorpay.com/v1/checkout.js`).
- **REST Orders API:** Native HTTP Basic Auth integration calling `https://api.razorpay.com/v1/orders` without heavy external dependencies.
- **Payment Verification:** Cryptographic HMAC-SHA256 signature calculation (`hmac.new(secret, order_id + "|" + payment_id, sha256)`).
- **Cryptographic Webhooks:** Signed webhook receiver (`X-Razorpay-Signature`) with idempotency table (`processed_webhooks`) preventing duplicate subscriptions or invoice generation.
- **Dual Mode Operation:** Instant switching between `test` mode (sandbox) and `live` mode via Admin Portal or `.env`.

### Merchant Setup Steps:
1. Log in to the Admin Portal using `admin@bihar-etender.in` / `Admin@Eproc2026`.
2. Go to **Admin Portal &rarr; 💳 Razorpay Merchant Settings**.
3. Enter your Razorpay **Key ID** (`rzp_test_...` or `rzp_live_...`) and **Key Secret**.
4. Click **"⚡ Test Razorpay Connection"** to verify that credentials authenticate with Razorpay servers.
5. Click **"💾 Save Merchant Credentials"**.
6. In Razorpay Dashboard &rarr; Settings &rarr; Webhooks, add your Webhook URL (`http://<your-domain>/api/payments/webhook`) and subscribe to `payment.captured`, `payment.failed`, and `refund.processed`.

---

## 💾 Database Backups & Restore

The platform stores all state in SQLite (`etender_subscription.db`) with Write-Ahead Logging (`WAL` mode) enabled for high concurrency.

### How to Backup Database:
```powershell
# Online safe backup without stopping the server:
python -c "import sqlite3; con = sqlite3.connect('etender_subscription.db'); bck = sqlite3.connect('backup_etender_subscription.db'); con.backup(bck); bck.close(); con.close(); print('Backup complete!')"
```

### How to Restore Database:
1. Stop the server (`Ctrl+C` or kill process).
2. Replace `etender_subscription.db` with your backup file:
   ```powershell
   Copy-Item backup_etender_subscription.db -Destination etender_subscription.db -Force
   ```
3. Restart `python server.py`.

---

## 💼 TDSC Employee Portal & Contractor Database CRM

Accessible to authorized staff (`role = 'employee'`) and Administrators:
1. **Employee Login:**
   - **Default Employee:** `employee@tdsc.in` | `Employee@Tdsc2026` (`Rahul Kumar - Tender Executive`).
   - Unrestricted staff access: all 1,454+ tenders, exact PAC values, EMD, and BOQ downloads are unlocked without requiring a paid subscription.
   - Admin can add any number of employee accounts from **Admin Portal &rarr; Staff & Employees (कर्मचारी)**.
2. **Contractor CRM Database:**
   - Pre-seeded with 10+ verified Bihar contractors across primary infrastructure sectors.
   - **🔄 Sync TDSC Office System:** 1-click synchronization to fetch and update contractors directly from the TDSC Office Management System.
   - **➕ Add / Edit Contractor:** Manage company name, contact person, mobile (WhatsApp enabled), email, district, work category, class (Class 1, 2, 3), and capacity slab (`₹0-25L`, `₹25L-50L`, `₹50L-2Cr`, `₹2Cr-5Cr`, `> ₹5Cr`).
3. **Category-wise Contractor Shortlisting:**
   - 7 primary sector filter chips:
     - 🛣️ Roads & Highways (सड़क एवं राजमार्ग)
     - 🏢 Building Construction (भवन निर्माण)
     - 🌉 Bridges & Culverts (पुल एवं पुलिया)
     - 💧 Water Supply & PHE (जल आपूर्ति एवं लोक स्वास्थ्य)
     - ⚡ Electrical Works (विद्युत एवं सौर ऊर्जा)
     - 🌾 Irrigation & Drainage (सिंचाई एवं जल निकासी)
     - 🏗️ General Civil (सामान्य निर्माण कार्य)
4. **Smart Bidirectional Tender Matcher & WhatsApp Dispatch:**
   - **Contractor to Tenders:** Select any contractor &rarr; System scans all 1,454+ live Bihar eProc tenders against their category, district, and budget capacity with instant affinity match scores.
   - **Tender to Contractors:** Each tender in the main table has a **`🎯 Suggest to Contractor`** button for staff. Clicking it automatically shortlists matching contractors and opens a pre-formatted WhatsApp dispatch modal.
   - **Dispatch History Log:** Full chronological record of all tender suggestions dispatched via WhatsApp or Email.

---

## ⚡ High-Performance Live eProc2 Check & Smart Delta Sync Engine

1. **🔄 Tender-by-Tender Live Check (`Live Check eProc2`):**
   - Each tender row features a dedicated **`🔄 Live Check eProc2`** button in the search table and in the Action modal.
   - Clicking it queries `eproc2.bihar.gov.in` for **only that specific tender ID** (`previewTenderByTenderId`).
   - Instantly updates latest **Due Date extension**, **PAC cost**, **EMD**, and **Corrigendum / BOQ attachments** in 1 second.
   - Updates only that single table row via DOM manipulation with green highlight flash — **zero page refresh and 0% CPU spike on the server**.
2. **Smart Delta Sync (System Tender No Deduplication):**
   - Bulk refresh compares incoming tenders against existing `systemTenderNo` and `tenderId`.
   - Existing tenders are checked for date/cost updates and skipped if identical, preserving all local records.
   - Brand-new tenders are prepended to the top of the master list.
   - Heavy openpyxl Excel rebuild on every sync is eliminated; `.xlsx` files are generated on-demand when downloading.
3. **In-Memory RAM Caching:**
   - Master tender dataset is served from server RAM (`MASTER_TENDERS_VIEW_CACHE`) in 0–2 ms with zero disk read overhead.

---

## 🧪 Automated Test Verification

All flows can be verified end-to-end at any time:
```powershell
python test_single_tender_live_check_and_delta_sync.py
python test_employee_contractor_crm.py
python test_plan_hide_unhide_switch.py
python test_complete_platform_flows.py
```
Verification covers:
- Single tender live check from eProc2 and granular DOM updating
- Smart Delta Deduplication by System Tender No
- In-memory RAM cache speed ($\le 5$ ms response)
- Employee login & staff entitlements
- Contractor CRM directory & TDSC Office System sync
- Bidirectional tender matcher & WhatsApp dispatching
- Public Plans API, dynamic pricing & plan-level Hide/Unhide switches
- Free tier locking & server-side entitlement checks
- Razorpay Merchant checkout simulation & webhook verification
- Expired subscription access revocation (HTTP 403) & audit trail logging
