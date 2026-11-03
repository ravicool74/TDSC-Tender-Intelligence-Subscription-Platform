"""
Bihar e-Procurement Subscription Platform - Entitlements & Access Control
Enforces server-side permissions:
- Free vs Paid plan feature gating (Masked PAC & EMD, locked BOQ/NIT downloads, locked Excel exports)
- Server-side entitlement checks on every API endpoint and download request
- Official source URLs and tender metadata (last checked time, stale/extended status)
- Authenticated secure token generation for document downloads
"""

import hmac
import hashlib
import time
import os
from datetime import datetime

SECRET_KEY = os.environ.get("FILE_AUTH_SECRET", "file_sec_bihar_eproc_2026_auth")

OFFICIAL_SOURCE_BASE = "https://eproc2.bihar.gov.in/EPSV2Web/openarea/tenderListingPage.action#latestTenders"

def generate_download_token(tender_id: int, user_id: int, file_type: str = "boq", expiry_seconds: int = 3600) -> str:
    """Generates a tamper-proof HMAC token for authenticated document download."""
    expiry_ts = int(time.time()) + expiry_seconds
    msg = f"{tender_id}:{user_id}:{file_type}:{expiry_ts}".encode('utf-8')
    sig = hmac.new(SECRET_KEY.encode('utf-8'), msg, hashlib.sha256).hexdigest()
    return f"{expiry_ts}_{sig}"

def verify_download_token(tender_id: int, user_id: int, file_type: str, token_str: str) -> bool:
    """Verifies tamper-proof HMAC token."""
    if not token_str or '_' not in token_str:
        return False
    try:
        parts = token_str.split('_', 1)
        expiry_ts = int(parts[0])
        sig = parts[1]
        if time.time() > expiry_ts:
            return False  # Expired
        msg = f"{tender_id}:{user_id}:{file_type}:{expiry_ts}".encode('utf-8')
        expected_sig = hmac.new(SECRET_KEY.encode('utf-8'), msg, hashlib.sha256).hexdigest()
        return hmac.compare_digest(expected_sig, sig)
    except Exception:
        return False

def filter_tenders_for_user(tenders: list, user_dict: dict = None, last_checked: str = None) -> list:
    """
    Applies server-side entitlement masking:
    - Paid subscribers with active plan see 100% full unmasked data.
    - Free plan / Guest users see only 5 full tender previews; subsequent tenders
      have PAC, EMD, and direct download links masked and locked on the backend.
    """
    is_paid_active = False
    allow_boq = False
    user_id = 0
    plan_code = "free"

    if user_dict and user_dict.get("subscription"):
        sub = user_dict["subscription"]
        is_paid_active = bool(sub.get("isPaidActive"))
        plan_code = sub.get("plan", {}).get("code", "free")
        allow_boq = bool(sub.get("entitlements", {}).get("allowBoqDownload"))
        user_id = user_dict.get("id", 0)

    # Admin, Super Admin, and Employee always have full access to all tenders
    if user_dict and user_dict.get("role") in ["admin", "super_admin", "employee"]:
        is_paid_active = True
        allow_boq = True

    now_iso = last_checked or datetime.now().strftime("%d-%m-%Y %I:%M %p")
    output = []

    for idx, t in enumerate(tenders):
        tender_copy = dict(t)
        tid = tender_copy.get("tenderId") or tender_copy.get("systemTenderNo")
        
        # Metadata attachment
        tender_copy["officialSourceUrl"] = OFFICIAL_SOURCE_BASE
        tender_copy["lastCheckedTime"] = now_iso
        
        # Determine status (extended/active/stale)
        due_str = str(tender_copy.get("dueDate", "")).lower()
        if "corrigendum" in due_str or "extended" in due_str:
            tender_copy["portalStatus"] = "extended"
        elif "withdrawn" in due_str or "cancelled" in due_str:
            tender_copy["portalStatus"] = "withdrawn"
        else:
            tender_copy["portalStatus"] = "active"

        # Check access permission:
        # Free users get only the first 5 tenders unmasked
        is_free_quota_exceeded = (not is_paid_active) and (idx >= 5)

        if is_free_quota_exceeded:
            tender_copy["isLocked"] = True
            tender_copy["lockReason"] = "Free Plan Limit: Upgrade to Starter or Professional to view exact PAC, EMD, and download BOQ files."
            tender_copy["estimatedValue"] = None
            tender_copy["estimatedValueFormatted"] = "₹ •••••••• (Locked on Free Plan)"
            tender_copy["isEstimatedValueHidden"] = True
            tender_copy["emdAmount"] = None
            tender_copy["emdAmountFormatted"] = "₹ •••••• (Upgrade Plan to View)"
            tender_copy["termsFiles"] = [{"locked": True, "title": "Terms & Conditions (Upgrade to View)"}]
            tender_copy["officialFiles"] = [{"locked": True, "label": "NIT Notice (Upgrade to View)"}]
            tender_copy["boqFiles"] = [{"locked": True, "label": "BOQ File (Upgrade to View)"}]
            tender_copy["hasDownloaded"] = False
        else:
            tender_copy["isLocked"] = False
            # If user has paid active plan, grant secure download URL
            if is_paid_active and user_id:
                token = generate_download_token(int(tid), user_id, "all")
                tender_copy["authenticatedDownloadUrl"] = f"/api/download_document?id={tid}&token={token}"

        output.append(tender_copy)

    return output

def check_entitlement(user_dict: dict, action: str) -> tuple[bool, str]:
    """
    Validates user entitlement for paid actions:
    - 'boq_download'
    - 'excel_export'
    - 'whatsapp_share'
    - 'saved_search'
    """
    if not user_dict:
        return False, "Authentication required. Please login or register to access this feature."

    # Admin, Super Admin, and Employee have unrestricted entitlement to all actions
    if user_dict.get("role") in ["admin", "super_admin", "employee"]:
        return True, ""

    sub = user_dict.get("subscription", {})
    status = sub.get("status", "none")
    if status == "expired":
        return False, "Your subscription has EXPIRED. Please renew your plan to restore full access."
    elif status == "cancelled":
        return False, "Your subscription was cancelled. Please subscribe to a plan to continue."
    elif status not in ["active"]:
        return False, "A paid subscription is required for this feature. Please view our plans."

    entitlements = sub.get("entitlements", {})
    plan = sub.get("plan", {})

    if action == "boq_download":
        if not entitlements.get("allowBoqDownload"):
            return False, f"BOQ and document downloads are not included in your {plan.get('name', 'current plan')}. Please upgrade to Starter or Professional."
        return True, ""

    elif action == "excel_export":
        if not entitlements.get("allowExcelExport"):
            return False, f"Comprehensive Excel Export (.xlsx) is available on Professional and Business plans. Please upgrade your plan."
        return True, ""

    elif action == "whatsapp_share":
        if not entitlements.get("allowWhatsappAlert"):
            return False, f"Automated WhatsApp alerts are available on Professional and Business plans."
        return True, ""

    elif action == "saved_search":
        if not entitlements.get("allowSavedSearches") and not sub.get("isPaidActive"):
            return False, "Saving search profiles is a premium feature. Please upgrade your plan."
        return True, ""

    return True, ""
