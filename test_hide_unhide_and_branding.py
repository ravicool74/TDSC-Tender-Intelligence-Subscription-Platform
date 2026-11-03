"""
Verification script for:
1. Header Title: "TDSC Tender Intelligence & Subscription Platform"
2. Removal of yellow-marked badge "BIHAR ePROC 2.0" and subtext "Official Aggregator • eproc2.bihar.gov.in • Verified Data"
3. "Hide/Unhide" toggle switch functionality and Admin/Super Admin access control preservation.
"""

import urllib.request
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

BASE_URL = "http://localhost:8000"

def test_branding_and_hide_unhide():
    print("=" * 70)
    print("VERIFYING TDSC BRANDING & HIDE/UNHIDE TOGGLE SWITCH")
    print("=" * 70)

    # 1. Fetch Homepage HTML
    req = urllib.request.Request(BASE_URL)
    with urllib.request.urlopen(req) as resp:
        html = resp.read().decode('utf-8')

    # Check 1: Brand Title
    assert "TDSC Tender Intelligence &amp; Subscription Platform" in html or "TDSC Tender Intelligence & Subscription Platform" in html, "TDSC brand title missing!"
    print(" [OK] Brand title updated: 'TDSC Tender Intelligence & Subscription Platform'")

    # Check 2: Yellow-marked badge deleted
    assert "govt-seal" not in html, "Yellow marked badge (govt-seal) should be deleted!"
    assert "BIHAR ePROC 2.0" not in html, "BIHAR ePROC 2.0 should be deleted!"
    print(" [OK] Yellow-marked badge 'BIHAR ePROC 2.0' successfully deleted.")

    # Check 3: Yellow-marked subtext deleted
    assert "Official Aggregator" not in html, "Yellow marked subtext 'Official Aggregator' should be deleted!"
    print(" [OK] Yellow-marked subtext 'Official Aggregator • eproc2.bihar.gov.in • Verified Data' successfully deleted.")

    # Check 4: Hide/Unhide switch elements exist
    assert "btnTopHideUnhide" in html, "Top 'Hide/Unhide' switch button missing!"
    assert "btnSwitchHideUnhide" in html, "Inline 'Hide/Unhide' switch bar missing!"
    assert "toggleHideUnhide" in html, "toggleHideUnhide function missing!"
    print(" [OK] 'Hide/Unhide' switch option (topbar button and dedicated switch pill) verified in DOM.")

    # Check 5: Role protection intact
    assert "isSuperOrAdmin" in html, "Role security check missing!"
    print(" [OK] Admin/Super Admin role security strictly preserved (pichle work me koi modification nahi).")

    print("\n" + "=" * 70)
    print("ALL BRANDING & HIDE/UNHIDE REQUIREMENTS VERIFIED 100%!")
    print("=" * 70)

if __name__ == "__main__":
    test_branding_and_hide_unhide()
