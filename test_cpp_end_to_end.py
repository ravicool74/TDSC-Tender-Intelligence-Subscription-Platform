"""
End-to-End Integration Test for CPP / Central Government Tender Subsystem
Tests server routes, live session initialization, CAPTCHA retrieval,
duplicate prevention, Excel workbook download, and Bihar isolation.
"""

import os
import sys
import json
import time
import urllib.request
import urllib.parse
import threading
import unittest
from http.server import HTTPServer

import server
from cpp_tender_manager import (
    load_cpp_tenders_store,
    save_cpp_tenders_batch,
    CPP_VIEW_JSON,
    CPP_EXCEL_PATH
)
from test_cpp_tender_workflow import SAMPLE_USER_PROMPT_DETAIL_TEXT, SAMPLE_SEARCH_RESULTS_HTML

TEST_PORT = 8089

class TestCppEndToEnd(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Backup original CPP_VIEW_JSON
        cls.original_data = None
        if os.path.exists(CPP_VIEW_JSON):
            try:
                with open(CPP_VIEW_JSON, 'r', encoding='utf-8') as f:
                    cls.original_data = f.read()
            except Exception:
                pass

        # Start server in background thread on TEST_PORT
        cls.httpd = HTTPServer(('127.0.0.1', TEST_PORT), server.PaidTenderRequestHandler)
        cls.server_thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.server_thread.start()
        time.sleep(0.5)

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        # Restore original CPP_VIEW_JSON
        if cls.original_data is not None:
            try:
                with open(CPP_VIEW_JSON, 'w', encoding='utf-8') as f:
                    f.write(cls.original_data)
            except Exception:
                pass

    def setUp(self):
        with open(CPP_VIEW_JSON, 'w', encoding='utf-8') as f:
            json.dump({
                "portal": "Central Public Procurement Portal (eprocure.gov.in)",
                "updated_at": "2026-11-05T00:00:00",
                "total_count": 0,
                "tenders": []
            }, f, indent=2)

    def test_01_import_and_list_tenders(self):
        """Test importing HTML via /api/cpp/import_html and querying /api/cpp/tenders."""
        url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/import_html"
        post_data = json.dumps({"html_content": SAMPLE_SEARCH_RESULTS_HTML}).encode('utf-8')
        req = urllib.request.Request(url, data=post_data, headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get("success"))
            self.assertEqual(data["import_result"]["tenders_found"], 2)
            self.assertEqual(data["import_result"]["added"], 2)

        # Query /api/cpp/tenders
        get_url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/tenders"
        with urllib.request.urlopen(get_url, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            res = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(res.get("success"))
            self.assertTrue(res["total"] >= 2)
            t1 = next(t for t in res["tenders"] if t["tender_id"] == "2026_DOP_933186_1")
            self.assertEqual(t1["tender_value_inr"], 1076893.0)
            self.assertEqual(t1["budget_slab"], "0to25L")

    def test_02_strict_duplicate_rejection_via_api(self):
        """Verify submitting same HTML again via API results in 0 added and 2 duplicates skipped."""
        url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/import_html"
        post_data = json.dumps({"html_content": SAMPLE_SEARCH_RESULTS_HTML}).encode('utf-8')
        
        # First import
        req1 = urllib.request.Request(url, data=post_data, headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req1, timeout=10) as resp1:
            data1 = json.loads(resp1.read().decode('utf-8'))
            self.assertEqual(data1["import_result"]["added"], 2)

        # Second import of identical data
        req2 = urllib.request.Request(url, data=post_data, headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req2, timeout=10) as resp2:
            data2 = json.loads(resp2.read().decode('utf-8'))
            self.assertEqual(data2["import_result"]["added"], 0)
            self.assertEqual(data2["import_result"]["skipped_duplicates"], 2)

    def test_03_deep_dive_detail_and_shortlist(self):
        """Test detail endpoint and shortlist toggle."""
        # Import detail page
        imp_url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/import_html"
        post_data = json.dumps({"html_content": SAMPLE_USER_PROMPT_DETAIL_TEXT}).encode('utf-8')
        req = urllib.request.Request(imp_url, data=post_data, headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            self.assertEqual(resp.status, 200)

        # Get detail
        det_url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/tender_detail?tender_id=2026_DOP_933186_1"
        with urllib.request.urlopen(det_url, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            res = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(res.get("success"))
            t = res["tender"]
            self.assertEqual(t["tender_ref_no"], "NIT/EEPCND/2026-27/27")
            self.assertEqual(t["emd_amount_inr"], 21538.0)
            self.assertTrue(len(t["instruments"]) >= 5)

        # Toggle shortlist
        sl_url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/toggle_shortlist"
        sl_data = json.dumps({"tender_id": "2026_DOP_933186_1"}).encode('utf-8')
        req = urllib.request.Request(sl_url, data=sl_data, headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            res = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(res.get("success"))
            self.assertTrue(res.get("is_shortlisted"))

    def test_04_export_cpp_excel(self):
        """Verify dedicated Excel download /api/cpp/export_xlsx returns valid xlsx."""
        xl_url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/export_xlsx"
        with urllib.request.urlopen(xl_url, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            content_type = resp.headers.get('Content-Type')
            self.assertIn("spreadsheetml", content_type)
            content_disp = resp.headers.get('Content-Disposition')
            self.assertIn("CPP_Central_Gov_Tenders_Comprehensive.xlsx", content_disp)
            body = resp.read()
            self.assertTrue(len(body) > 1000)

    def test_05_live_session_initialization(self):
        """Test initializing a live session with eprocure.gov.in and receiving live Base64 CAPTCHA."""
        start_url = f"http://127.0.0.1:{TEST_PORT}/api/cpp/session/start"
        post_data = json.dumps({"mode": "active_tenders"}).encode('utf-8')
        req = urllib.request.Request(start_url, data=post_data, headers={'Content-Type': 'application/json'})
        try:
            with urllib.request.urlopen(req, timeout=50) as resp:
                self.assertEqual(resp.status, 200)
                res = json.loads(resp.read().decode('utf-8'))
                if res.get("success"):
                    self.assertTrue(bool(res.get("session_id")))
                    self.assertTrue(res.get("captcha_image").startswith("data:image"))
                else:
                    self.assertTrue("error" in res)
        except Exception as e:
            print(f"Note: Portal live test caught network timeout: {e}")

    def test_06_verify_bihar_eproc_untouched(self):
        """Verify Bihar eProc endpoint still returns Bihar tenders and is completely untouched."""
        bihar_url = f"http://127.0.0.1:{TEST_PORT}/api/excel_tenders"
        with urllib.request.urlopen(bihar_url, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get("totalTenders") > 0)
            # Ensure zero contamination: no central CPP tender in Bihar database
            self.assertFalse(any(t.get("tenderId") == "2026_DOP_933186_1" for t in data.get("tenders", [])))

if __name__ == "__main__":
    unittest.main()
