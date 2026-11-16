"""
Automated Test Suite for Past & Expired Tenders Subsystem
Verifies:
1. Multi-format date parsing
2. Accurate expiration calculation
3. Central CPPP status filtering (Active vs Past)
4. Bihar eProc status filtering (Active vs Past)
5. /api/past_tenders unified API
6. /api/past_tenders/export_xlsx workbook export
7. Zero regression on master data integrity
"""

import unittest
import json
import os
import time
import threading
import urllib.request
import urllib.parse
from datetime import datetime
from http.server import HTTPServer

import server
from cpp_tender_manager import (
    parse_tender_due_date,
    is_tender_expired,
    filter_tenders_by_status,
    generate_past_tenders_excel_workbook,
    load_cpp_tenders_store,
    CPP_VIEW_JSON,
    PAST_EXCEL_PATH
)

TEST_PORT = 8095
BASE_URL = f"http://127.0.0.1:{TEST_PORT}"

class TestPastTendersSubsystem(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd = HTTPServer(('127.0.0.1', TEST_PORT), server.PaidTenderRequestHandler)
        cls.server_thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.server_thread.start()
        time.sleep(0.5)

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def test_01_date_parsing_across_formats(self):
        d1 = parse_tender_due_date("04-Nov-2026 06:30 PM")
        self.assertIsNotNone(d1)
        self.assertEqual((d1.year, d1.month, d1.day, d1.hour, d1.minute), (2026, 11, 4, 18, 30))

        d2 = parse_tender_due_date("01-11-2026 03:00 PM")
        self.assertIsNotNone(d2)
        self.assertEqual((d2.year, d2.month, d2.day, d2.hour, d2.minute), (2026, 11, 1, 15, 0))

        d3 = parse_tender_due_date("2026-10-26 15:00")
        self.assertIsNotNone(d3)
        self.assertEqual((d3.year, d3.month, d3.day, d3.hour, d3.minute), (2026, 10, 26, 15, 0))

        d4 = parse_tender_due_date("31-Dec-2026 02:00 PM")
        self.assertIsNotNone(d4)
        self.assertEqual((d4.year, d4.month, d4.day, d4.hour, d4.minute), (2026, 12, 31, 14, 0))

        self.assertIsNone(parse_tender_due_date("N/A"))
        self.assertIsNone(parse_tender_due_date("None"))
        self.assertIsNone(parse_tender_due_date("-"))
        self.assertIsNone(parse_tender_due_date(""))

    def test_02_is_tender_expired_logic(self):
        ref_now = datetime(2026, 11, 15, 12, 0, 0)
        past_tender = {"bid_closing_date": "04-Nov-2026 06:30 PM"}
        future_tender = {"bid_closing_date": "26-Nov-2026 06:00 PM"}
        bihar_past = {"dueDate": "2026-10-26 15:00"}
        bihar_future = {"dueDate": "31-Dec-2026 03:00 PM"}

        self.assertTrue(is_tender_expired(past_tender, ref_now))
        self.assertFalse(is_tender_expired(future_tender, ref_now))
        self.assertTrue(is_tender_expired(bihar_past, ref_now))
        self.assertFalse(is_tender_expired(bihar_future, ref_now))

    def test_03_filter_tenders_by_status(self):
        ref_now = datetime(2026, 11, 15, 12, 0, 0)
        items = [
            {"tender_id": "T1", "bid_closing_date": "04-Nov-2026 06:30 PM"},
            {"tender_id": "T2", "bid_closing_date": "26-Nov-2026 06:00 PM"},
            {"tender_id": "T3", "dueDate": "2026-10-26 15:00"},
            {"tender_id": "T4", "dueDate": "31-Dec-2026 03:00 PM"}
        ]
        active = filter_tenders_by_status(items, "active", ref_now)
        past = filter_tenders_by_status(items, "past", ref_now)
        all_items = filter_tenders_by_status(items, "all", ref_now)

        self.assertEqual(len(active), 2)
        self.assertEqual([x["tender_id"] for x in active], ["T2", "T4"])
        self.assertEqual(len(past), 2)
        self.assertEqual([x["tender_id"] for x in past], ["T1", "T3"])
        self.assertEqual(len(all_items), 4)

    def test_04_api_past_tenders_endpoint(self):
        url = f"{BASE_URL}/api/past_tenders?limit=all"
        req = urllib.request.urlopen(url, timeout=10)
        self.assertEqual(req.status, 200)
        data = json.loads(req.read().decode('utf-8'))

        self.assertTrue(data.get("success"))
        tenders = data.get("tenders", [])
        self.assertGreater(len(tenders), 2000, "Should have over 2,000 past tenders combined")

        stats = data.get("stats", {})
        self.assertGreater(stats.get("total_past", 0), 2000)
        self.assertGreater(stats.get("cpp_past", 0), 1000)
        self.assertGreater(stats.get("bihar_past", 0), 1000)
        self.assertGreater(stats.get("total_value_cr", 0), 1000)

        # Test portal filtering
        req_cpp = urllib.request.urlopen(f"{BASE_URL}/api/past_tenders?portal=cpp&limit=20")
        d_cpp = json.loads(req_cpp.read().decode('utf-8'))
        for t in d_cpp.get("tenders", []):
            self.assertEqual(t.get("portal"), "cpp")

        req_bihar = urllib.request.urlopen(f"{BASE_URL}/api/past_tenders?portal=bihar&limit=20")
        d_bihar = json.loads(req_bihar.read().decode('utf-8'))
        for t in d_bihar.get("tenders", []):
            self.assertEqual(t.get("portal"), "bihar")

    def test_05_api_cpp_tenders_status_filter(self):
        # 1. Default / Active only
        req_act = urllib.request.urlopen(f"{BASE_URL}/api/cpp/tenders?status=active&limit=100")
        d_act = json.loads(req_act.read().decode('utf-8'))
        self.assertTrue(d_act.get("success"))
        for t in d_act.get("tenders", []):
            self.assertFalse(t.get("is_past"), f"Tender {t.get('tender_id')} should be active")

        # 2. Past only
        req_past = urllib.request.urlopen(f"{BASE_URL}/api/cpp/tenders?status=past&limit=100")
        d_past = json.loads(req_past.read().decode('utf-8'))
        self.assertTrue(d_past.get("success"))
        for t in d_past.get("tenders", []):
            self.assertTrue(t.get("is_past"), f"Tender {t.get('tender_id')} should be expired")

    def test_06_api_excel_tenders_status_filter(self):
        # 1. Active status
        req_act = urllib.request.urlopen(f"{BASE_URL}/api/excel_tenders?status=active")
        d_act = json.loads(req_act.read().decode('utf-8'))
        self.assertEqual(d_act.get("statusFilter"), "active")
        self.assertGreater(len(d_act.get("tenders", [])), 50)
        for t in d_act.get("tenders", []):
            self.assertFalse(t.get("is_past"))

        # Default is strictly active
        req_def = urllib.request.urlopen(f"{BASE_URL}/api/excel_tenders")
        d_def = json.loads(req_def.read().decode('utf-8'))
        self.assertEqual(d_def.get("statusFilter"), "active")
        self.assertEqual(len(d_def.get("tenders", [])), len(d_act.get("tenders", [])))

        # 2. All status (explicit ?status=all)
        req_all = urllib.request.urlopen(f"{BASE_URL}/api/excel_tenders?status=all")
        d_all = json.loads(req_all.read().decode('utf-8'))
        self.assertGreaterEqual(len(d_all.get("tenders", [])), 1454)
        self.assertGreater(d_all.get("totalActive", 0), 50)
        self.assertGreater(d_all.get("totalPast", 0), 1000)

    def test_07_past_tenders_excel_export(self):
        url = f"{BASE_URL}/api/past_tenders/export_xlsx"
        req = urllib.request.urlopen(url, timeout=20)
        self.assertEqual(req.status, 200)
        content_type = req.headers.get("Content-Type")
        self.assertIn("spreadsheetml.sheet", content_type)
        body = req.read()
        self.assertGreater(len(body), 50000, "Exported past Excel workbook should be substantial")
        # Starts with PK zip header for .xlsx
        self.assertTrue(body.startswith(b"PK\x03\x04"))

    def test_08_bihar_and_cpp_raw_data_integrity(self):
        with open('excel_tenders_view.json', 'r', encoding='utf-8') as f:
            b_data = json.load(f)
        self.assertGreaterEqual(len(b_data.get('tenders', [])), 1454, "Bihar tenders must remain at least 1454 intact")

        store = load_cpp_tenders_store()
        self.assertEqual(len(store.get('tenders', [])), 2631, "CPP tenders must remain 2631 intact")

if __name__ == '__main__':
    unittest.main()
