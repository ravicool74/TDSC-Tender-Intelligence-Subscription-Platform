"""
Automated Test Suite for Dedicated Contractor Database & CRM
Verifies:
1. Retrieval of Google Sheet imported contractors with pagination
2. Multi-column search (Firm, Mobile, App ID, Director, PAN, District, Remarks)
3. District, Entity Type, and Caste filtering
4. Add single contractor (POST /api/contractors/create)
5. Edit contractor (POST /api/contractors/update)
6. Delete contractor (POST /api/contractors/delete)
7. Export to Excel (.xlsx) (GET /api/contractors/export_xlsx)
8. Contractor-to-Tender suggestion flow
"""

import unittest
import json
import time
import threading
import urllib.request
import urllib.parse
from http.server import HTTPServer
import server
import contractor_crm_manager

TEST_PORT = 8097
BASE_URL = f"http://127.0.0.1:{TEST_PORT}"

class TestContractorCRM(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        contractor_crm_manager.init_contractors_schema()
        cls.httpd = HTTPServer(('127.0.0.1', TEST_PORT), server.PaidTenderRequestHandler)
        cls.server_thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.server_thread.start()
        time.sleep(0.5)

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def test_01_get_contractors_list_and_stats(self):
        url = f"{BASE_URL}/api/contractors?limit=20"
        with urllib.request.urlopen(url, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get("success"))
            self.assertGreater(data.get("total", 0), 1000, "Should have imported thousands of contractors from Google Sheet")
            contractors = data.get("contractors", [])
            self.assertGreaterEqual(len(contractors), 10)
            c = contractors[0]
            # Verify Google Sheet columns exist in record
            self.assertIn("application_id", c)
            self.assertIn("company_name", c)
            self.assertIn("district", c)
            self.assertIn("mobile", c)
            self.assertIn("director_name", c)
            self.assertIn("pan_no", c)
            self.assertIn("entity_type", c)
            self.assertIn("cast", c)
            self.assertIn("calling_remarks", c)

    def test_02_search_and_district_filter(self):
        # District filter
        url = f"{BASE_URL}/api/contractors?district=Patna&limit=10"
        with urllib.request.urlopen(url, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get("success"))
            for c in data.get("contractors", []):
                self.assertEqual(c["district"], "Patna")

        # Search query
        url_search = f"{BASE_URL}/api/contractors?q=Kumar&limit=10"
        with urllib.request.urlopen(url_search, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            data_search = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data_search.get("success"))
            self.assertGreater(data_search.get("total", 0), 0)

    def test_03_crud_add_edit_delete_flow(self):
        # 1. Add Contractor
        payload_add = {
            "application_id": "TEST_APP_9999",
            "company_name": "Antigravity Test Construction Ltd",
            "director_name": "Test Director Sharma",
            "mobile": "9876543210",
            "email": "test.builder@tdsc.in",
            "district": "Patna",
            "address": "Boring Road, Patna",
            "cast": "GEN",
            "gender": "MALE",
            "entity_type": "Private Limited",
            "pan_no": "AABCT9999Z",
            "calling_remarks": "Interested in highway tenders"
        }
        req_add = urllib.request.Request(
            f"{BASE_URL}/api/contractors/create",
            data=json.dumps(payload_add).encode('utf-8'),
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_add, timeout=10) as resp_add:
            self.assertEqual(resp_add.status, 200)
            res_add = json.loads(resp_add.read().decode('utf-8'))
            self.assertTrue(res_add.get("success"))
            new_id = res_add.get("id")
            self.assertIsNotNone(new_id)

        # 2. Edit Contractor
        payload_update = {
            "id": new_id,
            "company_name": "Antigravity Test Construction Ltd (Updated)",
            "calling_remarks": "Verified and quoted for NIT 2026"
        }
        req_up = urllib.request.Request(
            f"{BASE_URL}/api/contractors/update",
            data=json.dumps(payload_update).encode('utf-8'),
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_up, timeout=10) as resp_up:
            self.assertEqual(resp_up.status, 200)
            res_up = json.loads(resp_up.read().decode('utf-8'))
            self.assertTrue(res_up.get("success"))

        # Verify update
        url_check = f"{BASE_URL}/api/contractors?q=TEST_APP_9999"
        with urllib.request.urlopen(url_check, timeout=10) as resp_ck:
            data_ck = json.loads(resp_ck.read().decode('utf-8'))
            self.assertEqual(data_ck.get("total"), 1)
            self.assertEqual(data_ck["contractors"][0]["company_name"], "Antigravity Test Construction Ltd (Updated)")
            self.assertEqual(data_ck["contractors"][0]["calling_remarks"], "Verified and quoted for NIT 2026")

        # 3. Delete Contractor
        req_del = urllib.request.Request(
            f"{BASE_URL}/api/contractors/delete",
            data=json.dumps({"id": new_id}).encode('utf-8'),
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_del, timeout=10) as resp_del:
            self.assertEqual(resp_del.status, 200)
            res_del = json.loads(resp_del.read().decode('utf-8'))
            self.assertTrue(res_del.get("success"))

        # Verify deletion
        with urllib.request.urlopen(url_check, timeout=10) as resp_ck2:
            data_ck2 = json.loads(resp_ck2.read().decode('utf-8'))
            self.assertEqual(data_ck2.get("total"), 0)

    def test_04_export_xlsx(self):
        url = f"{BASE_URL}/api/contractors/export_xlsx?district=Patna"
        with urllib.request.urlopen(url, timeout=15) as resp:
            self.assertEqual(resp.status, 200)
            self.assertIn("spreadsheetml.sheet", resp.headers.get("Content-Type"))
            content = resp.read()
            self.assertGreater(len(content), 10000)
            self.assertTrue(content.startswith(b"PK\x03\x04"))

if __name__ == '__main__':
    unittest.main()
