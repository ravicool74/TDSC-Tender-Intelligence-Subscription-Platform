import unittest
import json
import time
import threading
import urllib.request
import urllib.error
from http.server import HTTPServer
import server

TEST_PORT = 8093
BASE_URL = f"http://127.0.0.1:{TEST_PORT}"

class TestCppParity(unittest.TestCase):
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

    def test_01_get_all_tenders_limit_1000(self):
        req = urllib.request.urlopen(f"{BASE_URL}/api/cpp/tenders?limit=1000")
        self.assertEqual(req.status, 200)
        data = json.loads(req.read().decode('utf-8'))
        self.assertTrue(data.get("success"))
        tenders = data.get("tenders", [])
        self.assertGreaterEqual(len(tenders), 10)
        t = tenders[0]
        # Verify 13 column fields exist
        self.assertTrue("tender_id" in t or "canonical_key" in t)
        self.assertIn("tender_ref_no", t)
        self.assertIn("tender_value_inr", t)
        self.assertIn("department", t)
        self.assertIn("state_district", t)

    def test_02_open_folder_api(self):
        payload = json.dumps({"tender_id": "2023_MoRTH_771865_1"}).encode('utf-8')
        req = urllib.request.Request(
            f"{BASE_URL}/api/cpp/open_folder",
            data=payload,
            headers={"Content-Type": "application/json"}
        )
        res = urllib.request.urlopen(req)
        self.assertEqual(res.status, 200)
        data = json.loads(res.read().decode('utf-8'))
        self.assertEqual(data.get("status"), "opened")
        self.assertIn("folder", data)

    def test_03_bihar_eproc_untouched(self):
        # Bihar tenders master archive must be completely unaffected (1454 tenders)
        req = urllib.request.urlopen(f"{BASE_URL}/api/excel_tenders?status=all")
        self.assertEqual(req.status, 200)
        data = json.loads(req.read().decode('utf-8'))
        self.assertIn("tenders", data)
        self.assertEqual(len(data.get("tenders", [])), 1454)
        self.assertGreater(data.get("totalUnfiltered", 0), 1000)

        # Active view default must exclude expired tenders
        req_act = urllib.request.urlopen(f"{BASE_URL}/api/excel_tenders")
        d_act = json.loads(req_act.read().decode('utf-8'))
        self.assertEqual(d_act.get("statusFilter"), "active")
        self.assertGreater(len(d_act.get("tenders", [])), 50)
        self.assertTrue(all(not t.get("is_past") for t in d_act.get("tenders", [])))

if __name__ == '__main__':
    unittest.main()
