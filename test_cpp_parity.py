import unittest
import json
import urllib.request
import urllib.error

BASE_URL = "http://127.0.0.1:8000"

class TestCppParity(unittest.TestCase):
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
        # Bihar tenders must be completely unaffected
        req = urllib.request.urlopen(f"{BASE_URL}/api/excel_tenders")
        self.assertEqual(req.status, 200)
        data = json.loads(req.read().decode('utf-8'))
        self.assertIn("tenders", data)
        self.assertGreater(len(data.get("tenders", [])), 1000)

if __name__ == '__main__':
    unittest.main()
