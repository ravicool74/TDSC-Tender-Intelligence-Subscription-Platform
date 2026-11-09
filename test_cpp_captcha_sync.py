import unittest
import json
import os
import shutil
import time
import threading
import urllib.request
import urllib.parse
from http.server import HTTPServer
import server
from cpp_tender_manager import (
    load_cpp_tenders_store, save_cpp_tenders_batch,
    sync_cpp_tender_with_captcha, download_all_cpp_tender_attachments,
    CPP_DOWNLOADS_DIR
)

TEST_PORT = 8092

class TestCppCaptchaSync(unittest.TestCase):
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

    def setUp(self):
        self.base_url = f"http://127.0.0.1:{TEST_PORT}"
        self.store = load_cpp_tenders_store()
        self.tenders = self.store.get("tenders", [])
        self.assertTrue(len(self.tenders) > 0, "CPP store should have tenders")

    def test_01_get_tender_sync_captcha_endpoint(self):
        url = f"{self.base_url}/api/cpp/get_tender_sync_captcha"
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get("success"), "get_tender_sync_captcha should succeed")
            self.assertIn("session_id", data)
            self.assertTrue(bool(data.get("captcha_image")), "captcha_image should be present")

    def test_02_sync_with_invalid_captcha(self):
        # Initialize session first
        url_init = f"{self.base_url}/api/cpp/get_tender_sync_captcha"
        with urllib.request.urlopen(url_init, timeout=15) as resp:
            data_init = json.loads(resp.read().decode('utf-8'))
            sid = data_init["session_id"]

        first_tender = self.tenders[0]
        tid = first_tender.get("tender_id") or first_tender.get("canonical_key")

        # Submit deliberately invalid captcha
        post_data = json.dumps({
            "session_id": sid,
            "captcha_text": "WRONG99",
            "tender_id": tid
        }).encode('utf-8')
        req = urllib.request.Request(f"{self.base_url}/api/cpp/sync_with_captcha", data=post_data, headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req, timeout=20) as resp:
            res = json.loads(resp.read().decode('utf-8'))
            self.assertFalse(res.get("success"), "Should fail on invalid captcha")
            self.assertTrue(res.get("captcha_error"), "Should set captcha_error flag")
            self.assertIn("captcha_image", res, "Should provide new captcha image on retry")

    def test_03_downloaded_file_viewer(self):
        # Create a mock downloaded file in a test tender directory
        test_tid = "TEST_TID_99999"
        test_dir = os.path.join(CPP_DOWNLOADS_DIR, f"tender_{test_tid}")
        os.makedirs(test_dir, exist_ok=True)
        test_file = os.path.join(test_dir, "Tendernotice_test.pdf")
        with open(test_file, 'wb') as f:
            f.write(b"%PDF-1.4 test dummy pdf content for tdsc platform")

        try:
            view_url = f"{self.base_url}/api/cpp/view_file?tender_id={test_tid}&filename=Tendernotice_test.pdf"
            with urllib.request.urlopen(view_url, timeout=10) as resp:
                self.assertEqual(resp.status, 200)
                content_type = resp.headers.get('Content-Type')
                self.assertEqual(content_type, 'application/pdf')
                body = resp.read()
                self.assertTrue(body.startswith(b"%PDF-1.4"))
        finally:
            if os.path.exists(test_dir):
                shutil.rmtree(test_dir, ignore_errors=True)

    def test_04_download_all_detects_existing_folder_files(self):
        first_tender = self.tenders[0]
        tid = first_tender.get("tender_id") or first_tender.get("canonical_key")
        clean_tid = tid.replace('/', '_').replace('-', '_')
        test_dir = os.path.join(CPP_DOWNLOADS_DIR, f"tender_{clean_tid}")
        os.makedirs(test_dir, exist_ok=True)
        mock_file = os.path.join(test_dir, "NIT_Notice_1.pdf")
        with open(mock_file, 'wb') as f:
            f.write(b"%PDF-1.4 test file")

        try:
            res = download_all_cpp_tender_attachments(tid)
            self.assertTrue(res.get("success"))
            self.assertTrue(res.get("already_downloaded"))
            self.assertGreater(res.get("downloadedCount", 0), 0)
        finally:
            if os.path.exists(test_dir):
                shutil.rmtree(test_dir, ignore_errors=True)

    def test_05_bihar_eproc2_unmodified_isolation(self):
        with open('excel_tenders_view.json', 'r', encoding='utf-8') as f:
            bihar_data = json.load(f)
        self.assertEqual(len(bihar_data.get('tenders', [])), 1454, "Bihar eProc2 records must remain exactly 1,454")
        self.assertTrue(os.path.exists('Bihar_eProc_Tenders_Comprehensive.xlsx'))

if __name__ == '__main__':
    unittest.main()
