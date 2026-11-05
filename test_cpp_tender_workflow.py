"""
Automated Test Suite for CPP Central Government Tender Subsystem
Tests isolation, duplicate prevention, HTML parsing, Excel generation, and Bihar zero-regression.
"""

import os
import sys
import json
import unittest

from cpp_tender_manager import (
    parse_cpp_search_results_page,
    parse_cpp_tender_detail_page,
    save_cpp_tenders_batch,
    load_cpp_tenders_store,
    toggle_cpp_shortlist,
    generate_cpp_excel_workbook,
    CPP_VIEW_JSON,
    CPP_EXCEL_PATH,
    CPP_DOWNLOADS_DIR
)

SAMPLE_USER_PROMPT_DETAIL_TEXT = """Organisation Chain
Department of Posts||Delhi Circle, Dept. of Posts||Off. of EE(C),Delhi Circle,POST||Postal Civil Division,EE(C),Delhi Circle
Tender Reference Number
NIT/EEPCND/2026-27/27
Tender ID
2026_DOP_933186_1
S.No
Instrument Type
1
Demand Draft
2
R-T-G-S
3
FDR
4
NEFT
5
Bankers Cheque
6
Bank Guarantee
Tender Fee Details, [Total Fee in ₹ * - 0.00]
Tender Fee in ₹
0.00
Fee Payable To
Nil
Fee Payable At
Nil
Tender Fee Exemption Allowed
No
EMD Fee Details:
EMD Amount in ₹
21,538
EMD Exemption Allowed
No
EMD Fee Type
fixed
EMD Percentage
NA
EMD Payable To
Accounts Officer Postal Civil Division New Delhi
EMD Payable At
Accounts Officer Postal Civil Division New Delhi
Work Item Details:-
Title
Provision of day to day maintenance of Postal staff quarters, AO buildings and PO buildings under Postal Civil Enquiry, Sarojini Nagar, New Delhi.
Work Description
Provision of day to day maintenance of Postal staff quarters, AO buildings and PO buildings under Postal Civil Enquiry, Sarojini Nagar, New Delhi.
Tender Value in ₹
10,76,893
Contract Type
Tender
Critical Dates:
Published Date
04-Nov-2026 06:55 PM
Bid Opening Date
13-Nov-2026 03:30 PM
Document Download / Sale Start Date
04-Nov-2026 06:55 PM
Document Download / Sale End Date
12-Nov-2026 03:00 PM
Clarification Start Date
NA
Clarification End Date
NA
Bid Submission Start Date
04-Nov-2026 06:55 PM
Bid Submission End Date
12-Nov-2026 03:00 PM
Tenders Documents:
NIT Document
S.No
Document Name
Description
Document Size (in KB)
1
Tendernotice_1.pdf
NIT
1247.79
Download as zip file
"""

SAMPLE_SEARCH_RESULTS_HTML = """
<html>
<body>
<table id="table" class="list_table">
  <tr>
    <th>S.No</th>
    <th>e-Published Date</th>
    <th>Bid Submission Closing Date</th>
    <th>Tender Opening Date</th>
    <th>Title and Ref.No./Tender ID</th>
    <th>Organisation Chain</th>
    <th>Tender Value in ₹</th>
  </tr>
  <tr>
    <td>1</td>
    <td>04-Nov-2026 06:55 PM</td>
    <td>12-Nov-2026 03:00 PM</td>
    <td>13-Nov-2026 03:30 PM</td>
    <td><a href="/eprocure/app?page=FrontEndTenderDetails&service=direct&session=T&sp=sp1">Provision of day to day maintenance of Postal staff quarters<br>NIT/EEPCND/2026-27/27 [2026_DOP_933186_1]</a></td>
    <td>Department of Posts||Delhi Circle</td>
    <td>10,76,893</td>
  </tr>
  <tr>
    <td>2</td>
    <td>03-Nov-2026 05:00 PM</td>
    <td>26-Nov-2026 01:00 PM</td>
    <td>27-Nov-2026 03:30 PM</td>
    <td><a href="/eprocure/app?page=FrontEndTenderDetails&service=direct&session=T&sp=sp2">Civil Maintenance work at IIT Roorkee<br>IWD/IE/2026-27/E Tender/C-89 [2026_IITR_882912_1]</a></td>
    <td>IIT Roorkee||Estate Civil</td>
    <td>45,00,000</td>
  </tr>
</table>
<div class="pagination">
  <a href="/eprocure/app?component=%24TablePages.linkPage&page=FrontEndAdvancedSearchResult&service=direct&session=T&sp=AFrontEndAdvancedSearchResult%2Ctable&sp=2">2</a>
  <a href="/eprocure/app?component=%24TablePages.linkPage&page=FrontEndAdvancedSearchResult&service=direct&session=T&sp=AFrontEndAdvancedSearchResult%2Ctable&sp=3">3</a>
  <a href="/eprocure/app?component=%24TablePages.linkPage&page=FrontEndAdvancedSearchResult&service=direct&session=T&sp=AFrontEndAdvancedSearchResult%2Ctable&sp=2">Next &gt;</a>
</div>
</body>
</html>
"""

class TestCppTenderWorkflow(unittest.TestCase):

    def setUp(self):
        # Reset CPP store for clean isolated test runs
        with open(CPP_VIEW_JSON, 'w', encoding='utf-8') as f:
            json.dump({
                "portal": "Central Public Procurement Portal (eprocure.gov.in)",
                "updated_at": "2026-11-05T00:00:00",
                "total_count": 0,
                "tenders": []
            }, f, indent=2)

    def test_01_parse_search_results(self):
        """Test parsing of search results table with 7 columns and pagination."""
        res = parse_cpp_search_results_page(SAMPLE_SEARCH_RESULTS_HTML)
        tenders = res.get("tenders", [])
        self.assertEqual(len(tenders), 2)
        
        t1 = tenders[0]
        self.assertEqual(t1["tender_id"], "2026_DOP_933186_1")
        self.assertEqual(t1["tender_value_inr"], 1076893.0)
        self.assertIn("1,076,893", t1["tender_value_formatted"])
        self.assertEqual(t1["budget_slab"], "0to25L")
        self.assertIn("Delhi", t1["state_district"])
        self.assertTrue(t1["detail_url"].startswith("https://eprocure.gov.in"))

        pagination = res.get("pagination", {})
        self.assertTrue(len(pagination.get("available_pages", [])) >= 2)
        self.assertTrue(bool(pagination.get("next_page_url")))

    def test_02_parse_detail_page(self):
        """Test parsing of detailed tender page text as specified by user."""
        detail = parse_cpp_tender_detail_page(SAMPLE_USER_PROMPT_DETAIL_TEXT)
        self.assertEqual(detail["tender_id"], "2026_DOP_933186_1")
        self.assertEqual(detail["tender_ref_no"], "NIT/EEPCND/2026-27/27")
        self.assertEqual(detail["contract_type"], "Tender")
        self.assertEqual(detail["tender_value_inr"], 1076893.0)
        self.assertEqual(detail["emd_amount_inr"], 21538.0)
        self.assertEqual(detail["emd_payable_to"], "Accounts Officer Postal Civil Division New Delhi")
        self.assertEqual(detail["tender_fee_inr"], 0.0)
        
        # Verify instruments
        inst_names = [i["instrument_type"] for i in detail["instruments"]]
        self.assertTrue(any("Demand Draft" in n for n in inst_names))
        self.assertTrue(any("R-T-G-S" in n for n in inst_names))
        self.assertTrue(any("Bank Guarantee" in n for n in inst_names))

        # Verify critical dates
        self.assertEqual(detail["published_date"], "04-Nov-2026 06:55 PM")
        self.assertEqual(detail["bid_opening_date"], "13-Nov-2026 03:30 PM")
        self.assertEqual(detail["bid_submission_end_date"], "12-Nov-2026 03:00 PM")

    def test_03_strict_duplicate_prevention(self):
        """Test that importing duplicate records is strictly skipped."""
        detail = parse_cpp_tender_detail_page(SAMPLE_USER_PROMPT_DETAIL_TEXT)
        
        # First save
        res1 = save_cpp_tenders_batch([detail], allow_update=False)
        self.assertEqual(res1["added"], 1)
        self.assertEqual(res1["skipped_duplicates"], 0)

        # Second save of exact same tender
        res2 = save_cpp_tenders_batch([detail], allow_update=False)
        self.assertEqual(res2["added"], 0)
        self.assertEqual(res2["skipped_duplicates"], 1)

    def test_04_shortlist_and_excel(self):
        """Test shortlist toggle and Excel generation with 4 sheets."""
        detail = parse_cpp_tender_detail_page(SAMPLE_USER_PROMPT_DETAIL_TEXT)
        save_cpp_tenders_batch([detail], allow_update=False)

        tid = "2026_DOP_933186_1"
        res = toggle_cpp_shortlist(tid)
        self.assertTrue(res["success"])
        self.assertTrue(res["is_shortlisted"])

        # Check excel file generated
        xl_path = generate_cpp_excel_workbook()
        self.assertTrue(os.path.exists(xl_path))

        # Verify sheets
        import openpyxl
        wb = openpyxl.load_workbook(xl_path, data_only=True)
        sheet_names = wb.sheetnames
        self.assertIn("CPP Main Tenders", sheet_names)
        self.assertIn("Payment Instruments", sheet_names)
        self.assertIn("Tender Documents", sheet_names)
        self.assertIn("Shortlisted Tenders", sheet_names)

        ws_sl = wb["Shortlisted Tenders"]
        self.assertTrue(ws_sl.max_row >= 2)  # Header + at least 1 row

    def test_05_zero_regression_on_bihar_eproc(self):
        """Verify Bihar eProc files remain completely untouched and valid."""
        bihar_excel = os.path.join(os.path.dirname(CPP_VIEW_JSON), "Bihar_eProc_Tenders_Comprehensive.xlsx")
        bihar_view = os.path.join(os.path.dirname(CPP_VIEW_JSON), "excel_tenders_view.json")
        
        self.assertTrue(os.path.exists(bihar_excel))
        self.assertTrue(os.path.exists(bihar_view))

        with open(bihar_view, 'r', encoding='utf-8') as f:
            b_data = json.load(f)
            self.assertIn("tenders", b_data)
            self.assertTrue(len(b_data["tenders"]) > 0)
            # Ensure no CPP tender got mixed into Bihar view
            self.assertFalse(any(t.get("tenderId") == "2026_DOP_933186_1" for t in b_data["tenders"]))

    def test_06_pagination_loadnext_and_total_records(self):
        """Verify loadNext pagination links and Total records extraction."""
        html = """
        <table id="table" class="list_table">
          <tr><th>S.No</th><th>e-Published Date</th><th>Bid Submission Closing Date</th><th>Tender Opening Date</th><th>Title and Ref.No./Tender ID</th><th>Organisation Chain</th><th>Tender Value in ₹</th></tr>
          <tr><td>1</td><td>04-Nov-2026</td><td>26-Nov-2026</td><td>27-Nov-2026</td><td><a href="/eprocure/app?page=FrontEndTenderDetails&service=direct&sp=XYZ">Supply of Spinning Wheel [CWDB/2026][2026_CWDB_933178_1]</a></td><td>Central Wool Development Board</td><td>1,00,000</td></tr>
        </table>
        <div>Total records: 2605</div>
        <a href="/eprocure/app?component=loadPrevious&page=FrontEndLatestActiveTenders&service=direct&session=T">&lt; Previous</a>
        <a href="/eprocure/app?component=loadNext&page=FrontEndLatestActiveTenders&service=direct&session=T">Next &gt;</a>
        """
        res = parse_cpp_search_results_page(html, base_url="https://eprocure.gov.in/eprocure/app")
        self.assertEqual(len(res["tenders"]), 1)
        pag = res["pagination"]
        self.assertEqual(pag["total_records_portal"], 2605)
        self.assertTrue("component=loadNext" in pag["next_page_url"])
        self.assertTrue(pag["next_page_url"].startswith("https://eprocure.gov.in/eprocure/app?component=loadNext"))

if __name__ == "__main__":
    unittest.main()
