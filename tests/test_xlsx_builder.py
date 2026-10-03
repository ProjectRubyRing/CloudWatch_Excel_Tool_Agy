"""
Unit tests for xlsx_builder.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from lib.xlsx_builder import (
    XlsxWorkbook,
    STYLE_TITLE,
    STYLE_META_LABEL,
    STYLE_META_VALUE,
    STYLE_HEADER_CENTER,
    STYLE_HEADER_LEFT,
    STYLE_STREAM_BANNER,
    STYLE_DATA_CENTER,
    STYLE_DATA_CENTER_ZEBRA,
    STYLE_DATA_LEFT,
    STYLE_DATA_LEFT_ZEBRA,
    STYLE_DATA_MSG,
    STYLE_DATA_MSG_ZEBRA,
    STYLE_HIGHLIGHT_CENTER,
    STYLE_HIGHLIGHT_LEFT,
    STYLE_HIGHLIGHT_MSG,
    STYLE_SUMMARY_HEADER_CENTER,
    STYLE_SUMMARY_HEADER_LEFT,
    STYLE_SECTION_HEADER
)

class TestXlsxBuilder(unittest.TestCase):
    def test_build_and_validate(self):
        output_path = os.path.join(os.path.dirname(__file__), "test_output.xlsx")
        wb = XlsxWorkbook()

        # Sheet 1: Summary
        ws_sum = wb.add_worksheet("Summary", freeze_row=4)
        ws_sum.set_cell(1, 1, "CloudWatch Logs Export Report", STYLE_TITLE)
        ws_sum.merge_cells("A1:D1")
        ws_sum.set_row_height(1, 28)

        ws_sum.set_cell(2, 1, "Report Date:", STYLE_META_LABEL)
        ws_sum.set_cell(2, 2, "2026-10-04 06:30:00 JST", STYLE_META_VALUE)
        ws_sum.set_cell(3, 1, "Time Range:", STYLE_META_LABEL)
        ws_sum.set_cell(3, 2, "2026-10-04 00:00:00 ~ 2026-10-04 06:00:00 JST", STYLE_META_VALUE)

        headers = ["No", "Log Group Name", "Stream Count", "Event Count"]
        ws_sum.set_row_height(5, 24)
        for c, h in enumerate(headers, 1):
            st = STYLE_SUMMARY_HEADER_CENTER if c in (1, 3, 4) else STYLE_SUMMARY_HEADER_LEFT
            ws_sum.set_cell(5, c, h, st)

        ws_sum.set_cell(6, 1, 1, STYLE_DATA_CENTER)
        ws_sum.set_cell(6, 2, "/aws/lambda/user-service", STYLE_DATA_LEFT)
        ws_sum.set_cell(6, 3, 2, STYLE_DATA_CENTER)
        ws_sum.set_cell(6, 4, 150, STYLE_DATA_CENTER)

        # Sheet 2: Log group sheet with invalid chars and long name
        ws_group = wb.add_worksheet("/aws/lambda/very-long-log-group-name-with-invalid:chars?*", freeze_row=4)
        self.assertLessEqual(len(ws_group.name), 31)
        self.assertNotIn(":", ws_group.name)
        self.assertNotIn("?", ws_group.name)
        self.assertNotIn("*", ws_group.name)

        ws_group.set_cell(1, 1, "Log Group: /aws/lambda/user-service", STYLE_TITLE)
        ws_group.set_row_height(1, 26)

        # Stream Banner
        ws_group.set_cell(3, 1, "Stream: 2026/10/04/[$LATEST]abc123def456 (Events: 2)", STYLE_STREAM_BANNER)
        ws_group.merge_cells("A3:E3")
        ws_group.set_row_height(3, 22)

        data_headers = ["No", "Timestamp (JST)", "Log Stream", "Event ID", "Message"]
        ws_group.set_row_height(4, 22)
        for c, h in enumerate(data_headers, 1):
            st = STYLE_HEADER_CENTER if c in (1, 2, 4) else STYLE_HEADER_LEFT
            ws_group.set_cell(4, c, h, st)

        # Normal row
        ws_group.set_cell(5, 1, 1, STYLE_DATA_CENTER)
        ws_group.set_cell(5, 2, "2026-10-04 06:01:23.456", STYLE_DATA_CENTER)
        ws_group.set_cell(5, 3, "2026/10/04/[$LATEST]abc123def456", STYLE_DATA_LEFT)
        ws_group.set_cell(5, 4, "36000000000000001", STYLE_DATA_CENTER)
        ws_group.set_cell(5, 5, "INFO: Request processed successfully.", STYLE_DATA_MSG)

        # Highlight row (e.g. ERROR)
        ws_group.set_cell(6, 1, 2, STYLE_HIGHLIGHT_CENTER)
        ws_group.set_cell(6, 2, "2026-10-04 06:01:24.789", STYLE_HIGHLIGHT_CENTER)
        ws_group.set_cell(6, 3, "2026/10/04/[$LATEST]abc123def456", STYLE_HIGHLIGHT_LEFT)
        ws_group.set_cell(6, 4, "36000000000000002", STYLE_HIGHLIGHT_CENTER)
        ws_group.set_cell(6, 5, "ERROR: Database connection timeout occurred!", STYLE_HIGHLIGHT_MSG)

        wb.save(output_path)
        self.assertTrue(os.path.exists(output_path))

        # Validate with openpyxl if available
        try:
            import openpyxl
            wb_read = openpyxl.load_workbook(output_path)
            self.assertEqual(len(wb_read.sheetnames), 2)
            self.assertEqual(wb_read.sheetnames[0], "Summary")
            ws_r = wb_read["Summary"]
            self.assertEqual(ws_r["A1"].value, "CloudWatch Logs Export Report")
            print("Validation with openpyxl succeeded!")
        except ImportError:
            print("openpyxl not available, standard test completed.")

if __name__ == "__main__":
    unittest.main()
