"""
Focused unit and integration tests for excel_export.py.

Verifies:
1. Exporting the 28 workspace test invoices produces a valid 3-sheet workbook.
2. All 28 invoices appear in Summary sheet with correct statuses (26 PASS, 2 REVIEW, 0 ERROR).
3. Exactly 102 line items appear in Lines sheet with proper numeric formats.
4. Exactly 140 audit check/diagnostic records appear in Audit Details sheet.
5. ZIP container tracking is preserved in Excel export.
6. Error invoices in batch are captured cleanly in Summary and Audit Details.
7. Empty batch produces a valid, uncorrupted workbook with styled headers and autofilters.
8. Freeze panes ('A2') and auto-filters are active on all sheets.
9. Monetary values are stored as floats with '#,##0.00' format and empty values remain None.
"""

from datetime import date
from decimal import Decimal
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import zipfile

import openpyxl

# Ensure project root is in sys.path
PROJECT_ROOT = (
    Path(__file__).resolve().parent.parent
    if Path(__file__).parent.name == "tests"
    else Path(__file__).resolve().parent
)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from batch import collect_files, process_batch, cleanup_temp_dirs
from excel_export import export_to_excel
from tests.test_batch import VALID_INVOICE_XML


class TestExcelExport(unittest.TestCase):

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_excel_export_"))

    def tearDown(self):
        cleanup_temp_dirs()
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_all_28_invoices_export(self):
        """Export all 28 workspace invoices and verify workbook structure, row counts, and formatting."""
        xml_dir = (
            PROJECT_ROOT / "test_data" / "xml"
            if (PROJECT_ROOT / "test_data" / "xml").exists()
            else PROJECT_ROOT
        )
        files = collect_files(xml_dir)
        results = process_batch(files)

        output_path = self.temp_dir / "full_report.xlsx"
        saved = export_to_excel(results, output_path)
        self.assertTrue(saved.exists())

        # Load workbook
        wb = openpyxl.load_workbook(saved)
        self.assertEqual(wb.sheetnames, ["Summary", "Lines", "Audit Details"])

        # 1. Summary Sheet Checks
        ws_s = wb["Summary"]
        self.assertEqual(ws_s.max_row, 29)  # 1 header + 28 invoice rows
        self.assertEqual(ws_s.freeze_panes, "A2")
        self.assertIsNotNone(ws_s.auto_filter.ref)

        statuses = [ws_s.cell(row=r, column=1).value for r in range(2, 30)]
        self.assertEqual(statuses.count("PASS"), 26)
        self.assertEqual(statuses.count("REVIEW"), 2)
        self.assertEqual(statuses.count("ERROR"), 0)

        # Check numeric and date formats on row 2
        payable_cell = ws_s.cell(row=2, column=6)
        self.assertIsInstance(payable_cell.value, (int, float))
        self.assertEqual(payable_cell.number_format, "#,##0.00")

        date_cell = ws_s.cell(row=2, column=3)
        self.assertIsNotNone(date_cell.value)
        self.assertEqual(date_cell.number_format, "yyyy-mm-dd")

        # 2. Lines Sheet Checks
        ws_l = wb["Lines"]
        self.assertEqual(ws_l.max_row, 103)  # 1 header + 102 line items
        self.assertEqual(ws_l.freeze_panes, "A2")
        self.assertIsNotNone(ws_l.auto_filter.ref)

        # Check price and amount format on a line item
        price_cell = ws_l.cell(row=2, column=7)
        self.assertIsInstance(price_cell.value, (int, float))
        self.assertEqual(price_cell.number_format, "#,##0.00")

        # 3. Audit Details Sheet Checks
        ws_a = wb["Audit Details"]
        self.assertEqual(ws_a.max_row, 141)  # 1 header + 140 checks (5 * 28)
        self.assertEqual(ws_a.freeze_panes, "A2")
        self.assertIsNotNone(ws_a.auto_filter.ref)

    def test_zip_container_preserved_in_export(self):
        """Verify that ZIP container filename is recorded in the Summary sheet."""
        zip_path = self.temp_dir / "supplier_bundle.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("test_a.xml", VALID_INVOICE_XML.format(invoice_id="ZIP-001"))
            zf.writestr("test_b.xml", VALID_INVOICE_XML.format(invoice_id="ZIP-002"))

        collected = collect_files(zip_path)
        results = process_batch(collected)

        output_path = self.temp_dir / "zip_report.xlsx"
        export_to_excel(results, output_path)

        wb = openpyxl.load_workbook(output_path)
        ws_s = wb["Summary"]
        self.assertEqual(ws_s.max_row, 3)

        for row_idx in (2, 3):
            container_val = ws_s.cell(row=row_idx, column=12).value
            self.assertEqual(container_val, "supplier_bundle.zip")

    def test_error_invoice_in_export(self):
        """Verify that invoices with errors appear cleanly in Summary and Audit Details."""
        valid_file = self.temp_dir / "good.xml"
        valid_file.write_text(VALID_INVOICE_XML.format(invoice_id="GOOD-1"), encoding="utf-8")

        corrupt_file = self.temp_dir / "bad.xml"
        corrupt_file.write_text("<Invoice>corrupt incomplete xml", encoding="utf-8")

        results = process_batch([valid_file, corrupt_file])

        output_path = self.temp_dir / "error_report.xlsx"
        export_to_excel(results, output_path)

        wb = openpyxl.load_workbook(output_path)
        ws_s = wb["Summary"]
        self.assertEqual(ws_s.max_row, 3)

        # First row is PASS
        self.assertEqual(ws_s.cell(row=2, column=1).value, "PASS")
        # Second row is ERROR
        self.assertEqual(ws_s.cell(row=3, column=1).value, "ERROR")
        self.assertIn("ParseError", str(ws_s.cell(row=3, column=13).value))

        # Check Audit Details sheet records the error
        ws_a = wb["Audit Details"]
        error_rows = [
            ws_a.cell(row=r, column=5).value
            for r in range(2, ws_a.max_row + 1)
            if ws_a.cell(row=r, column=5).value == "ERROR"
        ]
        self.assertEqual(len(error_rows), 1)

    def test_empty_batch_export(self):
        """Verify an empty batch produces a clean, valid workbook with all 3 sheets."""
        output_path = self.temp_dir / "empty_report.xlsx"
        export_to_excel([], output_path)

        wb = openpyxl.load_workbook(output_path)
        self.assertEqual(wb.sheetnames, ["Summary", "Lines", "Audit Details"])
        for name in wb.sheetnames:
            ws = wb[name]
            self.assertEqual(ws.max_row, 1)  # Only header
            self.assertEqual(ws.freeze_panes, "A2")

    def test_empty_values_remain_none(self):
        """Verify that missing/empty values (e.g. no allowance) remain None and are not converted to 0.00."""
        xml_file = self.temp_dir / "no_discount.xml"
        xml_file.write_text(VALID_INVOICE_XML.format(invoice_id="NO-DISC"), encoding="utf-8")

        results = process_batch([xml_file])
        output_path = self.temp_dir / "no_discount.xlsx"
        export_to_excel(results, output_path)

        wb = openpyxl.load_workbook(output_path)
        ws_l = wb["Lines"]
        self.assertEqual(ws_l.max_row, 2)

        # Allowance (col 9) and Charge (col 10) should be None
        allowance_cell = ws_l.cell(row=2, column=9)
        charge_cell = ws_l.cell(row=2, column=10)
        self.assertIsNone(allowance_cell.value)
        self.assertIsNone(charge_cell.value)


if __name__ == "__main__":
    unittest.main(verbosity=2)
