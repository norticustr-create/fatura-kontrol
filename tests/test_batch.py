"""
Focused unit and integration tests for batch.py.

Verifies:
1. Single XML file input
2. Folder containing multiple XML files
3. ZIP archive containing multiple XML files
4. Fault tolerance: invalid/missing XML among valid files
5. Empty folder and empty ZIP handling
6. Non-XML files are ignored
7. Full 28-file baseline verification (26 PASS, 2 REVIEW, 0 ERROR)
"""

from decimal import Decimal
import io
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import zipfile

# Ensure project root is in sys.path
PROJECT_ROOT = (
    Path(__file__).resolve().parent.parent
    if Path(__file__).parent.name == "tests"
    else Path(__file__).resolve().parent
)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from batch import collect_files, process_batch, batch_summary, cleanup_temp_dirs


# Minimal valid UBL invoice XML template for test fixtures
VALID_INVOICE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
         xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
         xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
  <cbc:ID>{invoice_id}</cbc:ID>
  <cbc:UUID>11111111-2222-3333-4444-555555555555</cbc:UUID>
  <cbc:IssueDate>2026-03-01</cbc:IssueDate>
  <cbc:IssueTime>10:00:00</cbc:IssueTime>
  <cbc:InvoiceTypeCode>SATIS</cbc:InvoiceTypeCode>
  <cbc:DocumentCurrencyCode>TRY</cbc:DocumentCurrencyCode>
  <cbc:LineCountNumeric>1</cbc:LineCountNumeric>
  <cac:AccountingSupplierParty>
    <cac:Party>
      <cac:PartyName>
        <cbc:Name>Test Supplier Ltd</cbc:Name>
      </cac:PartyName>
    </cac:Party>
  </cac:AccountingSupplierParty>
  <cac:TaxTotal>
    <cbc:TaxAmount>20.00</cbc:TaxAmount>
  </cac:TaxTotal>
  <cac:LegalMonetaryTotal>
    <cbc:LineExtensionAmount>100.00</cbc:LineExtensionAmount>
    <cbc:TaxExclusiveAmount>100.00</cbc:TaxExclusiveAmount>
    <cbc:TaxInclusiveAmount>120.00</cbc:TaxInclusiveAmount>
    <cbc:PayableAmount>120.00</cbc:PayableAmount>
  </cac:LegalMonetaryTotal>
  <cac:InvoiceLine>
    <cbc:ID>1</cbc:ID>
    <cbc:InvoicedQuantity unitCode="C62">1</cbc:InvoicedQuantity>
    <cac:Item>
      <cbc:Name>Sample Item</cbc:Name>
    </cac:Item>
    <cac:Price>
      <cbc:PriceAmount>100.00</cbc:PriceAmount>
    </cac:Price>
    <cbc:LineExtensionAmount>100.00</cbc:LineExtensionAmount>
    <cac:TaxTotal>
      <cbc:TaxAmount>20.00</cbc:TaxAmount>
      <cac:TaxSubtotal>
        <cbc:Percent>20.0</cbc:Percent>
      </cac:TaxSubtotal>
    </cac:TaxTotal>
  </cac:InvoiceLine>
</Invoice>"""


class TestBatchProcessing(unittest.TestCase):

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_batch_workspace_"))

    def tearDown(self):
        cleanup_temp_dirs()
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_single_xml(self):
        """Verify collect_files and process_batch on a single XML file."""
        xml_file = self.temp_dir / "single_inv.xml"
        xml_file.write_text(VALID_INVOICE_XML.format(invoice_id="INV-001"), encoding="utf-8")

        collected = collect_files(xml_file)
        self.assertEqual(len(collected), 1)
        self.assertEqual(collected[0].resolve(), xml_file.resolve())

        results = process_batch(collected)
        self.assertEqual(len(results), 1)
        r = results[0]
        self.assertTrue(r["success"])
        self.assertEqual(r["status"], "PASS")
        self.assertEqual(r["invoice_number"], "INV-001")
        self.assertEqual(r["supplier_name"], "Test Supplier Ltd")
        self.assertEqual(r["payable_amount"], Decimal("120.00"))
        self.assertEqual(r["failed_checks"], [])

    def test_progress_callback(self):
        """Verify progress_callback is called with index, total, and result dictionary."""
        (self.temp_dir / "inv1.xml").write_text(VALID_INVOICE_XML.format(invoice_id="CB-1"), encoding="utf-8")
        (self.temp_dir / "inv2.xml").write_text(VALID_INVOICE_XML.format(invoice_id="CB-2"), encoding="utf-8")
        files = collect_files(self.temp_dir)

        callback_calls = []
        def my_callback(cur, tot, res):
            callback_calls.append((cur, tot, res["invoice_number"], res["status"]))

        results = process_batch(files, progress_callback=my_callback)
        self.assertEqual(len(callback_calls), 2)
        self.assertEqual(callback_calls[0], (1, 2, "CB-1", "PASS"))
        self.assertEqual(callback_calls[1], (2, 2, "CB-2", "PASS"))

    def test_folder_with_multiple_xml(self):
        """Verify collect_files finds XML files in folders and subfolders."""
        # Root level
        (self.temp_dir / "inv1.xml").write_text(VALID_INVOICE_XML.format(invoice_id="INV-1"), encoding="utf-8")
        (self.temp_dir / "inv2.xml").write_text(VALID_INVOICE_XML.format(invoice_id="INV-2"), encoding="utf-8")
        # Subfolder level
        sub_dir = self.temp_dir / "subfolder"
        sub_dir.mkdir()
        (sub_dir / "inv3.xml").write_text(VALID_INVOICE_XML.format(invoice_id="INV-3"), encoding="utf-8")
        # Non-XML files
        (self.temp_dir / "notes.txt").write_text("ignore this", encoding="utf-8")
        (sub_dir / "data.xlsx").write_text("binary dummy", encoding="utf-8")

        collected = collect_files(self.temp_dir)
        self.assertEqual(len(collected), 3)

        results = process_batch(collected)
        self.assertEqual(len(results), 3)
        summary = batch_summary(results)
        self.assertEqual(summary["total"], 3)
        self.assertEqual(summary["pass"], 3)
        self.assertEqual(summary["error"], 0)

    def test_zip_with_multiple_xml(self):
        """Verify collect_files extracts XML files from a ZIP archive."""
        zip_path = self.temp_dir / "invoices.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("invoice_a.xml", VALID_INVOICE_XML.format(invoice_id="ZIP-A"))
            zf.writestr("sub/invoice_b.xml", VALID_INVOICE_XML.format(invoice_id="ZIP-B"))
            zf.writestr("readme.txt", "documentation")
            zf.writestr("image.png", b"\x89PNG\r\n\x1a\n")

        collected = collect_files(zip_path)
        self.assertEqual(len(collected), 2)

        results = process_batch(collected)
        self.assertEqual(len(results), 2)
        summary = batch_summary(results)
        self.assertEqual(summary["pass"], 2)

        # Verify container tracking
        for r in results:
            self.assertEqual(r["container"], str(zip_path.resolve()))
            self.assertIn(r["original_name"], ["invoice_a.xml", "sub/invoice_b.xml"])

    def test_invalid_xml_fault_tolerance(self):
        """Verify process_batch records errors per file without halting the batch."""
        valid_file = self.temp_dir / "valid.xml"
        valid_file.write_text(VALID_INVOICE_XML.format(invoice_id="VALID-01"), encoding="utf-8")

        corrupt_file = self.temp_dir / "corrupt.xml"
        corrupt_file.write_text("<Invoice><cbc:ID>broken unclosed xml", encoding="utf-8")

        missing_file = self.temp_dir / "non_existent.xml"

        results = process_batch([valid_file, corrupt_file, missing_file])
        self.assertEqual(len(results), 3)

        # 1. Valid file
        self.assertTrue(results[0]["success"])
        self.assertEqual(results[0]["status"], "PASS")

        # 2. Corrupt XML
        self.assertFalse(results[1]["success"])
        self.assertEqual(results[1]["status"], "ERROR")
        self.assertIn("ParseError", results[1]["error"])
        self.assertIsNone(results[1]["invoice"])

        # 3. Missing file
        self.assertFalse(results[2]["success"])
        self.assertEqual(results[2]["status"], "ERROR")
        self.assertIn("FileNotFoundError", results[2]["error"])

        summary = batch_summary(results)
        self.assertEqual(summary["pass"], 1)
        self.assertEqual(summary["error"], 2)
        self.assertEqual(summary["total"], 3)

    def test_empty_folder_and_empty_zip(self):
        """Verify empty folder and empty zip return empty lists without error."""
        empty_dir = self.temp_dir / "empty_dir"
        empty_dir.mkdir()
        self.assertEqual(collect_files(empty_dir), [])

        empty_zip = self.temp_dir / "empty.zip"
        with zipfile.ZipFile(empty_zip, "w"):
            pass
        self.assertEqual(collect_files(empty_zip), [])

        zero_byte_zip = self.temp_dir / "zero.zip"
        zero_byte_zip.touch()
        self.assertEqual(collect_files(zero_byte_zip), [])

        results = process_batch([])
        self.assertEqual(results, [])
        self.assertEqual(batch_summary([]), {"total": 0, "pass": 0, "review": 0, "error": 0})

    def test_non_xml_files_ignored(self):
        """Verify non-XML files are ignored by collect_files."""
        txt_file = self.temp_dir / "document.txt"
        txt_file.write_text("Hello world", encoding="utf-8")
        self.assertEqual(collect_files(txt_file), [])

        folder = self.temp_dir / "non_xml_folder"
        folder.mkdir()
        (folder / "file.txt").write_text("txt", encoding="utf-8")
        (folder / "file.xlsx").write_text("xlsx", encoding="utf-8")
        self.assertEqual(collect_files(folder), [])

    def test_missing_source_raises_file_not_found(self):
        """Verify collect_files raises FileNotFoundError on non-existent path."""
        with self.assertRaises(FileNotFoundError):
            collect_files("totally_missing_source_12345.xyz")

    def test_all_28_workspace_invoices_baseline(self):
        """Verify batch processing against the 28 workspace test invoices matches 26 PASS, 2 REVIEW, 0 ERROR."""
        xml_dir = (
            PROJECT_ROOT / "test_data" / "xml"
            if (PROJECT_ROOT / "test_data" / "xml").exists()
            else PROJECT_ROOT
        )
        files = [
            p for p in xml_dir.glob("test*.xml")
        ] + [
            p for p in xml_dir.glob("TEST_*.xml")
        ]
        # Deduplicate
        seen = set()
        unique_files = []
        for f in files:
            if f.name.lower() not in seen:
                seen.add(f.name.lower())
                unique_files.append(f)

        self.assertEqual(len(unique_files), 28)

        results = process_batch(unique_files)
        summary = batch_summary(results)

        self.assertEqual(summary["total"], 28)
        self.assertEqual(summary["pass"], 26)
        self.assertEqual(summary["review"], 2)
        self.assertEqual(summary["error"], 0)

        # Verify specific expected REVIEW items
        review_files = {r["file_name"]: r["failed_checks"] for r in results if r["status"] == "REVIEW"}
        self.assertIn("test21.xml", review_files)
        self.assertEqual(review_files["test21.xml"], ["Payable total"])
        self.assertIn("TEST_BAD_VAT.xml", review_files)
        self.assertEqual(set(review_files["TEST_BAD_VAT.xml"]), {"Tax total calculation", "Line VAT total"})


if __name__ == "__main__":
    unittest.main(verbosity=2)
