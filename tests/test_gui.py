"""
Automated unit tests for gui.py without requiring interactive user input.
"""

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch, MagicMock

import tkinter as tk

# Ensure project root is on sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from gui import InvoiceAuditorGUI, set_dpi_awareness


class TestInvoiceAuditorGUI(unittest.TestCase):
    """Unit tests for the InvoiceAuditorGUI class."""

    @classmethod
    def setUpClass(cls):
        set_dpi_awareness()

    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the window during testing
        self.app = InvoiceAuditorGUI(self.root)
        self.root.update()

    def tearDown(self):
        try:
            self.root.destroy()
        except Exception:
            pass

    def test_gui_initial_state(self):
        """Verify initial UI variables and defaults."""
        self.assertEqual(self.app.source_var.get(), "")
        self.assertEqual(self.app.output_var.get(), "invoice_audit_report.xlsx")
        self.assertEqual(self.app.total_count_var.get(), "0")
        self.assertEqual(self.app.pass_count_var.get(), "0")
        self.assertEqual(self.app.review_count_var.get(), "0")
        self.assertEqual(self.app.error_count_var.get(), "0")
        self.assertFalse(self.app.is_processing)
        self.assertIsNone(self.app.last_saved_path)

    def test_suggest_output_path_for_file(self):
        """Verify output path suggestion when a single file is selected."""
        file_path = "C:/test_dir/my_sample_invoice.xml"
        self.app._suggest_output_path(file_path)
        expected = str(Path("C:/test_dir/my_sample_invoice_denetim_raporu.xlsx"))
        self.assertEqual(self.app.output_var.get(), expected)

    def test_suggest_output_path_for_folder(self):
        """Verify output path suggestion when a directory is selected."""
        folder_path = "C:/test_dir/invoices_folder"
        self.app._suggest_output_path(folder_path)
        expected = str(Path("C:/test_dir/invoices_folder/fatura_denetim_raporu.xlsx"))
        self.assertEqual(self.app.output_var.get(), expected)

    @patch("gui.messagebox.showwarning")
    def test_start_processing_empty_source_shows_warning(self, mock_warn):
        """Attempting to process with an empty source triggers a warning message."""
        self.app.source_var.set("")
        self.app.start_processing()
        mock_warn.assert_called_once()
        self.assertFalse(self.app.is_processing)

    @patch("gui.messagebox.showerror")
    def test_start_processing_invalid_source_shows_error(self, mock_err):
        """Attempting to process with a non-existent path triggers an error message."""
        self.app.source_var.set("C:/non_existent_path_xyz_123.xml")
        self.app.start_processing()
        mock_err.assert_called_once()
        self.assertFalse(self.app.is_processing)

    @patch("gui.messagebox.showinfo")
    def test_process_worker_execution(self, mock_info):
        """Test _process_worker with an actual sample fixture directly."""
        sample_xml = PROJECT_ROOT / "test_data" / "xml" / "test-metro.xml"
        self.assertTrue(sample_xml.exists(), "test-metro.xml fixture must exist")

        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "gui_test_report.xlsx"

            # Execute worker logic synchronously
            self.app._process_worker(str(sample_xml), str(out_file))
            self.root.update()

            self.assertTrue(out_file.exists())
            self.assertEqual(self.app.total_count_var.get(), "1")
            self.assertEqual(self.app.pass_count_var.get(), "1")
            self.assertEqual(self.app.review_count_var.get(), "0")
            self.assertEqual(self.app.error_count_var.get(), "0")
            self.assertEqual(self.app.last_saved_path, str(out_file))


if __name__ == "__main__":
    unittest.main()
