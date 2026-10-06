# Supplier Invoice Auditor

A high-performance Python utility for bulk parsing, auditing, and reporting Turkish e-Invoices (e-Fatura / e-Arşiv in **UBL 2.1** XML format).

The tool accepts a single XML invoice, a directory of invoices, or a ZIP archive, audits all mathematical and tax consistency rules in memory, and generates a structured, multi-sheet Excel report (`.xlsx`).

---

## Key Features

* **Flexible Batch Inputs**: Process a single XML file, a folder containing XML invoices (recursive by default), or a `.zip` archive containing invoice files.
* **In-Memory Speed & Resilience**: Processes dozens of invoices per second. If an invalid or corrupted XML file is encountered, the error is recorded in the report while processing of remaining files continues uninterrupted.
* **Automated Audit Checks**:
  * **Tax Total Calculation**: Verifies that declared Tax Exclusive Total + Tax Total equals Tax Inclusive Total (tolerance: $\pm 0.02$).
  * **Payable Total**: Verifies that declared Tax Inclusive Total equals Payable Amount ($\pm 0.02$).
  * **Line VAT Total**: Sums line-item VAT amounts and verifies consistency with declared invoice Tax Total ($\pm 0.02$).
  * **Diagnostics**: Cross-checks declared line count vs. parsed line count and verifies sum of line amounts against declared Line Extension Total.
* **Structured Excel Output**: Generates a clean 3-sheet `.xlsx` report featuring frozen headers, active autofilters, and Excel numeric/date formatting.

---

## Excel Report Structure

The generated workbook contains three interconnected sheets:

### 1. `Summary` (One row per invoice)
* **Status**: Visual badge (`PASS` in green, `REVIEW` in yellow, `ERROR` in red).
* **Header Metadata**: Invoice Number, Issue Date, Supplier Name, Currency.
* **Financial Totals**: Payable Amount, Tax Exclusive Total, Tax Total, Parsed Line Count.
* **Audit Outcome**: Comma-separated list of failed check names (empty if `PASS`).
* **Source Tracking**: Original filename, parent ZIP archive container name (if applicable), and parse error messages (if `ERROR`).

### 2. `Lines` (One row per line item)
* Flattened line items across all invoices: Line ID, Item/Product Name, Invoiced Quantity, Unit Code, Unit Price, Line Extension Amount, Allowance/Discount, Surcharge, VAT Rate (%), VAT Amount, and Currency.
* Empty allowances/charges remain blank cells rather than misleading zeros.

### 3. `Audit Details` (One row per check per invoice)
* Trace of every core check and diagnostic: Category (`Core Check` / `Diagnostic`), Check Name, Status (`PASS` / `FAIL` / `WARNING` / `NOT AVAILABLE`), Calculated Value, Declared Value, Difference, and full explanation message.

---

## Requirements & Installation

### Requirements
* **Python 3.10+** (Python 3.13 tested)
* **openpyxl** (for Excel workbook generation)

### Installation
Clone or download this repository, then install the single runtime dependency:
```bash
pip install -r requirements.txt
```

*(Note: All core XML parsing, audit calculations, and ZIP processing use the Python Standard Library.)*

---

## CLI Usage

The primary entry point is `main.py`:

```bash
python main.py <source> [output.xlsx] [options]
```

### Arguments & Options
* `source` *(required)*: Path to an invoice XML file, folder, or ZIP archive.
* `output` *(optional)*: Destination Excel file path. Defaults to `invoice_audit_report.xlsx`.
* `-o, --output-file`: Explicit output file path flag.
* `--no-recursive`: Disable recursive folder traversal when source is a directory.
* `-v, --verbose`: Print detailed check marks (`✓`, `✗`) for each invoice to console.
* `-h, --help`: Display help and usage examples.

---

## Examples

### 1. Audit a Single Invoice
```bash
python main.py test_data/xml/test-metro.xml
```
*Creates `invoice_audit_report.xlsx` with results for the single invoice.*

### 2. Audit a Folder of Invoices
```bash
python main.py test_data/xml/ monthly_audit_2026.xlsx
```
*Recursively discovers all `.xml` files in `test_data/xml/` and exports to `monthly_audit_2026.xlsx`.*

### 3. Audit a ZIP Archive Directly
```bash
python main.py Q1_Invoices.zip -o Q1_report.xlsx
```
*Extracts XML files to a temporary workspace in memory, runs the audit, tracks the ZIP archive name in the Summary sheet, and automatically cleans up temporary files upon completion.*

---

## Exit Codes

* `0`: Processing succeeded and Excel report was generated (invoices requiring `REVIEW` are standard audit findings and exit with code `0`).
* `1`: Input or runtime error (source path not found, no XML files found in source, or file write permission error).

---

## Current Scope & Limitations

* **Supported Standard**: Designed for UBL 2.1 Turkish e-Invoice schemas (`cac:` and `cbc:` components).
* **Buyer Details**: Focuses on supplier invoice auditing; buyer identification (`cac:AccountingCustomerParty`) is not currently exported to the Summary sheet.
* **Tevkifat (Withholding)**: Withholding tax sub-types are currently included in standard VAT checks; specialized withholding factor validation will be added in future releases.
* **Standalone Executable**: Currently runs via Python. A standalone Windows `.exe` build configuration using PyInstaller is planned for distribution.

---

## Testing

To run the complete test suite:
```bash
# Run baseline regression suite across all 28 XML test files
python -X utf8 tests/run_tests.py

# Run batch processing tests
python -X utf8 tests/test_batch.py

# Run Excel export tests
python -X utf8 tests/test_excel_export.py

# Run CLI entry point tests
python -X utf8 tests/test_main.py
```
Expected baseline outcome: `26 PASS, 2 REVIEW, 0 ERROR`.
