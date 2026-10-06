"""
Excel export module for Supplier Invoice Auditor.

Converts batch invoice processing results into a structured, professionally formatted
Excel (.xlsx) workbook containing:
1. Summary: One row per invoice with status, key metadata, financial totals, and error details.
2. Lines: Flattened invoice line items with product descriptions, quantities, prices, and taxes.
3. Audit Details: Detailed audit check and diagnostic calculation breakdowns per invoice.
"""

from datetime import datetime, date
from decimal import Decimal
from pathlib import Path
import sys
from typing import List, Dict, Any, Union, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from batch import collect_files, process_batch, batch_summary


# ------------------------------------------------------------
# STYLING DEFINITIONS
# ------------------------------------------------------------

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

STATUS_PASS_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
STATUS_PASS_FONT = Font(name="Calibri", size=10, bold=True, color="276A3C")

STATUS_REVIEW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
STATUS_REVIEW_FONT = Font(name="Calibri", size=10, bold=True, color="B25E00")

STATUS_ERROR_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
STATUS_ERROR_FONT = Font(name="Calibri", size=10, bold=True, color="C00000")

STATUS_NA_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
STATUS_NA_FONT = Font(name="Calibri", size=10, color="595959")

THIN_BORDER_SIDE = Side(style="thin", color="D9D9D9")
CELL_BORDER = Border(
    left=THIN_BORDER_SIDE,
    right=THIN_BORDER_SIDE,
    top=THIN_BORDER_SIDE,
    bottom=THIN_BORDER_SIDE,
)

CURRENCY_FORMAT = "#,##0.00"
PERCENT_FORMAT = "0.00"
DATE_FORMAT = "yyyy-mm-dd"
QTY_FORMAT = "#,##0.00"


# ------------------------------------------------------------
# VALUE FORMATTING HELPERS
# ------------------------------------------------------------

def _to_float(val: Any) -> Optional[float]:
    """Converts Decimal, int, or float string to float for openpyxl numeric cells. Returns None if val is None."""
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _to_date(val: Any) -> Any:
    """Parses YYYY-MM-DD date strings into datetime.date objects for genuine Excel date formatting."""
    if not val:
        return None
    if isinstance(val, (date, datetime)):
        return val if isinstance(val, date) else val.date()
    try:
        return datetime.strptime(str(val).strip(), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return str(val)


# ------------------------------------------------------------
# TURKISH DISPLAY MAPPINGS (EXCEL PRESENTATION LAYER)
# ------------------------------------------------------------

STATUS_DISPLAY_MAP = {
    "PASS": "BAŞARILI",
    "REVIEW": "İNCELEME GEREKLİ",
    "ERROR": "HATA",
    "FAIL": "BAŞARISIZ",
    "WARNING": "UYARI",
    "NOT AVAILABLE": "MEVCUT DEĞİL",
    "N/A": "MEVCUT DEĞİL",
}

CATEGORY_DISPLAY_MAP = {
    "Core Check": "Temel Denetim",
    "Diagnostic": "Teşhis",
    "File Error": "Dosya Hatası",
}

CHECK_NAME_DISPLAY_MAP = {
    "Tax total calculation": "Vergi Dahil Toplam Doğrulaması",
    "Payable total": "Ödenecek Tutar Mutabakatı",
    "Line VAT total": "Kalem KDV Toplamı Uyumu",
    "Line count": "Kalem Sayısı Tutarlılığı",
    "Invoice line total": "Kalem Satır Toplamı Uyumu",
    "Parse / Load": "Ayrıştırma / Yükleme",
}

CHECK_MESSAGE_DISPLAY_MAP = {
    # Tax total calculation
    "Required tax totals are not fully available.": "Gerekli vergi toplamları faturada eksiksiz mevcut değil.",
    "Tax exclusive plus tax total matches tax inclusive total.": "Vergi hariç tutar ile vergi toplamı, vergi dahil tutarla tam uyuşuyor.",
    "Tax exclusive plus tax total does not match tax inclusive total.": "Vergi hariç tutar ile vergi toplamı, vergi dahil tutarla uyuşmuyor.",
    # Payable total
    "Tax inclusive or payable total is not available.": "Vergi dahil toplam veya ödenecek tutar faturada mevcut değil.",
    "Tax inclusive total matches payable amount.": "Vergi dahil toplam tutar, ödenecek tutarla tam uyuşuyor.",
    "Tax inclusive total does not match payable amount.": "Vergi dahil toplam tutar, ödenecek tutarla uyuşmuyor.",
    # Line VAT total
    "Line-level VAT amounts are not fully available.": "Kalem bazlı KDV tutarları faturada eksiksiz mevcut değil.",
    "Line VAT total matches declared tax total.": "Kalem KDV tutarları toplamı, beyan edilen vergi toplamıyla tam uyuşuyor.",
    "Line VAT total does not match declared tax total.": "Kalem KDV tutarları toplamı, beyan edilen vergi toplamıyla uyuşmuyor.",
    # Line count
    "Expected line count is not available.": "Beyan edilen beklenen kalem sayısı faturada mevcut değil.",
    "Parsed line count matches declared line count.": "Ayrıştırılan kalem sayısı, beyan edilen kalem sayısıyla tam uyuşuyor.",
    "Parsed line count differs from declared line count.": "Ayrıştırılan kalem sayısı, beyan edilen kalem sayısından farklı.",
    # Invoice line total
    "Line totals are not fully available.": "Kalem satır tutarları faturada eksiksiz mevcut değil.",
    "Calculated line total matches declared line extension.": "Hesaplanan kalem satırları toplamı, beyan edilen mal/hizmet toplamıyla tam uyuşuyor.",
    "Calculated line total differs from declared line extension.": "Hesaplanan kalem satırları toplamı, beyan edilen mal/hizmet toplamından farklı.",
}


def _apply_status_style(cell, status_text: str):
    """Applies contextual color fills to status cells (supports both Turkish display and internal status codes)."""
    status_upper = str(status_text).upper().strip()
    if status_upper in ("PASS", "BAŞARILI"):
        cell.fill = STATUS_PASS_FILL
        cell.font = STATUS_PASS_FONT
    elif status_upper in ("REVIEW", "FAIL", "WARNING", "İNCELEME GEREKLİ", "BAŞARISIZ", "UYARI"):
        cell.fill = STATUS_REVIEW_FILL
        cell.font = STATUS_REVIEW_FONT
    elif status_upper in ("ERROR", "HATA"):
        cell.fill = STATUS_ERROR_FILL
        cell.font = STATUS_ERROR_FONT
    elif status_upper in ("NOT AVAILABLE", "N/A", "MEVCUT DEĞİL"):
        cell.fill = STATUS_NA_FILL
        cell.font = STATUS_NA_FONT


def _auto_fit_columns(ws, min_width: int = 12, max_width: int = 60):
    """Calculates column widths based on cell contents and headers."""
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = cell.value
            if val is not None:
                if isinstance(val, (datetime, date)):
                    cell_len = 10
                elif isinstance(val, float):
                    cell_len = len(f"{val:,.2f}")
                else:
                    cell_len = len(str(val))
                if cell_len > max_len:
                    max_len = cell_len
        header_len = len(str(col[0].value or ""))
        chosen_width = max(max_len + 3, header_len + 3, min_width)
        ws.column_dimensions[col_letter].width = min(chosen_width, max_width)


# ------------------------------------------------------------
# EXCEL EXPORT CORE
# ------------------------------------------------------------

def export_to_excel(
    batch_results: List[Dict[str, Any]],
    output_path: Union[str, Path] = "invoice_audit_report.xlsx",
) -> Path:
    """
    Exports a list of batch invoice results into a single multi-sheet Excel (.xlsx) workbook.

    Sheets:
    1. Özet - High-level invoice status, totals, and audit outcomes.
    2. Fatura Kalemleri - Flattened invoice line items with quantities, pricing, and taxes.
    3. Denetim Detayları - Detailed results for every audit check and diagnostic.

    Returns the Path to the saved workbook.
    """
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()

    # --------------------------------------------------------
    # 1. SUMMARY SHEET (ÖZET)
    # --------------------------------------------------------
    ws_summary = wb.active
    ws_summary.title = "Özet"

    summary_headers = [
        "Durum",
        "Fatura No",
        "Fatura Tarihi",
        "Tedarikçi",
        "Para Birimi",
        "Ödenecek Tutar",
        "Vergi Hariç Toplam",
        "Hesaplanan KDV",
        "Kalem Sayısı",
        "Uyumsuz Denetimler",
        "Dosya Adı",
        "Kaynak Arşiv (ZIP)",
        "Hata Detayı",
    ]

    ws_summary.append(summary_headers)

    for item in batch_results:
        inv = item.get("invoice")
        totals = inv.get("totals", {}) if inv else {}
        doc = inv.get("document", {}) if inv else {}
        supp = inv.get("supplier", {}) if inv else {}
        valid = inv.get("validation", {}) if inv else {}

        raw_status = item.get("status", "ERROR")
        display_status = STATUS_DISPLAY_MAP.get(raw_status, raw_status)

        inv_num = item.get("invoice_number") or doc.get("invoice_number")
        raw_date = item.get("issue_date") or doc.get("issue_date")
        issue_dt = _to_date(raw_date)
        supplier = item.get("supplier_name") or supp.get("name")
        currency = item.get("currency") or doc.get("currency")
        payable = _to_float(item.get("payable_amount") or totals.get("payable"))
        tax_exclusive = _to_float(totals.get("tax_exclusive"))
        tax_total = _to_float(totals.get("tax_total"))
        line_count = valid.get("parsed_line_count")

        failed_checks_raw = item.get("failed_checks", [])
        failed_checks_str = (
            ", ".join(CHECK_NAME_DISPLAY_MAP.get(c, c) for c in failed_checks_raw)
            if failed_checks_raw
            else None
        )

        filename = item.get("original_name") or item.get("file_name")
        container = item.get("container")
        container_display = Path(container).name if container else None
        error_details = item.get("error")

        row = [
            display_status,
            inv_num,
            issue_dt,
            supplier,
            currency,
            payable,
            tax_exclusive,
            tax_total,
            line_count,
            failed_checks_str,
            filename,
            container_display,
            error_details,
        ]
        ws_summary.append(row)

        current_row = ws_summary.max_row

        # Format Status cell
        cell_status = ws_summary.cell(row=current_row, column=1)
        _apply_status_style(cell_status, display_status)
        cell_status.alignment = Alignment(horizontal="center", vertical="center")

        # Format Date cell
        cell_date = ws_summary.cell(row=current_row, column=3)
        if isinstance(cell_date.value, date):
            cell_date.number_format = DATE_FORMAT
            cell_date.alignment = Alignment(horizontal="center", vertical="center")

        # Format Currency code
        cell_curr = ws_summary.cell(row=current_row, column=5)
        cell_curr.alignment = Alignment(horizontal="center", vertical="center")

        # Format Numeric Currency cells
        for col_idx in (6, 7, 8):
            cell_num = ws_summary.cell(row=current_row, column=col_idx)
            if cell_num.value is not None:
                cell_num.number_format = CURRENCY_FORMAT

        # Format Line Count cell
        cell_lc = ws_summary.cell(row=current_row, column=9)
        cell_lc.alignment = Alignment(horizontal="right", vertical="center")

        # Add borders to all row cells
        for c in range(1, len(summary_headers) + 1):
            ws_summary.cell(row=current_row, column=c).border = CELL_BORDER

    # Style Header Row
    ws_summary.row_dimensions[1].height = 26
    for col_idx in range(1, len(summary_headers) + 1):
        cell = ws_summary.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = CELL_BORDER

    ws_summary.freeze_panes = "A2"
    ws_summary.auto_filter.ref = ws_summary.dimensions
    _auto_fit_columns(ws_summary)

    # --------------------------------------------------------
    # 2. LINES SHEET (FATURA KALEMLERİ)
    # --------------------------------------------------------
    ws_lines = wb.create_sheet(title="Fatura Kalemleri")

    line_headers = [
        "Fatura No",
        "Dosya Adı",
        "Kalem No",
        "Ürün / Hizmet Adı",
        "Miktar",
        "Birim",
        "Birim Fiyat",
        "Satır Tutarı",
        "İskonto Tutarı",
        "Fiyat Artırımı",
        "KDV Oranı (%)",
        "KDV Tutarı",
        "Para Birimi",
    ]

    ws_lines.append(line_headers)

    for item in batch_results:
        inv = item.get("invoice")
        if not inv:
            continue

        inv_num = item.get("invoice_number") or inv.get("document", {}).get("invoice_number")
        filename = item.get("original_name") or item.get("file_name")
        currency = item.get("currency") or inv.get("document", {}).get("currency")

        for line in inv.get("lines", []):
            product = line.get("product", {})
            prod_name = product.get("name")
            line_id = line.get("line_id")
            qty = _to_float(line.get("quantity"))
            unit = line.get("unit")
            unit_price = _to_float(line.get("price_amount"))
            line_amount = _to_float(line.get("line_total"))
            allowance = _to_float(line.get("allowance"))
            charge = _to_float(line.get("charge"))
            vat_rate = _to_float(line.get("vat_rate"))
            vat_amount = _to_float(line.get("vat_amount"))

            row = [
                inv_num,
                filename,
                line_id,
                prod_name,
                qty,
                unit,
                unit_price,
                line_amount,
                allowance,
                charge,
                vat_rate,
                vat_amount,
                currency,
            ]
            ws_lines.append(row)

            current_row = ws_lines.max_row

            # Alignments & Number formats
            ws_lines.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")
            ws_lines.cell(row=current_row, column=6).alignment = Alignment(horizontal="center", vertical="center")
            ws_lines.cell(row=current_row, column=13).alignment = Alignment(horizontal="center", vertical="center")

            # Qty
            cell_qty = ws_lines.cell(row=current_row, column=5)
            if cell_qty.value is not None:
                cell_qty.number_format = QTY_FORMAT

            # Price, Line Amount, Allowance, Charge, VAT Amount
            for col_idx in (7, 8, 9, 10, 12):
                cell_num = ws_lines.cell(row=current_row, column=col_idx)
                if cell_num.value is not None:
                    cell_num.number_format = CURRENCY_FORMAT

            # VAT Rate
            cell_vat_rate = ws_lines.cell(row=current_row, column=11)
            if cell_vat_rate.value is not None:
                cell_vat_rate.number_format = PERCENT_FORMAT

            for c in range(1, len(line_headers) + 1):
                ws_lines.cell(row=current_row, column=c).border = CELL_BORDER

    # Style Header Row
    ws_lines.row_dimensions[1].height = 26
    for col_idx in range(1, len(line_headers) + 1):
        cell = ws_lines.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = CELL_BORDER

    ws_lines.freeze_panes = "A2"
    ws_lines.auto_filter.ref = ws_lines.dimensions
    _auto_fit_columns(ws_lines)

    # --------------------------------------------------------
    # 3. AUDIT DETAILS SHEET (DENETİM DETAYLARI)
    # --------------------------------------------------------
    ws_audit = wb.create_sheet(title="Denetim Detayları")

    audit_headers = [
        "Fatura No",
        "Dosya Adı",
        "Kategori",
        "Denetim Adı",
        "Durum",
        "Hesaplanan",
        "Beyan Edilen",
        "Fark",
        "Açıklama / Mesaj",
    ]

    ws_audit.append(audit_headers)

    for item in batch_results:
        inv_num = item.get("invoice_number")
        filename = item.get("original_name") or item.get("file_name")
        audit = item.get("audit")

        # If file failed completely to parse
        if not item.get("success"):
            row = [
                inv_num,
                filename,
                CATEGORY_DISPLAY_MAP.get("File Error", "File Error"),
                CHECK_NAME_DISPLAY_MAP.get("Parse / Load", "Parse / Load"),
                STATUS_DISPLAY_MAP.get("ERROR", "HATA"),
                None,
                None,
                None,
                item.get("error"),
            ]
            ws_audit.append(row)
            current_row = ws_audit.max_row
            _apply_status_style(ws_audit.cell(row=current_row, column=5), STATUS_DISPLAY_MAP.get("ERROR", "HATA"))
            ws_audit.cell(row=current_row, column=5).alignment = Alignment(horizontal="center", vertical="center")
            ws_audit.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")
            for c in range(1, len(audit_headers) + 1):
                ws_audit.cell(row=current_row, column=c).border = CELL_BORDER
            continue

        if not audit:
            continue

        # Core Checks
        for chk in audit.get("checks", []):
            name = chk.get("name")
            display_name = CHECK_NAME_DISPLAY_MAP.get(name, name)
            status_bool = chk.get("status")
            if status_bool is True:
                raw_status_str = "PASS"
            elif status_bool is False:
                raw_status_str = "FAIL"
            else:
                raw_status_str = "NOT AVAILABLE"
            display_status = STATUS_DISPLAY_MAP.get(raw_status_str, raw_status_str)

            calc = _to_float(chk.get("calculated"))
            decl = _to_float(chk.get("declared"))
            diff = _to_float(chk.get("difference"))
            msg = chk.get("message")
            display_msg = CHECK_MESSAGE_DISPLAY_MAP.get(msg, msg)

            row = [
                inv_num,
                filename,
                CATEGORY_DISPLAY_MAP.get("Core Check", "Core Check"),
                display_name,
                display_status,
                calc,
                decl,
                diff,
                display_msg,
            ]
            ws_audit.append(row)
            current_row = ws_audit.max_row

            cell_status = ws_audit.cell(row=current_row, column=5)
            _apply_status_style(cell_status, display_status)
            cell_status.alignment = Alignment(horizontal="center", vertical="center")
            ws_audit.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")

            for col_idx in (6, 7, 8):
                cell_num = ws_audit.cell(row=current_row, column=col_idx)
                if cell_num.value is not None:
                    cell_num.number_format = CURRENCY_FORMAT

            for c in range(1, len(audit_headers) + 1):
                ws_audit.cell(row=current_row, column=c).border = CELL_BORDER

        # Diagnostics
        for diag in audit.get("diagnostics", []):
            name = diag.get("name")
            display_name = CHECK_NAME_DISPLAY_MAP.get(name, name)
            status_bool = diag.get("status")
            if status_bool is True:
                raw_status_str = "PASS"
            elif status_bool is False:
                raw_status_str = "WARNING"
            else:
                raw_status_str = "NOT AVAILABLE"
            display_status = STATUS_DISPLAY_MAP.get(raw_status_str, raw_status_str)

            calc = _to_float(diag.get("calculated"))
            decl = _to_float(diag.get("declared"))
            diff = _to_float(diag.get("difference"))
            msg = diag.get("message")
            display_msg = CHECK_MESSAGE_DISPLAY_MAP.get(msg, msg)

            row = [
                inv_num,
                filename,
                CATEGORY_DISPLAY_MAP.get("Diagnostic", "Diagnostic"),
                display_name,
                display_status,
                calc,
                decl,
                diff,
                display_msg,
            ]
            ws_audit.append(row)
            current_row = ws_audit.max_row

            cell_status = ws_audit.cell(row=current_row, column=5)
            _apply_status_style(cell_status, display_status)
            cell_status.alignment = Alignment(horizontal="center", vertical="center")
            ws_audit.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")

            # If line count diagnostic, keep as integers without decimal places if whole number
            is_line_count = "count" in str(name).lower() or "kalem sayısı" in str(display_name).lower()
            num_fmt = "#,##0" if is_line_count else CURRENCY_FORMAT

            for col_idx in (6, 7, 8):
                cell_num = ws_audit.cell(row=current_row, column=col_idx)
                if cell_num.value is not None:
                    cell_num.number_format = num_fmt

            for c in range(1, len(audit_headers) + 1):
                ws_audit.cell(row=current_row, column=c).border = CELL_BORDER

    # Style Header Row
    ws_audit.row_dimensions[1].height = 26
    for col_idx in range(1, len(audit_headers) + 1):
        cell = ws_audit.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = CELL_BORDER

    ws_audit.freeze_panes = "A2"
    ws_audit.auto_filter.ref = ws_audit.dimensions
    _auto_fit_columns(ws_audit)

    # Save workbook
    wb.save(output_path)
    return output_path


# ------------------------------------------------------------
# CLI ENTRY POINT
# ------------------------------------------------------------

def main():
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    if len(sys.argv) < 2:
        print("Usage: python excel_export.py <source_path> [output_report.xlsx]")
        sys.exit(1)

    source = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "invoice_audit_report.xlsx"

    try:
        files = collect_files(source)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    print("=" * 70)
    print("SUPPLIER INVOICE AUDITOR - EXCEL EXPORTER")
    print("=" * 70)
    print(f"Source: {source}")
    print(f"Collected: {len(files)} XML file(s)")

    if not files:
        print("No XML files found to process.")
        sys.exit(0)

    results = process_batch(files)
    summary = batch_summary(results)

    print(f"Processed: {summary['total']} invoices ({summary['pass']} PASS, {summary['review']} REVIEW, {summary['error']} ERROR)")
    print(f"Exporting to Excel: {output_file}...")

    saved_path = export_to_excel(results, output_file)
    print(f"Report successfully saved to: {saved_path}")
    print("=" * 70)


if __name__ == "__main__":
    main()
