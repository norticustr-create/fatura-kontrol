import pandas as pd
import sys
from pathlib import Path


# ============================================================
# AI SUPPLIER INVOICE AUDITOR
# v0.2 - Basic PO / Invoice Comparison Engine
# ============================================================


def read_document(filepath):
    """Read an Excel document."""
    df = pd.read_excel(filepath, header=None)
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def find_value(df, label):
    """Find a value next to a label such as Supplier or PO number."""
    for _, row in df.iterrows():
        values = row.tolist()

        for i, value in enumerate(values):
            if pd.isna(value):
                continue

            if str(value).strip().lower() == label.lower():
                if i + 1 < len(values) and not pd.isna(values[i + 1]):
                    return str(values[i + 1]).strip()

    return None


def find_header_row(df):
    """Find the row containing the product table headers."""

    for index, row in df.iterrows():
        values = [
            str(value).strip().lower()
            for value in row.tolist()
            if not pd.isna(value)
        ]

        if "product" in values and "quantity" in values and "unit price" in values:
            return index

    return None


def extract_lines(df):
    """Extract product lines from the document."""

    header_row = find_header_row(df)

    if header_row is None:
        return []

    headers = [
        str(value).strip().lower()
        if not pd.isna(value)
        else ""
        for value in df.iloc[header_row].tolist()
    ]

    column_map = {}

    for index, header in enumerate(headers):
        if header == "product":
            column_map["product"] = index
        elif header == "sku":
            column_map["sku"] = index
        elif header == "quantity":
            column_map["quantity"] = index
        elif header == "unit price":
            column_map["unit_price"] = index

    lines = []

    for row_index in range(header_row + 1, len(df)):
        row = df.iloc[row_index]

        product = row.iloc[column_map["product"]]

        if pd.isna(product):
            continue

        line = {
            "product": str(product).strip(),
            "sku": (
                str(row.iloc[column_map["sku"]]).strip()
                if "sku" in column_map and not pd.isna(row.iloc[column_map["sku"]])
                else None
            ),
            "quantity": float(row.iloc[column_map["quantity"]]),
            "unit_price": float(row.iloc[column_map["unit_price"]]),
        }

        lines.append(line)

    return lines


def compare_documents(po, invoice):
    """Compare PO lines with invoice lines."""

    discrepancies = []

    # --------------------------------------------------------
    # Supplier check
    # --------------------------------------------------------

    po_supplier = find_value(po, "Supplier")
    invoice_supplier = find_value(invoice, "Supplier")

    if po_supplier != invoice_supplier:
        discrepancies.append(
            f"Supplier mismatch: PO='{po_supplier}', "
            f"Invoice='{invoice_supplier}'"
        )

    # --------------------------------------------------------
    # PO number check
    # --------------------------------------------------------

    po_number = find_value(po, "Purchase Order")
    invoice_po_reference = find_value(invoice, "PO Reference")

    if po_number != invoice_po_reference:
        discrepancies.append(
            f"PO reference mismatch: PO='{po_number}', "
            f"Invoice='{invoice_po_reference}'"
        )

    # --------------------------------------------------------
    # Extract product lines
    # --------------------------------------------------------

    po_lines = extract_lines(po)
    invoice_lines = extract_lines(invoice)

    # Index PO lines by SKU
    po_by_sku = {
        line["sku"]: line
        for line in po_lines
        if line["sku"] is not None
    }

    invoice_by_sku = {
        line["sku"]: line
        for line in invoice_lines
        if line["sku"] is not None
    }

    # --------------------------------------------------------
    # Check every invoice line
    # --------------------------------------------------------

    for invoice_line in invoice_lines:

        sku = invoice_line["sku"]

        if sku not in po_by_sku:

            discrepancies.append(
                f"Product not found in PO: "
                f"{invoice_line['product']} ({sku})"
            )

            continue

        po_line = po_by_sku[sku]

        # Quantity check

        if invoice_line["quantity"] != po_line["quantity"]:

            discrepancies.append(
                f"Quantity mismatch for {invoice_line['product']} "
                f"({sku}): "
                f"PO={po_line['quantity']}, "
                f"Invoice={invoice_line['quantity']}"
            )

        # Unit price check

        if invoice_line["unit_price"] != po_line["unit_price"]:

            difference = (
                invoice_line["unit_price"]
                - po_line["unit_price"]
            )

            discrepancies.append(
                f"Unit price mismatch for {invoice_line['product']} "
                f"({sku}): "
                f"PO={po_line['unit_price']:.2f}, "
                f"Invoice={invoice_line['unit_price']:.2f}, "
                f"Difference={difference:+.2f}"
            )

    # --------------------------------------------------------
    # Check for PO items missing from invoice
    # --------------------------------------------------------

    for po_line in po_lines:

        sku = po_line["sku"]

        if sku not in invoice_by_sku:

            discrepancies.append(
                f"PO item missing from invoice: "
                f"{po_line['product']} ({sku})"
            )

    # --------------------------------------------------------
    # Final result
    # --------------------------------------------------------

    if discrepancies:
        return "REVIEW", discrepancies

    return "PASS", []


def print_report(po, invoice, result, discrepancies):
    """Print the audit report."""

    print()
    print("=" * 60)
    print("SUPPLIER INVOICE AUDIT REPORT")
    print("=" * 60)

    po_number = find_value(po, "Purchase Order")
    invoice_number = find_value(invoice, "Invoice")
    supplier = find_value(invoice, "Supplier")

    print(f"\nSupplier      : {supplier}")
    print(f"PO            : {po_number}")
    print(f"Invoice       : {invoice_number}")

    print("\n" + "-" * 60)

    if result == "PASS":

        print("OVERALL RESULT: PASS")
        print()
        print("No discrepancies found.")

    else:

        print("OVERALL RESULT: REVIEW")
        print()
        print(f"{len(discrepancies)} discrepancy(ies) found:")

        for number, discrepancy in enumerate(discrepancies, start=1):
            print(f"\n{number}. {discrepancy}")

    print("\n" + "=" * 60)


def main():

    if len(sys.argv) < 3:

        print("Usage:")
        print("python auditor.py PO_FILE INVOICE_FILE")
        return

    po_file = sys.argv[1]
    invoice_file = sys.argv[2]

    print("=" * 60)
    print("AI SUPPLIER INVOICE AUDITOR")
    print("v0.2 - Basic Comparison Engine")
    print("=" * 60)

    print("\nLoading documents...")

    po = read_document(po_file)
    invoice = read_document(invoice_file)

    print(f"PO loaded      : {Path(po_file).name}")
    print(f"Invoice loaded : {Path(invoice_file).name}")

    print("\nAnalyzing...")

    result, discrepancies = compare_documents(po, invoice)

    print_report(
        po,
        invoice,
        result,
        discrepancies
    )


if __name__ == "__main__":
    main()