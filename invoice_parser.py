import xml.etree.ElementTree as ET
import sys
from decimal import Decimal, InvalidOperation

from audit_engine import audit_invoice


# ============================================================
# INVOICE PARSER v0.4 (Refactored)
# UBL XML -> NORMALIZED INVOICE
# ============================================================

NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
}


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def text_or_none(element):
    if element is None or element.text is None:
        return None

    text = element.text.strip()
    return text if text else None


def decimal_or_none(value):
    if value is None:
        return None

    try:
        return Decimal(value)
    except (InvalidOperation, ValueError):
        return None


def money(value):
    if value is None:
        return None

    return Decimal(value).quantize(Decimal("0.01"))


def find_text(parent, path):
    element = parent.find(path, NS)
    return text_or_none(element)


def almost_equal(a, b, tolerance=Decimal("0.02")):
    if a is None or b is None:
        return False

    return abs(a - b) <= tolerance


# ------------------------------------------------------------
# PARSER
# ------------------------------------------------------------

def parse_invoice(xml_path):
    """
    Parses a UBL 2.1 e-invoice XML file and returns a normalized dictionary.
    Raises FileNotFoundError if file is missing, or ET.ParseError if XML is invalid.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Document type
    if root.tag.endswith("Invoice"):
        document_type = "UBL Invoice"
    else:
        document_type = "Unknown XML structure"

    # Header
    invoice_number = find_text(root, "cbc:ID")
    uuid = find_text(root, "cbc:UUID")
    issue_date = find_text(root, "cbc:IssueDate")
    issue_time = find_text(root, "cbc:IssueTime")
    invoice_type = find_text(root, "cbc:InvoiceTypeCode")
    currency = find_text(root, "cbc:DocumentCurrencyCode")
    expected_line_count = decimal_or_none(find_text(root, "cbc:LineCountNumeric"))

    # Supplier
    supplier_name = find_text(
        root,
        "cac:AccountingSupplierParty/"
        "cac:Party/"
        "cac:PartyName/"
        "cbc:Name"
    )

    if not supplier_name:
        first_name = find_text(
            root,
            "cac:AccountingSupplierParty/"
            "cac:Party/"
            "cac:Person/"
            "cbc:FirstName"
        )
        family_name = find_text(
            root,
            "cac:AccountingSupplierParty/"
            "cac:Party/"
            "cac:Person/"
            "cbc:FamilyName"
        )
        person_parts = [part for part in (first_name, family_name) if part]
        if person_parts:
            supplier_name = " ".join(person_parts)

    # References
    despatch_number = find_text(root, "cac:DespatchDocumentReference/cbc:ID")
    despatch_date = find_text(root, "cac:DespatchDocumentReference/cbc:IssueDate")
    order_number = find_text(root, "cac:OrderReference/cbc:ID")
    order_date = find_text(root, "cac:OrderReference/cbc:IssueDate")

    # Lines
    lines = []
    invoice_lines = root.findall("cac:InvoiceLine", NS)

    for line in invoice_lines:
        line_id = find_text(line, "cbc:ID")

        quantity_element = line.find("cbc:InvoicedQuantity", NS)
        quantity = None
        unit_code = None

        if quantity_element is not None:
            quantity = decimal_or_none(quantity_element.text)
            unit_code = quantity_element.get("unitCode")

        # Product
        item_name = find_text(line, "cac:Item/cbc:Name")
        seller_item_id = find_text(
            line,
            "cac:Item/cac:SellersItemIdentification/cbc:ID"
        )
        manufacturer_item_id = find_text(
            line,
            "cac:Item/cac:ManufacturersItemIdentification/cbc:ID"
        )
        standard_item_id = find_text(
            line,
            "cac:Item/cac:StandardItemIdentification/cbc:ID"
        )

        # Price
        price_amount = decimal_or_none(
            find_text(line, "cac:Price/cbc:PriceAmount")
        )

        # Line total
        line_extension = decimal_or_none(
            find_text(line, "cbc:LineExtensionAmount")
        )

        # Line discount / charge (aggregating multiple AllowanceCharge elements if present)
        line_allowance = None
        line_charge = None

        for line_allowance_charge in line.findall("cac:AllowanceCharge", NS):
            charge_indicator = find_text(
                line_allowance_charge,
                "cbc:ChargeIndicator"
            )
            amount = decimal_or_none(
                find_text(
                    line_allowance_charge,
                    "cbc:Amount"
                )
            )

            if amount is not None:
                if charge_indicator == "false":
                    line_allowance = (line_allowance or Decimal("0")) + amount
                elif charge_indicator == "true":
                    line_charge = (line_charge or Decimal("0")) + amount

        # VAT
        vat_rate = None
        vat_amount = None

        tax_total_elem = line.find("cac:TaxTotal", NS)
        if tax_total_elem is not None:
            vat_amount = decimal_or_none(
                find_text(tax_total_elem, "cbc:TaxAmount")
            )
            tax_subtotal = tax_total_elem.find("cac:TaxSubtotal", NS)
            if tax_subtotal is not None:
                vat_rate = decimal_or_none(
                    find_text(tax_subtotal, "cbc:Percent")
                )

        normalized_line = {
            "line_id": line_id,
            "product": {
                "name": item_name,
                "seller_item_id": seller_item_id,
                "manufacturer_item_id": manufacturer_item_id,
                "standard_item_id": standard_item_id,
            },
            "quantity": quantity,
            "unit": unit_code,
            "price_amount": price_amount,
            "line_total": line_extension,
            "allowance": line_allowance,
            "charge": line_charge,
            "vat_rate": vat_rate,
            "vat_amount": vat_amount,
        }

        lines.append(normalized_line)

    # Totals
    tax_total = decimal_or_none(
        find_text(root, "cac:TaxTotal/cbc:TaxAmount")
    )
    line_extension_total = decimal_or_none(
        find_text(root, "cac:LegalMonetaryTotal/cbc:LineExtensionAmount")
    )
    tax_exclusive_total = decimal_or_none(
        find_text(root, "cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount")
    )
    tax_inclusive_total = decimal_or_none(
        find_text(root, "cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount")
    )
    allowance_total = decimal_or_none(
        find_text(root, "cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount")
    )
    charge_total = decimal_or_none(
        find_text(root, "cac:LegalMonetaryTotal/cbc:ChargeTotalAmount")
    )
    payable_total = decimal_or_none(
        find_text(root, "cac:LegalMonetaryTotal/cbc:PayableAmount")
    )

    # LegalMonetaryTotal is authoritative for invoice-level allowance / charge classification.
    invoice_allowance = allowance_total
    invoice_charge = charge_total

    normalized_invoice = {
        "document": {
            "type": document_type,
            "invoice_number": invoice_number,
            "uuid": uuid,
            "issue_date": issue_date,
            "issue_time": issue_time,
            "invoice_type": invoice_type,
            "currency": currency,
        },
        "supplier": {
            "name": supplier_name,
        },
        "references": {
            "despatch_number": despatch_number,
            "despatch_date": despatch_date,
            "order_number": order_number,
            "order_date": order_date,
        },
        "invoice_adjustments": {
            "allowance": invoice_allowance,
            "charge": invoice_charge,
        },
        "lines": lines,
        "totals": {
            "line_extension": line_extension_total,
            "tax_exclusive": tax_exclusive_total,
            "tax_total": tax_total,
            "tax_inclusive": tax_inclusive_total,
            "allowance_total": allowance_total,
            "charge_total": charge_total,
            "payable": payable_total,
        },
        "validation": {
            "expected_line_count": (
                int(expected_line_count)
                if expected_line_count is not None
                else None
            ),
            "parsed_line_count": len(lines),
        },
    }

    return normalized_invoice


# ------------------------------------------------------------
# PRESENTATION / OUTPUT
# ------------------------------------------------------------

def print_invoice_report(invoice, audit_result=None, xml_file=None):
    """
    Prints human-readable invoice details and audit results to standard output.
    """
    if audit_result is None:
        audit_result = audit_invoice(invoice)

    print("=" * 60)
    print("INVOICE PARSER v0.4")
    print("=" * 60)
    print()
    if xml_file:
        print(f"Reading: {xml_file}")
        print()

    print("AUDIT ENGINE")
    print("-" * 60)

    for check in audit_result["checks"]:
        name = check["name"]
        result = check["status"]

        if result is True:
            print(f"✓ {name}")
        elif result is False:
            print(f"✗ {name}")
            print(f"  Calculated : {check['calculated']}")
            print(f"  Declared   : {check['declared']}")
            print(f"  Difference : {check['difference']}")
            print(f"  Reason     : {check['message']}")
        else:
            print(f"- {name}: NOT AVAILABLE")
            print(f"  Reason     : {check['message']}")

    print()
    print("DIAGNOSTICS")
    print("-" * 60)

    for diagnostic in audit_result["diagnostics"]:
        name = diagnostic["name"]
        result = diagnostic["status"]

        if result is True:
            print(f"✓ {name}")
        elif result is False:
            print(f"⚠ {name}")
            print(f"  Calculated : {diagnostic['calculated']}")
            print(f"  Declared   : {diagnostic['declared']}")
            print(f"  Difference : {diagnostic['difference']}")
            print(f"  Reason     : {diagnostic['message']}")
        else:
            print(f"- {name}: NOT AVAILABLE")
            print(f"  Reason     : {diagnostic['message']}")

    print()
    if audit_result["overall_ok"]:
        print("AUDIT RESULT: PASS")
    else:
        print("AUDIT RESULT: REVIEW")

    doc = invoice["document"]
    print()
    print("DOCUMENT")
    print("-" * 60)
    print(f"Type           : {doc['type']}")
    print(f"Invoice number : {doc['invoice_number']}")
    print(f"UUID           : {doc['uuid']}")
    print(f"Date           : {doc['issue_date']}")
    print(f"Time           : {doc['issue_time']}")
    print(f"Type code      : {doc['invoice_type']}")
    print(f"Supplier       : {invoice['supplier']['name']}")
    print(f"Currency       : {doc['currency']}")

    refs = invoice["references"]
    print()
    print("REFERENCES")
    print("-" * 60)
    print(f"Order No.      : {refs['order_number']}")
    print(f"Order Date     : {refs['order_date']}")
    print(f"Despatch No.   : {refs['despatch_number']}")
    print(f"Despatch Date  : {refs['despatch_date']}")

    adj = invoice["invoice_adjustments"]
    print()
    print("INVOICE ADJUSTMENTS")
    print("-" * 60)
    print(f"Invoice allowance : {adj['allowance']}")
    print(f"Invoice charge    : {adj['charge']}")

    print()
    print("LINES")
    print("-" * 60)
    print(f"Expected lines : {invoice['validation']['expected_line_count']}")
    print(f"Parsed lines   : {invoice['validation']['parsed_line_count']}")

    for line in invoice["lines"]:
        product = line["product"]
        print()
        print(f"{line['line_id']} | {product['name']}")
        print(
            f"   Product IDs  : "
            f"seller={product['seller_item_id']} | "
            f"manufacturer={product['manufacturer_item_id']} | "
            f"standard={product['standard_item_id']}"
        )
        print(f"   Quantity     : {line['quantity']} {line['unit']}")
        print(f"   PriceAmount  : {line['price_amount']}")
        print(f"   Line total   : {line['line_total']}")
        print(f"   Allowance    : {line['allowance']}")
        print(f"   Charge       : {line['charge']}")
        print(f"   VAT rate     : {line['vat_rate']}%")
        print(f"   VAT amount   : {line['vat_amount']}")

    totals = invoice["totals"]
    print()
    print("TOTALS")
    print("-" * 60)
    print(f"Line extension : {totals['line_extension']}")
    print(f"Tax exclusive  : {totals['tax_exclusive']}")
    print(f"Tax total      : {totals['tax_total']}")
    print(f"Tax inclusive  : {totals['tax_inclusive']}")
    print(f"Allowance      : {totals['allowance_total']}")
    print(f"Charge         : {totals['charge_total']}")
    print(f"Payable        : {totals['payable']}")
    print()
    print()
    print("=" * 60)


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
        print("Usage: python invoice_parser.py <path_to_invoice_xml>")
        sys.exit(1)

    xml_file = sys.argv[1]

    try:
        invoice = parse_invoice(xml_file)
    except FileNotFoundError:
        print(f"ERROR: File not found: {xml_file}")
        sys.exit(1)
    except ET.ParseError as e:
        print("ERROR: Invalid XML file")
        print(e)
        sys.exit(1)

    audit_result = audit_invoice(invoice)
    print_invoice_report(invoice, audit_result, xml_file=xml_file)


if __name__ == "__main__":
    main()