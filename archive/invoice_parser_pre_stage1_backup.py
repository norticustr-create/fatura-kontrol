import xml.etree.ElementTree as ET
import sys
from decimal import Decimal, InvalidOperation


# ============================================================
# INVOICE PARSER v0.4
# UBL XML → NORMALIZED INVOICE
# ============================================================

XML_FILE = sys.argv[1] if len(sys.argv) > 1 else "metro_test.xml"

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
# LOAD XML
# ------------------------------------------------------------

print("=" * 60)
print("INVOICE PARSER v0.4")
print("=" * 60)
print()
print(f"Reading: {XML_FILE}")
print()

try:

    tree = ET.parse(XML_FILE)
    root = tree.getroot()

except FileNotFoundError:

    print(f"ERROR: File not found: {XML_FILE}")
    sys.exit(1)

except ET.ParseError as e:

    print("ERROR: Invalid XML file")
    print(e)
    sys.exit(1)


# ------------------------------------------------------------
# DOCUMENT TYPE
# ------------------------------------------------------------

if root.tag.endswith("Invoice"):

    document_type = "UBL Invoice"

else:

    document_type = "Unknown XML structure"


# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------

invoice_number = find_text(
    root,
    "cbc:ID"
)

uuid = find_text(
    root,
    "cbc:UUID"
)

issue_date = find_text(
    root,
    "cbc:IssueDate"
)

issue_time = find_text(
    root,
    "cbc:IssueTime"
)

invoice_type = find_text(
    root,
    "cbc:InvoiceTypeCode"
)

currency = find_text(
    root,
    "cbc:DocumentCurrencyCode"
)

expected_line_count = decimal_or_none(
    find_text(
        root,
        "cbc:LineCountNumeric"
    )
)


# ------------------------------------------------------------
# SUPPLIER
# ------------------------------------------------------------

supplier_name = find_text(
    root,
    "cac:AccountingSupplierParty/"
    "cac:Party/"
    "cac:PartyName/"
    "cbc:Name"
)


# ------------------------------------------------------------
# REFERENCES
# ------------------------------------------------------------

despatch_number = find_text(
    root,
    "cac:DespatchDocumentReference/cbc:ID"
)

despatch_date = find_text(
    root,
    "cac:DespatchDocumentReference/cbc:IssueDate"
)

order_number = find_text(
    root,
    "cac:OrderReference/cbc:ID"
)

order_date = find_text(
    root,
    "cac:OrderReference/cbc:IssueDate"
)


# ------------------------------------------------------------
# INVOICE LEVEL ALLOWANCE / CHARGE
# ------------------------------------------------------------

invoice_allowance = None
invoice_charge = None

invoice_allowance_charge = root.find(
    "cac:AllowanceCharge",
    NS
)

if invoice_allowance_charge is not None:

    amount = decimal_or_none(
        find_text(
            invoice_allowance_charge,
            "cbc:Amount"
        )
    )

    # ChargeIndicator is kept as source information,
    # but LegalMonetaryTotal will be authoritative
    # for the final invoice-level classification.

    charge_indicator = find_text(
        invoice_allowance_charge,
        "cbc:ChargeIndicator"
    )


# ------------------------------------------------------------
# INVOICE LINES
# ------------------------------------------------------------

lines = []

invoice_lines = root.findall(
    "cac:InvoiceLine",
    NS
)


for line in invoice_lines:

    # --------------------------------------------------------
    # BASIC LINE DATA
    # --------------------------------------------------------

    line_id = find_text(
        line,
        "cbc:ID"
    )

    quantity_element = line.find(
        "cbc:InvoicedQuantity",
        NS
    )

    quantity = None
    unit_code = None

    if quantity_element is not None:

        quantity = decimal_or_none(
            quantity_element.text
        )

        unit_code = quantity_element.get(
            "unitCode"
        )


    # --------------------------------------------------------
    # PRODUCT
    # --------------------------------------------------------

    item_name = find_text(
        line,
        "cac:Item/cbc:Name"
    )

    seller_item_id = find_text(
        line,
        "cac:Item/"
        "cac:SellersItemIdentification/"
        "cbc:ID"
    )

    manufacturer_item_id = find_text(
        line,
        "cac:Item/"
        "cac:ManufacturersItemIdentification/"
        "cbc:ID"
    )

    standard_item_id = find_text(
        line,
        "cac:Item/"
        "cac:StandardItemIdentification/"
        "cbc:ID"
    )


    # --------------------------------------------------------
    # PRICE
    # --------------------------------------------------------

    price_amount = decimal_or_none(
        find_text(
            line,
            "cac:Price/cbc:PriceAmount"
        )
    )


    # --------------------------------------------------------
    # LINE TOTAL
    # --------------------------------------------------------

    line_extension = decimal_or_none(
        find_text(
            line,
            "cbc:LineExtensionAmount"
        )
    )


    # --------------------------------------------------------
    # LINE DISCOUNT / CHARGE
    # --------------------------------------------------------

    line_allowance = None
    line_charge = None

    line_allowance_charge = line.find(
        "cac:AllowanceCharge",
        NS
    )

    if line_allowance_charge is not None:

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

        if charge_indicator == "false":

            line_allowance = amount

        elif charge_indicator == "true":

            line_charge = amount


    # --------------------------------------------------------
    # VAT
    # --------------------------------------------------------

    vat_rate = None
    vat_amount = None

    tax_total = line.find(
        "cac:TaxTotal",
        NS
    )

    if tax_total is not None:

        vat_amount = decimal_or_none(
            find_text(
                tax_total,
                "cbc:TaxAmount"
            )
        )

        tax_subtotal = tax_total.find(
            "cac:TaxSubtotal",
            NS
        )

        if tax_subtotal is not None:

            vat_rate = decimal_or_none(
                find_text(
                    tax_subtotal,
                    "cbc:Percent"
                )
            )


    # --------------------------------------------------------
    # NORMALIZED LINE
    # --------------------------------------------------------

    normalized_line = {

        "line_id": line_id,

        "product": {

            "name": item_name,

            "seller_item_id":
                seller_item_id,

            "manufacturer_item_id":
                manufacturer_item_id,

            "standard_item_id":
                standard_item_id,
        },

        "quantity":
            quantity,

        "unit":
            unit_code,

        "price_amount":
            price_amount,

        "line_total":
            line_extension,

        "allowance":
            line_allowance,

        "charge":
            line_charge,

        "vat_rate":
            vat_rate,

        "vat_amount":
            vat_amount,
    }

    lines.append(
        normalized_line
    )


# ------------------------------------------------------------
# TOTALS
# ------------------------------------------------------------

tax_total = decimal_or_none(
    find_text(
        root,
        "cac:TaxTotal/cbc:TaxAmount"
    )
)

line_extension_total = decimal_or_none(
    find_text(
        root,
        "cac:LegalMonetaryTotal/"
        "cbc:LineExtensionAmount"
    )
)

tax_exclusive_total = decimal_or_none(
    find_text(
        root,
        "cac:LegalMonetaryTotal/"
        "cbc:TaxExclusiveAmount"
    )
)

tax_inclusive_total = decimal_or_none(
    find_text(
        root,
        "cac:LegalMonetaryTotal/"
        "cbc:TaxInclusiveAmount"
    )
)

allowance_total = decimal_or_none(
    find_text(
        root,
        "cac:LegalMonetaryTotal/"
        "cbc:AllowanceTotalAmount"
    )
)

charge_total = decimal_or_none(
    find_text(
        root,
        "cac:LegalMonetaryTotal/"
        "cbc:ChargeTotalAmount"
    )
)

payable_total = decimal_or_none(
    find_text(
        root,
        "cac:LegalMonetaryTotal/"
        "cbc:PayableAmount"
    )
)
# LegalMonetaryTotal is authoritative for
# invoice-level allowance / charge classification.

invoice_allowance = allowance_total
invoice_charge = charge_total


# ============================================================
# NORMALIZED INVOICE
# ============================================================

normalized_invoice = {

    "document": {

        "type":
            document_type,

        "invoice_number":
            invoice_number,

        "uuid":
            uuid,

        "issue_date":
            issue_date,

        "issue_time":
            issue_time,

        "invoice_type":
            invoice_type,

        "currency":
            currency,
    },


    "supplier": {

        "name":
            supplier_name,
    },


    "references": {

        "despatch_number":
            despatch_number,

        "despatch_date":
            despatch_date,

        "order_number":
            order_number,

        "order_date":
            order_date,
    },


    "invoice_adjustments": {

        "allowance":
            invoice_allowance,

        "charge":
            invoice_charge,
    },


    "lines":
        lines,


    "totals": {

        "line_extension":
            line_extension_total,

        "tax_exclusive":
            tax_exclusive_total,

        "tax_total":
            tax_total,

        "tax_inclusive":
            tax_inclusive_total,

        "allowance_total":
            allowance_total,

        "charge_total":
            charge_total,

        "payable":
            payable_total,
    },


    "validation": {

    "expected_line_count":
        (
            int(expected_line_count)
            if expected_line_count is not None
            else None
        ),

    "parsed_line_count":
        len(lines),
}
}

# ============================================================
# AUDIT ENGINE TEST
# ============================================================

from audit_engine import audit_invoice

audit_result = audit_invoice(normalized_invoice)

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

# ------------------------------------------------------------
# DIAGNOSTICS
# ------------------------------------------------------------

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

# ============================================================
# OUTPUT
# ============================================================

print()
print("DOCUMENT")
print("-" * 60)

print(
    f"Type           : {document_type}"
)

print(
    f"Invoice number : {invoice_number}"
)

print(
    f"UUID           : {uuid}"
)

print(
    f"Date           : {issue_date}"
)

print(
    f"Time           : {issue_time}"
)

print(
    f"Type code      : {invoice_type}"
)

print(
    f"Supplier       : {supplier_name}"
)

print(
    f"Currency       : {currency}"
)


print()
print("REFERENCES")
print("-" * 60)

print(
    f"Order No.      : {order_number}"
)

print(
    f"Order Date     : {order_date}"
)

print(
    f"Despatch No.   : {despatch_number}"
)

print(
    f"Despatch Date  : {despatch_date}"
)


print()
print("INVOICE ADJUSTMENTS")
print("-" * 60)

print(
    f"Invoice allowance : "
    f"{invoice_allowance}"
)

print(
    f"Invoice charge    : "
    f"{invoice_charge}"
)


print()
print("LINES")
print("-" * 60)

print(
    f"Expected lines : "
    f"{normalized_invoice['validation']['expected_line_count']}"
)

print(
    f"Parsed lines   : "
    f"{normalized_invoice['validation']['parsed_line_count']}"
)


for line in lines:

    product = line["product"]

    print()

    print(
        f"{line['line_id']} | "
        f"{product['name']}"
    )

    print(
        f"   Product IDs  : "
        f"seller={product['seller_item_id']} | "
        f"manufacturer={product['manufacturer_item_id']} | "
        f"standard={product['standard_item_id']}"
    )

    print(
        f"   Quantity     : "
        f"{line['quantity']} "
        f"{line['unit']}"
    )

    print(
        f"   PriceAmount  : "
        f"{line['price_amount']}"
    )

    print(
        f"   Line total   : "
        f"{line['line_total']}"
    )

    print(
        f"   Allowance    : "
        f"{line['allowance']}"
    )

    print(
        f"   Charge       : "
        f"{line['charge']}"
    )

    print(
        f"   VAT rate     : "
        f"{line['vat_rate']}%"
    )

    print(
        f"   VAT amount   : "
        f"{line['vat_amount']}"
    )


print()
print("TOTALS")
print("-" * 60)

print(
    f"Line extension : "
    f"{line_extension_total}"
)

print(
    f"Tax exclusive  : "
    f"{tax_exclusive_total}"
)

print(
    f"Tax total      : "
    f"{tax_total}"
)

print(
    f"Tax inclusive  : "
    f"{tax_inclusive_total}"
)

print(
    f"Allowance      : "
    f"{allowance_total}"
)

print(
    f"Charge         : "
    f"{charge_total}"
)

print(
    f"Payable        : "
    f"{payable_total}"
)

print()

print()
print("=" * 60)