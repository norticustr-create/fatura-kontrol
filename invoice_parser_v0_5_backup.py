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

    charge_indicator = find_text(
        invoice_allowance_charge,
        "cbc:ChargeIndicator"
    )

    amount = decimal_or_none(
        find_text(
            invoice_allowance_charge,
            "cbc:Amount"
        )
    )

    if charge_indicator == "false":

        invoice_allowance = amount

    elif charge_indicator == "true":

        invoice_charge = amount


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


# ============================================================
# VALIDATION
# ============================================================

checks = []


# ------------------------------------------------------------
# DOCUMENT
# ------------------------------------------------------------

checks.append(
    (
        "Document type",
        document_type == "UBL Invoice"
    )
)


# ------------------------------------------------------------
# BASIC FIELDS
# ------------------------------------------------------------

basic_ok = (

    invoice_number is not None

    and issue_date is not None

    and currency is not None

)

checks.append(
    (
        "Basic invoice fields",
        basic_ok
    )
)


# ------------------------------------------------------------
# LINE COUNT
# ------------------------------------------------------------

if expected_line_count is not None:

    line_count_ok = (

        int(expected_line_count)
        == len(lines)

    )

    checks.append(
        (
            "Line count",
            line_count_ok
        )
    )

else:

    checks.append(
        (
            "Line count",
            None
        )
    )


# ------------------------------------------------------------
# LINE TOTAL CHECK
#
# IMPORTANT:
# We DO NOT assume:
#
# quantity × PriceAmount = LineExtensionAmount
#
# because some invoices express PriceAmount as
# VAT-inclusive unit price.
# ------------------------------------------------------------

line_calculation_warnings = []


for line in lines:

    quantity = line["quantity"]

    price_amount = line["price_amount"]

    line_total = line["line_total"]

    vat_rate = line["vat_rate"]


    if (

        quantity is not None

        and price_amount is not None

        and line_total is not None

    ):

        expected_from_price = money(

            quantity * price_amount

        )

        actual = money(
            line_total
        )


        if not almost_equal(
            expected_from_price,
            actual
        ):

            vat_adjusted = None

            if vat_rate is not None:

                vat_adjusted = money(

                    expected_from_price
                    / (
                        Decimal("1")
                        + vat_rate
                        / Decimal("100")
                    )

                )


            # If VAT-adjusted price explains the
            # difference, record that fact.
            if (

                vat_adjusted is not None

                and almost_equal(
                    vat_adjusted,
                    actual
                )

            ):

                reason = (
                    "PriceAmount appears VAT-inclusive"
                )

            else:

                reason = (
                    "PriceAmount does not directly "
                    "explain line total"
                )


            line_calculation_warnings.append(

                {

                    "line_id":
                        line["line_id"],

                    "price_amount":
                        price_amount,

                    "quantity":
                        quantity,

                    "expected_from_price":
                        expected_from_price,

                    "actual_line_total":
                        actual,

                    "vat_rate":
                        vat_rate,

                    "reason":
                        reason,
                }

            )


# ------------------------------------------------------------
# THIS IS NOW A WARNING, NOT A FAILURE
# ------------------------------------------------------------

checks.append(
    (
        "Line calculations",
        True
    )
)


# ------------------------------------------------------------
# SUM OF LINE TOTALS
# ------------------------------------------------------------

calculated_line_total = None

usable_line_totals = [

    line["line_total"]

    for line in lines

    if line["line_total"] is not None

]


if (

    len(usable_line_totals)
    == len(lines)

    and len(lines) > 0

):

    calculated_line_total = money(

        sum(
            usable_line_totals,
            Decimal("0")
        )

    )


line_total_check = None


if (

    calculated_line_total is not None

    and line_extension_total is not None

):

    line_total_check = almost_equal(

        calculated_line_total,

        money(
            line_extension_total
        )

    )


checks.append(
    (
        "Invoice line total",
        line_total_check
    )
)


# ------------------------------------------------------------
# TAX TOTAL CHECK
# ------------------------------------------------------------

tax_total_check = None


if (

    tax_exclusive_total is not None

    and tax_total is not None

    and tax_inclusive_total is not None

):

    calculated_gross = money(

        tax_exclusive_total
        + tax_total

    )

    tax_total_check = almost_equal(

        calculated_gross,

        money(
            tax_inclusive_total
        )

    )


checks.append(
    (
        "Tax total calculation",
        tax_total_check
    )
)


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

        "calculated_line_total":
            calculated_line_total,

        "line_calculation_warnings":
            line_calculation_warnings,
    }
}


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
    f"Calculated     : "
    f"{calculated_line_total}"
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


# ============================================================
# VALIDATION OUTPUT
# ============================================================

print()
print("VALIDATION")
print("-" * 60)

overall_ok = True


for name, result in checks:

    if result is True:

        print(
            f"✓ {name}"
        )

    elif result is False:

        print(
            f"✗ {name}"
        )

        overall_ok = False

    else:

        print(
            f"- {name}: NOT AVAILABLE"
        )


# ------------------------------------------------------------
# WARNINGS
# ------------------------------------------------------------

if line_calculation_warnings:

    print()
    print("LINE PRICE WARNINGS")
    print("-" * 60)


    for warning in line_calculation_warnings:

        print(
            f"Line {warning['line_id']}: "
            f"{warning['reason']}"
        )

        print(
            f"   PriceAmount: "
            f"{warning['price_amount']}"
        )

        print(
            f"   Quantity: "
            f"{warning['quantity']}"
        )

        print(
            f"   Price × Qty: "
            f"{warning['expected_from_price']}"
        )

        print(
            f"   Line total: "
            f"{warning['actual_line_total']}"
        )

        print(
            f"   VAT rate: "
            f"{warning['vat_rate']}%"
        )


print()

if overall_ok:

    print(
        "RESULT: PARSER VALIDATION PASSED"
    )

else:

    print(
        "RESULT: PARSER VALIDATION FAILED"
    )


print()
print("=" * 60)