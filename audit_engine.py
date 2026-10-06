from decimal import Decimal

TOLERANCE = Decimal("0.02")


def audit_invoice(invoice):
    checks = []
    diagnostics = []

    # ------------------------------------------------------------
    # CORE CHECK: TAX TOTAL CALCULATION
    # ------------------------------------------------------------

    tax_exclusive = invoice["totals"]["tax_exclusive"]
    tax_total = invoice["totals"]["tax_total"]
    tax_inclusive = invoice["totals"]["tax_inclusive"]

    if (
        tax_exclusive is None
        or tax_total is None
        or tax_inclusive is None
    ):
        tax_check = {
            "name": "Tax total calculation",
            "status": None,
            "calculated": None,
            "declared": tax_inclusive,
            "difference": None,
            "message": "Required tax totals are not fully available."
        }
    else:
        calculated_gross = tax_exclusive + tax_total
        difference = calculated_gross - tax_inclusive
        passed = abs(difference) <= TOLERANCE

        tax_check = {
            "name": "Tax total calculation",
            "status": passed,
            "calculated": calculated_gross,
            "declared": tax_inclusive,
            "difference": difference,
            "message": (
                "Tax exclusive plus tax total matches tax inclusive total."
                if passed
                else
                "Tax exclusive plus tax total does not match tax inclusive total."
            )
        }

    checks.append(tax_check)

    # ------------------------------------------------------------
    # CORE CHECK: PAYABLE TOTAL
    # ------------------------------------------------------------

    tax_inclusive = invoice["totals"]["tax_inclusive"]
    payable = invoice["totals"]["payable"]

    if tax_inclusive is None or payable is None:
        payable_check = {
            "name": "Payable total",
            "status": None,
            "calculated": tax_inclusive,
            "declared": payable,
            "difference": None,
            "message": "Tax inclusive or payable total is not available."
        }
    else:
        difference = tax_inclusive - payable
        passed = abs(difference) <= TOLERANCE

        payable_check = {
            "name": "Payable total",
            "status": passed,
            "calculated": tax_inclusive,
            "declared": payable,
            "difference": difference,
            "message": (
                "Tax inclusive total matches payable amount."
                if passed
                else
                "Tax inclusive total does not match payable amount."
            )
        }

    checks.append(payable_check)

    # ------------------------------------------------------------
    # CORE CHECK: LINE VAT TOTAL
    # ------------------------------------------------------------

    line_vat_amounts = [
        line["vat_amount"]
        for line in invoice["lines"]
        if line["vat_amount"] is not None
    ]

    declared_tax_total = invoice["totals"]["tax_total"]

    if (
        len(line_vat_amounts) != len(invoice["lines"])
        or declared_tax_total is None
    ):
        line_vat_check = {
            "name": "Line VAT total",
            "status": None,
            "calculated": None,
            "declared": declared_tax_total,
            "difference": None,
            "message": "Line-level VAT amounts are not fully available."
        }
    else:
        calculated_vat_total = sum(line_vat_amounts)
        difference = calculated_vat_total - declared_tax_total
        passed = abs(difference) <= TOLERANCE

        line_vat_check = {
            "name": "Line VAT total",
            "status": passed,
            "calculated": calculated_vat_total,
            "declared": declared_tax_total,
            "difference": difference,
            "message": (
                "Line VAT total matches declared tax total."
                if passed
                else
                "Line VAT total does not match declared tax total."
            )
        }

    checks.append(line_vat_check)

    # ------------------------------------------------------------
    # DIAGNOSTIC: LINE COUNT
    # ------------------------------------------------------------

    expected_lines = invoice["validation"]["expected_line_count"]
    actual_lines = invoice["validation"]["parsed_line_count"]

    if expected_lines is None:
        diagnostics.append({
            "name": "Line count",
            "status": None,
            "calculated": actual_lines,
            "declared": None,
            "difference": None,
            "message": "Expected line count is not available."
        })
    else:
        difference = actual_lines - expected_lines
        passed = actual_lines == expected_lines

        diagnostics.append({
            "name": "Line count",
            "status": passed,
            "calculated": actual_lines,
            "declared": expected_lines,
            "difference": difference,
            "message": (
                "Parsed line count matches declared line count."
                if passed
                else
                "Parsed line count differs from declared line count."
            )
        })

    # ------------------------------------------------------------
    # DIAGNOSTIC: INVOICE LINE TOTAL
    # ------------------------------------------------------------

    line_totals = [
        line["line_total"]
        for line in invoice["lines"]
        if line["line_total"] is not None
    ]

    declared_line_total = invoice["totals"]["line_extension"]

    if len(line_totals) != len(invoice["lines"]) or declared_line_total is None:
        diagnostics.append({
            "name": "Invoice line total",
            "status": None,
            "calculated": None,
            "declared": declared_line_total,
            "difference": None,
            "message": "Line totals are not fully available."
        })
    else:
        calculated_line_total = sum(line_totals)
        difference = calculated_line_total - declared_line_total
        passed = abs(difference) <= TOLERANCE

        diagnostics.append({
            "name": "Invoice line total",
            "status": passed,
            "calculated": calculated_line_total,
            "declared": declared_line_total,
            "difference": difference,
            "message": (
                "Calculated line total matches declared line extension."
                if passed
                else
                "Calculated line total differs from declared line extension."
            )
        })

    # ------------------------------------------------------------
    # OVERALL RESULT
    # ------------------------------------------------------------

    overall_ok = all(
        check["status"] is not False
        for check in checks
    )

    return {
        "checks": checks,
        "diagnostics": diagnostics,
        "overall_ok": overall_ok
    }
