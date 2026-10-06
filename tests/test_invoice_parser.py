"""
Regression and unit tests for supplier name extraction in invoice_parser.py.

Verifies:
1. Existing PartyName/Name extraction takes precedence when present.
2. If PartyName/Name is missing or empty, fallback to cac:Person is triggered.
3. Person FirstName and FamilyName are combined with a single space, ignoring whichever is missing.
4. If neither PartyName/Name nor Person names are available, returns None.
5. Real-world regression tests on actual fixtures (test13.xml, test26.xml, test16.xml).
"""

from pathlib import Path
import sys
import tempfile
import unittest

# Ensure project root is on sys.path
PROJECT_ROOT = (
    Path(__file__).resolve().parent.parent
    if Path(__file__).parent.name == "tests"
    else Path(__file__).resolve().parent
)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from invoice_parser import parse_invoice


UBL_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
         xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
         xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
  <cbc:ID>TEST-001</cbc:ID>
  <cbc:UUID>11111111-2222-3333-4444-555555555555</cbc:UUID>
  <cbc:IssueDate>2026-03-01</cbc:IssueDate>
  <cbc:DocumentCurrencyCode>TRY</cbc:DocumentCurrencyCode>
  <cac:AccountingSupplierParty>
    <cac:Party>
{party_body}
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
      <cbc:Name>Item</cbc:Name>
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


class TestSupplierNameFallback(unittest.TestCase):
    """Unit and regression tests for supplier name extraction with Person fallback."""

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_supp_fallback_"))

    def tearDown(self):
        import shutil
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _parse_xml_string(self, party_body: str):
        xml_content = UBL_TEMPLATE.format(party_body=party_body)
        xml_path = self.temp_dir / "invoice.xml"
        xml_path.write_text(xml_content, encoding="utf-8")
        return parse_invoice(xml_path)

    # ------------------------------------------------------------
    # 1. Real-world regression tests
    # ------------------------------------------------------------

    def _find_fixture(self, filename: str) -> Path:
        candidates = [
            PROJECT_ROOT / filename,
            PROJECT_ROOT / "test_data" / "xml" / filename,
        ]
        for c in candidates:
            if c.exists():
                return c
        self.skipTest(f"Fixture {filename} not found")

    def test_real_world_sole_proprietorship_test13(self):
        """In test13.xml, PartyName is missing and cac:Person is present."""
        fixture = self._find_fixture("test13.xml")
        inv = parse_invoice(fixture)
        self.assertEqual(
            inv["supplier"]["name"],
            "Ferah GÜNER ALTUĞ / Ferah Eczanesi"
        )

    def test_real_world_sole_proprietorship_test26(self):
        """In test26.xml, PartyName is missing and cac:Person is present."""
        fixture = self._find_fixture("test26.xml")
        inv = parse_invoice(fixture)
        self.assertEqual(
            inv["supplier"]["name"],
            "Ferah GÜNER ALTUĞ / Ferah Eczanesi"
        )

    def test_real_world_party_name_precedence_test16(self):
        """In test16.xml, both PartyName and cac:Person are present; PartyName takes precedence."""
        fixture = self._find_fixture("test16.xml")
        inv = parse_invoice(fixture)
        self.assertEqual(
            inv["supplier"]["name"],
            "SÜLEYMAN TUNAHAN ÖZÇAKIR"
        )

    # ------------------------------------------------------------
    # 2. Minimal UBL fixture tests
    # ------------------------------------------------------------

    def test_party_name_present_without_person(self):
        """When PartyName is present, it is extracted directly."""
        party_body = """
      <cac:PartyName>
        <cbc:Name>Acme Corporation</cbc:Name>
      </cac:PartyName>"""
        inv = self._parse_xml_string(party_body)
        self.assertEqual(inv["supplier"]["name"], "Acme Corporation")

    def test_party_name_present_with_person_precedence(self):
        """When PartyName is present, it takes precedence even if cac:Person is also present."""
        party_body = """
      <cac:PartyName>
        <cbc:Name>Acme Corporation</cbc:Name>
      </cac:PartyName>
      <cac:Person>
        <cbc:FirstName>John</cbc:FirstName>
        <cbc:FamilyName>Doe</cbc:FamilyName>
      </cac:Person>"""
        inv = self._parse_xml_string(party_body)
        self.assertEqual(inv["supplier"]["name"], "Acme Corporation")

    def test_party_name_missing_fallback_both_person_names(self):
        """When PartyName is absent, combine Person FirstName + FamilyName with a single space."""
        party_body = """
      <cac:Person>
        <cbc:FirstName>John</cbc:FirstName>
        <cbc:FamilyName>Doe</cbc:FamilyName>
      </cac:Person>"""
        inv = self._parse_xml_string(party_body)
        self.assertEqual(inv["supplier"]["name"], "John Doe")

    def test_party_name_empty_tag_fallback_to_person(self):
        """When PartyName/cbc:Name is empty string, fall back to cac:Person."""
        party_body = """
      <cac:PartyName>
        <cbc:Name></cbc:Name>
      </cac:PartyName>
      <cac:Person>
        <cbc:FirstName>John</cbc:FirstName>
        <cbc:FamilyName>Doe</cbc:FamilyName>
      </cac:Person>"""
        inv = self._parse_xml_string(party_body)
        self.assertEqual(inv["supplier"]["name"], "John Doe")

    def test_party_name_whitespace_only_fallback_to_person(self):
        """When PartyName/cbc:Name is whitespace only, fall back to cac:Person."""
        party_body = """
      <cac:PartyName>
        <cbc:Name>   </cbc:Name>
      </cac:PartyName>
      <cac:Person>
        <cbc:FirstName>John</cbc:FirstName>
        <cbc:FamilyName>Doe</cbc:FamilyName>
      </cac:Person>"""
        inv = self._parse_xml_string(party_body)
        self.assertEqual(inv["supplier"]["name"], "John Doe")

    def test_person_only_first_name(self):
        """When FamilyName is missing, FirstName alone is used without extra spaces."""
        party_body = """
      <cac:Person>
        <cbc:FirstName>John</cbc:FirstName>
      </cac:Person>"""
        inv = self._parse_xml_string(party_body)
        self.assertEqual(inv["supplier"]["name"], "John")

    def test_person_only_family_name(self):
        """When FirstName is missing, FamilyName alone is used without extra spaces."""
        party_body = """
      <cac:Person>
        <cbc:FamilyName>Doe</cbc:FamilyName>
      </cac:Person>"""
        inv = self._parse_xml_string(party_body)
        self.assertEqual(inv["supplier"]["name"], "Doe")

    def test_person_empty_first_and_family_name(self):
        """When Person tags are empty, returns None."""
        party_body = """
      <cac:Person>
        <cbc:FirstName></cbc:FirstName>
        <cbc:FamilyName>   </cbc:FamilyName>
      </cac:Person>"""
        inv = self._parse_xml_string(party_body)
        self.assertIsNone(inv["supplier"]["name"])

    def test_neither_party_name_nor_person(self):
        """When neither PartyName nor Person is present, returns None."""
        party_body = """
      <cac:PostalAddress>
        <cbc:CityName>Istanbul</cbc:CityName>
      </cac:PostalAddress>"""
        inv = self._parse_xml_string(party_body)
        self.assertIsNone(inv["supplier"]["name"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
