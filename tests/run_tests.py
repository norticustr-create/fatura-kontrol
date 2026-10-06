import subprocess
import re
import sys
from pathlib import Path

# Locate project root and test XML directory
PROJECT_ROOT = (
    Path(__file__).resolve().parent.parent
    if Path(__file__).parent.name == "tests"
    else Path(__file__).resolve().parent
)
XML_DIR = (
    PROJECT_ROOT / "test_data" / "xml"
    if (PROJECT_ROOT / "test_data" / "xml").exists()
    else PROJECT_ROOT
)
PARSER_SCRIPT = PROJECT_ROOT / "invoice_parser.py"

FILES = sorted(
    XML_DIR.glob("test*.xml"),
    key=lambda p: p.name.lower()
)

FILES += sorted(
    XML_DIR.glob("TEST_*.xml"),
    key=lambda p: p.name.lower()
)

# Duplicate filenames'i kaldır
unique_files = []
seen = set()

for file in FILES:
    if file.name.lower() not in seen:
        unique_files.append(file)
        seen.add(file.name.lower())


print("=" * 70)
print("SUPPLIER INVOICE AUDITOR - TEST RUNNER")
print("=" * 70)
print(f"Fixture directory: {XML_DIR}")
print(f"Invoices found   : {len(unique_files)}")
print()

results = []


for file in unique_files:

    print(f"Testing: {file.name}")

    result = subprocess.run(
        ["python", "-X", "utf8", str(PARSER_SCRIPT), str(file)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(PROJECT_ROOT),
    )

    output = result.stdout

    match = re.search(
        r"AUDIT RESULT:\s*(PASS|REVIEW)",
        output
    )

    audit_result = match.group(1) if match else "ERROR"

    failed_checks = re.findall(
        r"✗\s*(.+)",
        output
    )

    unavailable_checks = re.findall(
        r"-\s*(.+?): NOT AVAILABLE",
        output
    )

    results.append(
        {
            "file": file.name,
            "result": audit_result,
            "failed": failed_checks,
            "unavailable": unavailable_checks,
        }
    )


print()
print("=" * 70)
print("SUMMARY")
print("=" * 70)
print()

for item in results:

    if item["failed"]:
        reason = ", ".join(item["failed"])
    elif item["unavailable"]:
        reason = "N/A: " + ", ".join(item["unavailable"])
    else:
        reason = "-"

    print(
        f"{item['file']:<25} "
        f"{item['result']:<7} "
        f"{reason}"
    )

print()
print("=" * 70)

pass_count = sum(
    1 for item in results
    if item["result"] == "PASS"
)

review_count = sum(
    1 for item in results
    if item["result"] == "REVIEW"
)

error_count = sum(
    1 for item in results
    if item["result"] == "ERROR"
)

print(f"PASS   : {pass_count}")
print(f"REVIEW : {review_count}")
print(f"ERROR  : {error_count}")
print(f"TOTAL  : {len(results)}")
print("=" * 70)
