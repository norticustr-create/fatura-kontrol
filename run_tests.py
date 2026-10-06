import subprocess
import re
from pathlib import Path


FILES = sorted(
    Path(".").glob("test*.xml"),
    key=lambda p: p.name.lower()
)

FILES += sorted(
    Path(".").glob("TEST_*.xml"),
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
print()

results = []


for file in unique_files:

    print(f"Testing: {file.name}")

    result = subprocess.run(
    ["python", "-X", "utf8", "invoice_parser.py", str(file)],
    capture_output=True,
    text=True,
    encoding="utf-8",
    errors="replace"
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