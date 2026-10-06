"""
Supplier Invoice Auditor - Unified Command Line Interface.

A single entry point for bulk parsing, auditing, and Excel reporting of
Turkish e-invoice (UBL 2.1) XML files.

Usage:
    python main.py <source> [output.xlsx]
    python main.py <source> -o [output.xlsx]
"""

import argparse
from pathlib import Path
import sys
from typing import List, Optional

from batch import collect_files, process_batch, batch_summary
from excel_export import export_to_excel


def build_parser() -> argparse.ArgumentParser:
    """Builds and returns the argparse parser for the CLI."""
    parser = argparse.ArgumentParser(
        prog="invoice-auditor",
        description="Supplier Invoice Auditor - Bulk Turkish E-Invoice (UBL 2.1) XML Auditor & Excel Exporter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py test-metro.xml
  python main.py test-metro.xml my_report.xlsx
  python main.py ./invoices_folder/
  python main.py ./invoices_folder/ -o monthly_audit.xlsx
  python main.py invoices.zip audit_report.xlsx
        """,
    )

    parser.add_argument(
        "source",
        help="Path to an invoice XML file, a folder containing XML files, or a ZIP archive containing XML files.",
    )

    parser.add_argument(
        "output",
        nargs="?",
        default=None,
        help="Optional destination path for the output Excel report (.xlsx). Defaults to 'invoice_audit_report.xlsx'.",
    )

    parser.add_argument(
        "-o",
        "--output-file",
        dest="output_flag",
        default=None,
        help="Destination path for the output Excel report (.xlsx) (alternative to positional argument).",
    )

    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Do not search subfolders recursively when source is a directory.",
    )

    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Print per-invoice audit check details in the console output.",
    )

    return parser


def run_cli(args_list: Optional[List[str]] = None) -> int:
    """
    Executes the command-line interface.

    Args:
        args_list: Optional list of argument strings (e.g. for testing).
                   If None, reads from sys.argv[1:].

    Returns:
        int: Exit status code (0 for success, non-zero for CLI/input failures).
    """
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if sys.stderr and hasattr(sys.stderr, "reconfigure"):
        try:
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = build_parser()

    try:
        args = parser.parse_args(args_list)
    except SystemExit as e:
        return e.code if isinstance(e.code, int) else 1

    source_input = args.source
    output_target = args.output_flag or args.output or "invoice_audit_report.xlsx"

    # Ensure .xlsx extension
    if not output_target.lower().endswith(".xlsx"):
        output_target += ".xlsx"

    # 1. Collect files
    try:
        files = collect_files(source_input, recursive=not args.no_recursive)
    except FileNotFoundError as e:
        print(f"ERROR: Source path not found: '{source_input}'", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"ERROR: Failed to read source '{source_input}': {e}", file=sys.stderr)
        return 1

    if not files:
        print(f"ERROR: No XML invoice files found in source: '{source_input}'", file=sys.stderr)
        return 1

    print("=" * 70)
    print("SUPPLIER INVOICE AUDITOR - BULK AUDIT & EXCEL EXPORTER")
    print("=" * 70)
    print(f"Source   : {source_input}")
    print(f"Output   : {output_target}")
    print(f"Invoices : Found {len(files)} XML file(s)")
    print("-" * 70)

    # 2. Process batch
    results = process_batch(files)
    summary = batch_summary(results)

    # Print summary rows
    for r in results:
        status = r["status"]
        name = r["original_name"] or r["file_name"]
        if status == "PASS":
            mark = "✓ PASS  "
            details = "-"
        elif status == "REVIEW":
            mark = "⚠ REVIEW"
            details = ", ".join(r["failed_checks"]) if r["failed_checks"] else "Discrepancy detected"
        else:
            mark = "✗ ERROR "
            details = r["error"] or "Parse failure"

        print(f"{mark} | {name:<32} | {details}")

        if args.verbose and r["audit"]:
            for chk in r["audit"].get("checks", []):
                sym = "  ✓" if chk["status"] is True else ("  ✗" if chk["status"] is False else "  -")
                print(f"    {sym} {chk['name']}")

    # 3. Export to Excel
    try:
        saved_path = export_to_excel(results, output_target)
    except Exception as e:
        print(f"ERROR: Failed to write Excel report to '{output_target}': {e}", file=sys.stderr)
        return 1

    print("-" * 70)
    print("AUDIT SUMMARY")
    print("-" * 70)
    print(f"PASS     : {summary['pass']}")
    print(f"REVIEW   : {summary['review']}")
    print(f"ERROR    : {summary['error']}")
    print(f"TOTAL    : {summary['total']}")
    print("-" * 70)
    print(f"Report successfully saved to:\n  {saved_path}")
    print("=" * 70)

    # Note: REVIEW invoices are an intentional business audit outcome, not a CLI failure
    return 0


def entrypoint(args_list: Optional[List[str]] = None) -> int:
    """
    Unified entry point dispatcher.

    - No arguments (or '--gui'): Launches the native Windows GUI.
    - Command-line arguments: Executes the command-line interface.

    Args:
        args_list: Optional list of CLI argument strings. If None, reads sys.argv[1:].

    Returns:
        int: Process exit code.
    """
    raw_args = sys.argv[1:] if args_list is None else args_list

    if not raw_args or raw_args == ["--gui"]:
        from gui import launch_gui
        launch_gui()
        return 0

    return run_cli(raw_args)


def main():
    exit_code = entrypoint()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
