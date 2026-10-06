"""
Batch processing module for Supplier Invoice Auditor.

Collects XML invoice files from single files, directories, or ZIP archives,
and runs them through parse_invoice() and audit_invoice().
"""

import atexit
from pathlib import Path
import shutil
import sys
import tempfile
import zipfile
from typing import List, Dict, Any, Union, Optional, Callable

from audit_engine import audit_invoice
from invoice_parser import parse_invoice


_REGISTERED_TEMP_DIRS: List[Path] = []
_SOURCE_CONTAINERS: Dict[str, str] = {}
_ORIGINAL_NAMES: Dict[str, str] = {}


def cleanup_temp_dirs():
    """Removes any temporary directories created during zip extraction."""
    while _REGISTERED_TEMP_DIRS:
        temp_dir = _REGISTERED_TEMP_DIRS.pop()
        try:
            if temp_dir.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass


atexit.register(cleanup_temp_dirs)


def collect_files(
    source: Union[str, Path, List[Union[str, Path]]],
    recursive: bool = True
) -> List[Path]:
    """
    Collects XML invoice files from a source.

    Supports:
    - Single XML file path
    - Directory containing XML files (recursive search by default)
    - ZIP file containing XML files (extracted to temporary location)
    - List of any combination of the above

    Ignores non-XML files.
    Returns a sorted list of Path objects pointing to XML files.
    Raises FileNotFoundError if source path does not exist.
    """
    if isinstance(source, (list, tuple)):
        all_files: List[Path] = []
        for s in source:
            all_files.extend(collect_files(s, recursive=recursive))
        # Deduplicate while preserving order
        seen = set()
        unique_files: List[Path] = []
        for f in all_files:
            if str(f).lower() not in seen:
                seen.add(str(f).lower())
                unique_files.append(f)
        return unique_files

    src_path = Path(source)

    if not src_path.exists():
        raise FileNotFoundError(f"Source path not found: {source}")

    collected: List[Path] = []

    # 1. Directory
    if src_path.is_dir():
        pattern = "**/*" if recursive else "*"
        for item in src_path.glob(pattern):
            if (
                item.is_file()
                and item.suffix.lower() == ".xml"
                and not item.name.startswith("._")
            ):
                collected.append(item.resolve())

    # 2. ZIP archive
    elif zipfile.is_zipfile(src_path) or (
        src_path.suffix.lower() == ".zip" and src_path.is_file()
    ):
        if src_path.stat().st_size == 0:
            return []

        temp_dir = Path(tempfile.mkdtemp(prefix="inv_batch_zip_"))
        _REGISTERED_TEMP_DIRS.append(temp_dir)

        with zipfile.ZipFile(src_path, "r") as zf:
            for member in zf.namelist():
                lower_name = member.lower()
                # Check for XML extension and skip directory entries or macOS metadata
                if (
                    lower_name.endswith(".xml")
                    and not member.endswith("/")
                    and not any(
                        p.startswith("__MACOSX") or p.startswith("._")
                        for p in member.split("/")
                    )
                ):
                    extracted_target = temp_dir / member
                    extracted_target.parent.mkdir(parents=True, exist_ok=True)
                    with zf.open(member) as src_file, open(extracted_target, "wb") as dst_file:
                        shutil.copyfileobj(src_file, dst_file)

                    resolved_path = extracted_target.resolve()
                    collected.append(resolved_path)
                    _SOURCE_CONTAINERS[str(resolved_path)] = str(src_path.resolve())
                    _ORIGINAL_NAMES[str(resolved_path)] = member

    # 3. Single file
    elif src_path.is_file():
        if src_path.suffix.lower() == ".xml":
            collected.append(src_path.resolve())
        # Non-XML files are ignored, returning []

    return sorted(collected, key=lambda p: str(p).lower())


def process_batch(
    paths: Union[str, Path, List[Union[str, Path]]],
    progress_callback: Optional[Callable[[int, int, Dict[str, Any]], None]] = None,
) -> List[Dict[str, Any]]:
    """
    Processes a list of XML invoice file paths.
    Each file is parsed independently using parse_invoice() and audited using audit_invoice().
    Errors on individual files are recorded in the result dict rather than halting execution.

    If progress_callback is provided, it is invoked as:
        progress_callback(current_index, total_count, result_dict)
    after each file is processed.

    Returns a list of structured result dictionaries.
    """
    if isinstance(paths, (str, Path)):
        p = Path(paths)
        if p.is_dir() or zipfile.is_zipfile(p):
            paths = collect_files(p)
        else:
            paths = [p]

    results: List[Dict[str, Any]] = []
    total_count = len(paths)

    for idx, item in enumerate(paths, start=1):
        path_obj = Path(item).resolve() if Path(item).exists() else Path(item)
        container = _SOURCE_CONTAINERS.get(str(path_obj))
        original_name = _ORIGINAL_NAMES.get(str(path_obj), path_obj.name)

        try:
            invoice = parse_invoice(path_obj)
            audit_result = audit_invoice(invoice)
            status = "PASS" if audit_result["overall_ok"] else "REVIEW"
            failed_checks = [
                c["name"] for c in audit_result["checks"] if c["status"] is False
            ]

            doc = invoice.get("document", {})
            supp = invoice.get("supplier", {})
            totals = invoice.get("totals", {})

            result = {
                "file_path": str(path_obj),
                "file_name": path_obj.name,
                "original_name": original_name,
                "container": container,
                "success": True,
                "status": status,
                "invoice": invoice,
                "audit": audit_result,
                "error": None,
                "failed_checks": failed_checks,
                "invoice_number": doc.get("invoice_number"),
                "issue_date": doc.get("issue_date"),
                "supplier_name": supp.get("name"),
                "payable_amount": totals.get("payable"),
                "currency": doc.get("currency"),
            }
        except Exception as e:
            result = {
                "file_path": str(path_obj),
                "file_name": path_obj.name,
                "original_name": original_name,
                "container": container,
                "success": False,
                "status": "ERROR",
                "invoice": None,
                "audit": None,
                "error": f"{type(e).__name__}: {str(e)}",
                "failed_checks": [],
                "invoice_number": None,
                "issue_date": None,
                "supplier_name": None,
                "payable_amount": None,
                "currency": None,
            }

        results.append(result)

        if progress_callback is not None:
            try:
                progress_callback(idx, total_count, result)
            except Exception:
                pass

    return results


def batch_summary(results: List[Dict[str, Any]]) -> Dict[str, int]:
    """Returns a count summary of batch results."""
    return {
        "total": len(results),
        "pass": sum(1 for r in results if r["status"] == "PASS"),
        "review": sum(1 for r in results if r["status"] == "REVIEW"),
        "error": sum(1 for r in results if r["status"] == "ERROR"),
    }


def main():
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    if len(sys.argv) < 2:
        print("Usage: python batch.py <file.xml | folder | archive.zip>")
        sys.exit(1)

    source = sys.argv[1]
    try:
        files = collect_files(source)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    print("=" * 70)
    print("SUPPLIER INVOICE AUDITOR - BATCH PROCESSOR")
    print("=" * 70)
    print(f"Source: {source}")
    print(f"Collected: {len(files)} XML file(s)")
    print()

    if not files:
        print("No XML files found to process.")
        sys.exit(0)

    results = process_batch(files)

    for r in results:
        failed = (
            ", ".join(r["failed_checks"])
            if r["failed_checks"]
            else ("-" if r["success"] else r["error"])
        )
        print(f"{r['file_name']:<30} {r['status']:<8} {failed}")

    summary = batch_summary(results)
    print()
    print("-" * 70)
    print(f"PASS   : {summary['pass']}")
    print(f"REVIEW : {summary['review']}")
    print(f"ERROR  : {summary['error']}")
    print(f"TOTAL  : {summary['total']}")
    print("=" * 70)


if __name__ == "__main__":
    main()
