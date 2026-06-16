#!/usr/bin/env python3
import argparse
import glob
import json
import os
from collections import Counter

try:
    from openpyxl import Workbook
except ImportError as exc:  # pragma: no cover
    raise SystemExit("openpyxl is required to create the Excel export.") from exc


DEFAULT_REPORT_GLOB = "downloads_absa/_reports/*.json"
DEFAULT_OUTPUT_PATH = "downloads_absa/_reports/no_matter_found_accounts.xlsx"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export 'no matter found' accounts from ABSA JSON reports into Excel."
    )
    parser.add_argument(
        "--report-glob",
        default=DEFAULT_REPORT_GLOB,
        help=f"Glob pattern used to find JSON report files (default: {DEFAULT_REPORT_GLOB}).",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT_PATH,
        help=f"Output XLSX path (default: {DEFAULT_OUTPUT_PATH}).",
    )
    parser.add_argument(
        "--include-prefix-mismatch",
        action="store_true",
        help="Also include rows where multiple matters existed but none matched the allowed FileRef prefixes.",
    )
    parser.add_argument(
        "--dedupe",
        action="store_true",
        help="Keep only the first occurrence for each account number.",
    )
    return parser.parse_args(argv)


def load_report(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def is_target_failure(item: dict, include_prefix_mismatch: bool) -> bool:
    if item.get("status") != "failed":
        return False
    reason = str(item.get("reason") or "")
    if reason == "matter_not_found":
        return True
    return include_prefix_mismatch and reason == "fileref_prefix_not_found"


def collect_rows(report_paths: list[str], include_prefix_mismatch: bool) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for report_path in report_paths:
        report = load_report(report_path)
        report_date = str(report.get("date") or "")
        for item in report.get("results", []):
            if not is_target_failure(item, include_prefix_mismatch):
                continue
            rows.append(
                {
                    "report_date": report_date,
                    "account_number": str(item.get("account_number") or ""),
                    "reason": str(item.get("reason") or ""),
                    "screen_id": item.get("screen_id"),
                    "source_file": str(item.get("source_file") or ""),
                    "report_file": report_path,
                }
            )
    rows.sort(key=lambda row: (row["report_date"], row["account_number"], row["source_file"]))
    return rows


def dedupe_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    seen: set[str] = set()
    deduped: list[dict[str, object]] = []
    for row in rows:
        account_number = str(row["account_number"])
        if account_number in seen:
            continue
        seen.add(account_number)
        deduped.append(row)
    return deduped


def autosize_worksheet(worksheet) -> None:
    for column_cells in worksheet.columns:
        values = [str(cell.value or "") for cell in column_cells]
        max_length = max((len(value) for value in values), default=0)
        column_letter = column_cells[0].column_letter
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 60)


def write_workbook(output_path: str, rows: list[dict[str, object]], reason_counts: Counter) -> None:
    workbook = Workbook()

    detail_sheet = workbook.active
    detail_sheet.title = "No Matter Found"
    detail_headers = [
        "Report Date",
        "Account Number",
        "Reason",
        "Screen ID",
        "Source File",
        "Report File",
    ]
    detail_sheet.append(detail_headers)
    for row in rows:
        detail_sheet.append(
            [
                row["report_date"],
                row["account_number"],
                row["reason"],
                row["screen_id"],
                row["source_file"],
                row["report_file"],
            ]
        )

    summary_sheet = workbook.create_sheet("Summary")
    summary_sheet.append(["Metric", "Value"])
    summary_sheet.append(["Total Rows", len(rows)])
    summary_sheet.append(["Distinct Accounts", len({str(row['account_number']) for row in rows})])
    for reason, count in sorted(reason_counts.items()):
        summary_sheet.append([f"Reason: {reason}", count])

    autosize_worksheet(detail_sheet)
    autosize_worksheet(summary_sheet)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    workbook.save(output_path)
    workbook.close()


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    report_paths = sorted(glob.glob(args.report_glob))
    if not report_paths:
        raise SystemExit(f"No report files matched: {args.report_glob}")

    rows = collect_rows(report_paths, include_prefix_mismatch=args.include_prefix_mismatch)
    if args.dedupe:
        rows = dedupe_rows(rows)

    reason_counts = Counter(str(row["reason"]) for row in rows)
    write_workbook(args.output, rows, reason_counts)

    print(
        f"Exported {len(rows)} row(s) from {len(report_paths)} report file(s) to {args.output}"
    )
    if rows:
        print(
            "Distinct accounts: "
            f"{len({str(row['account_number']) for row in rows})}"
        )
    else:
        print("No matching failures were found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
