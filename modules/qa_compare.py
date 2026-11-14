#!/usr/bin/env python3
"""
QA PSV comparison tool v5

UI flow:
1. Ask user to select OIT1 Excel file.
2. Ask user to select webDM2 PSV Excel file.
3. Compare OIT Feature ID with the third Device Listing field, Device Name.
4. Save MMDD_QApsv_result_PTxxxx.xlsx beside OIT1 and open the result.

Required package:
    openpyxl
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from datetime import datetime
from html import unescape
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# -----------------------------
# Text normalization
# -----------------------------
def clean(value: Any) -> str:
    """Remove HTML tags/entities and normalize whitespace."""
    text = unescape(str(value or "")).replace("\xa0", " ")
    text = re.sub(r"<[^>]*>", "", text)
    return re.sub(r"\s+", " ", text).strip()


def key(value: Any) -> str:
    return clean(value).casefold()


def normalize_masks(value: Any) -> str:
    """Compare comma-separated masks without considering their order."""
    text = clean(value)
    if key(text) in {"", "n/a", "na", "-"}:
        return ""
    return ",".join(sorted(x.strip().casefold() for x in text.split(",") if x.strip()))


def normalize_oit_restriction(value: Any) -> str:
    """OIT Yes = webDM Restricted offering; OIT No/n/a = blank."""
    text = key(value)
    if text == "yes":
        return "restricted offering"
    if text in {"no", "", "n/a", "na", "-"}:
        return ""
    return text


def normalize_webdm_restriction(value: Any) -> str:
    text = key(value)
    if text in {"", "n/a", "na", "-", "no"}:
        return ""
    if text in {"yes", "restricted", "restricted offering"}:
        return "restricted offering"
    return text


def maturity_matches(oit_value: Any, webdm_value: Any) -> bool:
    """
    Required logic:
      OIT Yes = webDM Qualified
      OIT No  = webDM Development or Frozen
    """
    oit = key(oit_value)
    webdm = key(webdm_value)
    if oit == "yes":
        return webdm == "qualified"
    if oit == "no":
        return webdm in {"development", "frozen"}
    return oit == webdm


# -----------------------------
# Input readers
# -----------------------------
def read_oit(path: Path) -> dict[str, dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    if "OIT_result" not in workbook.sheetnames:
        raise ValueError("OIT1 must contain a sheet named 'OIT_result'.")

    sheet = workbook["OIT_result"]
    # OIT may now have a new leading column A = Feature Group.
    # Read by header name (not fixed column positions), so shifted columns remain compatible.
    headers = {clean(sheet.cell(2, c).value): c for c in range(1, sheet.max_column + 1)}

    # Accept header spelling used by the refined OIT template.
    header_aliases = {
        "Feature Catergory": "Feature Category",
        "Catergory (0-99)": "Category Range (0-99)",
        "Pcell name": "Pcell Name",
    }
    for source_name, canonical_name in header_aliases.items():
        if source_name in headers and canonical_name not in headers:
            headers[canonical_name] = headers[source_name]

    required = [
        "Feature Category",
        "Feature ID",
        "Feature Description",
        "Additional Mask List",
        "Restricted (Yes/No)",
        "Qualified (Yes/No)",
    ]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError("Missing OIT1 column(s): " + ", ".join(missing))

    records: dict[str, dict[str, Any]] = {}
    for row_num in range(3, sheet.max_row + 1):
        feature_id = clean(sheet.cell(row_num, headers["Feature ID"]).value)
        if not feature_id:
            continue
        records[key(feature_id)] = {
            "id": feature_id,
            "row": row_num,
            "category": clean(sheet.cell(row_num, headers["Feature Category"]).value),
            "description": clean(sheet.cell(row_num, headers["Feature Description"]).value),
            "masks": clean(sheet.cell(row_num, headers["Additional Mask List"]).value),
            "restrictions": clean(sheet.cell(row_num, headers["Restricted (Yes/No)"]).value),
            "maturity": clean(sheet.cell(row_num, headers["Qualified (Yes/No)"]).value),
        }
    workbook.close()
    return records


def read_webdm(path: Path) -> dict[str, dict[str, Any]]:
    """Read webDM2 as raw .psv, PSV-export Excel, or normal Excel Device Listing.
    Match key and all comparison logic remain unchanged.
    """
    # Raw PSV text file: parse the same pipe-separated Device Listing rows.
    if path.suffix.lower() == ".psv":
        records: dict[str, dict[str, Any]] = {}
        headers: dict[str, int] | None = None
        text = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = path.read_text(encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError(f"Cannot decode PSV file: {path.name}")
        for row_num, raw in enumerate(text.splitlines(), start=1):
            if not raw.startswith("|Device Listing|"):
                continue
            fields = [clean(x) for x in raw.split("|")[1:]]
            if fields and fields[-1] == "":
                fields.pop()
            if len(fields) > 2 and fields[2] == "Device Name":
                headers = {name: index for index, name in enumerate(fields)}
                continue
            if not headers:
                continue
            def get_field(name: str) -> str:
                index = headers.get(name)
                return fields[index] if index is not None and index < len(fields) else ""
            device_name = get_field("Device Name")
            if device_name:
                records[key(device_name)] = {
                    "id": device_name, "row": row_num,
                    "category": get_field("Device Class"),
                    "description": get_field("Device Description"),
                    "masks": get_field("Additional Masks"),
                    "restrictions": get_field("Restrictions"),
                    "maturity": get_field("Maturity"),
                }
        if not records:
            raise ValueError("Could not find Device Listing in raw PSV file.")
        return records

    workbook = load_workbook(path, read_only=True, data_only=True)
    records: dict[str, dict[str, Any]] = {}

    # Format A: PSV-export workbook. Device Listing rows are pipe-separated in column A.
    for sheet in workbook.worksheets:
        headers: dict[str, int] | None = None
        for row_num, row in enumerate(
            sheet.iter_rows(min_col=1, max_col=1, values_only=True), start=1
        ):
            raw = row[0]
            if not isinstance(raw, str) or not raw.startswith("|Device Listing|"):
                continue
            fields = [clean(x) for x in raw.split("|")[1:]]
            if fields and fields[-1] == "":
                fields.pop()
            if len(fields) > 2 and fields[2] == "Device Name":
                headers = {name: index for index, name in enumerate(fields)}
                continue
            if not headers:
                continue
            def get_field(name: str) -> str:
                index = headers.get(name)
                return fields[index] if index is not None and index < len(fields) else ""
            device_name = get_field("Device Name")
            if device_name:
                records[key(device_name)] = {
                    "id": device_name, "row": row_num,
                    "category": get_field("Device Class"),
                    "description": get_field("Device Description"),
                    "masks": get_field("Additional Masks"),
                    "restrictions": get_field("Restrictions"),
                    "maturity": get_field("Maturity"),
                }
        if records:
            workbook.close()
            return records

    # Format B: normal Excel Device Listing table with columns in separate cells.
    required = {"Device Name", "Device Class", "Device Description",
                "Additional Masks", "Restrictions", "Maturity"}
    for sheet in workbook.worksheets:
        # Search first 30 rows for the Device Listing header row.
        for header_row in range(1, min(sheet.max_row, 30) + 1):
            headers = {clean(sheet.cell(header_row, c).value): c
                       for c in range(1, sheet.max_column + 1)
                       if clean(sheet.cell(header_row, c).value)}
            if not required.issubset(headers):
                continue
            for row_num in range(header_row + 1, sheet.max_row + 1):
                device_name = clean(sheet.cell(row_num, headers["Device Name"]).value)
                if not device_name:
                    continue
                records[key(device_name)] = {
                    "id": device_name, "row": row_num,
                    "category": clean(sheet.cell(row_num, headers["Device Class"]).value),
                    "description": clean(sheet.cell(row_num, headers["Device Description"]).value),
                    "masks": clean(sheet.cell(row_num, headers["Additional Masks"]).value),
                    "restrictions": clean(sheet.cell(row_num, headers["Restrictions"]).value),
                    "maturity": clean(sheet.cell(row_num, headers["Maturity"]).value),
                }
            if records:
                workbook.close()
                return records

    workbook.close()
    raise ValueError("Could not find a Device Listing in PSV or normal Excel format.")


# -----------------------------
# Comparison and output
# -----------------------------
def compare_records(oit: dict, webdm: dict):
    all_rows = []
    matched_keys = sorted(set(oit) & set(webdm))

    for record_key in matched_keys:
        left = oit[record_key]
        right = webdm[record_key]
        checks = [
            (
                "Feature Category / Device Class",
                left["category"],
                right["category"],
                key(left["category"]) == key(right["category"]),
            ),
            (
                "Device Description",
                left["description"],
                right["description"],
                key(left["description"]) == key(right["description"]),
            ),
            (
                "Additional Masks",
                left["masks"],
                right["masks"],
                normalize_masks(left["masks"]) == normalize_masks(right["masks"]),
            ),
            (
                "Restrictions",
                left["restrictions"],
                right["restrictions"],
                normalize_oit_restriction(left["restrictions"])
                == normalize_webdm_restriction(right["restrictions"]),
            ),
            (
                "Maturity",
                left["maturity"],
                right["maturity"],
                maturity_matches(left["maturity"], right["maturity"]),
            ),
        ]
        for field, oit_value, webdm_value, same in checks:
            all_rows.append(
                [
                    left["id"],
                    field,
                    oit_value,
                    webdm_value,
                    "Same" if same else "Different",
                    left["row"],
                    right["row"],
                ]
            )

    differences = [row for row in all_rows if row[4] == "Different"]
    oit_only = sorted(set(oit) - set(webdm))
    webdm_only = sorted(set(webdm) - set(oit))
    return matched_keys, all_rows, differences, oit_only, webdm_only


def make_output_path(oit_path: Path) -> Path:
    mmdd = datetime.now().strftime("%m%d")
    match = re.search(r"PT\d+", oit_path.stem, flags=re.IGNORECASE)
    pt_number = match.group(0).upper() if match else "PTxxxx"
    return Path(__file__).resolve().parent / f"{mmdd}_QApsv_result_{pt_number}.xlsx"


def write_result(oit_path: Path, webdm_path: Path, output_path: Path) -> dict[str, int]:
    oit = read_oit(oit_path)
    webdm = read_webdm(webdm_path)
    matched, all_rows, differences, oit_only, webdm_only = compare_records(oit, webdm)

    workbook = Workbook()
    workbook.remove(workbook.active)
    navy, white, red, green = "1F4E78", "FFFFFF", "FCE4D6", "E2F0D9"

    def setup_sheet(sheet, title, headers):
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        title_cell = sheet.cell(1, 1, title)
        title_cell.fill = PatternFill("solid", fgColor=navy)
        title_cell.font = Font(color=white, bold=True, size=14)
        sheet.append(headers)
        for cell in sheet[2]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(color=white, bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.freeze_panes = "A3"
        sheet.sheet_view.showGridLines = False

    # Summary
    sheet = workbook.create_sheet("Summary")
    setup_sheet(sheet, "OIT vs webDM Device Listing QA", ["Metric", "Count", "Rule / Note"])
    summary_rows = [
        ("OIT IDs", len(oit), oit_path.name),
        ("webDM IDs", len(webdm), webdm_path.name),
        ("Matched IDs", len(matched), "OIT Feature ID = webDM Device Name"),
        ("Different field records", len(differences), "See Differences tab"),
        ("Devices with differences", len({row[0] for row in differences}), "Unique IDs"),
        ("OIT-only IDs", len(oit_only), "See Unmatched tab"),
        ("webDM-only IDs", len(webdm_only), "See Unmatched tab"),
        ("Maturity mapping", "", "Yes = Qualified; No = Development/Frozen"),
    ]
    for item in summary_rows:
        sheet.append(item)
    for column, width in zip("ABC", [30, 16, 78]):
        sheet.column_dimensions[column].width = width

    detail_headers = [
        "Feature ID / Device Name",
        "Field",
        "OIT Value",
        "webDM Value",
        "Result",
        "OIT Row",
        "webDM PSV Row",
    ]
    for sheet_name, data in [("Differences", differences), ("All_Comparisons", all_rows)]:
        sheet = workbook.create_sheet(sheet_name)
        setup_sheet(sheet, sheet_name.replace("_", " "), detail_headers)
        for item in data:
            sheet.append(item)
        for row in sheet.iter_rows(min_row=3):
            row[4].fill = PatternFill("solid", fgColor=red if row[4].value == "Different" else green)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        if sheet.max_row > 2:
            table = Table(
                displayName=sheet_name.replace("_", "") + "Table",
                ref=f"A2:G{sheet.max_row}",
            )
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
            )
            sheet.add_table(table)
        for index, width in enumerate([30, 32, 60, 60, 12, 10, 15], start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width

    sheet = workbook.create_sheet("Unmatched")
    setup_sheet(sheet, "Unmatched IDs", ["Source", "ID", "Source Row", "Comment"])
    for record_key in oit_only:
        sheet.append(["OIT only", oit[record_key]["id"], oit[record_key]["row"], "Not found in webDM"])
    for record_key in webdm_only:
        sheet.append(["webDM only", webdm[record_key]["id"], webdm[record_key]["row"], "Not found in OIT"])
    for column, width in zip("ABCD", [18, 40, 14, 48]):
        sheet.column_dimensions[column].width = width

    sheet = workbook.create_sheet("Read_Me")
    setup_sheet(sheet, "Comparison Rules", ["OIT Field", "webDM Field", "Comparison Rule"])
    rules = [
        ("Feature Category", "Device Class", "HTML/whitespace normalized"),
        ("Feature Description", "Device Description", "HTML/whitespace normalized"),
        ("Additional Mask List", "Additional Masks", "Mask order ignored; blank/n/a equivalent"),
        ("Restricted (Yes/No)", "Restrictions", "Yes = Restricted offering; No/n/a = blank"),
        ("Qualified (Yes/No)", "Maturity", "Yes = Qualified; No = Development/Frozen"),
    ]
    for item in rules:
        sheet.append(item)
    for column, width in zip("ABC", [32, 32, 78]):
        sheet.column_dimensions[column].width = width

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(output_path)

    return {
        "oit": len(oit),
        "webdm": len(webdm),
        "matched": len(matched),
        "differences": len(differences),
        "devices_with_differences": len({row[0] for row in differences}),
        "oit_only": len(oit_only),
        "webdm_only": len(webdm_only),
    }


# -----------------------------
# UI and file opening
# -----------------------------
def select_files_from_command() -> tuple[Path, Path]:
    """List Excel files in the script folder and ask for OIT1 then webDM2 by number."""
    script_dir = Path(__file__).resolve().parent

    def list_excel_files(exclude: set[Path] | None = None) -> list[Path]:
        exclude = exclude or set()
        return sorted(
            p for p in script_dir.iterdir()
            if p.is_file()
            and (p.suffix.lower() in {".xlsx", ".xlsm", ".psv"}
                 or p.name.lower().endswith(".psv.xlsx"))
            and not p.name.startswith("~$")
            and p not in exclude
            and "_QApsv_result_" not in p.name
        )

    def choose_file(label: str, files: list[Path]) -> Path:
        if not files:
            raise FileNotFoundError(f"No Excel file available for {label} in: {script_dir}")
        print(f"\nWhich file is {label}?")
        for i, file_path in enumerate(files, 1):
            print(f"  {i}. {file_path.name}")
        while True:
            selection = input("Enter file number: ").strip()
            if selection.isdigit() and 1 <= int(selection) <= len(files):
                return files[int(selection) - 1]
            print("Invalid selection. Try again.")

    print("=" * 64)
    print("QA PSV Comparison")
    print("Put OIT1, webDM2, and this Python file in the SAME folder.")
    print(f"Working folder: {script_dir}")
    print("=" * 64)

    oit_path = choose_file("OIT1", list_excel_files())
    print(f"Selected OIT1: {oit_path.name}")

    # List files again for webDM2, excluding the OIT1 file just selected.
    webdm_path = choose_file("webDM2", list_excel_files(exclude={oit_path}))
    print(f"Selected webDM2: {webdm_path.name}")
    return oit_path, webdm_path


def open_result(path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def show_completion(output_path: Path, stats: dict[str, int]) -> None:
    message = (
        f"Completed.\n\n"
        f"Matched IDs: {stats['matched']}\n"
        f"Different field records: {stats['differences']}\n"
        f"Devices with differences: {stats['devices_with_differences']}\n"
        f"OIT-only IDs: {stats['oit_only']}\n"
        f"webDM-only IDs: {stats['webdm_only']}\n\n"
        f"Saved result:\n{output_path}"
    )
    print(message)


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare OIT1 with webDM2 Device Listing.")
    parser.add_argument("--oit", help="OIT1 Excel path")
    parser.add_argument("--webdm", help="webDM2 PSV Excel path")
    parser.add_argument("--output", help="Optional result Excel path")
    parser.add_argument("--no-open", action="store_true", help="Do not open result automatically")
    args = parser.parse_args()

    try:
        if args.oit and args.webdm:
            oit_path, webdm_path = Path(args.oit), Path(args.webdm)
        else:
            oit_path, webdm_path = select_files_from_command()

        if not oit_path.is_file():
            raise FileNotFoundError(f"OIT1 not found: {oit_path}")
        if not webdm_path.is_file():
            raise FileNotFoundError(f"webDM2 not found: {webdm_path}")

        output_path = Path(args.output) if args.output else make_output_path(oit_path)
        stats = write_result(oit_path, webdm_path, output_path)
        show_completion(output_path, stats)
        if not args.no_open:
            open_result(output_path)
    except SystemExit:
        raise
    except Exception as exc:
        error_text = f"QA PSV failed:\n{exc}"
        print(error_text, file=sys.stderr)
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showerror("QA PSV error", error_text, parent=root)
            root.destroy()
        except Exception:
            pass
        raise SystemExit(1)


# Streamlit wrapper calls write_result() directly; CLI entry point intentionally disabled.
