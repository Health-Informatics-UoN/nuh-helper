import csv
from pathlib import Path

import openpyxl
import pytest

from nuh_helper import generate_scan_report


@pytest.fixture
def simple_csv(tmp_path: Path) -> Path:
    """CSV with columns: id, name, dob, score."""
    path = tmp_path / "patients.csv"
    with open(path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["id", "name", "dob", "score"])
        writer.writerow(["1", "Alice", "1980-01-01", "10"])
        writer.writerow(["2", "Bob", "1990-06-15", "20"])
    return path


def load_sheet_headers(wb: openpyxl.Workbook, sheet_name: str) -> list[str]:
    ws = wb[sheet_name]
    return [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]


def field_overview_fields(wb: openpyxl.Workbook) -> list[str]:
    ws = wb["Field Overview"]
    return [row[1].value for row in ws.iter_rows(min_row=2) if row[1].value]


def value_sheet_columns(wb: openpyxl.Workbook, sheet_name: str) -> list[str]:
    """Return the field names (every other header) from a value sheet."""
    ws = wb[sheet_name]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    return [headers[i] for i in range(0, len(headers), 2)]


def test_no_excluded_columns(simple_csv: Path, tmp_path: Path) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report([str(simple_csv)], output_path=str(out))
    wb = openpyxl.load_workbook(out)
    fields = field_overview_fields(wb)
    assert fields == ["id", "name", "dob", "score"]


def test_excluded_columns_removed_from_field_overview(
    simple_csv: Path, tmp_path: Path
) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)], output_path=str(out), excluded_columns=["dob"]
    )
    wb = openpyxl.load_workbook(out)
    fields = field_overview_fields(wb)
    assert "dob" not in fields
    assert fields == ["id", "name", "score"]


def test_excluded_columns_removed_from_value_sheet(
    simple_csv: Path, tmp_path: Path
) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)], output_path=str(out), excluded_columns=["dob"]
    )
    wb = openpyxl.load_workbook(out)
    columns = value_sheet_columns(wb, "patients.csv")
    assert "dob" not in columns
    assert columns == ["id", "name", "score"]


def test_multiple_excluded_columns(simple_csv: Path, tmp_path: Path) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)], output_path=str(out), excluded_columns=["dob", "id"]
    )
    wb = openpyxl.load_workbook(out)
    fields = field_overview_fields(wb)
    assert fields == ["name", "score"]
    columns = value_sheet_columns(wb, "patients.csv")
    assert columns == ["name", "score"]


def test_excluded_nonexistent_column_is_ignored(
    simple_csv: Path, tmp_path: Path
) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)], output_path=str(out), excluded_columns=["nonexistent"]
    )
    wb = openpyxl.load_workbook(out)
    fields = field_overview_fields(wb)
    assert fields == ["id", "name", "dob", "score"]
