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


@pytest.fixture
def second_csv(tmp_path: Path) -> Path:
    """CSV with columns: visit_id, patient_id, visit_date, result."""
    path = tmp_path / "visits.csv"
    with open(path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["visit_id", "patient_id", "visit_date", "result"])
        writer.writerow(["101", "1", "2024-01-10", "normal"])
        writer.writerow(["102", "2", "2024-02-20", "abnormal"])
    return path


def field_overview_fields(wb: openpyxl.Workbook, table_name: str) -> list[str]:
    ws = wb["Field Overview"]
    return [
        row[1].value
        for row in ws.iter_rows(min_row=2)
        if row[0].value == table_name and row[1].value
    ]


def value_sheet_columns(wb: openpyxl.Workbook, sheet_name: str) -> list[str]:
    ws = wb[sheet_name]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    return [headers[i] for i in range(0, len(headers), 2)]


def value_sheet_data(
    wb: openpyxl.Workbook, sheet_name: str, column: str
) -> list[str]:
    """Return the values listed under a given column in a value sheet."""
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    col_index = next(i for i, v in enumerate(header) if v == column)
    return [row[col_index] for row in rows[1:] if row[col_index] not in (None, "")]


def test_no_excluded_columns(simple_csv: Path, tmp_path: Path) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report([str(simple_csv)], output_path=str(out))
    wb = openpyxl.load_workbook(out)
    assert field_overview_fields(wb, "patients.csv") == ["id", "name", "dob", "score"]
    assert value_sheet_columns(wb, "patients.csv") == ["id", "name", "dob", "score"]


def test_excluded_column_still_in_field_overview(
    simple_csv: Path, tmp_path: Path
) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)],
        output_path=str(out),
        excluded_columns={"patients.csv": ["dob"]},
    )
    wb = openpyxl.load_workbook(out)
    assert "dob" in field_overview_fields(wb, "patients.csv")


def test_excluded_column_still_in_value_sheet_header(
    simple_csv: Path, tmp_path: Path
) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)],
        output_path=str(out),
        excluded_columns={"patients.csv": ["dob"]},
    )
    wb = openpyxl.load_workbook(out)
    assert "dob" in value_sheet_columns(wb, "patients.csv")


def test_excluded_column_has_no_values(simple_csv: Path, tmp_path: Path) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)],
        output_path=str(out),
        excluded_columns={"patients.csv": ["dob"]},
    )
    wb = openpyxl.load_workbook(out)
    assert value_sheet_data(wb, "patients.csv", "dob") == []


def test_non_excluded_column_still_has_values(simple_csv: Path, tmp_path: Path) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)],
        output_path=str(out),
        excluded_columns={"patients.csv": ["dob"]},
    )
    wb = openpyxl.load_workbook(out)
    assert value_sheet_data(wb, "patients.csv", "name") != []


def test_exclusions_are_per_table(
    simple_csv: Path, second_csv: Path, tmp_path: Path
) -> None:
    """Exclusions on one table must not affect another."""
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv), str(second_csv)],
        output_path=str(out),
        excluded_columns={"patients.csv": ["dob"]},
    )
    wb = openpyxl.load_workbook(out)
    assert value_sheet_data(wb, "patients.csv", "dob") == []
    assert value_sheet_data(wb, "visits.csv", "visit_date") != []


def test_excluded_nonexistent_column_is_ignored(
    simple_csv: Path, tmp_path: Path
) -> None:
    out = tmp_path / "report.xlsx"
    generate_scan_report(
        [str(simple_csv)],
        output_path=str(out),
        excluded_columns={"patients.csv": ["nonexistent"]},
    )
    wb = openpyxl.load_workbook(out)
    assert value_sheet_data(wb, "patients.csv", "name") != []
