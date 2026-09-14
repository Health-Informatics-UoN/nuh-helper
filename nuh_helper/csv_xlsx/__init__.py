import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def append_page(wb: Workbook, sheet_name: str, sheet_data: list[list]) -> None:
    ws = wb.create_sheet(sheet_name)
    overwrite_page(ws, sheet_data)


def overwrite_page(ws: Worksheet, sheet_data: list[list]) -> None:
    for row_idx, row_data in enumerate(sheet_data, start=1):
        for col_idx, cell_data in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=cell_data)


def csvs_into_xlsx(xlsx: Path, csvs: None | Path | list[Path] = None) -> None:

    if csvs is None:
        csvs = xlsx.parent

    if isinstance(csvs, Path):
        csvs = list(csvs.glob("*.csv"))

    assert isinstance(csvs, list)

    assert not xlsx.is_dir()

    workbook = Workbook()
    first_page = True

    for csv_file in csvs:
        assert csv_file.is_file()
        csv_name: str = csv_file.stem

        with csv_file.open() as file:
            data = list(csv.reader(file.readlines()))

        # write teh data
        if first_page:
            page: Worksheet = workbook.active
            page.title = csv_name
            overwrite_page(page, data)
            first_page = False
        else:
            append_page(workbook, csv_name, data)

    assert not first_page
    workbook.save(xlsx)


__all__ = [
    csvs_into_xlsx,
]
