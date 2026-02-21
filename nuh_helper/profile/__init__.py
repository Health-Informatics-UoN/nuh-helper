import csv
from pathlib import Path
from collections import Counter, defaultdict
from openpyxl import Workbook
from datetime import datetime

SCAN_REPORT_FILE_NAME = "ScanReport.xlsx"

FIELD_OVERVIEW_HEADERS = [
    "Table",
    "Field",
    "Description",
    "Type",
    "Max length",
    "N rows",
]

TABLE_OVERVIEW_HEADERS = [
    "Table",
    "Description",
    "N rows",
    "N rows checked",
    "N fields",
    "N fields empty",
]


def index_table_names(table_names):
    indexed = {}
    counts = defaultdict(int)

    for name in table_names:
        indexed[name] = name if counts[name] == 0 else f"{name}_{counts[name]}"
        counts[name] += 1

    return indexed


def read_csv_header(csv_path):
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        return next(reader)


def scan_csv_values(csv_path, min_cell_count):
    counters = defaultdict(Counter)
    row_count = 0

    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            row_count += 1
            for field, value in row.items():
                value = value if value is not None else ""
                counters[field][value] += 1

    filtered = {}
    for field, counter in counters.items():
        filtered[field] = [
            (val, cnt) for val, cnt in counter.most_common() if cnt >= min_cell_count
        ]

    return filtered, row_count


def generate_scan_report(
    csv_files, output_path=SCAN_REPORT_FILE_NAME, min_cell_count=1
):
    tables = []

    for csv_file in csv_files:
        csv_file = Path(csv_file)
        header = read_csv_header(csv_file)
        tables.append({"name": csv_file.name, "path": csv_file, "fields": header})

    tables.sort(key=lambda t: t["name"])
    indexed_names = index_table_names([t["name"] for t in tables])

    wb = Workbook()
    wb.remove(wb.active)

    # FIELD_OVERVIEW
    field_sheet = wb.create_sheet("Field Overview")
    field_sheet.append(FIELD_OVERVIEW_HEADERS)

    for table in tables:
        table_name_indexed = indexed_names[table["name"]]
        for field in table["fields"]:
            field_sheet.append([table_name_indexed, field, "", "STRING", "", ""])
        field_sheet.append([""])

    # TABLE_OVERVIEW
    table_sheet = wb.create_sheet("Table Overview")
    table_sheet.append(TABLE_OVERVIEW_HEADERS)

    table_value_data = {}

    for table in tables:
        table_name_indexed = indexed_names[table["name"]]
        value_data, row_count = scan_csv_values(table["path"], min_cell_count)
        table_value_data[table_name_indexed] = value_data

        table_sheet.append(
            [table_name_indexed, "", row_count, row_count, len(table["fields"]), -1]
        )

    # VALUE SHEETS
    for table in tables:
        table_name_indexed = indexed_names[table["name"]]
        value_sheet = wb.create_sheet(table_name_indexed)

        fields = table["fields"]
        value_data = table_value_data[table_name_indexed]

        header = []
        for field in fields:
            header.append(field)
            header.append("Frequency")
        value_sheet.append(header)

        max_len = max((len(value_data.get(field, [])) for field in fields), default=0)

        for i in range(max_len):
            row = []
            for field in fields:
                values = value_data.get(field, [])
                if i < len(values):
                    row.append(values[i][0])
                    row.append(values[i][1])
                else:
                    row.append("")
                    row.append("")
            value_sheet.append(row)

    # META
    meta_sheet = wb.create_sheet("_")
    meta_sheet.append(["Key", "Value"])
    meta_sheet.append(["Version", "python-whiterabbit-lite"])
    meta_sheet.append(["Scan started at", datetime.now().isoformat()])
    meta_sheet.append(["Scan finished at", datetime.now().isoformat()])
    meta_sheet.append(["sourceType", "CSV_FILES"])
    meta_sheet.append(["scanValues", True])
    meta_sheet.append(["minCellCount", min_cell_count])

    wb.save(output_path)
    return output_path
