"""Pre-shift auditing for unconfigured date columns."""

import logging
from typing import Any, cast

import pandas as pd

from nuh_helper.date_shift import _excel, _parse

logger = logging.getLogger(__name__)


def audit_date_columns(
    input_file: str,
    sheet_configs: dict[str, dict[str, Any]],
    threshold: float = 0.5,
) -> dict[str, list[str]]:
    """
    Detect columns that appear to contain dates but are not configured for shifting.

    Reads every sheet in the workbook and checks each column that is not listed
    in sheet_configs date_columns. A column is flagged when at least `threshold`
    of its non-null values parse as dates.

    Useful for validating configuration before running shift_excel_dates or
    shift_excel_dates_inplace — helps catch "date escape" where dates in
    unconfigured columns would pass through unshifted.

    Args:
        input_file: Path to the Excel file to inspect.
        sheet_configs: The same sheet_configs dict passed to the shift functions,
                       describing which columns are already configured for shifting.
        threshold: Fraction of non-null values that must parse as dates for a
                   column to be flagged (default: 0.5). Lower for more sensitivity.

    Returns:
        Dict mapping sheet names to lists of column names that appear to contain
        dates but are not configured for shifting. An empty dict means no
        unconfigured date columns were detected.
    """
    logger.info("Auditing date columns in '%s'", input_file)
    findings: dict[str, list[str]] = {}

    excel_file = pd.ExcelFile(input_file, engine="openpyxl")

    for sheet_name in excel_file.sheet_names:
        sheet_name = cast(str, sheet_name)
        config = sheet_configs.get(sheet_name, {})
        configured_date_cols: set[str] = set(config.get("date_columns", []))
        patient_id_col: str = cast(str, config.get("patient_id_col", ""))
        header_row: int = cast(int, config.get("header_row", 0))
        skip_rows: list[int] | None = config.get("skip_rows_after_header")

        if sheet_name not in sheet_configs:
            logger.warning(
                "Sheet '%s' has no entry in sheet_configs — auditing with "
                "header_row=0 and no skip rows. Add it to sheet_configs with "
                "at least 'header_row' set if the sheet has a non-standard layout.",
                sheet_name,
            )

        df, _, _, _ = _excel._read_sheet_with_structure(
            excel_file,
            sheet_name=sheet_name,
            header_row=header_row,
            input_file=input_file,
            skip_rows_after_header=skip_rows,
        )

        suspect_cols: list[str] = []
        for col in df.columns:
            col = str(col)
            if (
                col in configured_date_cols
                or col == patient_id_col
                or col.startswith("Unnamed:")
            ):
                continue

            non_null = df[col].dropna()
            if len(non_null) == 0:
                continue

            date_count = sum(
                1 for v in non_null if _parse._parse_date_value(v) is not None
            )
            if date_count == 0:
                continue

            ratio = date_count / len(non_null)
            if ratio >= threshold:
                suspect_cols.append(col)
                logger.warning(
                    "Sheet '%s', column '%s': %d/%d non-null values look like "
                    "dates but are not configured for shifting",
                    sheet_name,
                    col,
                    date_count,
                    len(non_null),
                )

        if suspect_cols:
            findings[sheet_name] = suspect_cols

    if findings:
        total = sum(len(cols) for cols in findings.values())
        logger.warning(
            "%d unconfigured date column(s) found across %d sheet(s) — "
            "these dates will not be shifted",
            total,
            len(findings),
        )
    else:
        logger.info("No unconfigured date columns detected in '%s'", input_file)

    return findings
