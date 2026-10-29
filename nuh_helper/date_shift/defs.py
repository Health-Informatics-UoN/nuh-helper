"""
collects the exceptions and sheet_configs_t. these are used across the package and need
to be in one "root" file.
"""

from datetime import datetime
from typing import Literal

# each entry in a sheet config will match this
_sheet_config_row_t = str | int | list[str] | dict[str, set[str] | list[str]]

# a single sheet will consist of this config
_sheet_config_single_t = dict[str, _sheet_config_row_t] | Literal["skip"]

# the whole configuration for a shift workbook will have this type
sheet_configs_t = dict[str, _sheet_config_single_t]


class PatientColumnMissing(Exception):
    """raised when the patient column is missing from any sheet"""

    def __init__(self, page_name: str, column_name: str) -> None:
        message = f"The patient {column_name=} is not present on {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_name = column_name


class DateColumnsMissing(Exception):
    def __init__(self, page_name: str, column_names: list[str]) -> None:
        message = f"date {column_names=} is missing from the cdm {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_names = column_names


class TextColumnsMissing(Exception):
    def __init__(self, page_name: str, column_names: list[str]) -> None:
        message = f"text {column_names=} is missing from the cdm {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_names = column_names


class ExtraColumns(Exception):
    """raised when unknown columns appear in a page we're shifting.

    could mean a column is named wrong, or, that the sheet_config is incomplete"""

    def __init__(self, page_name: str, column_names: list[str]) -> None:
        message = f"{page_name=} has {column_names=} that're neither ignored or shifted"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_names = column_names


class UnknownPatient(Exception):
    def __init__(self, page: str, id: str) -> None:
        message = f"Unknown {id=} on {page=}"
        super().__init__(message)
        self._message = message


class MissingPatient(ValueError):
    """raised when a row has data but no patient id"""

    def __init__(self, page: str, row: int) -> None:
        message = f"{page=} has date with no patient on {row=}"

        super().__init__(message)

        self._message = message
        self._page = page
        self._row = row


class ShiftFoundNonDate(Exception):
    def __init__(
        self, page: str, row: int, col: int, col_name: str, value: str
    ) -> None:
        # the spaces and newline in the message make the value easier to copy/paste back
        message = f"{page=}[{row}, {col} @   {col_name=}  ]\n\t'{value}'"
        super().__init__(message)
        self._message: str = message

        self._page: str = page
        self._row: int = row
        self._col: int = col
        self._col_name: str = col_name
        self._value: str = value


class HiddenDate(ValueError):
    perform_checks: bool = True

    def __init__(
        self,
        sheet_name: str,
        row: int,
        col: int,
        col_name: str,
        value: str,
        found: datetime,
    ) -> None:
        message = (
            f"hidden date in [{sheet_name=}, {row}, {col} @ {col_name}] {value=} "
            + f"// {found=}"
        )
        super().__init__(message)
        self._message = message

        self._sheet_name = sheet_name
        self._row = row
        self._col = col
        self._col_name = col_name
        self._value = value
        self._found = found


class BlankColumnHasData(Exception):
    """raised when a column with a blank name has data.

    prevents data being hidden in the wrong part of the CDM"""

    def __init__(self, page: str, row: int, col: int, value: any) -> None:
        message = f"[{page=} @ {row}, {col}] is a blank column with data {value=}"
        super().__init__(message)
        self._message = message

        self._page: str = page
        self._row: int = row
        self._col: int = col
        self._value: any = value


class FoundMissingPages(Exception):
    def __init__(self, page_names: list[str]) -> None:
        message = f"page missing from the source cdm {page_names=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_names


class FoundExtraPages(Exception):
    def __init__(self, page_names: list[str], expected: list[str]) -> None:
        message = (
            f"unexpected page(s) found in the source cdm {page_names=} ({expected=})"
        )
        super().__init__(message)
        self._message = message

        self._page_names: list[str] = page_names
        self._config_names: list[str] = expected
