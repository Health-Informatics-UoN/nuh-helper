from abc import ABC, abstractmethod
from collections.abc import Generator as gen
from collections.abc import Iterable as iterator
from datetime import date, datetime
from functools import cached_property
from pathlib import Path

cell_val_t = None | str | date | datetime
row_val_t = list[cell_val_t]


class Page(ABC):
    def __init__(self, name: str, page: object) -> None:
        self._name = name
        self._page = page

    @abstractmethod
    def _stream_rows_impl(self) -> gen[tuple[int, row_val_t]]:
        """performs the workbook specific yield/write of the row data"""
        ...

    def stream_rows(self) -> gen[tuple[int, row_val_t]]:
        """public method for reading a RO stream of rows from a specific page"""

        for row_idx, row_data in self._stream_rows_impl():
            if Book.normalize:
                for raw_idx, raw_val in enumerate(row_data):
                    row_data[raw_idx] = normalize_val(raw_val)

            yield row_idx, row_data

    @cached_property
    def name(self) -> str:
        return self._name


class Book(ABC):
    normalize: bool = True

    def __contains__(self, name: str) -> bool:
        return name in self.keys()

    @abstractmethod
    def __getitem__(self, name: str) -> gen[Page]: ...

    @abstractmethod
    def keys(self) -> iterator[str]: ...

    @abstractmethod
    def _stream_copy_impl(self, into: Path) -> gen[tuple[str, int, row_val_t]]:
        """performs the workbook specific yield/write of the row data"""
        ...

    def stream_copy(
        self,
        into: str | Path,
        date_format: None | str = None,
        time_format: None | str = None,
    ) -> gen[tuple[str, int, row_val_t]]:
        """public method for copying a stream of rows between two workbooks"""

        into = (into if isinstance(into, Path) else Path(into)).absolute()

        time_format = (
            time_format
            if time_format
            else (date_format if date_format else "%Y-%m-%d %H:%M:%S")
        )
        date_format = date_format if date_format else "%Y-%m-%d"

        for page_name, row_idx, row_data in self._stream_copy_impl(into):
            yield page_name, row_idx, row_data
            if Book.normalize:
                for raw_idx, raw_val in enumerate(row_data):
                    if isinstance(raw_val, date):
                        row_data[raw_idx] = raw_val.strftime(date_format)
                    elif isinstance(raw_val, datetime):
                        row_data[raw_idx] = raw_val.strftime(time_format)

    def __iter__(self) -> gen[Page]:
        for name in self.keys():
            with self[name] as page:
                yield page


def normalize_val(src: object) -> object:
    """parse to a datetime, .strip(), turn "" into None, retain bool"""

    if isinstance(src, (bool | float | int | datetime | date)) or not Book.normalize:
        return src

    if isinstance(src, str):
        src = src.strip()

    if not src:
        return None

    if isinstance(src, str):
        # Try a handful of common formats, including YYYY-DD-MM found in some feeds
        for fmt in ("%Y-%m-%d", "%Y-%d-%m", "%d-%m-%Y", "%m-%d-%Y"):
            try:
                return datetime.strptime(src, fmt)
            except (TypeError, ValueError):
                # If parsing with this specific format fails, silently try next format.
                pass

        # Fallback: try ISO 8601 parsing (handles many ambiguous/edge cases)
        try:
            return datetime.fromisoformat(src)
        except (TypeError, ValueError):
            pass

        # Last resort: try a few more relaxed formats
        for fmt in (
            "%Y/%m/%d",
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%d %H:%M:%S",
        ):
            try:
                return datetime.strptime(src, fmt)
            except (TypeError, ValueError):
                pass

    return src


def normalize_row(row: list[str]) -> list[str]:
    """modifies (and returns) the passed array to drop extra whitespace and replace
    fully blank entries with None. doesn't touch non-strings"""

    return [normalize_val(val) for val in row]
