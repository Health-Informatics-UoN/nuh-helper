import contextlib
import csv
from _collections_abc import dict_keys
from collections.abc import Generator as gen
from pathlib import Path

from nuh_helper.book_page.book import BasePage, BaseWorkBook, row_val_t


class CSVWorkBook(BaseWorkBook):
    def __init__(self, path: str | Path | list[Path]) -> None:
        self._page_files = {
            page.stem.lower(): page.absolute() for page in csv_files(path)
        }

    def keys(self) -> dict_keys:
        return self._page_files.keys()

    @contextlib.contextmanager
    def __getitem__(self, name: str) -> gen[BasePage]:

        class CSVPage(BasePage):
            def _stream_rows_impl(self) -> gen[tuple[int, row_val_t]]:
                assert isinstance(self._page, Path)
                with self._page.open() as file:
                    for row, line in enumerate(
                        csv.reader(
                            file,
                            delimiter="\t" if file.name.endswith(".tsv") else ",",
                        )
                    ):
                        # strip bom
                        if (
                            row == 0
                            and isinstance(line[0], str)
                            and line[0][:1] == "\ufeff"
                        ):
                            line[0] = line[0][1:]

                        yield (row, line)

        assert isinstance(self._page_files[name], Path)
        yield CSVPage(name, self._page_files[name])

    def _stream_copy_impl(self, into: Path) -> gen[tuple[str, int, row_val_t]]:

        if into.is_file():
            raise RuntimeError(f"can't write csvs to {into} because it's a file")

        # don't overwrite input
        if inputs_in_there := sorted(
            [
                input
                for input in self._page_files.values()
                if str(input).startswith(str(into))
            ]
        ):
            raise RuntimeError(
                f"can't write to {into} - there are inputs there {inputs_in_there=}"
            )

        into.mkdir(parents=True, exist_ok=True)

        for page in self:
            name = page.name
            file = self._page_files[name].name
            with (into / file).open("w") as data:
                writer = csv.writer(
                    data, delimiter="\t" if file.endswith(".tsv") else ","
                )

                for row_idx, row_cells in page.stream_rows():
                    yield name, row_idx, row_cells
                    writer.writerow(row_cells)


def csv_files(path: str | Path | list[Path]) -> list[Path]:
    """process whatever form the path takes into a list of Path objects"""
    if isinstance(path, str):
        path = Path(path)

    # get it to a list
    if isinstance(path, Path):
        if path.is_file():
            raise RuntimeError(f"need a dir or list of files, but, got a file {path=}")
        path = sorted(
            list(path.glob("*.csv")) + list(path.glob("*.tsv")),
            key=lambda file: file.stem.lower(),
        )

    # compute dupes
    collisions = sorted(
        [
            page
            for page in path
            if len([them for them in path if them.stem.lower() == page.stem.lower()])
            != 1
        ],
        key=lambda file: file.name.lower(),
    )
    # raise error if there are dupes
    if collisions:
        raise ValueError(
            f"the following csv files have name collisions\n    {
                '\n    '.join(
                    [
                        (file.stem.lower() + '\n        ' + str(file))
                        for file in collisions
                    ]
                )
            }"
        )

    return path
