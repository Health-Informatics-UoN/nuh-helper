from abc import ABC, abstractmethod
from collections.abc import Generator as gen
from collections.abc import Iterable as iterator
from functools import cached_property
from pathlib import Path


class Page(ABC):
    def __init__(self, name: str, page: object) -> None:
        self._name = name
        self._page = page

    @abstractmethod
    def stream_rows(self) -> gen[tuple[int, list[str]]]: ...

    @cached_property
    def name(self) -> str:
        return self._name


class Book(ABC):
    def __contains__(self, name: str) -> bool:
        return name in self.keys()

    @abstractmethod
    def __getitem__(self, name: str) -> gen[Page]: ...

    @abstractmethod
    def keys(self) -> iterator[str]: ...

    @abstractmethod
    def stream_copy(self, into: Path) -> gen[tuple[str, int, list[str]]]: ...

    def __iter__(self) -> gen[Page]:
        for name in self.keys():
            with self[name] as page:
                yield page
