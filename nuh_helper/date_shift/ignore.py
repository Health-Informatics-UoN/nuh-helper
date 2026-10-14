import logging
from pathlib import Path

import yaml

from nuh_helper.date_shift.defs import normalize_inplace

shift_ignore_t = dict[str, dict[str, set[str]]]

logger = logging.getLogger(__name__)


def shift_ignore_load(shift_ignore_yaml: None | Path | str) -> shift_ignore_t:
    """load the ignore rules. if file or value is missing, just use {}

    - shift_ignore_yaml; path to the file or None.
    """
    if shift_ignore_yaml is None:
        return {}

    if not isinstance(shift_ignore_yaml, Path):
        shift_ignore_yaml = Path(shift_ignore_yaml)

    if not shift_ignore_yaml.is_file():
        logger.warning(f"didn't find shift_ignore file {shift_ignore_yaml=}")
        return {}

    with open(shift_ignore_yaml) as file:
        data: dict = yaml.safe_load(file)
        shift_ignore_full = {}
        for page_name, page_data in data.items():
            shift_ignore_page = {}

            for col_name, col_ignored in page_data.items():
                shift_ignore_page[col_name.strip()] = set(
                    normalize_inplace(col_ignored + [None])
                )
            shift_ignore_full[page_name.strip()] = shift_ignore_page
        return shift_ignore_full
