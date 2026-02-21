import shutil
from pathlib import Path

import pytest

OUTPUT_DIR = Path(__file__).parent / "output"


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--save-output",
        action="store_true",
        default=False,
        help="Write e2e output files to tests/output/<test-name>/ for manual inspection.",
    )


@pytest.fixture
def output_path(request: pytest.FixtureRequest, tmp_path: Path) -> Path:
    """Directory for e2e test output files.

    Normally delegates to pytest's tmp_path (ephemeral).
    Pass --save-output to write to tests/output/<test-name>/ instead,
    so you can open the files afterwards for manual inspection.
    """
    if request.config.getoption("--save-output"):
        path = OUTPUT_DIR / request.node.name
        if path.exists():
            shutil.rmtree(path)
        path.mkdir(parents=True)
        return path
    return tmp_path
