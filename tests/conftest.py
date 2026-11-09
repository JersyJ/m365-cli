from collections.abc import Iterator
from pathlib import Path

import pytest

from m365_cli.core.config import Settings
from m365_cli.core.context import AppContext


@pytest.fixture
def fake_ctx(tmp_path: Path) -> Iterator[AppContext]:
    ctx = AppContext(Settings(backend="fake", fake_state=tmp_path / "state.json", data_dir=tmp_path))
    yield ctx
    ctx.close()
