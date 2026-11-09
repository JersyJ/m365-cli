from pathlib import Path

import pytest

from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.core.refs import BackendId, RefStore


def test_refs_are_short_stable_and_roundtrip(tmp_path: Path) -> None:
    store = RefStore(tmp_path / "refs.sqlite3")
    bid = BackendId("m", "00" * 70, "STORE")
    ref = store.put(bid)
    same_ref_again = store.put(bid)
    assert ref == same_ref_again
    assert ref.startswith("m:") and len(ref) == 9
    assert store.get(ref) == bid
    reopened = RefStore(tmp_path / "refs.sqlite3")
    assert reopened.get(ref) == bid


def test_ref_kind_is_enforced(tmp_path: Path) -> None:
    store = RefStore(tmp_path / "refs.sqlite3")
    ref = store.put(BackendId("e", "abc"))
    with pytest.raises(InvalidInputError):
        store.get(ref, kinds="m")
    assert store.get(ref, kinds="em").item_id == "abc"


def test_unknown_and_malformed_refs(tmp_path: Path) -> None:
    store = RefStore(tmp_path / "refs.sqlite3")
    with pytest.raises(NotFoundError):
        store.get("m:AAAAAAA")
    with pytest.raises(InvalidInputError):
        store.get("nonsense")
