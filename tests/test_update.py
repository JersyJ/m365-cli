from pathlib import Path

from m365_cli.core.update import is_newer, parse_version, pick_wheel, source_checkout


def test_parse_version():
    assert parse_version("v0.2.0") == (0, 2, 0)
    assert parse_version("1.10") == (1, 10)
    assert parse_version("nightly") is None


def test_is_newer():
    assert is_newer("v0.2.0", "0.1.0")
    assert not is_newer("v0.1.0", "0.1.0")
    assert not is_newer("v0.1.0", "0.2.0")
    assert is_newer("v0.10.0", "0.9.9")
    # Unparsable versions count as newer only when tag and installed differ.
    assert is_newer("nightly", "0.1.0")
    assert is_newer("v0.1.0", "unknown")
    assert not is_newer("nightly", "nightly")


def test_pick_wheel():
    assets = [
        {"name": "notes.txt", "url": "u1"},
        {"name": "m365_cli-0.2.0-py3-none-any.whl", "url": "u2"},
    ]
    picked = pick_wheel(assets)
    assert picked is not None and picked["url"] == "u2"
    assert pick_wheel([]) is None


def test_source_checkout_detects_this_repo():
    repo = source_checkout()
    assert repo is not None
    assert (Path(repo) / "pyproject.toml").is_file()
