import pytest

from m365_cli.core.errors import InvalidInputError
from m365_cli.core.paths import to_windows_path


def test_wsl_mount_paths_are_translated() -> None:
    assert to_windows_path("/mnt/c/Users/x/a b.pdf") == r"C:\Users\x\a b.pdf"
    assert to_windows_path("/mnt/d") == "D:\\"


def test_windows_paths_pass_through() -> None:
    assert to_windows_path(r"C:\tmp\a.txt") == r"C:\tmp\a.txt"


def test_linux_only_paths_are_rejected() -> None:
    with pytest.raises(InvalidInputError):
        to_windows_path("/home/me/a.txt")
