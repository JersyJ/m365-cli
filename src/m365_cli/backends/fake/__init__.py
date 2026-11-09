"""Fake backend for tests and evaluations: behaves like classic Outlook, but never sends.

Everything lives in one JSON file; send actions are recorded in `FakeState.outbox`.
"""

from m365_cli.backends.fake.state import FakeState
from m365_cli.backends.fake.store import FakeStore, write_state

__all__ = ["FakeState", "FakeStore", "write_state"]
