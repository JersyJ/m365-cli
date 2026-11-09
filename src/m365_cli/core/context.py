from __future__ import annotations

from functools import cached_property
from typing import TYPE_CHECKING

from m365_cli.core.audit import AuditLog
from m365_cli.core.config import Settings
from m365_cli.core.errors import InvalidInputError
from m365_cli.core.refs import RefStore
from m365_cli.core.timeutil import use_clock

if TYPE_CHECKING:
    from m365_cli.backends.fake import FakeStore
    from m365_cli.backends.outlook import OutlookSession
    from m365_cli.backends.teams import TeamsSession


class AppContext:
    """Application context shared by all features: settings, ref store, send log, lazy backends.

    Backends are created on first use, so a command that never touches Outlook never starts it.
    """

    def __init__(self, settings: Settings | None = None) -> None:
        """Use `settings` or load them.

        With the fake backend, the clock is frozen at the clock of the fake data.
        """
        self.settings = settings or Settings.load()
        if self.is_fake:
            use_clock(self.fake.state.clock)

    @property
    def is_fake(self) -> bool:
        """True when the fake backend is active."""
        return self.settings.backend == "fake"

    @cached_property
    def refs(self) -> RefStore:
        """Short-ref store."""
        return RefStore(self.settings.data_dir / "refs.sqlite3")

    @cached_property
    def audit(self) -> AuditLog:
        """Send log with the per-minute send limit."""
        return AuditLog(
            self.settings.data_dir / "audit.sqlite3", max_per_minute=self.settings.max_sends_per_minute
        )

    @cached_property
    def outlook(self) -> OutlookSession:
        """Classic Outlook session, started on first use."""
        from m365_cli.backends.outlook import OutlookSession

        return OutlookSession()

    @cached_property
    def teams(self) -> TeamsSession:
        """Teams session in the dedicated browser, started on first use."""
        from m365_cli.backends.teams import BrowserConfig, TeamsSession

        s = self.settings
        return TeamsSession(
            BrowserConfig(
                port=s.teams_port,
                profile_dir=s.teams_profile_dir or s.data_dir / "chrome-teams",
                executable=s.teams_browser,
            )
        )

    @cached_property
    def fake(self) -> FakeStore:
        """Fake backend store.

        Raises:
            InvalidInputError: The fake backend is not active.
        """
        if not self.is_fake:
            raise InvalidInputError("The fake backend is not active (set backend = 'fake').")
        from m365_cli.backends.fake import FakeStore

        return FakeStore(self.settings.fake_state or self.settings.data_dir / "fake-state.json")

    def close(self) -> None:
        """Stop running backends and restore the real clock."""
        if "outlook" in self.__dict__:
            self.outlook.close()
        if "teams" in self.__dict__:
            self.teams.close()
        if self.is_fake:
            use_clock(None)
