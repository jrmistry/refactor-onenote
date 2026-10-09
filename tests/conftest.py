"""Keep test authentication configuration away from real user caches/accounts."""

import pytest


@pytest.fixture(autouse=True)
def isolated_auth_environment(monkeypatch, tmp_path):
    """Real LOCALAPPDATA must not redirect mocked auth tests outside their temp root."""
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local-app-data"))
    monkeypatch.delenv("ONENOTE_CLIENT_ID", raising=False)
    monkeypatch.delenv("ONENOTE_TENANT", raising=False)
