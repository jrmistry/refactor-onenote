"""Graph tests use controlled HTTP responses rather than a Microsoft account."""

# Provider fixtures intentionally repeat wire fields across integration tests.
# pylint: disable=duplicate-code

import json
import pytest

from onenote_hierarchy import cli, graph
from onenote_hierarchy.graph import (
    GraphBackend,
    GraphClient,
    GraphError,
    SnapshotClient,
    token_provider,
)
from onenote_hierarchy.model import Node, render


class Response:  # pylint: disable=too-few-public-methods
    """A minimal requests response at the external HTTP boundary."""

    def __init__(self, payload, status=200, headers=None):
        self.payload = payload
        self.status_code = status
        self.headers = headers or {}

    def json(self):
        """Return the provider response."""
        return self.payload


class Session:  # pylint: disable=too-few-public-methods
    """Serve explicit method/path fixtures and retain HTTP request arguments."""

    def __init__(self, routes):
        self.routes = routes
        self.calls = []

    def request(self, method, url, **kwargs):
        """Return the next fixture for an exact HTTP request."""
        self.calls.append((method, url, kwargs))
        key = (method, url.split("/me/onenote/")[1])
        response = self.routes[key]
        if isinstance(response, list):
            return response.pop(0)
        return response


def client_for(routes):
    """Use the real Graph transport with controlled network responses."""

    return GraphClient(
        lambda: "fake-token", session=Session(routes), sleep=lambda _: None
    )


def test_graph_scan_paginates_nested_groups_and_orders_subpages():
    """Stopping at the first response or using modification order corrupts the tree."""

    client = client_for(
        {
            ("GET", "notebooks"): Response(
                {"value": [{"id": "n", "displayName": "N"}]}
            ),
            ("GET", "notebooks/n/sections"): Response({"value": []}),
            ("GET", "notebooks/n/sectionGroups"): Response(
                {"value": [{"id": "g", "displayName": "G"}]}
            ),
            ("GET", "sectionGroups/g/sections"): Response(
                {"value": [{"id": "s", "displayName": "S"}]}
            ),
            ("GET", "sectionGroups/g/sectionGroups"): Response(
                {"value": [{"id": "h", "displayName": "H"}]}
            ),
            ("GET", "sectionGroups/h/sections"): Response({"value": []}),
            ("GET", "sectionGroups/h/sectionGroups"): Response({"value": []}),
            ("GET", "sections/s/pages?pagelevel=true"): Response(
                {
                    "value": [{"id": "c", "title": "Child", "level": 1, "order": 1}],
                    "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/onenote/"
                    "sections/s/pages?pagelevel=true&$skip=1",
                }
            ),
            ("GET", "sections/s/pages?pagelevel=true&$skip=1"): Response(
                {
                    "value": [
                        {"id": "p", "title": "Parent", "level": 0, "order": 0},
                        {"id": "d", "title": "Detail", "level": 2, "order": 2},
                        {"id": "t", "title": "Tail", "level": 0, "order": 3},
                    ]
                }
            ),
        }
    )
    backend = GraphBackend(client)
    result = backend.scan(backend.list_notebooks())
    assert result.complete
    text = render(result)
    assert (
        "            [Page] Parent\n"
        "                [Page] Child\n"
        "                    [Page] Detail"
    ) in text
    assert "            [Page] Tail" in text
    assert "        [Section group] H" in text


@pytest.mark.parametrize(
    "metadata",
    [
        [{"level": 0, "order": 0}, {"level": 1, "order": 0}],
        [{"level": 0}, {"level": 1}],
        [{"order": 0}, {"order": 1}],
    ],
)
def test_ambiguous_graph_metadata_is_flat_and_incomplete(metadata):
    """Unavailable or repeated order/level values cannot establish parent pages."""

    client = client_for(
        {
            ("GET", "sections/s/pages?pagelevel=true"): Response(
                {
                    "value": [
                        {"id": str(i), "title": str(i), **data}
                        for i, data in enumerate(metadata)
                    ]
                }
            )
        }
    )
    section = Node("Section", "S", "s")
    GraphBackend(client).read_pages(section)
    assert not section.complete
    assert len(section.children) == 2
    assert all(not node.children for node in section.children)


def test_duplicate_graph_page_ids_never_create_parents():
    """A page moved during pagination must not appear as its own parent."""
    client = client_for(
        {
            ("GET", "sections/s/pages?pagelevel=true"): Response(
                {
                    "value": [
                        {"id": "a", "title": "A", "level": 0, "order": 0},
                        {"id": "a", "title": "A", "level": 1, "order": 1},
                        {"id": "b", "title": "B", "level": 2, "order": 2},
                    ]
                }
            )
        }
    )
    section = Node("Section", "S", "s")
    GraphBackend(client).read_pages(section)
    assert not section.complete
    assert len(section.children) == 3
    assert all(not page.children for page in section.children)


def test_graph_section_failure_does_not_hide_other_notebooks():
    """A forbidden section must be marked, rather than silently omitted."""

    client = client_for(
        {
            ("GET", "notebooks/n/sections"): Response(
                {"value": [{"id": "s", "displayName": "Locked"}]}
            ),
            ("GET", "notebooks/n/sectionGroups"): Response({"value": []}),
            ("GET", "sections/s/pages?pagelevel=true"): Response({}, status=403),
            ("GET", "notebooks/m/sections"): Response({"value": []}),
            ("GET", "notebooks/m/sectionGroups"): Response({"value": []}),
        }
    )
    result = GraphBackend(client).scan(
        [Node("Notebook", "N", "n"), Node("Notebook", "M", "m")]
    )
    assert not result.complete
    assert "Locked" in render(result) and "HTTP 403" in render(result)
    assert "[Notebook] M" in render(result)


@pytest.mark.parametrize(
    "url",
    ["https://evil.example/path", "http://graph.microsoft.com/v1.0/me/onenote/pages"],
)
def test_graph_rejects_untrusted_next_link_before_sending_token(url):
    """Provider pagination must not send bearer credentials to another origin."""

    client = client_for(
        {("GET", "notebooks"): Response({"value": [], "@odata.nextLink": url})}
    )
    with pytest.raises(GraphError, match="URL"):
        client.collection("notebooks")
    assert len(client.session.calls) == 1


def test_graph_detects_pagination_loop():
    """A repeated next link must terminate rather than hang or duplicate entities."""

    url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
    client = client_for(
        {("GET", "notebooks"): Response({"value": [], "@odata.nextLink": url})}
    )
    with pytest.raises(GraphError, match="loop"):
        client.collection("notebooks")


def test_graph_retries_read_throttle_but_never_retries_creation():
    """Replaying an accepted POST can duplicate sample notes."""

    client = client_for(
        {
            ("GET", "notebooks"): [Response({}, 429), Response({"value": []})],
            ("POST", "notebooks/n/sectionGroups"): Response({}, 503),
        }
    )
    assert not client.collection("notebooks")
    with pytest.raises(GraphError, match="HTTP 503"):
        client.request(
            "POST", "notebooks/n/sectionGroups", json={"displayName": "Demo"}
        )
    assert len(client.session.calls) == 3


def test_page_creation_escapes_html_title():
    """User-visible titles must not become arbitrary HTML in create-page requests."""

    client = client_for({("POST", "sections/s/pages"): Response({"id": "p"}, 201)})
    page = GraphBackend(client).create_page(Node("Section", "S", "s"), "A < B & C")
    assert page.id == "p"
    html = client.session.calls[0][2]["data"].decode("utf-8")
    assert "<title>A &lt; B &amp; C</title>" in html


def test_auth_missing_client_id_is_actionable(monkeypatch):
    """Missing app setup must fail before launching a sign-in flow."""

    monkeypatch.delenv("ONENOTE_CLIENT_ID", raising=False)
    with pytest.raises(ValueError, match="client ID"):
        token_provider(None, write=False)


def test_auth_rejects_client_id_path_traversal():
    """An invalid app identifier must not become an arbitrary cache-file path."""
    with pytest.raises(ValueError, match="client ID"):
        token_provider("../../outside", write=False)


def test_forced_login_remembers_chosen_account(monkeypatch, tmp_path):
    """Choosing among cached accounts must keep subsequent requests on that account."""

    accounts = [{"username": "a@example.com"}, {"username": "b@example.com"}]

    class AuthApp:
        """An MSAL boundary with two cached accounts."""

        def get_accounts(self, username=None):
            """Return all accounts or one selected account."""
            return [a for a in accounts if not username or a["username"] == username]

        def acquire_token_interactive(self, **_kwargs):
            """Select the second account through the browser."""
            return {
                "access_token": "b-token",
                "id_token_claims": {"preferred_username": "b@example.com"},
            }

        def acquire_token_silent(self, _scopes, account):
            """Refresh under the chosen identity."""
            return {"access_token": account["username"]}

    monkeypatch.setattr(graph.Path, "home", lambda: tmp_path)
    monkeypatch.setattr(graph, "build_encrypted_persistence", lambda _path: object())
    monkeypatch.setattr(graph, "PersistedTokenCache", lambda _cache: object())
    monkeypatch.setattr(
        graph.msal, "PublicClientApplication", lambda *_args, **_kwargs: AuthApp()
    )
    token = token_provider(
        "00000000-0000-0000-0000-000000000001", write=False, login=True
    )
    assert token() == "b-token"
    assert token() == "b@example.com"


def test_snapshot_missing_collection_is_incomplete(tmp_path):
    """Recorded MCP data must not treat absent collections as empty notebooks."""

    path = tmp_path / "snapshot.json"
    path.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "source": "OneNote MCP",
                "captured_at": "2026-10-08",
                "collections": {
                    "notebooks": {
                        "complete": True,
                        "value": [{"id": "n", "displayName": "N"}],
                    }
                },
            }
        ),
        encoding="utf-8",
    )
    backend = GraphBackend(SnapshotClient.from_path(path))
    result = backend.scan(backend.list_notebooks())
    assert not result.complete
    assert "snapshot" in render(result).lower()
    assert "missing" in render(result).lower()


@pytest.mark.parametrize(
    "tenant", ["organizations", "consumers", "11111111-2222-3333-4444-555555555555"]
)
def test_auth_authority_and_cache_are_separated(tenant, monkeypatch, tmp_path):
    """School and personal sign-ins must not share an authority/cache boundary."""
    captured = {}

    def persistence(path):
        """Avoid OS keychain writes; capture the production cache destination."""
        captured["path"] = path
        return object()

    def application(_client_id, **kwargs):
        """Capture the MSAL construction boundary without authenticating."""
        captured.update(kwargs)
        return object()

    monkeypatch.setattr(graph.Path, "home", lambda: tmp_path)
    monkeypatch.setattr(graph, "build_encrypted_persistence", persistence)
    monkeypatch.setattr(graph, "PersistedTokenCache", lambda _cache: object())
    monkeypatch.setattr(graph.msal, "PublicClientApplication", application)
    token_provider("00000000-0000-0000-0000-000000000001", write=False, tenant=tenant)
    assert captured["authority"] == f"https://login.microsoftonline.com/{tenant}"
    assert captured["path"].endswith(
        f"00000000-0000-0000-0000-000000000001-{tenant}.bin"
    )


def test_auth_invalid_tenant_has_no_cache_side_effect(monkeypatch, tmp_path):
    """A tenant identifier must not be interpreted as a cache path or URL."""
    monkeypatch.setattr(graph.Path, "home", lambda: tmp_path)
    with pytest.raises(ValueError, match="tenant"):
        token_provider(
            "00000000-0000-0000-0000-000000000001", write=False, tenant="../../other"
        )
    assert not list(tmp_path.iterdir())


def test_graph_inventory_retains_notebook_web_reference():
    """A returned web reference permits checking identity against the supplied link."""
    url = "https://school.sharepoint.com/Doc.aspx?sourcedoc=%7Breference%7D"
    backend = GraphBackend(
        client_for(
            {
                ("GET", "notebooks"): Response(
                    {
                        "value": [
                            {
                                "id": "graph-notebook-id",
                                "displayName": "School",
                                "links": {"oneNoteWebUrl": {"href": url}},
                            }
                        ]
                    }
                )
            }
        )
    )
    node = backend.list_notebooks()[0]
    assert node.web_url == url
    assert node.id == "graph-notebook-id"


def test_school_export_requests_only_read_permission(monkeypatch, tmp_path, capsys):
    """Export must not request write permission when school sign-in is introduced."""
    captured = {}

    class AuthApp:
        """A first-time browser authentication boundary."""

        def get_accounts(self):
            """No cached identity exists yet."""
            return []

        def acquire_token_silent(self, *_args, **_kwargs):
            """No silent login should be attempted for an absent account."""
            pytest.fail("unexpected silent sign-in")

        def acquire_token_interactive(self, **kwargs):
            """Record scopes before returning a token without user data."""
            captured.update(kwargs)
            return {"access_token": "school-token"}

    monkeypatch.setattr(graph.Path, "home", lambda: tmp_path)
    monkeypatch.setattr(graph, "build_encrypted_persistence", lambda _path: object())
    monkeypatch.setattr(graph, "PersistedTokenCache", lambda _cache: object())
    monkeypatch.setattr(
        graph.msal, "PublicClientApplication", lambda *_args, **_kwargs: AuthApp()
    )
    token = token_provider(
        "00000000-0000-0000-0000-000000000001", write=False, tenant="organizations"
    )
    assert token() == "school-token"
    assert captured["scopes"] == ["https://graph.microsoft.com/Notes.Read"]
    assert "work/school" in capsys.readouterr().err


@pytest.mark.parametrize(
    "path",
    [
        "notebooks/../../drive/root",
        "https://graph.microsoft.com/v1.0/me/onenote/../drive/root",
        "https://graph.microsoft.com/v1.0/me/onenote/%2e%2e/drive/root",
        "https://graph.microsoft.com/v1.0/me/onenote/%2E%2E%2Fdrive/root",
        "https://graph.microsoft.com/v1.0/me/onenote/%5c..%5cdrive/root",
    ],
)
def test_graph_rejects_paths_that_escape_the_onenote_root(path):
    """URL normalization must not turn an accepted path into another Graph resource."""
    with pytest.raises(GraphError, match="URL"):
        GraphClient.validated_url(path)


@pytest.mark.parametrize(
    "payload",
    [[], {"value": None}, {"value": [1]}, {"value": [], "@odata.nextLink": 42}],
)
def test_graph_malformed_responses_fail_instead_of_silent_empty_export(payload):
    """Provider shape failures must never be interpreted as a complete inventory."""
    client = client_for({("GET", "notebooks"): Response(payload)})
    with pytest.raises(GraphError):
        client.collection("notebooks")


def test_graph_invalid_section_title_keeps_other_notebooks_incomplete_report(
    tmp_path, monkeypatch
):
    """Malformed title metadata must not abort otherwise useful export output."""
    client = client_for(
        {
            ("GET", "notebooks/n/sections"): Response(
                {"value": [{"id": "s", "displayName": 123}]}
            ),
            ("GET", "notebooks/n/sectionGroups"): Response({"value": []}),
            ("GET", "notebooks/m/sections"): Response({"value": []}),
            ("GET", "notebooks/m/sectionGroups"): Response({"value": []}),
        }
    )
    result = GraphBackend(client).scan(
        [Node("Notebook", "N", "n"), Node("Notebook", "M", "m")]
    )
    assert not result.complete
    text = render(result)
    assert "title" in text.lower()
    assert "[Notebook] M" in text
    client.session.routes[("GET", "notebooks")] = Response(
        {"value": [{"id": "n", "displayName": "N"}, {"id": "m", "displayName": "M"}]}
    )
    monkeypatch.setattr(cli, "make_backend", lambda _args: GraphBackend(client))
    output = tmp_path / "hierarchy.txt"
    assert cli.main(["export", "--output", str(output)]) == 2
    assert "[Notebook] M" in output.read_text(encoding="utf-8")


def test_graph_null_page_title_uses_untitled_fallback_without_render_failure():
    """Nullable titles should use the existing node fallback consistently."""
    client = client_for(
        {
            ("GET", "sections/s/pages?pagelevel=true"): Response(
                {"value": [{"id": "p", "title": None, "level": 0, "order": 0}]}
            )
        }
    )
    section = Node("Section", "S", "s")
    GraphBackend(client).read_pages(section)
    text = render(graph.ScanResult("graph", [Node("Notebook", "N", "n", [section])]))
    assert "[Page] (untitled)" in text
    assert section.complete
