"""Test safe demo creation and public command behavior."""

import json

import pytest

from onenote_hierarchy import cli
from onenote_hierarchy.cli import main
from onenote_hierarchy.demo import create_demo
from onenote_hierarchy.model import Node, ScanResult


class Backend:
    """A local backend with real demo journal side effects."""

    name = "graph"

    def __init__(self, existing=False, fail=False):
        self.existing = existing
        self.fail = fail
        self.created = []

    def list_groups(self, _parent):
        """Return potential name collisions."""

        return [Node("Section group", "Hierarchy Demo", "old")] if self.existing else []

    def create_group(self, _parent, title):
        """Create test nodes with unique IDs."""

        node = Node("Section group", title, f"g{len(self.created)}")
        self.created.append(node)
        return node

    def create_section(self, _parent, title):
        """Create a section, optionally simulating an uncertain HTTP result."""

        if self.fail:
            raise RuntimeError("connection lost")
        node = Node("Section", title, f"s{len(self.created)}")
        self.created.append(node)
        return node

    def create_page(self, _parent, title):
        """Create a page node."""

        node = Node("Page", title, f"p{len(self.created)}")
        self.created.append(node)
        return node


def test_demo_records_created_ids_and_requires_ui_indentation(tmp_path):
    """Graph cannot report nested pages before a UI step actually sets their level."""

    manifest = tmp_path / "demo.json"
    result = create_demo(Backend(), Node("Notebook", "School", "n"), manifest)
    data = json.loads(manifest.read_text(encoding="utf-8"))
    assert data["status"] == "awaiting_ui_indentation"
    assert result["status"] == data["status"]
    assert len(data["items"]) == 9  # two groups, two sections, five pages
    assert [
        (item["title"], item.get("desired_level"))
        for item in data["items"]
        if item["kind"] == "Page"
    ] == [
        ("Guide", 0),
        ("Installing", 1),
        ("Requirements", 2),
        ("First example", 0),
        ("Notes", 0),
    ]


def test_duplicate_demo_refused_without_writes(tmp_path):
    """Rerunning a sample must not duplicate existing notes."""

    backend = Backend(existing=True)
    with pytest.raises(ValueError, match="already exists"):
        create_demo(backend, Node("Notebook", "School", "n"), tmp_path / "demo.json")
    assert not backend.created


def test_demo_failure_retains_ids_without_retrying(tmp_path):
    """Failed demo creation retains a journal and stops uncertain writes."""

    path = tmp_path / "demo.json"
    backend = Backend(fail=True)
    with pytest.raises(RuntimeError):
        create_demo(backend, Node("Notebook", "School", "n"), path)
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["status"] == "failed_review_required"
    assert data["items"][0]["id"] == "g0"
    assert len(backend.created) == 1


@pytest.mark.parametrize("failure", ["title", "levels"])
def test_desktop_update_failure_journals_exact_pending_operation(tmp_path, failure):
    """Created IDs and uncertain COM updates must both survive in the journal."""

    class DesktopWriter(Backend):
        """Simulate failure during a title or level update."""

        name = "desktop"

        def finish_page(self, _page):
            """Fail a title update if requested."""
            if failure == "title":
                raise RuntimeError("title update lost")

        def set_page_levels(self, _section, _pages):
            """Fail the nesting update if requested."""
            if failure == "levels":
                raise RuntimeError("level update lost")

    path = tmp_path / "demo.json"
    with pytest.raises(RuntimeError):
        create_demo(DesktopWriter(), Node("Notebook", "School", "n"), path)
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["status"] == "failed_review_required"
    assert data["pending"]["operation"] == (
        "set_page_title" if failure == "title" else "set_page_levels"
    )
    assert any(item["kind"] == "Page" and item["id"] for item in data["items"])


def test_cli_exports_incomplete_scan_with_nonzero_status(tmp_path, monkeypatch, capsys):
    """A written file must not imply complete coverage when a section is unavailable."""

    class Reader:
        """Provide a real tree with an unavailable node."""

        def list_notebooks(self):
            """List one notebook."""
            return [Node("Notebook", "N", "n")]

        def scan(self, notebooks):
            """Mark the selected notebook incomplete."""
            notebooks[0].warnings.append("unavailable")
            return ScanResult("graph", notebooks)

    monkeypatch.setattr(cli, "make_backend", lambda *_args: Reader())
    path = tmp_path / "tree.txt"
    assert cli.main(["export", "--output", str(path)]) == 2
    assert "Status: INCOMPLETE" in path.read_text(encoding="utf-8")
    assert "incomplete" in capsys.readouterr().out.lower()


def test_cli_requires_explicit_demo_destination(tmp_path):
    """No write command may fall back to the default notebook."""

    with pytest.raises(SystemExit) as exc:
        main(["create-demo", "--manifest", str(tmp_path / "demo.json")])
    assert exc.value.code == 2


def test_cli_backend_error_leaves_existing_export(tmp_path, monkeypatch):
    """Authentication failure must not replace a successful prior export."""

    path = tmp_path / "tree.txt"
    path.write_text("old", encoding="utf-8")

    def fail_backend(*_args):
        raise RuntimeError("not signed in")

    monkeypatch.setattr(cli, "make_backend", fail_backend)
    assert cli.main(["export", "--output", str(path)]) == 1
    assert path.read_text(encoding="utf-8") == "old"


def test_cli_exports_mcp_snapshot_without_sign_in(tmp_path, monkeypatch):
    """A complete captured MCP hierarchy should use the real Python export pipeline."""

    snapshot = tmp_path / "snapshot.json"
    snapshot.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "source": "OneNote MCP",
                "captured_at": "2026-10-08",
                "collections": {
                    "notebooks": {
                        "complete": True,
                        "value": [{"id": "n", "displayName": "N"}],
                    },
                    "notebooks/n/sections": {"complete": True, "value": []},
                    "notebooks/n/sectionGroups": {"complete": True, "value": []},
                },
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.delenv("ONENOTE_CLIENT_ID", raising=False)
    path = tmp_path / "tree.txt"
    assert (
        main(
            [
                "--backend",
                "graph",
                "export",
                "--snapshot",
                str(snapshot),
                "--output",
                str(path),
            ]
        )
        == 0
    )
    text = path.read_text(encoding="utf-8")
    assert "[Notebook] N" in text
    assert "OneNote MCP metadata snapshot" in text


@pytest.mark.parametrize("empty", [False, True])
def test_cli_list_reports_inventory_and_empty_state(empty, monkeypatch, capsys):
    """Inventory output supports selecting IDs and checking returned web references."""

    class Reader:  # pylint: disable=too-few-public-methods
        """Supply only the external inventory boundary."""

        def list_notebooks(self):
            """Return a Unicode title and web reference, or an empty account."""
            return (
                []
                if empty
                else [
                    Node(
                        "Notebook", "研究", "n", web_url="https://example.com/notebook"
                    )
                ]
            )

    monkeypatch.setattr(cli, "make_backend", lambda *_args: Reader())
    assert main(["list"]) == 0
    output = capsys.readouterr().out
    if empty:
        assert "No notebooks returned" in output
    else:
        assert "研究\n  ID: n" in output
        assert "Web: https://example.com/notebook" in output


@pytest.mark.parametrize("backend_name", ["graph", "desktop"])
def test_cli_demo_success_records_journal_and_next_step(
    backend_name, tmp_path, monkeypatch, capsys
):
    """Distinguish completed desktop updates from the required Graph UI work."""

    class Writer(Backend):
        """Supply notebook IO with successful desktop update boundaries."""

        name = backend_name

        def list_notebooks(self):
            """Expose the explicitly selected destination."""
            return [Node("Notebook", "Disposable Test", "n")]

        def finish_page(self, _page):
            """Successful title update boundary."""

        def set_page_levels(self, _section, _pages):
            """Successful nesting update boundary."""

    monkeypatch.setattr(cli, "make_backend", lambda *_args: Writer())
    path = tmp_path / "demo.json"
    assert (
        main(["create-demo", "--notebook", "Disposable Test", "--manifest", str(path)])
        == 0
    )
    data = json.loads(path.read_text(encoding="utf-8"))
    output = capsys.readouterr().out
    assert len(data["items"]) == 9
    assert data["status"] == (
        "awaiting_ui_indentation" if backend_name == "graph" else "created"
    )
    assert ("indent Installing" if backend_name == "graph" else "nesting set") in output
