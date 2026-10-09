"""Behavior tests for multi-notebook consolidation inventories."""

import csv
import io
import json
from pathlib import Path
import subprocess
import sys

import pytest

from onenote_hierarchy import cli, model
from onenote_hierarchy.model import Node, ScanResult


def inventory():
    """Two notebooks with duplicate paths and nested Unicode page titles."""
    page = Node("Page", "Café", "p", [Node("Page", "Details", "sub")])
    group = Node(
        "Section group",
        "Projects",
        "g",
        [Node("Section", "Same", "s", [page, Node("Page", "Café", "p2")])],
    )
    return [Node("Notebook", "A", "a", [group]), Node("Notebook", "B", "b")]


def test_report_counts_nested_nodes_and_capture_time():
    """A misleading count or absent capture time makes a migration hard to audit."""
    result = ScanResult("desktop", inventory())
    output = model.render(result)
    assert "Captured at:" in output
    assert "Notebooks: 2; Section groups: 1; Sections: 1; Pages: 3" in output


def test_csv_distinguishes_duplicate_paths_and_preserves_parentage():
    """Duplicate page titles must each receive a separate planning reference."""
    output = model.render_planning_csv(ScanResult("desktop", inventory()))
    rows = list(csv.DictReader(io.StringIO(output)))
    assert len(rows) == 7
    assert len({row["reference"] for row in rows}) == 7
    assert rows[4]["parent_reference"] == rows[3]["reference"]
    assert rows[3]["original_path"] == rows[5]["original_path"]
    assert "Café" in rows[3]["original_path"]
    assert all(
        not row["proposed_path"] and not row["action"] and not row["reason"]
        for row in rows
    )
    assert all(row["status"] == "COMPLETE" for row in rows)


@pytest.mark.parametrize(
    "option,values", [("--notebook", ["A", "B"]), ("--notebook-id", ["a", "b"])]
)
def test_cli_scans_every_selected_notebook(option, values, tmp_path, monkeypatch):
    """Repeated selectors must not silently export only the last notebook."""

    class Reader:
        """A metadata boundary; real selection and serialization run normally."""

        def list_notebooks(self):
            """Supply a selected pair and an unrelated notebook."""
            return inventory() + [Node("Notebook", "Other", "other")]

        def scan(self, notebooks):
            """Expose which notebooks the CLI selected."""
            return ScanResult("desktop", notebooks)

    monkeypatch.setattr(cli, "make_backend", lambda _args: Reader())
    tree, table = tmp_path / "tree.txt", tmp_path / "plan.csv"
    args = ["export", "--output", str(tree), "--planning-csv", str(table)]
    for value in values:
        args.extend([option, value])
    assert cli.main(args) == 0
    text = tree.read_text(encoding="utf-8")
    assert "[Notebook] A" in text and "[Notebook] B" in text
    assert "Other" not in text
    assert len(list(csv.DictReader(table.open(encoding="utf-8")))) == 7


def test_missing_selection_prevents_scan_and_file_changes(tmp_path, monkeypatch):
    """A typo must not silently yield a partial selected set."""

    class Reader:  # pylint: disable=too-few-public-methods
        """Selection fails before scan could be invoked."""

        def list_notebooks(self):
            """Return the known notebooks."""
            return inventory()

    monkeypatch.setattr(cli, "make_backend", lambda _args: Reader())
    target = tmp_path / "tree.txt"
    target.write_text("old", encoding="utf-8")
    assert (
        cli.main(
            [
                "export",
                "--notebook",
                "A",
                "--notebook",
                "Missing",
                "--output",
                str(target),
            ]
        )
        == 1
    )
    assert target.read_text(encoding="utf-8") == "old"


def test_same_tree_and_csv_destination_is_rejected(tmp_path, monkeypatch):
    """The CSV must not overwrite the tree when both paths alias the same file."""
    monkeypatch.setattr(cli, "make_backend", lambda _args: pytest.fail("connected"))
    path = str(tmp_path / "output.txt")
    assert cli.main(["export", "--output", path, "--planning-csv", path]) == 1


def test_csv_exposes_incomplete_scan():
    """Planning references must retain incomplete-scan uncertainty."""
    result = ScanResult("desktop", inventory(), warnings=["unavailable"])
    rows = list(csv.DictReader(io.StringIO(model.render_planning_csv(result))))
    assert all(row["status"] == "INCOMPLETE" for row in rows)


def test_help_and_desktop_selection_do_not_import_graph_dependencies():
    """Windows setup must work when optional cloud packages are unavailable."""
    code = """
import sys
# Make the non-Windows error assertion deterministic on Windows runners too.
sys.platform = "darwin"
for name in ("msal", "msal_extensions", "requests"):
    sys.modules[name] = None
from onenote_hierarchy.cli import main
raise SystemExit(main(["--backend", "desktop", "list"]))
"""
    result = subprocess.run(
        [sys.executable, "-c", code], capture_output=True, text=True, check=False
    )
    assert result.returncode == 1
    assert "requires Windows" in result.stderr
    assert "ModuleNotFoundError" not in result.stderr


def test_packaging_has_minimal_desktop_extra():
    """The installed desktop dependency set must not pull in cloud libraries."""
    import tomllib  # pylint: disable=import-outside-toplevel

    config = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))
    project = config["project"]
    assert not project.get("dependencies")
    assert len(project["optional-dependencies"]["desktop"]) == 1
    assert project["optional-dependencies"]["desktop"][0].startswith("pywin32")


def test_snapshot_retains_original_capture_time(tmp_path):
    """Re-rendering old data must not present it as a fresh notebook capture."""
    path = tmp_path / "snapshot.json"
    path.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "captured_at": "2025-01-01",
                "collections": {"notebooks": {"complete": True, "value": []}},
            }
        ),
        encoding="utf-8",
    )
    target = tmp_path / "tree.txt"
    assert (
        cli.main(
            [
                "--backend",
                "graph",
                "export",
                "--snapshot",
                str(path),
                "--output",
                str(target),
            ]
        )
        == 0
    )
    assert "Captured at: 2025-01-01" in target.read_text(encoding="utf-8")


def test_tenant_environment_and_cli_override(monkeypatch):
    """A command-specific school tenant must override the environment setting."""
    monkeypatch.setenv("ONENOTE_TENANT", "organizations")
    assert cli.parser().parse_args(["list"]).tenant == "organizations"
    assert (
        cli.parser().parse_args(["--tenant", "consumers", "list"]).tenant == "consumers"
    )


def test_new_case_alias_destinations_do_not_overwrite_each_other(tmp_path, monkeypatch):
    """A filesystem alias must fail before either export file is committed."""
    probe = tmp_path / "case-probe"
    probe.touch()
    insensitive = (tmp_path / "CASE-PROBE").exists()
    probe.unlink()
    if not insensitive:
        pytest.skip("Requires a case-insensitive filesystem")

    class Reader:
        """Return a complete inventory without reading real notebooks."""

        def list_notebooks(self):
            """Return the inventory."""
            return inventory()

        def scan(self, notebooks):
            """Produce the selected tree."""
            return ScanResult("desktop", notebooks)

    monkeypatch.setattr(cli, "make_backend", lambda _args: Reader())
    target = tmp_path / "Report.txt"
    assert (
        cli.main(
            [
                "export",
                "--output",
                str(target),
                "--planning-csv",
                str(tmp_path / "report.txt"),
            ]
        )
        == 1
    )
    assert not target.exists()
    assert not list(tmp_path.iterdir())


def test_csv_completeness_work_scales_linearly():
    """Adding rows must not repeatedly traverse the entire inventory."""
    visits = 0

    class CountingNode(Node):  # pylint: disable=too-few-public-methods
        """Count hierarchy visits, avoiding machine-dependent timing assertions."""

        @property
        def complete(self):
            nonlocal visits
            visits += 1
            return super().complete

    notebook = CountingNode(
        "Notebook",
        "N",
        "n",
        [CountingNode("Page", str(index), str(index)) for index in range(100)],
    )
    rows = list(
        csv.DictReader(
            io.StringIO(model.render_planning_csv(ScanResult("desktop", [notebook])))
        )
    )
    assert len(rows) == 101
    assert visits <= 2 * len(rows)
