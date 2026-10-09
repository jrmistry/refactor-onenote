"""Exercise the public diagnostics and Windows boundary without live notebooks."""

import subprocess
import sys
from types import SimpleNamespace

import pytest

from onenote_hierarchy import cli, desktop

XML = """<Notebooks xmlns="http://schemas.microsoft.com/office/onenote/2013/onenote">
<Notebook ID="n" name="Test"><Section ID="s" name="Notes">
<Page ID="p" name="Guide" pageLevel="1"/>
</Section></Notebook></Notebooks>"""


class ComFailure(Exception):
    """Match pywintypes.com_error's exposed error metadata."""

    def __init__(self, code, inner=None):
        super().__init__("private note text must not appear")
        self.hresult = code
        self.excepinfo = (0, "OneNote", "private title", None, 0, inner)


def install_boundary(monkeypatch, *, xml=XML, connect_error=None, read_error=None):
    """Replace only Windows/COM IO; keep production adapter and XML parser."""

    class App:  # pylint: disable=too-few-public-methods
        """Expose the read-only inventory method required by doctor."""

        def GetHierarchy(
            self, start, scope, *, xsSchema
        ):  # pylint: disable=invalid-name
            """Reject attempts to modify notes or silently skip the inventory read."""
            assert (start, scope, xsSchema) == ("", 2, 2)
            if read_error:
                raise read_error
            return xml

    def connect(prog_id):
        assert prog_id == "OneNote.Application"
        if connect_error:
            raise connect_error
        return App()

    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setattr(
        desktop.importlib,
        "import_module",
        lambda name: SimpleNamespace(gencache=SimpleNamespace(EnsureDispatch=connect)),
    )


def test_doctor_reads_inventory_without_graph_or_credentials(monkeypatch, capsys):
    """Wrong backend selection or skipped inventory must not report success."""
    install_boundary(monkeypatch)
    monkeypatch.delenv("ONENOTE_CLIENT_ID", raising=False)
    for name in ("msal", "msal_extensions", "requests", "onenote_hierarchy.graph"):
        monkeypatch.setitem(sys.modules, name, None)
    assert cli.main(["doctor"]) == 0  # Windows default
    output = capsys.readouterr()
    assert "PASS" in output.out
    assert "1 notebook" in output.out
    assert not output.err


def test_doctor_empty_inventory_warns_but_connection_succeeds(monkeypatch, capsys):
    """No open notebooks is a usable connection, not a false connection failure."""
    install_boundary(monkeypatch, xml="<Notebooks/>")
    assert cli.main(["--backend", "desktop", "doctor"]) == 0
    assert "No notebooks" in capsys.readouterr().out


def test_doctor_checks_inventory_after_com_activation(monkeypatch, capsys):
    """COM activation alone cannot prove that this user's notebooks are readable."""
    install_boundary(monkeypatch, read_error=ComFailure(-2147024891))
    assert cli.main(["doctor"]) == 1
    output = capsys.readouterr()
    assert "access denied" in output.err.lower()
    assert "PASS" not in output.out
    assert "private" not in output.err


def test_doctor_rejects_graph_without_attempting_authentication(monkeypatch, capsys):
    """A desktop diagnostic must not unexpectedly start cloud sign-in."""
    monkeypatch.setattr(cli, "make_backend", lambda _args: pytest.fail("connected"))
    assert cli.main(["--backend", "graph", "doctor"]) == 1
    assert "--backend desktop" in capsys.readouterr().err


def test_doctor_reports_wrong_platform(capsys, monkeypatch):
    """The Mac limitation must explain which machine can run the workflow."""
    monkeypatch.setattr(desktop.sys, "platform", "darwin")
    assert cli.main(["--backend", "desktop", "doctor"]) == 1
    assert "Windows" in capsys.readouterr().err


def test_missing_pywin32_has_pip_install_action(monkeypatch):
    """Dependency failures must tell the user how to fix this Python environment."""
    monkeypatch.setattr(desktop.sys, "platform", "win32")

    def missing(_name):
        raise ImportError("missing win32com")

    monkeypatch.setattr(desktop.importlib, "import_module", missing)
    with pytest.raises(RuntimeError, match=r"python -m pip install.*desktop"):
        desktop.DesktopBackend()


def test_doctor_import_cache_failure_keeps_private_path_out_of_diagnostic(
    monkeypatch, capsys
):
    """pywin32 can initialize its cache during import, before EnsureDispatch."""
    monkeypatch.setattr(desktop.sys, "platform", "win32")

    def blocked_cache(_name):
        raise PermissionError(
            13, "Permission denied", "private-folder/gen_py/dicts.dat"
        )

    monkeypatch.setattr(desktop.importlib, "import_module", blocked_cache)
    assert cli.main(["--backend", "desktop", "doctor"]) == 1
    output = capsys.readouterr()
    assert "writable" in output.err
    assert "private-folder" not in output.err
    assert "PASS" not in output.out


@pytest.mark.parametrize(
    "error,expected",
    [
        (ComFailure(-2147221164), "desktop OneNote"),  # class not registered
        (ComFailure(-2147221005), "desktop OneNote"),  # invalid class string
        (ComFailure(-2147024891), "access denied"),
        (ComFailure(-2147352567, -2147024891), "access denied"),
        (PermissionError("private filename"), "writable"),
        (ComFailure(-2146959355), "Open desktop OneNote"),
        (ComFailure(-1), "0xFFFFFFFF"),
    ],
)
def test_connection_errors_are_actionable_without_private_details(
    error, expected, monkeypatch
):
    """Failures identify useful next steps without echoing COM payloads/titles."""
    install_boundary(monkeypatch, connect_error=error)
    with pytest.raises(RuntimeError) as caught:
        desktop.DesktopBackend()
    message = str(caught.value)
    assert expected.lower() in message.lower()
    assert "private" not in message


@pytest.mark.parametrize(
    "inner,expected",
    [
        (ComFailure(-2147024891), "access denied"),
        (PermissionError("private cache path"), "writable"),
    ],
)
def test_doctor_preserves_wrapped_failure_without_private_payload(
    inner, expected, monkeypatch, capsys
):
    """pywin32 wrapper failures must not hide an underlying Windows error."""
    error = TypeError("private wrapper payload")
    error.__context__ = inner
    install_boundary(monkeypatch, connect_error=error)
    assert cli.main(["doctor"]) == 1
    output = capsys.readouterr()
    assert "TypeError" in output.err
    assert type(inner).__name__ in output.err
    assert expected in output.err
    if isinstance(inner, ComFailure):
        assert "0x80070005" in output.err
    assert "private" not in output.err
    assert "PASS" not in output.out


def test_non_com_failure_reports_type_without_promising_missing_hresult():
    """An ordinary Python failure must be diagnosable without exposing its text."""
    message = desktop.desktop_error(TypeError("private wrapper payload"), "connection")
    assert "TypeError" in message
    assert "no HRESULT available" in message
    assert "HRESULT shown here" not in message
    assert "private" not in message


def test_inventory_error_in_list_has_actionable_message(monkeypatch):
    """The normal list command needs the same context as diagnostics."""
    install_boundary(monkeypatch, read_error=ComFailure(-2147024891))
    with pytest.raises(RuntimeError, match="access denied"):
        desktop.DesktopBackend().list_notebooks()


def test_export_read_failure_preserves_incomplete_status(tmp_path, monkeypatch, capsys):
    """A per-notebook COM denial must retain warnings and produce exit 2."""

    class App:  # pylint: disable=too-few-public-methods
        """Return inventory, then deny the detailed hierarchy scan."""

        def GetHierarchy(
            self, start, scope, *, xsSchema
        ):  # pylint: disable=invalid-name
            """Exercise actual export CLI selection, adapter, parser and renderer."""
            assert xsSchema == 2
            if (start, scope) == ("", 2):
                return XML
            assert (start, scope) == ("n", 4)
            raise ComFailure(-2147024891)

    monkeypatch.setattr(
        cli, "make_backend", lambda _args: desktop.DesktopBackend(App())
    )
    path = tmp_path / "tree.txt"
    assert cli.main(["--backend", "desktop", "export", "--output", str(path)]) == 2
    report = path.read_text(encoding="utf-8")
    assert "INCOMPLETE" in report
    assert "access denied" in report
    assert "private" not in report
    assert "incomplete" in capsys.readouterr().out


def test_windows_cli_help_and_doctor_without_optional_cloud_modules():
    """Run entry-point parsing in a fresh Windows-shaped interpreter boundary."""
    code = """
import sys
from types import SimpleNamespace
for name in ("msal", "msal_extensions", "requests"):
    sys.modules[name] = None
from onenote_hierarchy import cli, desktop
class App:
    def GetHierarchy(self, start, scope, *, xsSchema):
        assert (start, scope, xsSchema) == ("", 2, 2)
        return "<Notebooks/>"
desktop.importlib.import_module = lambda name: SimpleNamespace(
    gencache=SimpleNamespace(EnsureDispatch=lambda name: App()))
sys.platform = "win32"
assert cli.parser().parse_args(["doctor"]).backend == "desktop"
raise SystemExit(cli.main(["doctor"]))
"""
    result = subprocess.run(
        [sys.executable, "-c", code], text=True, capture_output=True, check=False
    )
    assert result.returncode == 0, result.stderr
    assert "No notebooks" in result.stdout


def test_successful_desktop_export_reads_selected_notebook_and_writes_pair(
    tmp_path, monkeypatch, capsys
):
    """Real CLI/parser/renderer retain observed nesting without page-body calls."""
    nested = XML.replace(
        "</Section>",
        '<Page ID="q" name="子頁" pageLevel="2"/></Section>',
    )
    calls = []

    class App:  # pylint: disable=too-few-public-methods
        """The only supported COM call is read-only GetHierarchy."""

        def GetHierarchy(
            self, start, scope, *, xsSchema
        ):  # pylint: disable=invalid-name
            """Inventory and detail scopes are both required."""
            calls.append((start, scope, xsSchema))
            return nested

    monkeypatch.setattr(
        cli, "make_backend", lambda _args: desktop.DesktopBackend(App())
    )
    report = tmp_path / "hierarchy.txt"
    planning = tmp_path / "planning.csv"
    assert (
        cli.main(
            [
                "--backend",
                "desktop",
                "export",
                "--notebook",
                "Test",
                "--include-ids",
                "--output",
                str(report),
                "--planning-csv",
                str(planning),
            ]
        )
        == 0
    )
    assert calls == [("", 2, 2), ("n", 4, 2)]
    text = report.read_text(encoding="utf-8")
    assert "Status: COMPLETE" in text
    assert "            [Page] 子頁" in text
    assert "子頁" in planning.read_text(encoding="utf-8")
    assert "Saved complete" in capsys.readouterr().out
