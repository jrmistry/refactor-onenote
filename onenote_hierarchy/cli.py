"""Command-line entry points for listing, scanning, and demo creation."""

import argparse
import os
from pathlib import Path
import sys

from .demo import create_demo
from .desktop import DesktopBackend
from .model import (
    atomic_write,
    clean_title,
    ensure_distinct_outputs,
    render,
    render_planning_csv,
    select_notebooks,
    summary,
)


def parser() -> argparse.ArgumentParser:
    """Define explicit write destinations and the small public command interface."""
    result = argparse.ArgumentParser(
        description="Export OneNote titles and nesting to a text tree."
    )
    result.add_argument(
        "--backend",
        choices=["graph", "desktop"],
        default="desktop" if sys.platform == "win32" else "graph",
    )
    result.add_argument(
        "--client-id",
        default=os.getenv("ONENOTE_CLIENT_ID"),
        help="Graph public-client application ID (or ONENOTE_CLIENT_ID)",
    )
    result.add_argument(
        "--login", action="store_true", help="Choose a Graph account in the browser"
    )
    result.add_argument(
        "--tenant",
        default=os.getenv("ONENOTE_TENANT", "consumers"),
        help="Graph authority: consumers, organizations, or tenant UUID "
        "(ONENOTE_TENANT)",
    )
    commands = result.add_subparsers(dest="command", required=True)
    commands.add_parser(
        "doctor", help="Check the Windows desktop connection without editing notes"
    )
    commands.add_parser("list", help="List notebook names and backend-specific IDs")
    export = commands.add_parser(
        "export", help="Export all available notebooks or selected notebooks"
    )
    demo = commands.add_parser(
        "create-demo", help="Create an isolated Hierarchy Demo test group"
    )
    for command in [export, demo]:
        selection = command.add_mutually_exclusive_group(required=command is demo)
        selection.add_argument(
            "--notebook",
            action="append" if command is export else "store",
            help="Exact notebook name, case-insensitive; repeat for export",
        )
        selection.add_argument(
            "--notebook-id",
            action="append" if command is export else "store",
            help="Exact ID from list; repeat for export",
        )
    export.add_argument("--output", type=Path, default=Path("exports/hierarchy.txt"))
    export.add_argument("--include-ids", action="store_true")
    export.add_argument(
        "--planning-csv", type=Path, help="Optional consolidation worksheet"
    )
    export.add_argument(
        "--snapshot",
        type=Path,
        help="Export a recorded OneNote MCP metadata snapshot without signing in",
    )
    demo.add_argument(
        "--manifest",
        type=Path,
        default=Path("exports/demo-manifest.json"),
        help="New journal path; existing files are never replaced",
    )
    return result


def make_backend(args):
    """Connect only to the selected backend, with command-specific Graph scopes."""
    if args.backend == "desktop":
        return DesktopBackend()
    try:
        # Cloud packages are optional; Windows export/help never imports them.
        from .graph import (  # pylint: disable=import-outside-toplevel
            GraphBackend,
            GraphClient,
            SnapshotClient,
            token_provider,
        )
    except ImportError as exc:
        raise RuntimeError("Graph packages missing. Install '.[graph]'.") from exc
    if args.command == "export" and args.snapshot:
        if args.backend != "graph":
            raise ValueError("Use --backend graph when exporting an MCP snapshot.")
        return GraphBackend(SnapshotClient.from_path(args.snapshot))
    token = token_provider(
        args.client_id,
        write=args.command == "create-demo",
        login=args.login,
        tenant=args.tenant,
    )
    return GraphBackend(GraphClient(token))


def validate_args(args):
    """Reject incompatible requests before any backend connection."""
    if args.command == "doctor" and args.backend != "desktop":
        raise ValueError(
            "doctor checks the Windows connection. Use --backend desktop doctor "
            "on the Windows laptop; it does not configure Graph authentication."
        )
    if args.command == "export" and args.snapshot and args.backend != "graph":
        raise ValueError("Use --backend graph when exporting an MCP snapshot.")
    if (
        args.command == "export"
        and args.planning_csv
        and args.planning_csv.expanduser().resolve()
        == args.output.expanduser().resolve()
    ):
        raise ValueError("Hierarchy and planning CSV require different output paths.")


def report_diagnostics(notebooks):
    """Show success only after the desktop connection and inventory read succeeded."""
    print("PASS: Windows, pywin32, and OneNote COM connection.")
    print(f"PASS: inventory readable ({len(notebooks)} notebook(s) open).")
    if not notebooks:
        print("WARNING: No notebooks are open. Open and sync them in OneNote.")
    print("Next: run list, then export. No notes were edited.")


def main(argv: list[str] | None = None) -> int:
    """Return 0 for success, 1 for failure, or 2 for an incomplete export."""
    args = parser().parse_args(argv)
    try:
        validate_args(args)
        backend = make_backend(args)
        notebooks = backend.list_notebooks()
        if args.command == "doctor":
            report_diagnostics(notebooks)
            return 0
        if args.command == "list":
            for notebook in notebooks:
                print(f"{clean_title(notebook.title)}\n  ID: {notebook.id}")
                if notebook.web_url:
                    print(f"  Web: {clean_title(notebook.web_url)}")
            if not notebooks:
                print(
                    "No notebooks returned. Check the account "
                    "or open notebooks in desktop OneNote."
                )
            return 0
        selected = select_notebooks(notebooks, args.notebook, args.notebook_id)
        if args.command == "create-demo":
            manifest = create_demo(backend, selected[0], args.manifest)
            print(f"Created sample. Journal: {args.manifest.expanduser().resolve()}")
            if manifest["status"] == "awaiting_ui_indentation":
                print(
                    "Next in OneNote: order Basics as Guide, Installing, Requirements; "
                    "indent Installing once and Requirements twice, "
                    "then sync and export."
                )
            else:
                print("Sample nesting set. Export to verify the observed hierarchy.")
            return 0
        snapshot = backend.scan(selected)
        if args.planning_csv:
            ensure_distinct_outputs(args.output, args.planning_csv)
        atomic_write(args.output, render(snapshot, include_ids=args.include_ids))
        if args.planning_csv:
            atomic_write(args.planning_csv, render_planning_csv(snapshot))
            print(f"Saved planning CSV: {args.planning_csv.expanduser().resolve()}")
        print(summary(snapshot))
        print(
            f"Saved {'complete' if snapshot.complete else 'incomplete'} hierarchy: "
            f"{args.output.expanduser().resolve()}"
        )
        return 0 if snapshot.complete else 2
    # Dynamic COM exception types must also produce a CLI failure status.
    except Exception as exc:  # pylint: disable=broad-exception-caught
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
