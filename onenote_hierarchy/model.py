"""Shared tree, page nesting, selection, and atomic UTF-8 output."""

from dataclasses import dataclass, field
from datetime import datetime, timezone
import csv
import io
import os
from pathlib import Path
import tempfile


@dataclass
class Node:
    """A notebook, section group, section, or page with backend-specific ID."""

    kind: str
    title: str
    id: str
    children: list["Node"] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    web_url: str | None = None

    @property
    def complete(self) -> bool:
        """Whether this node and all descendants were read without uncertainty."""
        return not self.warnings and all(child.complete for child in self.children)


@dataclass
class Page:
    """A sibling page record; level is zero-based and order is optional."""

    id: str
    title: str
    level: int | None
    order: int | None = None


@dataclass
class ScanResult:
    """An export snapshot; completeness is derived from every node."""

    backend: str
    notebooks: list[Node]
    warnings: list[str] = field(default_factory=list)
    source: str = "live backend"
    captured_at: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )

    @property
    def complete(self) -> bool:
        """Whether every available notebook was enumerated successfully."""
        return not self.warnings and all(node.complete for node in self.notebooks)


def clean_title(value: str) -> str:
    """Keep a title on one line, including when it contains control characters."""
    return " ".join(value.split()) or "(untitled)"


def select_notebooks(
    notebooks: list[Node],
    name: str | list[str] | None = None,
    notebook_id: str | list[str] | None = None,
) -> list[Node]:
    """Resolve every requested notebook, rejecting missing or ambiguous names."""
    if name is None and notebook_id is None:
        return notebooks
    if name is not None and notebook_id is not None:
        raise ValueError("Select by names or IDs, not both.")
    values = notebook_id if notebook_id is not None else name
    values = [values] if isinstance(values, str) else values
    selected = []
    for value in values or []:
        matches = [
            node
            for node in notebooks
            if (
                node.id == value
                if notebook_id is not None
                else clean_title(node.title).casefold() == clean_title(value).casefold()
            )
        ]
        if not matches:
            raise ValueError(
                "Notebook not found. Run list and check the selected account."
            )
        if len(matches) > 1:
            raise ValueError("Several notebooks match. Use --notebook-id from list.")
        if not any(node.id == matches[0].id for node in selected):
            selected.append(matches[0])
    return selected


def walk_nodes(notebooks: list[Node]):
    """Yield every node, its parent, and labeled original path in export order."""

    def visit(node, parent_reference, path):
        nonlocal index
        index += 1
        reference = f"N{index:06d}"
        path = path + [f"[{node.kind}] {clean_title(node.title)}"]
        yield reference, parent_reference, node, path
        for child in node.children:
            yield from visit(child, reference, path)

    index = 0
    for notebook in notebooks:
        yield from visit(notebook, "", [])


def summary(result: ScanResult) -> str:
    """Count all hierarchy nodes, including subpages, without fetching content."""
    counts = {"Notebook": 0, "Section group": 0, "Section": 0, "Page": 0}
    for _, _, node, _ in walk_nodes(result.notebooks):
        counts[node.kind] += 1
    return (
        f"Notebooks: {counts['Notebook']}; Section groups: {counts['Section group']}; "
        f"Sections: {counts['Section']}; Pages: {counts['Page']}"
    )


def render_planning_csv(result: ScanResult) -> str:
    """Create a mapping worksheet; references identify nodes only in this export."""
    stream = io.StringIO(newline="")
    writer = csv.writer(stream)
    writer.writerow(
        [
            "reference",
            "parent_reference",
            "kind",
            "original_path",
            "proposed_path",
            "action",
            "reason",
            "status",
        ]
    )
    status = "COMPLETE" if result.complete else "INCOMPLETE"
    for reference, parent, node, path in walk_nodes(result.notebooks):
        writer.writerow(
            [
                reference,
                parent,
                node.kind,
                " / ".join(path),
                "",
                "",
                "",
                status,
            ]
        )
    return stream.getvalue()


def ensure_distinct_outputs(first: Path, second: Path):
    """Check native filename aliasing, including nonexistent case/Unicode aliases."""
    first, second = first.expanduser().resolve(), second.expanduser().resolve()
    if first == second:
        raise ValueError("Hierarchy and planning CSV require different output paths.")
    for path in (first, second):
        path.parent.mkdir(parents=True, exist_ok=True)
    if not first.parent.samefile(second.parent):
        return
    # Probe in an owned temporary directory on the destination filesystem. This
    # uses its actual filename semantics without touching either destination.
    with tempfile.TemporaryDirectory(
        dir=first.parent, prefix=".onenote-alias-"
    ) as probe:
        (Path(probe) / first.name).touch(exist_ok=False)
        if (Path(probe) / second.name).exists():
            raise ValueError(
                "Hierarchy and planning CSV refer to the same output file."
            )


def nest_pages(section: Node, pages: list[Page], *, sort_by_order: bool = False):
    """Build page parents only when every level and ordering value is reliable."""
    problem = (
        "duplicate page IDs; scan changed during enumeration"
        if len({page.id for page in pages}) != len(pages)
        else None
    )
    if sort_by_order and pages:
        orders = [page.order for page in pages]
        if any(
            not isinstance(order, int) or isinstance(order, bool) or order < 0
            for order in orders
        ):
            problem = "page order missing or invalid"
        elif len(set(orders)) != len(orders):
            problem = "page order ambiguous"
        else:
            pages = sorted(pages, key=lambda page: page.order)
    previous_level = -1
    for page in pages:
        if (
            not isinstance(page.level, int)
            or isinstance(page.level, bool)
            or page.level not in {0, 1, 2}
        ):
            problem = problem or "page level missing or invalid"
            break
        if page.level > previous_level + 1:
            problem = problem or "page nesting has no preceding parent"
            break
        previous_level = page.level
    if problem:
        section.warnings.append(f"{problem}; pages listed flat, nesting unavailable")
        section.children.extend(Node("Page", page.title, page.id) for page in pages)
        return
    stack: list[Node] = []
    for page in pages:
        node = Node("Page", page.title, page.id)
        stack = stack[: page.level]
        parent = stack[-1] if stack else section
        parent.children.append(node)
        stack.append(node)


def render(result: ScanResult, *, include_ids: bool = False) -> str:
    """Render the scan as labeled lines, exposing uncertainty and backend scope."""
    scope = (
        "open desktop notebooks"
        if result.backend == "desktop"
        else "cloud notebooks returned for the signed-in account"
    )
    lines = [
        f"OneNote hierarchy — backend: {result.backend}",
        f"Status: {'COMPLETE' if result.complete else 'INCOMPLETE'}",
        f"Scope: {scope}",
        f"Source: {clean_title(result.source)}",
        f"Captured at: {clean_title(result.captured_at)}",
        summary(result),
    ]
    lines.extend(f"Warning: {clean_title(warning)}" for warning in result.warnings)
    lines.append("")

    def visit(node: Node, depth: int):
        line = "    " * depth + f"[{node.kind}] {clean_title(node.title)}"
        if include_ids:
            line += f"  (ID: {clean_title(node.id)})"
        for warning in node.warnings:
            line += f"  [INCOMPLETE: {clean_title(warning)}]"
        lines.append(line)
        for child in node.children:
            visit(child, depth + 1)

    for notebook in result.notebooks:
        visit(notebook, 0)
    if not result.notebooks:
        lines.append("(No notebooks returned.)")
    return "\n".join(lines) + "\n"


def atomic_write(path: Path, text: str):
    """Replace a file only after its complete UTF-8 contents are written."""
    path = path.expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=path.parent,
            prefix=f".{path.name}.",
            delete=False,
        ) as stream:
            temporary = Path(stream.name)
            stream.write(text)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
