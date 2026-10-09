"""Windows OneNote COM adapter and portable XML parser."""

from datetime import datetime, timezone
import importlib
import sys
import xml.etree.ElementTree as ET

from .model import Node, Page, ScanResult, nest_pages

SCHEMA = 2  # xs2013
NAMESPACE = "http://schemas.microsoft.com/office/onenote/2013/onenote"
LABELS = {
    "Notebook": "Notebook",
    "SectionGroup": "Section group",
    "Section": "Section",
    "Page": "Page",
}


def desktop_error(exc: Exception, operation: str) -> str:
    """Explain Windows failures without echoing note text or raw COM payloads."""
    code = getattr(exc, "hresult", None)
    details = getattr(exc, "excepinfo", None)
    if details and len(details) > 5 and details[5]:
        code = details[5]
    code = code & 0xFFFFFFFF if isinstance(code, int) else None
    label = f" (HRESULT 0x{code:08X})" if code is not None else ""
    prefix = f"OneNote desktop {operation} failed{label}: "
    if code == 0x80070005:
        return prefix + (
            "access denied. Use the same ordinary Windows user session as OneNote. "
            "If workplace policy blocks automation, ask IT; the exporter cannot "
            "bypass that restriction."
        )
    if isinstance(exc, PermissionError):
        return prefix + (
            "Python could not write its automation cache. Extract the project into "
            "a user-writable folder and create .venv there. If it remains blocked, "
            "ask IT which local folder is permitted."
        )
    if code in {0x80040154, 0x800401F3}:
        return prefix + (
            "desktop OneNote is unavailable to COM. Open the installed Microsoft "
            "365 desktop OneNote app, not the browser or old Windows 10 Store app. "
            "If that app is absent or cannot connect, ask IT to check its installation."
        )
    return prefix + (
        "Open desktop OneNote in this Windows user session, let it finish syncing, "
        "then retry. If it still fails, give IT the HRESULT shown here. "
        "No administrator or registry changes are part of this setup."
    )


def kind(element: ET.Element) -> str:
    """Return the local XML element name."""
    return element.tag.rsplit("}", 1)[-1]


def is_true(element: ET.Element, attribute: str) -> bool:
    """Interpret OneNote's boolean schema attributes."""
    return element.get(attribute, "").lower() in {"true", "1"}


def parse_node(element: ET.Element) -> Node:
    """Parse visible hierarchy metadata, including sibling page indentation."""
    node = Node(
        LABELS[kind(element)],
        element.get("nickname") or element.get("name") or "(untitled)",
        element.get("ID", ""),
    )
    if kind(element) == "Section":
        if is_true(element, "locked"):
            node.warnings.append("LOCKED; page listing may be incomplete")
        if element.get("areAllPagesAvailable", "").lower() in {"false", "0"}:
            node.warnings.append("some pages are unavailable")
    pages = []
    for child in element:
        if kind(child) not in LABELS:
            continue
        if is_true(child, "isRecycleBin") or is_true(child, "isDeletedPages"):
            continue
        if kind(child) == "Page":
            try:
                # A legacy isSubPage boolean cannot distinguish both subpage depths.
                level = int(child.get("pageLevel", "")) - 1
            except ValueError:
                level = None
            pages.append(Page(child.get("ID", ""), child.get("name", ""), level))
        else:
            node.children.append(parse_node(child))
    nest_pages(node, pages)
    return node


def parse_hierarchy(xml: str) -> ScanResult:
    """Parse a GetHierarchy response without requiring Windows."""
    root = ET.fromstring(xml)
    notebooks = [
        parse_node(element)
        for element in root.iter()
        if kind(element) == "Notebook" and not is_true(element, "isRecycleBin")
    ]
    return ScanResult("desktop", notebooks)


class DesktopBackend:
    """Read hierarchy or add isolated demo items through desktop OneNote."""

    name = "desktop"

    def __init__(self, app=None):
        if app is None:
            if sys.platform != "win32":
                raise RuntimeError(
                    "The desktop backend requires Windows and desktop OneNote. "
                    "Run this command on the Windows laptop with the notebooks open; "
                    "the Mac cannot use the Windows COM connection."
                )
            try:
                client = importlib.import_module("win32com.client")
            except ImportError as exc:
                raise RuntimeError(
                    "pywin32 could not be loaded in this Python environment. "
                    "From the extracted project folder run "
                    'python -m pip install ".[desktop]" '
                    "using the same .venv Python as this command."
                ) from exc
            except Exception as exc:  # pylint: disable=broad-exception-caught
                # pywin32 may initialize its generated-code cache during import.
                raise RuntimeError(desktop_error(exc, "dependency loading")) from exc
            try:
                app = client.gencache.EnsureDispatch("OneNote.Application")
            except Exception as exc:  # pylint: disable=broad-exception-caught
                raise RuntimeError(desktop_error(exc, "connection")) from exc
        self.app = app

    def list_notebooks(self) -> list[Node]:
        """List notebooks currently open in desktop OneNote."""
        try:
            xml = self.app.GetHierarchy("", 2, xsSchema=SCHEMA)
        except Exception as exc:  # pylint: disable=broad-exception-caught
            raise RuntimeError(desktop_error(exc, "inventory read")) from exc
        return parse_hierarchy(xml).notebooks

    def scan(self, notebooks: list[Node]) -> ScanResult:
        """Read selected notebooks separately so one failure cannot hide others."""
        result = ScanResult(self.name, [])
        for notebook in notebooks:
            try:
                parsed = parse_hierarchy(
                    self.app.GetHierarchy(notebook.id, 4, xsSchema=SCHEMA)
                )
                matches = [node for node in parsed.notebooks if node.id == notebook.id]
                if len(matches) != 1:
                    raise ValueError("Expected notebook missing in desktop response")
                result.notebooks.extend(matches)
            # COM error types depend on the installed Windows automation runtime.
            except Exception as exc:  # pylint: disable=broad-exception-caught
                notebook.warnings.append(desktop_error(exc, "notebook read"))
                result.notebooks.append(notebook)
        return result

    def list_groups(self, parent: Node) -> list[Node]:
        """List direct section groups for duplicate-demo detection."""
        root = ET.fromstring(self.app.GetHierarchy(parent.id, 1, xsSchema=SCHEMA))
        element = next(
            (item for item in root.iter() if item.get("ID") == parent.id), None
        )
        if element is None:
            raise ValueError("Parent missing in desktop response")
        return [
            parse_node(child)
            for child in element
            if kind(child) == "SectionGroup" and not is_true(child, "isRecycleBin")
        ]

    def create_group(self, parent: Node, title: str) -> Node:
        """Create one section-group folder under a known parent."""
        item_id = self.app.OpenHierarchy(title, parent.id, cftIfNotExist=2)
        return Node("Section group", title, item_id)

    def create_section(self, parent: Node, title: str) -> Node:
        """Create one .one section under a known parent."""
        item_id = self.app.OpenHierarchy(f"{title}.one", parent.id, cftIfNotExist=3)
        return Node("Section", title, item_id)

    def create_page(self, section: Node, title: str) -> Node:
        """Create a blank page and set its title without editing existing pages."""
        page_id = self.app.CreateNewPage(section.id, npsNewPageStyle=0)
        # Return the ID before the title update so the demo journal can retain it.
        return Node("Page", title, page_id)

    def finish_page(self, page: Node):
        """Set a newly created page's title after its ID has been journaled."""
        element = ET.Element(
            f"{{{NAMESPACE}}}Page",
            {
                "ID": page.id,
                "name": page.title,
                "dateTime": datetime.now(timezone.utc).isoformat(),
            },
        )
        title = ET.SubElement(element, f"{{{NAMESPACE}}}Title")
        outline = ET.SubElement(title, f"{{{NAMESPACE}}}OE")
        ET.SubElement(outline, f"{{{NAMESPACE}}}T").text = page.title
        self.app.UpdatePageContent(
            ET.tostring(element, encoding="unicode"), xsSchema=SCHEMA
        )

    def set_page_levels(self, section: Node, pages: list[tuple[str, int]]):
        """Update nesting and order of only the new demo pages."""
        root = ET.Element(f"{{{NAMESPACE}}}Section", {"ID": section.id})
        for page_id, level in pages:
            ET.SubElement(
                root,
                f"{{{NAMESPACE}}}Page",
                {"ID": page_id, "pageLevel": str(level + 1)},
            )
        self.app.UpdateHierarchy(ET.tostring(root, encoding="unicode"), SCHEMA)
