"""Tests for visible hierarchy and export behavior."""

import xml.etree.ElementTree as ET

import pytest

from onenote_hierarchy.desktop import DesktopBackend, parse_hierarchy
from onenote_hierarchy.model import (
    Node,
    Page,
    ScanResult,
    atomic_write,
    nest_pages,
    render,
    select_notebooks,
)

XML = """<Notebooks xmlns="http://schemas.microsoft.com/office/onenote/2013/onenote">
<Notebook ID="n" name="School">
 <SectionGroup ID="g" name="Projects">
  <Section ID="s" name="Café" locked="true" areAllPagesAvailable="false">
   <Page ID="p1" name="Overview" pageLevel="1"/>
   <Page ID="p2" name="Setup" pageLevel="2"/>
   <Page ID="p3" name="Details" pageLevel="3"/>
   <Page ID="p4" name="Next" pageLevel="1"/>
  </Section>
  <SectionGroup ID="nested" name="Archive"><Section ID="empty" name="Empty"/></SectionGroup>
 </SectionGroup>
 <SectionGroup ID="trash" name="Bin" isRecycleBin="true"><Section ID="deleted" name="Gone"/></SectionGroup>
 <Section ID="deleted-pages" name="Trash" isDeletedPages="true"/>
</Notebook></Notebooks>"""


def test_desktop_xml_preserves_nesting_and_marks_incomplete():
    """Flattening sibling page XML or omitting availability flags loses information."""

    result = parse_hierarchy(XML)
    output = render(result, include_ids=True)
    assert not result.complete
    assert "            [Page] Overview" in output
    assert "                [Page] Setup" in output
    assert "                    [Page] Details" in output
    assert "            [Page] Next" in output
    assert "[Section] Café" in output
    assert "LOCKED" in output and "unavailable" in output
    assert "[Section] Empty" in output
    assert "Gone" not in output and "Trash" not in output
    assert "ID: p3" in output


def test_empty_notebook_and_title_whitespace():
    """An empty notebook is complete; newlines in names must not forge tree entries."""

    result = ScanResult("graph", [Node("Notebook", "  Empty\n notebook  ", "n")])
    assert result.complete
    assert "[Notebook] Empty notebook\n" in render(result)
    assert "ID:" not in render(result)


def test_notebook_selection_disambiguates_names():
    """Picking the first matching name could write to the wrong notebook."""

    notebooks = [Node("Notebook", "School", "a"), Node("Notebook", "School", "b")]
    assert select_notebooks(notebooks, notebook_id="b")[0].id == "b"
    with pytest.raises(ValueError, match="Several"):
        select_notebooks(notebooks, name="school")
    with pytest.raises(ValueError, match="not found"):
        select_notebooks(notebooks, name="Missing")
    assert len(select_notebooks(notebooks)) == 2


@pytest.mark.parametrize("levels", [[1], [0, 2], [0, 3], [0, None]])
def test_invalid_page_levels_do_not_invent_parents(levels):
    """Orphaned or invalid page levels must yield an incomplete flat listing."""

    section = Node("Section", "Test", "s")
    nest_pages(section, [Page(str(i), str(i), level) for i, level in enumerate(levels)])
    assert section.warnings
    assert len(section.children) == len(levels)
    assert all(not page.children for page in section.children)


def test_legacy_subpage_flags_are_incomplete_and_deleted_pages_excluded():
    """A legacy boolean cannot establish either subpage's exact parent depth."""

    output = render(
        parse_hierarchy("""<Notebook ID="n" name="N"><Section ID="s" name="S">
    <Page ID="a" name="Parent" pageLevel="1"/>
    <Page ID="b" name="Child" isSubPage="true"/>
    <Page ID="d" name="Detail" isSubPage="true"/>
    <Page ID="c" name="Gone" isDeletedPages="true"/></Section></Notebook>""")
    )
    assert "Status: INCOMPLETE" in output
    assert "        [Page] Child" in output
    assert "        [Page] Detail" in output
    assert "            [Page] Child" not in output
    assert "Gone" not in output


def test_atomic_export_preserves_old_file_on_replace_failure(tmp_path, monkeypatch):
    """Failed replacement preserves the old export and removes temp files."""

    target = tmp_path / "hierarchy.txt"
    target.write_text("old", encoding="utf-8")

    def fail_replace(_source, _destination):
        raise OSError("replacement failed")

    monkeypatch.setattr("os.replace", fail_replace)
    with pytest.raises(OSError):
        atomic_write(target, "new")
    assert target.read_text(encoding="utf-8") == "old"
    assert list(tmp_path.iterdir()) == [target]


def test_desktop_creator_sets_only_new_page_levels():
    """Demo updates must contain only the new page IDs in the chosen section."""

    class App:  # pylint: disable=too-few-public-methods
        """Capture the desktop XML boundary without launching OneNote."""

        def __init__(self):
            self.changes = None

        def UpdateHierarchy(self, xml, _schema):  # pylint: disable=invalid-name
            """Capture hierarchy updates."""
            self.changes = ET.fromstring(xml)

    app = App()
    backend = DesktopBackend(app)
    backend.set_page_levels(Node("Section", "S", "s"), [("a", 0), ("b", 1), ("c", 2)])
    assert app.changes.attrib["ID"] == "s"
    assert [(page.get("ID"), page.get("pageLevel")) for page in app.changes] == [
        ("a", "1"),
        ("b", "2"),
        ("c", "3"),
    ]
