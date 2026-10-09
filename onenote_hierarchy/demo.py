"""Create a small isolated sample with a durable journal of completed writes."""

import json
from pathlib import Path

from .model import Node, atomic_write


def create_demo(backend, notebook: Node, manifest_path: Path) -> dict:
    """Create only new demo items, stopping rather than replaying uncertain writes."""
    if any(
        node.title.casefold() == "hierarchy demo"
        for node in backend.list_groups(notebook)
    ):
        raise ValueError(
            "Hierarchy Demo already exists; refusing to duplicate or modify it."
        )
    manifest_path = manifest_path.expanduser().resolve()
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest = {
        "backend": backend.name,
        "notebook_id": notebook.id,
        "notebook_title": notebook.title,
        "status": "in_progress",
        "items": [],
    }
    # Reserve the journal before the first external write; never overwrite an older run.
    with manifest_path.open("x", encoding="utf-8") as stream:
        json.dump(manifest, stream, ensure_ascii=False, indent=2)

    def save():
        atomic_write(
            manifest_path, json.dumps(manifest, ensure_ascii=False, indent=2) + "\n"
        )

    def update(operation, details, action):
        manifest["pending"] = {"operation": operation, **details}
        save()
        action()
        del manifest["pending"]
        save()

    def create(parent, title, node_kind, desired_level=None):
        manifest["pending"] = {
            "operation": "create",
            "parent_id": parent.id,
            "title": title,
            "kind": node_kind,
        }
        save()
        methods = {
            "Section group": backend.create_group,
            "Section": backend.create_section,
            "Page": backend.create_page,
        }
        node = methods[node_kind](parent, title)
        record = {
            "id": node.id,
            "parent_id": parent.id,
            "title": title,
            "kind": node_kind,
        }
        if desired_level is not None:
            record["desired_level"] = desired_level
        manifest["items"].append(record)
        del manifest["pending"]
        save()
        if node_kind == "Page" and backend.name == "desktop":
            update(
                "set_page_title",
                {"page_id": node.id, "title": title},
                lambda: backend.finish_page(node),
            )
        return node

    try:
        root = create(notebook, "Hierarchy Demo", "Section group")
        basics = create(root, "Basics", "Section")
        pages = [
            create(basics, title, "Page", level)
            for title, level in [("Guide", 0), ("Installing", 1), ("Requirements", 2)]
        ]
        if backend.name == "desktop":
            levels = [(page.id, level) for level, page in enumerate(pages)]
            update(
                "set_page_levels",
                {"section_id": basics.id, "pages": levels},
                lambda: backend.set_page_levels(basics, levels),
            )
        examples = create(root, "Examples", "Section group")
        walkthrough = create(examples, "Walkthrough", "Section")
        for title in ["First example", "Notes"]:
            create(walkthrough, title, "Page", 0)
        manifest["status"] = (
            "created" if backend.name == "desktop" else "awaiting_ui_indentation"
        )
        save()
    except Exception:
        manifest["status"] = "failed_review_required"
        save()
        raise
    return manifest
