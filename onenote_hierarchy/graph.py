"""Microsoft Graph traversal, public-client authentication, and demo writes."""

from html import escape
import json
import os
from pathlib import Path
import sys
import time
from urllib.parse import quote, unquote, urlsplit
from uuid import UUID

import msal
from msal_extensions import PersistedTokenCache, build_encrypted_persistence
import requests

from .model import Node, Page, ScanResult, nest_pages

BASE_URL = "https://graph.microsoft.com/v1.0/me/onenote/"


class GraphError(RuntimeError):
    """A Graph request could not establish a reliable result."""


def token_provider(
    client_id: str | None,
    *,
    write: bool,
    login: bool = False,
    tenant: str | None = None,
):
    """Return a refreshing token callback backed by an OS-encrypted MSAL cache."""
    client_id = client_id or os.getenv("ONENOTE_CLIENT_ID")
    if not client_id:
        raise ValueError(
            "Microsoft app client ID missing. Set ONENOTE_CLIENT_ID or --client-id; "
            "see README.md for app registration. The OneNote MCP sign-in cannot "
            "be reused by this standalone CLI."
        )
    try:
        client_id = str(UUID(client_id))
    except ValueError as exc:
        raise ValueError(
            "Microsoft app client ID must be a valid application UUID."
        ) from exc
    tenant = tenant if tenant is not None else os.getenv("ONENOTE_TENANT", "consumers")
    if tenant not in {"consumers", "organizations"}:
        try:
            tenant = str(UUID(tenant))
        except (ValueError, AttributeError) as exc:
            raise ValueError(
                "Microsoft tenant must be consumers, organizations, or a tenant UUID."
            ) from exc
    if sys.platform == "darwin":
        cache_dir = Path.home() / "Library" / "Caches" / "onenote-hierarchy"
    elif sys.platform == "win32":
        cache_dir = (
            Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "onenote-hierarchy"
        )
    else:
        raise ValueError("Graph sign-in is configured for macOS and Windows only.")
    cache_dir.mkdir(parents=True, exist_ok=True)
    # Encrypted persistence must succeed; there is no plaintext fallback.
    persistence = build_encrypted_persistence(
        str(cache_dir / f"{client_id}-{tenant}.bin")
    )
    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant}",
        token_cache=PersistedTokenCache(persistence),
    )
    scopes = [
        (
            "https://graph.microsoft.com/Notes.ReadWrite"
            if write
            else "https://graph.microsoft.com/Notes.Read"
        )
    ]
    force_login = login
    chosen_account = None

    def acquire():
        nonlocal force_login, chosen_account
        result = None
        accounts = app.get_accounts()
        if len(accounts) > 1 and not force_login and chosen_account is None:
            raise ValueError(
                "Multiple cached accounts. Use --login to choose an account."
            )
        if accounts and not force_login:
            chosen_account = chosen_account or accounts[0]
            result = app.acquire_token_silent(scopes, account=chosen_account)
        if not result or "access_token" not in result:
            print(
                "Sign in to your "
                + (
                    "personal Microsoft"
                    if tenant == "consumers"
                    else "work/school Microsoft"
                )
                + " account in the browser.",
                file=sys.stderr,
            )
            result = app.acquire_token_interactive(
                scopes=scopes, prompt="select_account", timeout=300
            )
            force_login = False
        if "access_token" not in result:
            raise GraphError(
                f"Microsoft sign-in failed ({result.get('error', 'no_token')})."
            )
        username = result.get("id_token_claims", {}).get("preferred_username")
        if username:
            matching = app.get_accounts(username=username)
            if len(matching) == 1:
                chosen_account = matching[0]
        return result["access_token"]

    return acquire


class GraphClient:
    """Small synchronous HTTP client; only reads are retried."""

    def __init__(self, token, session=None, sleep=time.sleep):
        self.token = token
        self.session = session or requests.Session()
        self.sleep = sleep

    @staticmethod
    def validated_url(path: str) -> str:
        """Restrict credential-bearing requests to the intended Graph root."""
        url = path if "://" in path else BASE_URL + path.lstrip("/")
        parsed = urlsplit(url)
        decoded_path = unquote(parsed.path)
        unsafe_path = "\\" in decoded_path or any(
            part in {".", ".."} for part in decoded_path.split("/")
        )
        if (
            parsed.scheme != "https"
            or parsed.netloc != "graph.microsoft.com"
            or not parsed.path.startswith("/v1.0/me/onenote/")
            or parsed.fragment
            or unsafe_path
        ):
            raise GraphError(
                "Unexpected Graph pagination/request URL; refusing to send token."
            )
        return url

    def request(self, method: str, path: str, **kwargs) -> dict:
        """Request JSON with bounded reads and no replay of uncertain writes."""
        url = self.validated_url(path)
        headers = {
            "Authorization": f"Bearer {self.token()}",
            "Accept": "application/json",
        }
        headers.update(kwargs.pop("headers", {}))
        for attempt in range(3 if method == "GET" else 1):
            try:
                response = self.session.request(
                    method,
                    url,
                    headers=headers,
                    timeout=30,
                    allow_redirects=False,
                    **kwargs,
                )
            except requests.RequestException as exc:
                # Keep URLs, provider bodies, and credentials out of errors.
                raise GraphError(
                    f"Graph network error ({type(exc).__name__}); request not replayed."
                ) from exc
            if (
                response.status_code in {429, 502, 503, 504}
                and method == "GET"
                and attempt < 2
            ):
                try:
                    delay = float(response.headers.get("Retry-After", 2**attempt))
                except ValueError:
                    delay = 2**attempt
                if delay > 60:
                    raise GraphError(
                        "Graph throttled; Retry-After exceeds 60 seconds. Try later."
                    )
                self.sleep(max(0, delay))
                continue
            if not 200 <= response.status_code < 300:
                raise GraphError(
                    f"Graph HTTP {response.status_code}; "
                    "check sign-in, permissions, and sync."
                )
            try:
                payload = response.json()
            except ValueError as exc:
                raise GraphError(
                    "Graph returned invalid JSON; request not replayed."
                ) from exc
            if not isinstance(payload, dict):
                raise GraphError("Graph returned an unexpected response shape.")
            return payload
        raise GraphError("Graph read retries exhausted.")

    def collection(self, path: str) -> list[dict]:
        """Enumerate every page, detecting broken or repeated pagination links."""
        items = []
        seen = set()
        while path:
            url = self.validated_url(path)
            if url in seen:
                raise GraphError("Graph pagination loop detected.")
            seen.add(url)
            payload = self.request("GET", url)
            value = payload.get("value")
            if not isinstance(value, list) or any(
                not isinstance(item, dict) for item in value
            ):
                raise GraphError("Graph collection missing a valid value array.")
            items.extend(value)
            path = payload.get("@odata.nextLink")
            if path is not None and not isinstance(path, str):
                raise GraphError("Graph pagination URL is invalid.")
        return items


class SnapshotClient:
    """Replay explicitly captured MCP metadata through the same Graph traversal."""

    def __init__(self, data: dict):
        if data.get("schema_version") != 1 or not isinstance(
            data.get("collections"), dict
        ):
            raise ValueError("Unsupported metadata snapshot format.")
        self.collections = data["collections"]
        self.captured_at = data.get("captured_at") or "(unknown time)"
        self.source = (
            f"{data.get('source', 'OneNote MCP')} metadata snapshot; "
            f"captured {self.captured_at}"
        )

    @classmethod
    def from_path(cls, path: Path):
        """Load a recorded snapshot; no authentication or network requests occur."""
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Metadata snapshot must be a JSON object.")
        return cls(data)

    def collection(self, path: str) -> list[dict]:
        """Require explicitly complete collections; absent data is never empty."""
        record = self.collections.get(path)
        if not isinstance(record, dict):
            raise GraphError("Collection missing in metadata snapshot.")
        if record.get("complete") is not True:
            raise GraphError("Collection incomplete in metadata snapshot.")
        value = record.get("value")
        if not isinstance(value, list) or any(
            not isinstance(item, dict) for item in value
        ):
            raise GraphError("Invalid collection in metadata snapshot.")
        return value


def graph_node(data: dict, node_kind: str, title: str | None = None) -> Node:
    """Require a provider ID before traversing or journaling an entity."""
    if not isinstance(data.get("id"), str) or not data["id"]:
        raise GraphError("Graph entity is missing its ID.")
    value = title if title is not None else data.get("displayName", data.get("title"))
    if value is not None and not isinstance(value, str):
        raise GraphError("Graph entity title is invalid.")
    node = Node(node_kind, value or "(untitled)", data["id"])
    if node_kind == "Notebook":
        links = data.get("links") or {}
        web = links.get("oneNoteWebUrl") or {}
        if isinstance(web.get("href"), str):
            node.web_url = web["href"]
    return node


def parent_path(parent: Node) -> str:
    """Build a known notebook or section-group path with its ID escaped."""
    if parent.kind not in {"Notebook", "Section group"}:
        raise ValueError("Expected a notebook or section group parent.")
    collection = "notebooks" if parent.kind == "Notebook" else "sectionGroups"
    return f"{collection}/{quote(parent.id, safe='')}"


class GraphBackend:
    """Traverse cloud notebooks and create additive sample content."""

    name = "graph"

    def __init__(self, client: GraphClient):
        self.client = client

    def list_notebooks(self) -> list[Node]:
        """List all notebooks returned for the signed-in account."""
        return [
            graph_node(item, "Notebook") for item in self.client.collection("notebooks")
        ]

    def list_groups(self, parent: Node) -> list[Node]:
        """List direct child groups, including all collection pages."""
        return [
            graph_node(item, "Section group")
            for item in self.client.collection(f"{parent_path(parent)}/sectionGroups")
            if not item.get("isRecycleBin")
        ]

    def read_pages(self, section: Node):
        """Read page metadata and reconstruct nesting only from reliable ordering."""
        try:
            items = self.client.collection(
                f"sections/{quote(section.id, safe='')}/pages?pagelevel=true"
            )
            pages = []
            for item in items:
                if item.get("isDeleted") or item.get("isDeletedPages"):
                    continue
                node = graph_node(item, "Page")
                pages.append(
                    Page(node.id, node.title, item.get("level"), item.get("order"))
                )
            nest_pages(section, pages, sort_by_order=True)
        except (GraphError, ValueError) as exc:
            section.warnings.append(str(exc))

    def read_children(self, parent: Node, seen: set[str]):
        """Walk groups recursively, preserving visible nodes on partial failures."""
        if parent.id in seen:
            parent.warnings.append("repeated group ID; recursive scan stopped")
            return
        seen.add(parent.id)
        sections = []
        groups = []
        try:
            sections = [
                graph_node(item, "Section")
                for item in self.client.collection(f"{parent_path(parent)}/sections")
                if not item.get("isDeletedPages")
            ]
        except (GraphError, ValueError) as exc:
            parent.warnings.append(f"section listing failed: {exc}")
        try:
            groups = self.list_groups(parent)
        except (GraphError, ValueError) as exc:
            parent.warnings.append(f"section-group listing failed: {exc}")
        parent.children = sorted(
            sections + groups, key=lambda node: (node.title.casefold(), node.id)
        )
        for child in parent.children:
            if child.kind == "Section":
                self.read_pages(child)
            else:
                self.read_children(child, seen)

    def scan(self, notebooks: list[Node]) -> ScanResult:
        """Read all selected cloud notebooks without fetching page bodies."""
        nodes = [
            Node("Notebook", notebook.title, notebook.id) for notebook in notebooks
        ]
        for node in nodes:
            self.read_children(node, set())
        result = ScanResult(
            self.name, nodes, source=getattr(self.client, "source", "live backend")
        )
        if hasattr(self.client, "captured_at"):
            result.captured_at = self.client.captured_at
        return result

    def create_group(self, parent: Node, title: str) -> Node:
        """Create one group without retrying an uncertain response."""
        return graph_node(
            self.client.request(
                "POST",
                f"{parent_path(parent)}/sectionGroups",
                json={"displayName": title},
            ),
            "Section group",
            title,
        )

    def create_section(self, parent: Node, title: str) -> Node:
        """Create one section under the selected parent."""
        return graph_node(
            self.client.request(
                "POST", f"{parent_path(parent)}/sections", json={"displayName": title}
            ),
            "Section",
            title,
        )

    def create_page(self, section: Node, title: str) -> Node:
        """Create a titled ordinary page; Graph cannot assign subpage levels."""
        html = (
            f"<!DOCTYPE html><html><head><title>{escape(title)}</title></head>"
            "<body><p>Sample created by onenote-hierarchy.</p></body></html>"
        )
        data = self.client.request(
            "POST",
            f"sections/{quote(section.id, safe='')}/pages",
            data=html.encode("utf-8"),
            headers={"Content-Type": "text/html"},
        )
        return graph_node(data, "Page", title)
