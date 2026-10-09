# OneNote hierarchy exporter

A local Python CLI for notebook → section group → section → page → subpage titles.
It exports a hierarchy map for planning manual notebook consolidation with an
approved workplace Copilot. It does not fetch page bodies, attachments, or headings
inside pages. Windows reads notebooks open in desktop OneNote; Mac uses Graph.
The source code is MIT licensed. Windows exports work without MCP, cloud app
registration, or AI access. [Download the source ZIP](https://github.com/jrmistry/refactor-onenote/archive/refs/heads/main.zip)
or browse [the repository](https://github.com/jrmistry/refactor-onenote).

## Windows workplace setup

Follow these steps once to set up the exporter. After that, use **Run it again
later** below whenever you need an updated hierarchy.

### 1. Check what you need

- A Windows laptop with **Microsoft 365 desktop OneNote** installed.
- **Python 3.11 or newer** and permission to install the `pywin32` package.
- A folder where you are allowed to save the project and its reports.

On a work laptop, use your employer's approved software installation process.
If Python or desktop OneNote is missing, ask IT to install it. The
[official Python Windows guide](https://docs.python.org/3/using/windows.html)
explains Python installation and its Windows commands.

You do not need Git or a GitHub account. This Windows workflow uses your existing
OneNote session; no MCP connection, cloud app registration, or additional sign-in
is required.

### 2. Download and extract the project

1. Open [the GitHub repository](https://github.com/jrmistry/refactor-onenote).
2. Click the green **Code** button, then **Download ZIP**.
3. In File Explorer, right-click `refactor-onenote-main.zip` and select **Extract All**.
4. Choose a user-writable, employer-approved folder and click **Extract**.
5. Open the extracted folders until you can see **`pyproject.toml`** and **`README.md`**.
   This is the project folder. Run the commands below from here, not inside the ZIP.

### 3. Open PowerShell and check Python

In File Explorer, while looking at the project folder, click the address bar at
the top, type `powershell`, and press **Enter**. A PowerShell window opens in
that folder. Use an ordinary window, not **Run as administrator**.

**Paste each command on its own, press Enter, and wait for it to finish before
running the next one.** Copy only the command text inside the boxes.

```powershell
python --version
```

You should see a version such as `Python 3.13.x`. It must be **3.11 or newer**.
If `python` is not recognized or opens the Microsoft Store instead, try:

```powershell
py -3 --version
```

If neither command shows a suitable installed version, stop and ask IT for help.

### 4. Create the virtual environment and install the exporter

A virtual environment is a `.venv` folder that keeps this project's Python
packages separate from other programs.

```powershell
python -m venv .venv
```

If only `py -3` worked in step 3, use this command **instead**:

```powershell
py -3 -m venv .venv
```

Creating `.venv` may finish without printing anything. Then install:

```powershell
.\.venv\Scripts\python.exe -m pip install ".[desktop]"
```

Wait for installation to finish successfully. All remaining commands use the
Python inside `.venv`, so you do not need to activate it, change PowerShell
execution policies, run `pywin32_postinstall`, or use an administrator shell.

### 5. Open your notebooks and check the connection

Open the **desktop OneNote app**. Open every notebook you want to export and let
OneNote finish syncing. Keep it open while running these commands:

```powershell
.\.venv\Scripts\python.exe -m onenote_hierarchy --backend desktop doctor
```

For example, with two notebooks open, a working connection reports:

```text
PASS: Windows, pywin32, and OneNote COM connection.
PASS: inventory readable (2 notebook(s) open).
Next: run list, then export. No notes were edited.
```

Your notebook count may differ. If it says no notebooks are open, open them in
OneNote and run `doctor` again. If you see an error, use the troubleshooting table
below before continuing. `doctor` checks the connection and notebook inventory;
the export checks whether individual sections can be read.

List the notebooks the exporter can see:

```powershell
.\.venv\Scripts\python.exe -m onenote_hierarchy --backend desktop list
```

Check that your intended notebook names appear.

### 6. Export all open notebooks

```powershell
.\.venv\Scripts\python.exe -m onenote_hierarchy --backend desktop export --output hierarchy.txt --planning-csv planning.csv
```

Immediately after the export, check its exit code:

```powershell
$LASTEXITCODE
```

| Code | What it means |
|---:|---|
| `0` | The export completed. Review the files against OneNote. |
| `1` | The command failed. Read the error; any existing output may be from an older run. |
| `2` | An incomplete report was saved. Read its warnings before using it for planning. |

The export reads titles and nesting; it does not change your notes or fetch page
bodies or attachments. The exporter does not upload reports or invoke AI. The
installed OneNote app may sync through its normal configuration.

### 7. Find and use your results

Return to the project folder in File Explorer. You will find:

| File | How to use it |
|---|---|
| `hierarchy.txt` | Open with Notepad to see notebook, group, section, page and subpage titles. |
| `planning.csv` | Open with Excel or your approved spreadsheet app to fill in proposed paths, actions and reasons. |

Compare some sections and subpage indentation with OneNote before planning a
consolidation. Keep company reports on your work laptop and use only approved
workplace tools: titles can contain confidential information.

Use [COPILOT_PROMPT.md](COPILOT_PROMPT.md) with your approved workplace Copilot.
It asks for a proposed combined structure and a mapping for each item, flags
uncertainty, and avoids claiming duplicate content from titles alone. Carry out
the plan with OneNote's built-in **Copy** controls, sync and verify content,
nesting, sharing and internal links before retiring any originals.

### Run it again later

Open PowerShell in the **same project folder**, open/sync your notebooks, and
repeat the export command in step 6 followed by `$LASTEXITCODE`. You do not need
to recreate `.venv` or reinstall. A successful export replaces the same output
files; copy or rename them first if you want to keep the previous report or any
planning edits you made in the CSV. Close the files in Excel/other apps before
exporting again.

### Optional: export only selected notebooks

Use the exact names shown by `list`. Replace `Notebook A` and `Notebook B` with
your names, keeping the quotation marks. To export just one, omit the second
`--notebook` argument.

```powershell
.\.venv\Scripts\python.exe -m onenote_hierarchy --backend desktop export --notebook "Notebook A" --notebook "Notebook B" --output hierarchy.txt --planning-csv planning.csv
$LASTEXITCODE
```

### Troubleshooting

| What you see | What to do |
|---|---|
| `python` is not recognized / opens the Store | Try `py -3 --version`; otherwise ask IT to install approved Python 3.11+. |
| `.venv\Scripts\python.exe` is not found | Check that you are in the folder containing `pyproject.toml`, then complete step 4. |
| pip cannot find `pyproject.toml` / the project | Fully extract the ZIP and open PowerShell in the project folder from step 2. |
| Package download is blocked / pywin32 could not be loaded | Ask IT about its approved package source; retry the step 4 pip command once installation is permitted. |
| Desktop OneNote unavailable to COM | Open Microsoft 365 desktop OneNote, not the browser or old Windows 10 Store app; ask IT to check its installation if needed. |
| Access denied / automation cache not writable | Use your ordinary OneNote user session and an approved writable folder; ask IT if the restriction persists. |
| No notebooks returned / a notebook is missing | Open and sync it in desktop OneNote, then rerun `doctor` and `list`. |
| Several notebooks match | Use `--notebook-id "ID from list"` instead of the name; repeat for multiple notebooks. |
| Output file cannot be written | Close it in Excel/other apps and check that the project folder is writable. |

If another connection error persists after syncing, give IT the error and any
reported HRESULT code. This exporter cannot bypass workplace automation policies.

## macOS / Graph setup

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e '.[graph]'
```

The independent Python CLI requires its own Microsoft application ID. It cannot
reuse a Codex OneNote connector's tokens or an existing browser/desktop sign-in.
Signing into the SharePoint web page alone does not configure this app.

### Work or school accounts

1. In [Microsoft Entra app registrations](https://entra.microsoft.com/), register
   a public-client app supporting work/school accounts. For a registration in the
   notebook's own tenant, single-tenant is sufficient: use that tenant UUID. For
   a registration in another tenant you administer, select **Accounts in any
   organizational directory** and use `organizations`. School policies may still
   require administrator consent or prohibit external apps. If registration is
   unavailable, request an approved public-client app from the school administrator.
2. Under **Authentication**, add **Mobile and desktop applications**, redirect URI
   `http://localhost`. Do not create a client secret.
3. Under **API permissions**, add Microsoft Graph **delegated** `Notes.Read`.
   List/export request only this read scope. No write or application permission is
   needed for consolidation planning. Complete any required administrator approval.
4. Copy the **Application (client) ID** and set it in your terminal (it is not a secret):

```bash
export ONENOTE_CLIENT_ID="your-application-client-id"
export ONENOTE_TENANT="organizations"
python -m onenote_hierarchy --backend graph --login list
```

For a single-tenant app, set `ONENOTE_TENANT` to its Directory (tenant) UUID instead.
Sign into the account that owns/can access the intended notebooks. `list` prints
names, Graph IDs, and returned web links when available. Confirm the notebook names
and returned links refer to the intended notebooks before exporting. A SharePoint
`sourcedoc` GUID identifies a document reference; do not pass it as `--notebook-id`.
If a link returned by Graph cannot establish identity, resolve that uncertainty
before using the export for consolidation.

For two notebooks returned by `list`:

```bash
.venv/bin/python -m onenote_hierarchy --backend graph --tenant organizations --login list
.venv/bin/python -m onenote_hierarchy \
  --backend graph --tenant organizations export \
  --notebook "Notebook A" \
  --notebook "Notebook B" \
  --output exports/hierarchy.txt \
  --planning-csv exports/planning.csv
```

The Graph backend traverses `/me/onenote`, not arbitrary shared group/site notebook
inventories. If the selected account does not return a notebook, stop and check
account/access; the Windows backend can instead read notebooks already open there.
Graph access remains unverified until app configuration and real sign-in succeed.

### Personal accounts

Use a public-client registration supporting personal Microsoft accounts with the
same localhost redirect, and `--tenant consumers` (the default). Request delegated
`Notes.Read` for export. The separate `create-demo` command requests `Notes.ReadWrite`;
do not grant write scope for the workplace planning workflow. Global options must
come **before** the subcommand:

```bash
python -m onenote_hierarchy --backend graph --client-id YOUR_APP_ID --tenant consumers list
```

`--client-id` overrides `ONENOTE_CLIENT_ID`; `--tenant` overrides `ONENOTE_TENANT`.
Tenant accepts only `consumers`, `organizations`, or a UUID. `--login` chooses an
account again; ambiguous cached accounts otherwise require explicit selection.
MSAL refreshes tokens when possible. Caches are encrypted through Mac Keychain or
Windows DPAPI outside the repo, keyed by client ID and authority:
`~/Library/Caches/onenote-hierarchy` or `%LOCALAPPDATA%\onenote-hierarchy`.
There is no plaintext fallback. The new authority-specific cache filenames mean
users of v0.1 may need to sign in again; old caches are not modified.

## Command behavior and output

Both `python -m onenote_hierarchy` and installed `onenote-hierarchy` are supported.
`list` inventories available notebooks. `export` is read-only. Repeat names or IDs
for a selected set; don't mix the two selector types. Case-insensitive names must
match uniquely; duplicate names require the exact backend IDs from `list`.
Selections are all resolved before any scan or output write. Repeated selection
of the same notebook exports it once. With no selection, all available notebooks
are exported. `create-demo` still requires exactly one explicit notebook.

```bash
python -m onenote_hierarchy --backend graph export --notebook-id ID_A --notebook-id ID_B --include-ids --output exports/hierarchy.txt --planning-csv exports/planning.csv
```

Default output is `exports/hierarchy.txt`. Text contains backend/scope/source,
capture time, completeness, counts, and labeled titles with four spaces per level.
Subpages are included in the page total. Graph sorts pages by returned order before
interpreting levels; groups and sections are sorted by name. Desktop preserves
returned XML order and normalizes levels 1–3 to 0–2. Deleted/recycle-bin content is
excluded. Unavailable sections or ambiguous metadata produce an INCOMPLETE report;
parent relationships are never guessed.

`--planning-csv PATH` adds one row per hierarchy node with reference,
parent_reference, kind, original_path, blank proposed_path/action/reason, and scan
status. References are unique within that export and are not permanent OneNote IDs.
The CSV accompanies the same run's text report, whose warnings describe incomplete
coverage. Paths are labeled display strings; titles may themselves contain slashes.

Outputs are UTF-8 and each file is replaced atomically. The two files are not a
single transaction: a CSV write failure returns failure even if the text file was
already written; rerun before treating them as a pair. Output paths must differ.
No notebook titles or notebook URLs are stored in a token cache. Generated files and `.env` files are
ignored by Git; workplace outputs should be saved in an approved local folder.

| Exit code | Meaning |
|---:|---|
| 0 | Listing or complete export succeeded |
| 1 | Setup, connection, selection, or output write failed |
| 2 | Output written, but scan incomplete |

Coverage means notebooks open in desktop OneNote or returned for the signed-in
Graph account. A traversal is not a transactional backup: avoid moving notes while
scanning. Only GET requests receive bounded retries; no creation request is replayed.

### Previously captured MCP metadata

```bash
python -m onenote_hierarchy --backend graph export --snapshot exports/onenote-mcp-snapshot.json --output exports/hierarchy-mcp.txt
```

This replays old metadata; it is not a fresh scan. `.[graph]` is required to load the
Graph adapter. The original capture time/source remain visible. Snapshot v1 contains
`schema_version`, `source`, `captured_at`, and `collections` keyed by relative Graph
paths. Each record requires `complete: true` and a fully enumerated `value` array;
missing records are errors, not empty collections. Never substitute replay or UI
observations for a requested live Python export.

## Optional sample creation

This is separate from the read-only workplace workflow. `create-demo` adds an isolated
**Hierarchy Demo** group only to a selected test notebook. It refuses an existing
group or journal and never deletes/overwrites notes or automatically resumes.

```bash
python -m onenote_hierarchy --backend desktop create-demo --notebook "Disposable Test"
```

The sample contains Basics → Guide → Installing → Requirements and nested
Examples → Walkthrough → First example / Notes. Windows sets page nesting. Graph
creates ordinary pages, then prints the manual indentation steps (Guide top level,
Installing one level, Requirements two levels), because Graph cannot write page
level/order. Sync and export to verify. Created IDs and pending writes are recorded
in a new `exports/demo-manifest.json`; inspect it on failure before another attempt.

## Dependency list for IT review

| Installation | Runtime packages | Access |
|---|---|---|
| Base | Python standard library | Local files and hierarchy rendering |
| `.[desktop]` | `pywin32` (Windows only) | Existing desktop OneNote COM API; no tool cloud connection |
| `.[graph]` | `requests`, `msal`, `msal-extensions` and their package dependencies | Microsoft Graph, browser authentication, OS encrypted cache |
| `.[dev]` | Black, Pylint, pytest, pytest-cov (coverage.py) and their package dependencies | Development only |

Installation via pip also uses the setuptools build backend. Review dependency
versions and distribution through your employer's approved package process. No
administrator installation or execution-policy change is part of these commands.

## Development and Windows smoke test

```bash
python -m pip install -e '.[desktop,graph,dev]'
python -m black --check onenote_hierarchy tests
python -m pylint onenote_hierarchy tests
python -m pytest -q
python -m pip check
```

Ordinary `pytest` measures the entire package, including both backends and demo
code, and fails below **80% combined line/branch coverage**. CI runs Python 3.11
and 3.13 on Windows and macOS, plus a separate Windows installation without Graph
packages. See [CONTRIBUTING.md](CONTRIBUTING.md), [AGENTS.md](AGENTS.md), and
[VALIDATION.md](VALIDATION.md) for architecture and verification limits.

Tests use controlled XML/HTTP/authentication boundaries; no real notebook is changed.
On Windows, run `doctor`, then open/sync two disposable notebooks containing nested
groups, Unicode
and repeated page titles, and subpages. Run `list`, select both in one export, compare
titles/nesting/counts with OneNote, and check every CSV row/reference. Include a
locked/unavailable section and verify exit 2. Rerun after changing a title and confirm
fresh output. Fixture tests on Mac do not establish live Windows COM compatibility.

## Building the source ZIP (maintainers only)

After checks pass and changes are committed, build a source-only archive:

```bash
mkdir -p dist
git archive --format=zip --prefix=refactor-onenote/ --output=dist/refactor-onenote-windows.zip HEAD
```

The ZIP uses committed files only. It contains the Python source, documentation,
license, and tests; it excludes `.git`, `.venv`, authentication caches, generated
reports, and local snapshots. The receiving Windows laptop does not need Git or
development tools. Dependencies are installed through its permitted pip source.

## References

- [OneNote desktop API](https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote)
- [Desktop enum values](https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference)
- [Graph content, pagination, and levels](https://learn.microsoft.com/en-us/graph/onenote-get-content)
- [Graph page properties](https://learn.microsoft.com/en-us/graph/api/resources/page?view=graph-rest-1.0)
- [MSAL browser sign-in](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens)
- [Encrypted token cache](https://github.com/AzureAD/microsoft-authentication-extensions-for-python)
- [OneNote Move or Copy controls](https://support.microsoft.com/en-us/onenote/organize-your-notes)
