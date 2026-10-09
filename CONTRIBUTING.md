# Contributing

Use Python 3.11+ in a virtual environment. From the source root:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[desktop,graph,dev]'
.venv/bin/python -m pytest -q
.venv/bin/python -m black --check onenote_hierarchy tests
.venv/bin/python -m pylint onenote_hierarchy tests
.venv/bin/python -m pip check
```

On Windows, create the environment with `python -m venv .venv` and replace
`.venv/bin/python` with `.\.venv\Scripts\python.exe`. No activation is needed.

Pytest enforces 80% combined line/branch coverage across **all** production modules.
Subprocess coverage is collected too. Terminal output lists uncovered lines.
Do not exclude a backend or demo code to raise the score.

Tests isolate external COM/HTTP/authentication boundaries. Keep fixtures synthetic
and temporary. A new failure should have a regression test; do not put notebook
exports, account IDs, tokens, or corporate titles in tests or commits.
Hosted CI cannot establish live OneNote access. The manual Windows smoke procedure
is in the README; report live checks separately from fixture checks.

See [AGENTS.md](AGENTS.md) for the architecture and review contract.
