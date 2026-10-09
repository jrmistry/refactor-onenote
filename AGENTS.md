# Contributor and LLM guide

## Architecture

- `cli.py`: argument parsing, lazy backend selection, diagnostic and exit codes.
- `model.py`: shared tree, exact selection, page nesting, references, rendering,
  counts and atomic file writes.
- `desktop.py`: portable XML parser plus optional Windows COM boundary.
- `graph.py`: optional MSAL authentication, validated Graph HTTP/pagination,
  recursive metadata traversal and explicit recorded-snapshot replay.
- `demo.py`: opt-in sample creation with an exclusive durable operation journal.
- `tests/`: synthetic XML/HTTP/auth/COM fixtures, no live notes or credentials.

## Review and verification

Run pytest, Black, Pylint and pip check using the README's development setup.
Ordinary pytest must enforce at least 80% combined line/branch coverage for the
entire package. Both `python -m onenote_hierarchy` and `onenote-hierarchy`
must work. Keep Windows desktop usable without importing Graph packages.

Resolve every notebook selector before scanning. Preserve incomplete warnings
and exit 2; never guess nesting when ordering or levels are ambiguous. Export/list/
doctor must remain read-only; page bodies and attachments are out of scope.
Allow credential-bearing Graph requests only under the validated OneNote root.
Never blindly retry writes. Preserve the pending-operation journal on failures.

Do not publish real snapshots, exports, authentication caches or personal account
identifiers. Keep examples generic. Prefer focused regression fixes over broad
refactoring. Mark fixture results separately from live notebook verification.
Pair output files are individually atomic, not a transaction; preserve the
documented failure behavior rather than claiming both were committed together.
