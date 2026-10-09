# Validation

The public project uses synthetic notebook titles and identifiers. Automated tests
do not read real OneNote notebooks or sign into Microsoft accounts.

## Automated checks

Ordinary pytest enforces **80% combined line/branch coverage** across the complete
`onenote_hierarchy` package, including Windows, Graph and demo modules. The
Windows/macOS CI matrix runs Python 3.11 and 3.13, Black, Pylint and pip check.
A separate Windows job installs desktop/development extras without Graph packages
and runs controlled COM diagnostics.

Fixtures cover nested groups and subpages, Unicode and duplicate names/paths,
empty notebooks, exact multi-selection, deleted/unavailable content, pagination,
unordered/missing page metadata, request failures, cache separation and write
journaling. CLI integration checks use real selection/rendering/file-writing code
with external notebook IO replaced. Authentication paths use temporary directories.

Publication review fixed URL normalization that could otherwise accept dot
segments outside the Graph OneNote root; regression tests cover literal and
encoded traversal. Requests reject redirects and foreign pagination origins.
Exports do not retrieve page bodies or attachments.

## Publication review results

Local macOS Python 3.13 verification: **86 tests passed, 89.98% combined
line/branch coverage**, Black passed, Pylint 10/10, and pip check passed.
The 80% gate includes every production module, with no coverage omissions.

Three passes covered correctness, independent whole-source review, and source
archive installation/publication delivery. The independent reviewer found
malformed Graph titles could abort partial output; two failing regressions
reproduced it before the fix. Invalid titles now produce scoped incomplete
warnings, and null page titles consistently use the existing untitled fallback.
The CLI regression checks incomplete exit 2 while retaining unaffected notebooks.
No critical findings or deferred minor findings were reported.

## Desktop diagnostic patch (0.3.1)

Three failing regressions reproduced lost exception types and wrapped Windows/
cache failures before the fix. Diagnostics now preserve those types and HRESULTs
without including raw exception descriptions. Errors without an HRESULT no
longer ask users to supply a nonexistent code. The original connection method
remains unchanged; these fixtures do not identify a particular laptop's failure.

## Live verification limits

A successful fixture test proves controlled behavior, not access to a particular
Microsoft account. Windows CI installs pywin32 and exercises COM fixtures; hosted
runners do not provide desktop OneNote with an authenticated notebook session.
macOS Graph requires an approved public-client ID and successful delegated consent.

Before using an export, compare groups, titles and subpage nesting against OneNote.
Record the actual command, exit status, capture time and incomplete warnings.
An old snapshot is explicitly labeled as replay, and cannot establish freshness.
No automatic consolidation or deletion is implemented. Originals remain under
human control throughout the manual copy/verification workflow.
