# Known failure modes and short recovery paths

## Imported references become Generic or titles disappear

EndNote XML Journal Article is `17`; the internal type table/SQLite value `0`
is not an interchange code. One real import with XML `0` produced unknown
internal type `-1`, and the custom style omitted article titles.

Correct the import XML and validate a small native import. Do not bulk-write
the database to compensate for a malformed import. Do not re-import into a
populated library without checking duplicates. Native correction or a backed-up
isolated rebuild is preferable to undocumented database changes.

## Internal IDs appear in the bibliography

A custom style may include `Label`, `Short Title`, or other optional fields.
Keep source IDs in the external mapping/audit, and omit tracking fields from
imports. Do not edit the user's style merely to hide an accidental internal tag.

## Author names or preprint identifiers are wrong

DOI title agreement does not prove that registry authors are complete or split
correctly. Seen examples include a middle initial becoming part of the surname,
question marks replacing diacritics, and a full author list being truncated.
Compare with the reviewed bibliography and the paper's publisher/arXiv header.
Keep authors on separate lines. Do not turn `et al.` into a literal author.

Keep preprint IDs out of numeric Pages. A style can interpret the dotted
identifier as a page expression and corrupt its display.

## Wrong paper selected or some temporary citations remain

Re-read the actual library IDs. A bibliography index is not an EndNote ID, and
multiple libraries can have the same record number. Compare DOI/title/year and
authors, not only `RecNum`. Preserve the database identity in existing traveling
records. Resolve duplicate/metadata conflicts before another format run.

## PowerShell cannot call GetActiveObject

Modern PowerShell uses a runtime where
`[Runtime.InteropServices.Marshal]::GetActiveObject` is absent. The bridge
re-enters Windows PowerShell 5.1 with safely quoted parameters. Do not repeatedly
create new Word instances or change system execution policy as a workaround.
Under a restrictive script-file policy, run the reviewed source as command text.

## Add-in Object is null or the macro is not found

The COM add-in can report `Connect=true` while its `.Object` is null. Use the
installed `EndNote Cwyw.dotm` and real `ENFormatBibliography` macro. Do not invent
`EndNoteFormatBibliography`, call `*Dlg` macros, or create fake ADDIN fields.
Read-only `Probe` shows connection/template state but does not prove formatting.

## Macro returned, but the saved DOCX has no bibliography

CWYW can continue asynchronously after `Application.Run` returns. Wait for a
non-empty `EN.REFLIST`, no temporary markers, expected counts and repeated stable
bibliography hashes before saving/exporting. A raw `Fields.Count` also includes
`EN.CITE.DATA`, hyperlinks and unrelated fields; it is not the citation count.

## XML comparison reports changes despite identical layout

Native Word may renumber relationship IDs and serialize namespaces differently.
Compare section geometry and resolved relationship targets, not raw `sectPr`
byte strings. Embedded figure bytes, citation identities and non-citation prose
are useful invariants. For modern OOXML, preserve declarations referenced by
`mc:Ignorable` or `mc:Choice/@Requires`; use prefix-aware serialization.

Native `w:fldData` can be Base64 containing NUL-terminated UTF-8 EndNote XML.
Strip only trailing NUL terminators before parsing. Do not decode it as ANSI,
count `EN.CITE.DATA` as another citation, or use permissive XML recovery to hide
damaged metadata. Nested `EN.CITE.DATA` payloads belong to the parent citation.

## UI inputs stop reaching the intended dialog

Re-list and select the current EndNote/Word window after a restart. The old
window handle can belong to an obsolete frame with the same title. Inspect the
fresh state before each state-derived input. Related screenshots may have
different origins and z-order; do not assume screenshot zero is always the main
window. Follow the installed computer-use API's coordinate/focus rules.

After the same input/focus failure recurs, refresh window selection once and
choose a supported native/COM route. Do not keep issuing menus, hidden dialogs,
or blind typing sequences. Leave pending modal decisions for explicit inspection.

## Existing EndNote database is unsupported

The read-only helper recognizes the verified `refs` SQLite layout only. Export
native EndNote XML when the engine/schema differs. Do not emulate proprietary
collations or `EN_MAKE_SORT_KEY`, remove triggers, or infer permission to perform
database surgery from a request to insert citations.
