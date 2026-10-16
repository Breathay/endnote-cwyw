---
name: endnote-cwyw
description: Build or extend an EndNote library and insert or update real Cite While You Write citations in a versioned Word manuscript. Use for EndNote建库、导入参考文献、Word插入EndNote引文、手工编号转EndNote、CWYW、Nature自定义样式 and figure-caption citations. Requires native Windows Word and EndNote for final field creation.
---

# EndNote library and Word CWYW

Deliver a new Word version containing native, editable `ADDIN EN.CITE` and one
`ADDIN EN.REFLIST` field, linked to verified EndNote records. Reuse existing
library records and live citations. Preserve the source manuscript and figures.

Read [references/workflow.md](references/workflow.md) for command schemas and
the first end-to-end run. Read [references/troubleshooting.md](references/troubleshooting.md)
only for an import, mapping, Word, or UI failure.

## Choose the shortest valid path

1. Inspect the requested source file and discover the running Word/EndNote
   context. Do not infer the source from whichever document is active.
2. If the document already has live EndNote fields, preserve them and add only
   requested citations. If it uses manual references, plan their conversion.
3. Prefer existing EndNote/library exports and reviewed bibliographic metadata.
   Fetch external metadata only to fill missing fields. Do not re-audit every
   paper or silently overwrite reviewed names and dates with a registry result.
4. Reconcile against the target library **before import**, import only missing
   records using native EndNote, and reconcile again afterward.
5. Prepare a new DOCX with temporary citations using the actual mapped record
   numbers. Format it with native CWYW, wait for completion, save, then verify.

For document authoring in Codex, resolve the bundled runtime through the
workspace dependency loader and follow the available documents skill's
artifact-operation and render/inspect requirements. The helper CLI is portable;
do not assume a particular local Python executable or package version.

Use an available Python 3.10+ interpreter to run the helper scripts. Before
editing a DOCX, check whether that interpreter can import `lxml` version 5 or
newer. Reuse the available document runtime when it already has the dependency.
If needed, install `requirements.txt` using that interpreter's `-m pip` in an
appropriate managed environment. Handle these checks as part of the task;
users should not have to run setup commands before requesting citation work.
Neither uv nor a repository `.venv` is required. Run commands from the skill
directory, or use absolute paths to its scripts and requirements file.

## Reference details

- EndNote XML uses `<ref-type name="Journal Article">17</ref-type>`.
  The local SQLite Journal Article value `0` belongs to a different namespace.
  The exporter deliberately supports only this verified XML type; for books,
  proceedings or other types, use validated native exports/RIS mappings.
- Record numbers are library-local IDs. Neither bibliography numbers nor XML
  input order prove the assigned ID. Never assume `#1..#N`, even in a new library.
- Keep full authors as separate `Family, Given` entries. Reject `et al.` as an
  author, corrupted Unicode and HTML. Registry surname splitting and incomplete
  author lists need source-based review; do not hard-code past paper corrections.
- Keep tracking IDs in the external mapping/audit. Do not fill `Label` or
  `Short Title` with internal markers: a custom output style may print them.
- Preprint identifiers such as `2025.11.11.123456v1` are text, not page ranges.
  Preserve a version-specific manuscript year when the cited version requires it.
- Create native fields through the installed CWYW plugin. XML construction of
  an `ADDIN` lookalike is not a substitute. `EN.CITE.DATA` is a storage field,
  not an additional citation group.
- Do not edit the EndNote database, custom collations, triggers or sort keys.
  The mapping helper opens supported databases with `mode=ro`. On an unsupported
  layout, use a native EndNote XML export. Catch import problems before bulk use.

## Reference and caption plan

Use `inspect` to obtain paragraph indexes and a SHA-256 source fingerprint.
Create an explicit plan with exact text ranges and source reference IDs.
Use the `prepare` helper for text outside existing fields; it preserves the rest
of the DOCX package. For a replacement that intersects a field, equation, drawing,
or an uncertain field boundary, use native Word range editing on the new version.
Do not flatten a whole paragraph to insert one citation.

If the user asks to remove separate `Sources:` blocks, move their citation
targets into the corresponding panel descriptions **before** removing the blocks.
Preserve panel-to-paper correspondence. Keep source/permission details in the
figure records when required; do not treat a citation as proof of reuse permission.
Do not remove image credits or manuscript prose unless the requested edit covers it.

## Native formatting

The installed output-style file must exist. Use its exact name; preserve an
existing document's `EN.Layout` preferences except for a requested style change.
Do not redistribute a user's custom `.ens` file with this skill.

Run `python scripts/endnote_workflow.py probe` with the resolved interpreter
to inspect the live connection without modifying Word. Use
`format <new-output.docx> --style <installed.ens>`
only on the authorized new output. Both commands call `scripts/endnote_word.ps1`,
which uses Windows PowerShell 5.1 for Word COM and the real macro
`ENFormatBibliography`. It never quits the user's applications.

Macro return does not prove completion. Require a non-empty native bibliography,
zero temporary markers, expected citation/reference counts, and stable state on
consecutive polls before saving. The polling timeout starts **after** the native
macro returns; a modal macro can still require UI inspection. Do not retry the
macro while a previous call or dialog is pending.

When UI is needed, use the available computer-use skill and current window
observations. Re-select windows after an EndNote restart: old handles can point
to obsolete frames with identical titles. Do not replay recorded coordinates.

## Check the result

Use `verify --audit ...` to check the original fingerprint, planned citation
identities, non-citation prose, native fields, unresolved markers, and embedded
image hashes. For numbered bibliographies, the helper also checks displayed
numbers against the bibliography. Verify that reference types and authors in
the library match the reviewed records. Counts alone cannot establish correctness.

Render the final saved document with native Word when available and inspect
every page for clipping, split captions, broken tables and missing glyphs. Keep
QA images/PDFs internal unless requested. Stop repeating checks once they pass;
re-run the affected checks after a concrete change.

Deliver the new DOCX and the library location. Explain any automatic renumbering
and any unresolved metadata conflict. An `.enl` and its same-named `.Data` folder
must stay together. A skill or README is not evidence that native application
formatting has succeeded.
