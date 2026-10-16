# Detailed workflow

## Environment and scope

Native field creation requires Windows, desktop Microsoft Word, EndNote, its
connected CWYW add-in and a locally installed `.ens` output style. Python 3.10+
handles inspection, normalized XML, mapping and verification. `lxml` is needed
for prefix-safe editing of modern DOCX files with `mc:Ignorable` declarations.
The PowerShell bridge automatically re-enters Windows PowerShell 5.1 from
PowerShell 7; it does not change execution policy.

The Word connection has been tested with EndNote 20 and 64-bit Word. If using
another version, check the connection and try a few citations in a document
copy before processing the full manuscript. Choose the installed output style
that the manuscript needs, including a custom `.ens` style if requested.

Run `uv sync --locked` from the repository or installed skill directory before
using the commands below. `pyproject.toml` declares the dependencies, `uv.lock`
records the resolved versions, and `.python-version` selects Python 3.12 by
default. Python 3.10+ is supported. To use a particular Python installation,
add `--python <path-to-python>` to both `uv sync` and `uv run`.

Before editing, record source and intended output paths, library, existing live
citations, requested style and any caption changes. Infer these from the user's
request and existing context; ask only for a decision that materially blocks work.
Do not restart the user's Word or EndNote to discover paths.

## 1. Inspect

```powershell
uv run scripts/endnote_workflow.py inspect manuscript.docx --output work/inspection.json
uv run scripts/endnote_workflow.py probe
```

The inspection includes paragraph indexes/text, source hash, field inventories,
traveling-record identities, image parts and table counts. All are read-only.

`probe` and `format` call the included PowerShell bridge through Windows
PowerShell 5.1. They pass the script as command text, so no change to the system
script execution policy is needed. The standalone `endnote_word.ps1` remains
available for existing PowerShell workflows.

## 2. Normalize reviewed references

`examples/records.json` shows the input format using fictional articles:

```json
{
  "records": [{
    "id": "paper-a",
    "title": "A synthetic article title",
    "authors": ["Example, Alice", "Example, Bob"],
    "year": "2026",
    "journal": "Example Journal",
    "volume": "1",
    "pages": "10-20",
    "doi": "",
    "url": "",
    "notes": "Source and verification notes belong here."
  }]
}
```

`id` is an external stable key. It is never an EndNote record number.
The CLI accepts an array directly or a `records` object, and accepts legacy
`number` instead of `id` when adapting an existing reviewed catalog.
The current XML exporter accepts `reference_type: journal-article` only.

Resolve DOI/title disagreement and malformed authors before import. Preserve
the cited edition/version. Crossref or DataCite can enrich missing full authors,
but title/author/version checks against a publisher or author-deposited paper
take precedence over blindly trusting one registry response.

If enrichment is needed, cache by DOI, use modest concurrency (for example 2–3
workers), honor `Retry-After`, and retry transient 429/connection errors a bounded
number of times. Avoid fetching all metadata again on the next retry.

## 3. Reconcile and import only missing references

```powershell
uv run scripts/endnote_workflow.py map work/records.json --library Library.enl --output work/mapping-before.json --missing-xml work/missing.xml
```

`map` matches a unique DOI first, then an exact normalized title and year.
It reports duplicate records, author differences and invalid types instead of
choosing an arbitrary record. `--missing-xml` contains missing records only;
conflicts must be resolved separately. A native EndNote exported `.xml` can be
supplied in place of `.enl` when direct read-only database inspection is unavailable.

In native EndNote, open the intended library and use:

1. **File > Import > File**.
2. Select the generated XML and **EndNote generated XML** as Import Option.
3. Import once. Observe the imported count and check the native reference types.

Do not re-import the same file to recover from an uncertain UI outcome. Read
the library again first. For a changed environment/import format, validate a
few representative records, including Unicode authors and a preprint, before
importing a large set.

```powershell
uv run scripts/endnote_workflow.py map work/records.json --library Library.enl --output work/mapping.json
```

Proceed only when `missing` and `conflicts` are empty. These are the actual IDs
that must appear in temporary citations. Back up an existing library through
native EndNote before a substantial mutation; keep `.enl` and `.Data` together.

## 4. Plan exact replacements in a new version

Offsets use Python Unicode character indexes in the inspected paragraph text.
They are half-open: `start` inclusive, `end` exclusive. Every operation contains
the exact original `old` text. Empty `old` with equal offsets requests insertion.

```json
{
  "source_sha256": "the exact hash from inspect",
  "replacements": [
    {"paragraph": 5, "start": 42, "end": 50, "old": "(ref. 1)", "references": ["paper-a"]},
    {"paragraph": 8, "start": 10, "end": 10, "old": "", "references": ["paper-a", "paper-b"]},
    {"paragraph": 9, "start": 0, "end": 7, "old": "Sources", "new_text": "Evidence"}
  ],
  "remove_paragraphs": [
    {"paragraph": 30, "old": "1. Exact manually typed bibliography entry to replace."}
  ]
}
```

Use either `references` or `new_text` in a replacement. New text is for an
explicitly requested prose/credit adjustment, not a blanket rewrite. If a
paragraph is removed, do not also replace text in it. Convert ranges such as
`(refs 9–12)` into the four intended source IDs in the plan. Keep `(Fig. 4c)`
or correction qualifiers separate from the temporary citation.

```powershell
uv run scripts/endnote_workflow.py prepare manuscript.docx --output manuscript-next.docx --mapping work/mapping.json --plan work/plan.json --audit work/prepare-audit.json
```

The helper refuses to overwrite the source or any existing output. It rejects
edits inside live fields and uncertain multi-run edits crossing fields/drawings.
Use native Word for these cases, on the new version, and retain an equivalent
expected-target/prose audit. Do not replace a live bibliography with manual text.

The prepared document is not finished. Its placeholders resemble
`{Example, 2026 #27; Other, 2025 #103}`. Native CWYW must resolve them.

## 5. Format through native CWYW

Ensure the target library is open in EndNote and the style file exists.

```powershell
uv run scripts/endnote_workflow.py format manuscript-next.docx --style 'C:/path/to/installed/Style.ens' --expected-citations 25 --expected-references 18 --pdf work/final.pdf
```

`--expected-citations` counts groups, not individual cited papers. `--expected-references`
counts traveling-library record IDs; the Python verifier additionally checks
distinct identities and, for numbered styles, the bibliography entries/numbers.
Supply the counts calculated from the actual plan, not example constants.

The bridge preserves an existing `EN.Layout`, selects the exact installed style,
calls `ENFormatBibliography`, waits for a stable non-empty bibliography and saves.
It leaves Word and EndNote open. `--timeout-seconds` controls the wait after the
macro returns (5–300 seconds, default 60). If Word is waiting for a dialog,
inspect that dialog before trying the command again.

## 6. Verify and render

```powershell
uv run scripts/endnote_workflow.py verify manuscript-next.docx --audit work/prepare-audit.json --expected-citations 25 --expected-references 18 --output work/verification.json
```

For an already formatted manuscript with no edit plan, omit `--audit` and use
the field/count/numbering checks. Inspect the returned `numbering_checked` flag:
numeric displays require a parseable numbered bibliography; author-year styles
have no numeric check. The CLI never creates a native field fixture as output.
With an audit, citation identities are checked at their paragraph and prose
offsets as well as by total counts. A field moved to a different claim fails
validation even if the total set of cited papers is unchanged.

Render the final saved DOCX using the document environment's renderer/native
Word PDF and inspect every page. A non-empty PDF or extracted text is not a
visual check. After a small final edit, compare rendered-page hashes and inspect
changed pages again. Deliver only requested final artifacts, keeping QA internal.

## Running the tests

From the repository directory:

```powershell
uv sync --locked
uv run --locked python -B -m unittest discover -s tests -v
```

Tests use fictional references and temporary documents. They do not edit an
open manuscript or personal EndNote library. CI also checks the uv entry points
on Windows and Linux. Changes to Word formatting still need a trial on a copy
with Word and EndNote installed, followed by a page-by-page visual check.

## Further reading

- [Clarivate XML DTD](https://support.clarivate.com/Endnote/s/article/EndNote-XML-Document-Type-Definition?language=en_US)
- [Native EndNote import](https://docs.endnote.com/docs/endnote/2025/v1/windows/en/content/08import/import_text_files-win.htm)
- [Native Word CWYW instructions](https://docs.endnote.com/docs/endnote/2025/v1/windows/en/content/09word/basic_instructions1.htm)
- [Clarivate RSServices API background](https://docs.endnote.com/docs/endnote/2025/v1/windows/en/content/appendices/apis_and_plug-ins.htm)
- [Open-source macro implementation reference](https://github.com/felimet/endnote-for-word)
- [Working with uv projects](https://docs.astral.sh/uv/guides/projects/)

The included tools were written for this project. These links explain the
formats and application interfaces they use.
