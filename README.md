# EndNote CWYW Skill

[简体中文](README.zh-CN.md) · [Skill instructions](SKILL.md) · [MIT license](LICENSE)

A reusable Codex skill and small command-line toolkit for moving manuscript
references into EndNote and inserting **real, editable Cite While You Write
citations in Microsoft Word**.

It turns a fragile sequence of ad hoc imports, manual numbering and repeated
UI attempts into a checked workflow:

```text
Reviewed references → reconcile library → import missing records
                    → map real record IDs → prepare a new DOCX
                    → native CWYW formatting → verify and render
```

The toolkit prepares data and verifies results. Native EndNote imports the
records; the installed Word CWYW plugin creates the citation and bibliography
fields. This is not an EndNote replacement or a library-free field generator.

## What is included

- A discoverable Agent Skill covering new libraries, manual-reference conversion,
  additional citations in existing manuscripts, captions and custom output styles.
- A Python CLI for DOCX inspection, normalized EndNote XML export, duplicate-aware
  library mapping, exact-range edits and native-field verification.
- A Windows PowerShell bridge for connection probing and native
  `ENFormatBibliography` execution, with completion checks before saving.
- Synthetic examples, offline tests, cross-platform CI and troubleshooting notes
  drawn from an actual EndNote 20/Word workflow.

The repository contains no manuscripts, personal EndNote databases, credentials,
publisher PDFs, or custom `.ens` style files.

## Requirements

- Python **3.10 or newer**.
- `lxml` for editing modern DOCX files while preserving namespace declarations.
  Inspection, XML export, mapping and verification use the standard library.
- For final formatting: **Windows, desktop Microsoft Word, EndNote and its enabled
  Cite While You Write add-in**, plus a locally installed output style.
- Windows PowerShell 5.1, included in supported Windows installations. The bridge
  can be invoked from PowerShell 7 and re-enters 5.1 automatically for Word COM.

Native integration was verified with EndNote 20 and 64-bit Word. Offline tests
do not establish compatibility with every EndNote/Word release. Probe and smoke
test a different native environment before a bulk operation.

## Install as a personal Codex skill

```powershell
git clone https://github.com/Breathay/endnote-cwyw.git
cd endnote-cwyw
python -m pip install -r requirements.txt
python scripts/install_skill.py
```

The installer copies the skill to `$CODEX_HOME/skills/endnote-cwyw`, or
`~/.codex/skills/endnote-cwyw` when `CODEX_HOME` is unset. It does not change global
settings. To update an existing installation, use `--replace`; the previous
skill is backed up outside the discovery directory, and no files are deleted.

```powershell
python scripts/install_skill.py --replace
```

Example request:

> Use $endnote-cwyw to import this manuscript's references into my EndNote
> library and create the next Word version using my installed output style.

中文也可以：

> 使用 $endnote-cwyw，把论文参考文献导入 EndNote，并在新版本 Word 中用
> Nature 自定义样式插入可更新引文。图注引用直接放在对应描述中。

## Command-line workflow

The detailed data contracts and recovery paths are in
[workflow.md](references/workflow.md) and
[troubleshooting.md](references/troubleshooting.md).

### Inspect the source and connection

```powershell
python scripts/endnote_workflow.py inspect manuscript.docx --output work/inspection.json
pwsh -NoProfile -File scripts/endnote_word.ps1 -Action Probe
```

`Probe` is read-only: it does not launch Word, invoke a macro or modify documents.
If PowerShell 7 is unavailable, invoke the bridge using Windows PowerShell 5.1
subject to the machine's existing script policy. The workflow reference explains
source-text invocation without changing that policy.

### Normalize and map references

Create reviewed `work/records.json` using the contract in
[examples/records.json](examples/records.json). The example records are fictional.
Keep full authors as separate `Family, Given` strings and keep source IDs outside
EndNote's printable `Label` and `Short Title` fields.

```powershell
python scripts/endnote_workflow.py map work/records.json --library Library.enl --output work/mapping-before.json --missing-xml work/missing.xml
```

Import missing records once through native EndNote:
**File → Import → File → EndNote generated XML**. Then map again:

```powershell
python scripts/endnote_workflow.py map work/records.json --library Library.enl --output work/mapping.json
```

A native EndNote XML export can replace `Library.enl` as the read-only mapping
source. Resolve missing and ambiguous records before preparing the manuscript.
The helper never writes to the EndNote database.

The bundled exporter supports the verified **Journal Article XML type 17** only.
Other reference types require validated native EndNote/RIS mappings. Internal
database type IDs are not interchangeable with XML type IDs.

### Prepare a new version

Use the inspected hash and paragraph offsets to create an exact edit plan,
as described in [workflow.md](references/workflow.md#4-plan-exact-replacements-in-a-new-version).

```powershell
python scripts/endnote_workflow.py prepare manuscript.docx --output manuscript-next.docx --mapping work/mapping.json --plan work/plan.json --audit work/prepare-audit.json
```

The source is preserved, and an existing output is never overwritten. The helper
keeps embedded images and unrelated package parts. It rejects replacements
inside existing fields or across uncertain field/drawing boundaries; use native
Word editing for these operations. Existing live citations elsewhere are retained.

### Format and verify

With the intended library open in EndNote:

```powershell
pwsh -NoProfile -File scripts/endnote_word.ps1 -Action Format -Document manuscript-next.docx -StylePath 'C:/path/to/installed/Style.ens' -ExpectedCitations 25 -ExpectedReferences 18 -PdfPath work/final.pdf
python scripts/endnote_workflow.py verify manuscript-next.docx --audit work/prepare-audit.json --expected-citations 25 --expected-references 18 --output work/verification.json
```

Replace the example counts with the plan's actual counts. A citation count means
**citation groups**, not the number of papers inside those groups.

The bridge waits for a stable, non-empty native bibliography and zero temporary
markers before saving. The verifier checks identities and, for numbered styles,
displayed citation numbers against the bibliography. With an audit it also
checks source preservation, planned prose, figure bytes and table/figure counts.

Finally, render the saved document and inspect all pages. Native application
formatting and visual layout cannot be proven by offline unit tests alone.

## Important boundaries

- This skill does not redistribute or install EndNote, Word or `.ens` styles.
- It does not fabricate native citation fields or edit SQLite records, triggers,
  proprietary collations or sort keys.
- External metadata is enrichment, not unquestioned authority: author splitting,
  Unicode, preprint versions and registry completeness can be wrong.
- Captions and image credits change only when requested; citations do not grant
  figure-reuse rights.
- The macro polling timeout begins after the COM macro returns. If a native
  dialog blocks the macro, inspect it instead of launching another format run.
- Keep an `.enl` file with its same-named `.Data` folder when moving a library.

## Tests and development

```powershell
python -m unittest discover -s tests -v
```

The offline suite checks XML type encoding, author validation, real ID mapping,
duplicate/conflict handling, read-only database access, split-run preservation,
existing-field protection, namespace preservation, native-field parsing and
display-number mismatches. Synthetic field fixtures exercise parsing only;
they are not valid substitutes for CWYW-generated user documents.

CI runs these tests on Linux and Windows and parses the PowerShell bridge on
Windows without requiring Office or an EndNote license. Changes to the native
bridge should also receive a separate authorized native smoke test on a copied
document. Keep all real manuscripts and libraries out of commits.

## License

[MIT](LICENSE), copyright © 2026 Breathay. EndNote and Microsoft Word remain
products of their respective owners; their software and custom styles are not
covered by this repository's license.
