# EndNote CWYW

[简体中文](README.zh-CN.md) · [Skill instructions](SKILL.md) · [MIT license](LICENSE)

Import manuscript references into EndNote and replace manually typed citations
with citations you can keep updating in Word. You can then add papers, reorder
citations or switch journal styles using the EndNote add-in, with the
bibliography updated alongside them.

This is a Codex skill for converting a whole manuscript or adding a few missing
citations. It keeps existing EndNote citations and saves changes in a new Word
version, leaving the original manuscript in place.

## Before you start

You need Windows, desktop Microsoft Word and EndNote, with the Cite While You
Write (CWYW) add-in enabled in Word. Have your manuscript, target EndNote library
and installed output style ready. Custom `.ens` styles work too.

To install the skill, you also need Git and
[uv](https://docs.astral.sh/uv/getting-started/installation/). The `uv run`
commands below set up Python and the required dependencies for you. Dependencies
are declared in `pyproject.toml`, with their versions recorded in `uv.lock`.

## Install

Run these commands in a terminal:

```powershell
git clone https://github.com/Breathay/endnote-cwyw.git
cd endnote-cwyw
uv sync --locked
uv run scripts/install_skill.py
```

The default location is `~/.codex/skills/endnote-cwyw`. If you have set
`CODEX_HOME`, the installer uses its `skills/endnote-cwyw` directory instead.

To update, run the following from the repository directory. The installer backs
up your previously installed skill first.

```powershell
git pull --ff-only
uv sync --locked
uv run scripts/install_skill.py --replace
```

## Use it in Codex

Tell Codex which manuscript, library and style to use. For example:

> Use $endnote-cwyw to import the references from manuscript-v12.docx into my
> EndNote library. Save a v13 Word copy and insert citations using my installed
> custom Nature style. Put figure-caption citations after the relevant
> descriptions instead of listing them in a separate Sources paragraph.

For a smaller change:

> Use $endnote-cwyw to add these three papers to this manuscript, which already
> has EndNote citations.

The skill checks which references are already in your library and imports the
missing ones. It then prepares a manuscript copy, uses Word's EndNote add-in to
generate the citations and bibliography, and checks the result. Duplicate
records or conflicting reference details are resolved before formatting.

## Run with uv

Most users can let Codex handle the steps. If you want to inspect a document or
run a step yourself, use these commands from the repository directory.

See the available commands and options:

```powershell
uv run scripts/endnote_workflow.py --help
```

Check the Word add-in connection without changing a document:

```powershell
uv run scripts/endnote_workflow.py probe
```

Inspect citations already present in a manuscript:

```powershell
uv run scripts/endnote_workflow.py inspect manuscript.docx --output work/inspection.json
```

Once references have been matched to the library and temporary citations have
been placed in a new manuscript version, format them with the add-in:

```powershell
uv run scripts/endnote_workflow.py format manuscript-next.docx --style 'C:/path/to/Style.ens'
uv run scripts/endnote_workflow.py verify manuscript-next.docx
```

`format` requires Windows, Word, EndNote and the target library open in EndNote.
It works on a prepared manuscript; it does not determine what each manually
typed reference number in the original document means. For the full sequence
and input-file examples, see the [detailed workflow](references/workflow.md).

`.python-version` selects Python 3.12 by default; Python 3.10 or newer is
supported. To use an existing interpreter, run
`uv sync --locked --python <path-to-python>`.
Include the same `--python <path-to-python>` option in later `uv run` commands.

## Common questions

**Can I keep editing the citations in Word?**

Yes. The installed EndNote add-in creates them, so you can continue adding
references and updating styles through the add-in.

**Can I use my own journal style?**

Yes. Provide the installed `.ens` file, including a Nature style you have
customized yourself.

**What about books and conference papers?**

The included preparation and matching tools currently handle journal articles.
For books and conference papers, organize the records in EndNote and insert
them using the Word add-in.

**What if the add-in is disconnected or citations do not update?**

Start with `probe`, then see [troubleshooting](references/troubleshooting.md).

## License

This project is licensed under [MIT](LICENSE). Word, EndNote and journal style
files are supplied separately and are not included in this repository.
