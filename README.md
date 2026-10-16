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

## Install

Give Codex the repository link and ask it to install the skill:

> Install the endnote-cwyw skill from https://github.com/Breathay/endnote-cwyw
> in my personal skills directory.

Codex checks for a suitable Python environment and reuses it when available,
installing missing dependencies as needed. You do not need to install uv or
create a virtual environment yourself.

If the skill is already installed, ask Codex to update it. The previous version
is backed up first. For manual installation or individual commands, see the
[detailed workflow](references/workflow.md).

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

Ask Codex to check the Word and EndNote connection. See
[troubleshooting](references/troubleshooting.md) for details.

## License

This project is licensed under [MIT](LICENSE). Word, EndNote and journal style
files are supplied separately and are not included in this repository.
