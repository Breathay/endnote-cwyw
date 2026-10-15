#!/usr/bin/env python3
"""Portable preparation and read-only verification for native EndNote CWYW.

This module does not create native EndNote fields or write EndNote databases.
Only EndNote's own importer and Word CWYW perform those operations.
"""
from __future__ import annotations

import argparse
import base64
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import sqlite3
import sys
import unicodedata
import xml.etree.ElementTree as ET
from zipfile import ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("w", W[1:-1])
TEMP = re.compile(r"\{[^{}]*#\d+[^{}]*\}")


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def normalized(value):
    return "".join(c for c in unicodedata.normalize("NFKC", value).casefold() if c.isalnum())


def doi_key(value):
    return re.sub(r"^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)", "", value.strip(), flags=re.I).casefold()


def require(condition, message):
    if not condition:
        raise ValueError(message)


def load_records(path):
    value = read_json(path)
    rows = value["records"] if isinstance(value, dict) else value
    require(isinstance(rows, list) and rows, "records must be a non-empty list")
    result = []
    seen = set()
    for raw in rows:
        row = dict(raw)
        source_id = row.get("id", row.get("number", ""))
        require(isinstance(source_id, (str, int)) and not isinstance(source_id, bool), "Source reference ID must be text or an integer")
        row["id"] = str(source_id)
        require(row["id"] and row["id"] not in seen, "Missing or duplicate source reference ID")
        seen.add(row["id"])
        for key in ("title", "year"):
            require(isinstance(row.get(key), (str, int)) and str(row[key]).strip(), f"{row['id']}: missing {key}")
            row[key] = str(row[key]).strip()
        require(re.fullmatch(r'\d{4}[a-z]?', row['year'], re.I) is not None, f"{row['id']}: year must contain a four-digit publication year")
        authors = row.get("authors")
        require(isinstance(authors, list) and authors, f"{row['id']}: authors must be a non-empty list")
        for author in authors:
            require(isinstance(author, str) and author.strip(), f"{row['id']}: empty author")
            require(not re.search(r"et\s*al\.?|[?\ufffd\r\n]", author, re.I), f"{row['id']}: incomplete or corrupted author: {author!r}")
            require(not re.search(r"<[^>]+>", author), f"{row['id']}: HTML in author")
        row["authors"] = [a.strip() for a in authors]
        require(not re.search(r"<[^>]+>", row["title"]), f"{row['id']}: HTML in title; normalize it first")
        require(row.get("reference_type", "journal-article") == "journal-article", f"{row['id']}: only verified Journal Article XML mapping is supported; use native export/RIS for other types")
        require(not row.get("label"), f"{row['id']}: Label is intentionally unsupported; a custom style can print it")
        for key in ("journal", "volume", "pages", "doi", "url", "notes"):
            row[key] = str(row.get(key, ""))
        row["doi"] = re.sub(r"^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)", "", row["doi"].strip(), flags=re.I)
        if row["doi"]:
            require(re.fullmatch(r"10\.\d{4,9}/\S+", row["doi"]) is not None, f"{row['id']}: malformed DOI")
        # Preprint identifiers are text, not numeric page ranges.
        if row["journal"].lower().startswith(("arxiv", "biorxiv", "medrxiv")):
            require(not re.search(r"\d{4}\.\d", row["pages"]), f"{row['id']}: put the preprint identifier in journal text or URL, not Pages")
        result.append(row)
    return result


def export_xml(rows, output):
    output = Path(output)
    require(not output.exists(), f"Refusing to overwrite {output}")
    root = ET.Element("xml")
    container = ET.SubElement(root, "records")
    exported = {}
    for row in rows:
        identity = record_identity(row)
        if identity in exported:
            prior = exported[identity]
            require(normalized(prior['title']) == normalized(row['title']) and prior['year'] == row['year'] and prior['authors'] == row['authors'], "Conflicting incoming metadata for " + identity)
            continue
        exported[identity] = row
        record = ET.SubElement(container, "record")
        # XML interchange 17 is Journal Article. Internal SQLite type 0 is a different namespace.
        ET.SubElement(record, "ref-type", {"name": "Journal Article"}).text = "17"
        people = ET.SubElement(ET.SubElement(record, "contributors"), "authors")
        for author in row["authors"]:
            ET.SubElement(people, "author").text = author
        titles = ET.SubElement(record, "titles")
        ET.SubElement(titles, "title").text = row["title"]
        if row["journal"]:
            ET.SubElement(titles, "secondary-title").text = row["journal"]
        ET.SubElement(ET.SubElement(record, "dates"), "year").text = row["year"]
        for key in ("volume", "pages", "notes"):
            if row[key]:
                ET.SubElement(record, key).text = row[key]
        if row["doi"]:
            ET.SubElement(record, "electronic-resource-num").text = row["doi"]
        url = row["url"] or ("https://doi.org/" + row["doi"] if row["doi"] else "")
        if url:
            ET.SubElement(ET.SubElement(ET.SubElement(record, "urls"), "related-urls"), "url").text = url
    ET.indent(root)
    output.parent.mkdir(parents=True, exist_ok=True)
    ET.ElementTree(root).write(output, encoding="utf-8", xml_declaration=True)
    return {"records": len(exported), "source_ids": len(rows), "xml": str(output.resolve()), "journal_article_xml_type": 17}


def library_records(library):
    library = Path(library).resolve()
    if library.suffix.lower() == ".xml":
        root = ET.parse(library).getroot()
        result = []
        for r in root.findall(".//record"):
            result.append({"record_number": int(r.findtext("rec-number", "0")), "title": r.findtext("titles/title", ""),
                           "year": r.findtext("dates/year", ""), "authors": [a.text or "" for a in r.findall("contributors/authors/author")],
                           "doi": r.findtext("electronic-resource-num", ""), "reference_type": r.findtext("ref-type", "")})
        return result
    require(library.suffix.lower() == ".enl", "--library must be an .enl or native EndNote exported XML")
    database = library.with_suffix(".Data") / "sdb" / "sdb.eni"
    require(database.is_file(), "SQLite library layout not found; export native EndNote XML instead")
    try:
        conn = sqlite3.connect(database.as_uri() + "?mode=ro", uri=True)
        try:
            conn.row_factory = sqlite3.Row
            columns = {r[1] for r in conn.execute("pragma table_info(refs)")}
            required = {"id", "title", "year", "author", "electronic_resource_number", "reference_type", "trash_state"}
            require(required <= columns, "Unsupported library schema; use native exported XML")
            return [{"record_number": r["id"], "title": r["title"], "year": r["year"], "authors": r["author"].splitlines(),
                     "doi": r["electronic_resource_number"], "reference_type": r["reference_type"]}
                    for r in conn.execute("select id,title,year,author,electronic_resource_number,reference_type from refs where trash_state=0 order by id")]
        finally:
            conn.close()
    except sqlite3.Error as exc:
        raise ValueError("Cannot read this EndNote database; use native XML export: " + str(exc)) from exc


def map_library(rows, library):
    catalog = library_records(library)
    result = {"library": str(Path(library).resolve()), "mappings": [], "missing": [], "conflicts": []}
    for row in rows:
        by_doi = [r for r in catalog if row["doi"] and doi_key(r["doi"]) == doi_key(row["doi"])]
        candidates = by_doi or [r for r in catalog if normalized(r["title"]) == normalized(row["title"]) and r["year"] == row["year"]]
        if not candidates:
            result["missing"].append(row["id"])
            continue
        if len(candidates) != 1:
            result["conflicts"].append({"id": row["id"], "reason": "ambiguous duplicate records", "record_numbers": [r["record_number"] for r in candidates]})
            continue
        found = candidates[0]
        problems = []
        if normalized(found["title"]) != normalized(row["title"]): problems.append("DOI/title disagreement")
        if found["year"] != row["year"]: problems.append("year disagreement")
        if row["doi"] and found["doi"] and doi_key(row["doi"]) != doi_key(found["doi"]): problems.append("DOI disagreement")
        if found["authors"] != row["authors"]: problems.append("author list differs; review before linking")
        if found["record_number"] <= 0: problems.append("missing native record number")
        if found["reference_type"] not in (0, "17"): problems.append("reference type differs from Journal Article; check the native import")
        if problems:
            result["conflicts"].append({"id": row["id"], "reason": "; ".join(problems), "record_numbers": [found["record_number"]]})
            continue
        result["mappings"].append({"id": row["id"], **found})
    return result


def read_docx(path):
    with ZipFile(path) as archive:
        require(archive.testzip() is None, "Corrupted DOCX archive")
        infos = archive.infolist()
        parts = {i.filename: archive.read(i.filename) for i in infos}
    return infos, parts, ET.fromstring(parts["word/document.xml"])


def paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def field_kind(code):
    match = re.match(r"\s*ADDIN\s+(EN\.[A-Z.]+)\b", code)
    return match[1] if match else "other"


def field_xml(field):
    if "<EndNote>" in field["code"]:
        return ET.fromstring(field["code"][field["code"].index("<EndNote>"):].strip())
    blobs = list(field["data"])
    def visit(children):
        for child in children:
            blobs.extend(child["data"])
            visit(child["children"])
    visit(field["children"])
    for blob in blobs:
        if blob.lstrip().startswith(b"<EndNote>"):
            # Native CWYW sometimes stores a NUL-terminated UTF-8 payload.
            # Remove the terminator, never repair malformed XML silently.
            return ET.fromstring(blob.rstrip(b'\x00'))
    raise ValueError("Native citation field has no readable EndNote payload")


def record_identity(record):
    if record["doi"]:
        return "doi:" + doi_key(record["doi"])
    return "title:" + normalized(record["title"]) + "|" + record["year"]


def scan_docx(root):
    paragraphs = list(root.iter(W + "p"))
    pids = {id(p): i for i, p in enumerate(paragraphs)}
    stack, fields = [], []
    free = ["" for _ in paragraphs]
    protected = set()
    serial = 0
    current_p = -1
    def visit(node):
        nonlocal serial, current_p
        prior = current_p
        if node.tag == W + "p": current_p = pids[id(node)]
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                serial += 1
                stack.append({"code": "", "data": [], "children": [], "result": "", "phase": "code", "paragraph": current_p, "serial": serial,
                              "plain_offset": len(free[current_p]) if current_p >= 0 else 0})
            elif kind == "separate":
                require(bool(stack), "Unbalanced Word field separator")
                stack[-1]["phase"] = "result"
            elif kind == "end":
                require(bool(stack), "Unbalanced Word field end")
                field = stack.pop(); fields.append(field)
                if stack: stack[-1]["children"].append(field)
        elif node.tag == W + "instrText" and stack: stack[-1]["code"] += node.text or ""
        elif node.tag == W + "fldData" and stack: stack[-1]["data"].append(base64.b64decode(node.text or "", validate=False))
        elif node.tag == W + "t":
            if stack: protected.add(id(node))
            elif current_p >= 0: free[current_p] += node.text or ""
            for field in stack:
                if field["phase"] == "result": field["result"] += node.text or ""
        # fldSimple isn't emitted by the validated CWYW path. Fail closed if it
        # contains EndNote data, rather than silently reporting no citations.
        elif node.tag == W + "fldSimple":
            require("EN.CITE" not in node.get(W + "instr", ""), "Unsupported simple EndNote field; inspect in native Word")
        for child in node: visit(child)
        if node.tag == W + "p":
            for field in stack:
                if field["phase"] == "result": field["result"] += "\n"
        current_p = prior
    visit(root)
    require(not stack, "Unbalanced Word fields")
    cites = []
    for field in sorted(fields, key=lambda f: f["serial"]):
        field["kind"] = field_kind(field["code"])
        if field["kind"] != "EN.CITE": continue
        data = field_xml(field)
        targets = []
        for cite in data.findall("Cite"):
            r = cite.find("record")
            require(r is not None, "Native citation is missing its traveling record")
            targets.append({"record_number": int(cite.findtext("RecNum", r.findtext("rec-number", "0"))),
                            "title": r.findtext("titles/title", ""), "year": r.findtext("dates/year", ""),
                            "authors": [a.text or "" for a in r.findall("contributors/authors/author")],
                            "doi": r.findtext("electronic-resource-num", "")})
        require(targets, "Native citation field has no targets")
        cites.append({"paragraph": field["paragraph"], "plain_offset": field["plain_offset"], "display": field["result"], "targets": targets})
    return {"paragraphs": paragraphs, "free": free, "fields": fields, "citations": cites, "protected": protected}


def inspect_docx(path):
    _, parts, root = read_docx(path)
    scan = scan_docx(root)
    return {"source": str(Path(path).resolve()), "source_sha256": sha256(path),
            "paragraphs": [{"index": i, "text": paragraph_text(p)} for i, p in enumerate(scan["paragraphs"])],
            "native_citations": len(scan["citations"]), "native_bibliographies": sum(f["kind"] == "EN.REFLIST" for f in scan["fields"]),
            "citations": scan["citations"], "drawings": len(list(root.iter(W + "drawing"))),
            "tables": len(list(root.iter(W + "tbl"))), "media_parts": sorted(n for n in parts if n.startswith("word/media/"))}


def replace_span(paragraph, start, end, new, protected):
    require(not any(n.tag == W+'p' for n in list(paragraph.iter())[1:]), "Target the innermost paragraph when a text box is present")
    require(not any(n.tag.endswith('}oMath') for n in paragraph.iter()), "Use native Word range editing for a paragraph with equations")
    nodes = list(paragraph.iter(W + "t"))
    text = paragraph_text(paragraph)
    require(0 <= start <= end <= len(text), "Text range is outside the paragraph")
    positions = []
    offset = 0
    for node in nodes:
        value = node.text or ""
        positions.append((node, offset, offset + len(value), value))
        offset += len(value)
    if start == end:
        for node, a, b, value in positions:
            if a <= start < b and id(node) not in protected:
                node.text = value[:start-a] + new + value[start-a:]
                node.set(XML_SPACE, "preserve")
                return
        require(not any(a < start < b and id(node) in protected for node,a,b,_ in positions), "Insertion falls inside an existing field")
        # Appending is safe only if the paragraph itself contains no fields.
        require(start == len(text) and not any(id(n) in protected for n in nodes) and not any(n.tag == W+'fldChar' for n in paragraph.iter()), "Use native Word for insertion at an existing field boundary")
        run = ET.SubElement(paragraph, W + "r")
        ET.SubElement(run, W + "t", {XML_SPACE: "preserve"}).text = new
        return
    touched = [(n,a,b,v) for n,a,b,v in positions if b > start and a < end]
    require(touched and not any(id(n) in protected for n,_,_,_ in touched), "Replacement intersects an existing Word field")
    # Don't cross invisible fields, equations or drawings between visible runs.
    if any(n.tag in (W+'fldChar', W+'drawing') for n in paragraph.iter()):
        require(len(touched) == 1, "Use native Word for a multi-run edit in a paragraph containing fields or drawings")
    first, a, _, value = touched[0]
    last, la, _, lv = touched[-1]
    prefix, suffix = value[:start-a], lv[end-la:]
    first.text = prefix + new + (suffix if first is last else "")
    for node,_,_,_ in touched[1:]: node.text = ""
    if first is not last: last.text = suffix
    for node,_,_,_ in touched: node.set(XML_SPACE, "preserve")


def temporary_citation(ids, mapping):
    values = []
    seen = set()
    for key in ids:
        key = str(key)
        require(key in mapping, f"Unmapped source reference: {key}")
        row = mapping[key]
        if row['record_number'] in seen: continue
        seen.add(row['record_number'])
        surname = row["authors"][0].split(",")[0]
        require(not re.search(r"[{}#;\\\r\n]", surname + row["year"]), "Unsafe temporary citation author/year")
        values.append(f"{surname}, {row['year']} #{row['record_number']}")
    require(values, "Citation replacement needs at least one source reference")
    return "{" + "; ".join(values) + "}"


def prepare_docx(source, output, mapping_path, plan_path, audit_path):
    source, output = Path(source).resolve(), Path(output).resolve()
    require(source != output and not output.exists(), "Output must be a new DOCX version")
    require(output.suffix.lower() == ".docx", "Output must be .docx")
    plan, mapping = read_json(plan_path), read_json(mapping_path)
    require(not mapping["missing"] and not mapping["conflicts"], "Resolve missing/ambiguous library records before authoring")
    require(plan["source_sha256"] == sha256(source), "Source changed since the plan was prepared")
    lookup = {str(r["id"]): r for r in mapping["mappings"]}
    infos, parts, root = read_docx(source)
    original_xml = parts['word/document.xml']
    scan = scan_docx(root)
    expected = [sorted(record_identity(t) for t in c["targets"]) for c in scan["citations"]]
    operations = {}
    for raw in plan.get("replacements", []):
        op = dict(raw)
        pi, start, end = int(op["paragraph"]), int(op["start"]), int(op["end"])
        require(0 <= pi < len(scan["paragraphs"]), "Unknown paragraph in edit plan")
        require(paragraph_text(scan["paragraphs"][pi])[start:end] == op["old"], "Planned text does not match the source")
        require(("references" in op) != ("new_text" in op), "Use either references or new_text in each replacement")
        if "references" in op:
            op["new_text"] = temporary_citation(op["references"], lookup)
            expected.append(sorted({record_identity(lookup[str(key)]) for key in op["references"]}))
        operations.setdefault(pi, []).append((start,end,op["new_text"]))
    deletions = {int(r["paragraph"]): r["old"] for r in plan.get("remove_paragraphs", [])}
    require(not (set(deletions) & set(operations)), "Do not edit and delete the same paragraph")
    for pi, edits in operations.items():
        edits.sort()
        require(all(a[1] <= b[0] and a[0] != b[0] for a,b in zip(edits, edits[1:])), "Overlapping replacements")
        for start,end,new in reversed(edits):
            replace_span(scan["paragraphs"][pi],start,end,new,scan["protected"])
    parents = {id(child): parent for parent in root.iter() for child in parent}
    for pi, text in deletions.items():
        require(0 <= pi < len(scan["paragraphs"]), "Unknown paragraph deletion")
        p = scan["paragraphs"][pi]
        require(paragraph_text(p) == text, "Deleted paragraph text differs from the plan")
        require(not any(n.tag in (W+'drawing', W+'sectPr', W+'fldChar') or n.tag.endswith('}oMath') for n in p.iter()), "Cannot remove a paragraph containing a figure, section, equation or field")
        require(not any(id(n) in scan["protected"] for n in p.iter(W+'t')), "Cannot remove a paragraph inside a live bibliography")
        parents[id(p)].remove(p)
    after = scan_docx(root)
    expected_free = [TEMP.sub("", value) for value in after["free"] if TEMP.sub("",value).strip()]
    spans = {i:list(TEMP.finditer(value)) for i,value in enumerate(after['free'])}
    def without_markers(pi, offset):
        return offset - sum(m.end()-m.start() for m in spans[pi] if m.end() <= offset)
    locations = [[c['paragraph'],without_markers(c['paragraph'],c['plain_offset']),
                  sorted(record_identity(t) for t in c['targets'])] for c in after['citations']]
    by_number = {r['record_number']:r for r in lookup.values()}
    for pi,matches in spans.items():
        for match in matches:
            numbers = [int(n) for n in re.findall(r'#(\d+)',match.group())]
            require(all(n in by_number for n in numbers), 'An existing temporary citation has no verified mapping')
            locations.append([pi,without_markers(pi,match.start()),sorted({record_identity(by_number[n]) for n in numbers})])
    parts["word/document.xml"] = ET.tostring(root,encoding="utf-8",xml_declaration=True)
    # ElementTree can rename prefixes referenced by mc:Ignorable. Restore
    # declarations from the source before serializing with a prefix-aware tool.
    # Refuse affected files instead of producing an invalid Office package.
    original_root = ET.fromstring(original_xml)
    ignorable = original_root.get("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable", "")
    if ignorable:
        try:
            from lxml import etree as LE
        except ImportError as exc:
            raise ValueError("DOCX uses mc:Ignorable; install lxml or use the bundled document runtime for prefix-safe authoring") from exc
        rebuilt = LE.fromstring(parts["word/document.xml"])
        original = LE.fromstring(original_xml)
        restored = LE.Element(original.tag, nsmap={**rebuilt.nsmap, **original.nsmap})
        restored.attrib.update(rebuilt.attrib)
        for child in rebuilt: restored.append(child)
        parts["word/document.xml"] = LE.tostring(restored,encoding="utf-8",xml_declaration=True,standalone=True)
    output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(output,"w") as archive:
        for info in infos: archive.writestr(info,parts[info.filename])
    report = {"source": str(source), "source_sha256": plan["source_sha256"], "prepared": str(output),
              "expected_citation_groups": expected, "expected_plain_paragraphs": expected_free,
              "expected_citation_locations": sorted(locations),
              "drawings": len(list(root.iter(W+'drawing'))), "tables": len(list(root.iter(W+'tbl'))),
              "media_sha256": {n:hashlib.sha256(b).hexdigest() for n,b in parts.items() if n.startswith('word/media/')},
              "status": "temporary citations only; native CWYW formatting still required"}
    write_json(audit_path,report)
    return {"output": str(output), "expected_citations": len(expected), "audit": str(Path(audit_path).resolve())}


def verify_docx(document, audit_path=None, expected_citations=None, expected_references=None):
    _, parts, root = read_docx(document)
    scan = scan_docx(root)
    errors = []
    cites = scan["citations"]
    ref_fields = [f for f in scan["fields"] if f["kind"] == "EN.REFLIST"]
    identities = {record_identity(t) for c in cites for t in c["targets"]}
    if len(ref_fields) != 1: errors.append(f"Expected one native bibliography field; found {len(ref_fields)}")
    if ref_fields and not ref_fields[0]["result"].strip(): errors.append("Bibliography is empty or still being written")
    if any(not c['display'].strip() for c in cites): errors.append('A native citation has an empty displayed result')
    if TEMP.search("".join(paragraph_text(p) for p in scan["paragraphs"])): errors.append("Unresolved temporary citations remain")
    if expected_citations is not None and len(cites) != expected_citations: errors.append("Citation group count differs")
    if expected_references is not None and len(identities) != expected_references: errors.append("Distinct cited-reference count differs")
    numbering_checked = False
    numeric = re.compile(r"\s*\d+(?:\s*[,;–—-]\s*\d+)*\s*")
    if cites and ref_fields and all(numeric.fullmatch(c['display']) for c in cites):
        rows = re.findall(r"(?m)^\s*(\d+)\.[\t ]*(.+)$", ref_fields[0]['result'])
        number_map = {}
        targets = {record_identity(t): t for c in cites for t in c['targets']}
        for number, row in rows:
            candidates = [t for t in targets.values() if normalized(t['title']) and normalized(t['title']) in normalized(row)]
            if not candidates:
                errors.append(f"Bibliography row {number} has no matching traveling record")
                continue
            target = max(candidates, key=lambda t:len(t['title']))
            identity = record_identity(target)
            if identity in number_map: errors.append("Duplicate bibliography entry: " + identity)
            number_map[identity] = int(number)
        if set(number_map) != identities or set(number_map.values()) != set(range(1,len(identities)+1)):
            errors.append("Numbered bibliography is incomplete or cannot be parsed")
        else:
            for cite in cites:
                displayed=set()
                for token in re.split(r'[,;]',cite['display']):
                    bounds=re.split(r'[–—-]',token.strip())
                    a=int(bounds[0]);b=int(bounds[-1])
                    if not 1 <= a <= b <= len(identities):
                        errors.append("Displayed citation number is outside the bibliography")
                        continue
                    displayed.update(range(a,b+1))
                if displayed != {number_map[record_identity(t)] for t in cite['targets']}:
                    errors.append("Displayed citation numbers point to different bibliography entries")
            numbering_checked=True
    if audit_path:
        audit = read_json(audit_path)
        if sha256(audit["source"]) != audit["source_sha256"]: errors.append("Original source changed")
        actual = Counter(tuple(sorted(record_identity(t) for t in c["targets"])) for c in cites)
        planned = Counter(tuple(row) for row in audit["expected_citation_groups"])
        if actual != planned: errors.append("Citation targets differ from the edit plan")
        if 'expected_citation_locations' in audit:
            locations=sorted([c['paragraph'],c['plain_offset'],sorted(record_identity(t) for t in c['targets'])] for c in cites)
            if locations != audit['expected_citation_locations']: errors.append('Citation locations differ from the edit plan')
        plain = [t for t in scan["free"] if t.strip()]
        if plain != audit["expected_plain_paragraphs"]: errors.append("Non-citation prose differs from the plan")
        if len(list(root.iter(W+'drawing'))) != audit['drawings']: errors.append("Figure count changed")
        if len(list(root.iter(W+'tbl'))) != audit['tables']: errors.append("Table count changed")
        for name, digest in audit["media_sha256"].items():
            if name not in parts or hashlib.sha256(parts[name]).hexdigest()!=digest: errors.append("Embedded image changed: "+name)
    style = None
    if 'word/settings.xml' in parts:
        settings=ET.fromstring(parts['word/settings.xml'])
        for variable in settings.iter(W+'docVar'):
            if variable.get(W+'name')=='EN.Layout':
                style=ET.fromstring(variable.get(W+'val')).findtext('Style')
    return {"document":str(Path(document).resolve()),"sha256":sha256(document),"passed":not errors,"errors":errors,
            "native_citations":len(cites),"native_bibliographies":len(ref_fields),"distinct_references":len(identities),
            "numbering_checked":numbering_checked,"style":style}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    sub=parser.add_subparsers(dest='command',required=True)
    p=sub.add_parser('inspect');p.add_argument('document');p.add_argument('--output',required=True)
    p=sub.add_parser('export');p.add_argument('records');p.add_argument('--output',required=True)
    p=sub.add_parser('map');p.add_argument('records');p.add_argument('--library',required=True);p.add_argument('--output',required=True);p.add_argument('--missing-xml')
    p=sub.add_parser('prepare');p.add_argument('source');p.add_argument('--output',required=True);p.add_argument('--mapping',required=True);p.add_argument('--plan',required=True);p.add_argument('--audit',required=True)
    p=sub.add_parser('verify');p.add_argument('document');p.add_argument('--audit');p.add_argument('--expected-citations',type=int);p.add_argument('--expected-references',type=int);p.add_argument('--output')
    args=parser.parse_args()
    try:
        if args.command=='inspect': result=inspect_docx(args.document);write_json(args.output,result);result={k:v for k,v in result.items() if k not in ('paragraphs','citations','media_parts')}
        elif args.command=='export':result=export_xml(load_records(args.records),args.output)
        elif args.command=='map':
            rows=load_records(args.records);result=map_library(rows,args.library);write_json(args.output,result)
            if args.missing_xml and result['missing']:
                export_xml([r for r in rows if r['id'] in result['missing']],args.missing_xml)
            result={'mapped':len(result['mappings']),'missing':result['missing'],'conflicts':result['conflicts'],'output':args.output}
        elif args.command=='prepare':result=prepare_docx(args.source,args.output,args.mapping,args.plan,args.audit)
        else:
            result=verify_docx(args.document,args.audit,args.expected_citations,args.expected_references)
            if args.output:write_json(args.output,result)
        print(json.dumps(result,ensure_ascii=False,indent=2))
        return 1 if result.get('passed') is False else 0
    except (ValueError,KeyError,ET.ParseError,OSError) as exc:
        print(json.dumps({'error':str(exc)},ensure_ascii=False),file=sys.stderr)
        return 2


if __name__=='__main__':
    raise SystemExit(main())
