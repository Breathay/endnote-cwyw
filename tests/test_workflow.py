"""Offline contract tests. Synthetic XML fields test the parser only, not CWYW."""
import base64
import contextlib
from copy import deepcopy
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('endnote_workflow',ROOT/'scripts/endnote_workflow.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
W=m.W


def fixture_doc(path,paragraphs,media=False):
    root=ET.Element(W+'document');body=ET.SubElement(root,W+'body')
    for paragraph in paragraphs:
        if isinstance(paragraph,str):
            p=ET.SubElement(body,W+'p');ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text=paragraph
        else:body.append(paragraph)
    with ZipFile(path,'w') as z:
        z.writestr('word/document.xml',ET.tostring(root,encoding='utf-8',xml_declaration=True))
        if media:z.writestr('word/media/synthetic.png',b'synthetic test bytes, not a real image')


def payload(records):
    root=ET.Element('EndNote')
    for r in records:
        cite=ET.SubElement(root,'Cite');ET.SubElement(cite,'RecNum').text=str(r['record_number'])
        record=ET.SubElement(cite,'record');ET.SubElement(record,'rec-number').text=str(r['record_number'])
        titles=ET.SubElement(record,'titles');ET.SubElement(titles,'title').text=r['title']
        ET.SubElement(ET.SubElement(record,'dates'),'year').text=r['year']
        authors=ET.SubElement(ET.SubElement(record,'contributors'),'authors')
        for author in r['authors']:ET.SubElement(authors,'author').text=author
        if r.get('doi'):ET.SubElement(record,'electronic-resource-num').text=r['doi']
    return ET.tostring(root,encoding='unicode')


def field_runs(p,code,result,data=None):
    begin=ET.SubElement(ET.SubElement(p,W+'r'),W+'fldChar',{W+'fldCharType':'begin'})
    if data:ET.SubElement(begin,W+'fldData').text=base64.b64encode(data.encode()).decode()
    ET.SubElement(ET.SubElement(p,W+'r'),W+'instrText').text=code
    ET.SubElement(ET.SubElement(p,W+'r'),W+'fldChar',{W+'fldCharType':'separate'})
    ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text=result
    ET.SubElement(ET.SubElement(p,W+'r'),W+'fldChar',{W+'fldCharType':'end'})


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.base=Path(self.temp.name)
        self.rows=m.load_records(ROOT/'examples/records.json')
        self.targets=[{**r,'record_number':n} for r,n in zip(self.rows,[35,91])]

    def dump(self,name,value):
        p=self.base/name;m.write_json(p,value);return p

    def make_database(self,records=None):
        library=self.base/'Test.enl';library.write_bytes(b'test')
        db=library.with_suffix('.Data')/'sdb/sdb.eni';db.parent.mkdir(parents=True)
        c=sqlite3.connect(db)
        c.execute('create table refs(id integer primary key, trash_state integer, reference_type integer, author text, year text, title text, electronic_resource_number text)')
        for r in records or self.targets:
            c.execute('insert into refs values(?,?,?,?,?,?,?)',(r['record_number'],0,r.get('reference_type',0),'\r'.join(r['authors']),r['year'],r['title'],r['doi']))
        c.commit();c.close();return library,db

    def test_export_uses_interchange_type_17_and_no_tracking_fields(self):
        self.rows[0]['authors'].append('Vučković, Jelena')
        output=self.base/'import.xml';m.export_xml(self.rows,output)
        root=ET.parse(output).getroot()
        self.assertEqual(root.findtext('records/record/ref-type'),'17')
        self.assertEqual(root.find('records/record/ref-type').get('name'),'Journal Article')
        self.assertFalse(root.findall('.//label'));self.assertFalse(root.findall('.//short-title'))
        self.assertIn('Vučković, Jelena',output.read_text(encoding='utf-8'))
        with self.assertRaises(ValueError):m.export_xml(self.rows,output)

    def test_corrupt_or_truncated_authors_are_rejected(self):
        for author in ['Example, Alice et al.','Vu?kovi?, Jelena','Example, A\nOther, B']:
            rows=deepcopy(self.rows);rows[0]['authors']=[author]
            with self.assertRaises(ValueError):m.load_records(self.dump('invalid.json',rows))

    def test_preprint_identifier_is_not_a_page_range(self):
        rows=deepcopy(self.rows);rows[0].update(journal='bioRxiv',pages='2026.01.01.123456v1')
        with self.assertRaises(ValueError):m.load_records(self.dump('invalid.json',rows))

    def test_duplicate_incoming_work_exports_once_and_conflicts_fail(self):
        rows=deepcopy(self.rows[:1]);alias=deepcopy(rows[0]);alias['id']='alias';rows.append(alias)
        output=self.base/'deduplicated.xml'
        self.assertEqual(m.export_xml(rows,output)['records'],1)
        rows[0]['doi']='10.1234/example';rows[1]['doi']='10.1234/example';rows[1]['title']='Conflicting title'
        with self.assertRaises(ValueError):m.export_xml(rows,self.base/'conflict.xml')

    def test_aliases_do_not_duplicate_a_temporary_citation(self):
        alias={**self.targets[0],'id':'alias'}
        mapping={r['id']:r for r in [self.targets[0],alias]}
        self.assertEqual(m.temporary_citation(['paper-a','alias'],mapping).count('#35'),1)

    def test_mapping_reads_actual_nonsequential_ids_without_writing(self):
        library,db=self.make_database();before=m.sha256(db)
        mapped=m.map_library(self.rows,library)
        self.assertEqual([r['record_number'] for r in mapped['mappings']],[35,91])
        self.assertFalse(mapped['missing']);self.assertFalse(mapped['conflicts'])
        self.assertEqual(before,m.sha256(db))
        self.assertIn('#91',m.temporary_citation(['paper-b'],{r['id']:r for r in mapped['mappings']}))

    def test_duplicate_doi_is_not_arbitrarily_selected(self):
        rows=deepcopy(self.rows);rows[0]['doi']='10.1234/synthetic'
        records=[{**rows[0],'record_number':35},{**rows[0],'record_number':36},self.targets[1]]
        library,_=self.make_database(records)
        mapped=m.map_library(rows,library)
        self.assertEqual(len(mapped['conflicts']),1)
        self.assertEqual(mapped['conflicts'][0]['record_numbers'],[35,36])

    def test_doi_title_mismatch_and_unknown_type_are_conflicts(self):
        rows=deepcopy(self.rows);rows[0]['doi']='10.1234/synthetic'
        first={**rows[0],'title':'An entirely different paper','record_number':35,'reference_type':-1}
        library,_=self.make_database([first,self.targets[1]])
        result=m.map_library(rows,library)
        self.assertIn('DOI/title disagreement',result['conflicts'][0]['reason'])
        self.assertIn('reference type',result['conflicts'][0]['reason'])

    def test_native_export_mapping_uses_record_numbers(self):
        root=ET.Element('xml');container=ET.SubElement(root,'records')
        for target in self.targets:
            r=ET.fromstring(payload([target])).find('Cite/record')
            ET.SubElement(r,'ref-type',{'name':'Journal Article'}).text='17';container.append(r)
        path=self.base/'native.xml';ET.ElementTree(root).write(path,encoding='utf-8')
        self.assertEqual(len(m.map_library(self.rows,path)['mappings']),2)

    def test_multirun_replacement_preserves_source_formatting_and_media(self):
        p=ET.Element(W+'p')
        first=ET.SubElement(p,W+'r');ET.SubElement(ET.SubElement(first,W+'rPr'),W+'b');ET.SubElement(first,W+'t').text='A bold opening. '
        for text in ['Evidence (ref. ','1',') remains.']:ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text=text
        source=self.base/'source.docx';fixture_doc(source,[p,'References','1. Manual entry.'],media=True)
        mapping=self.dump('map.json',{'mappings':self.targets,'missing':[],'conflicts':[]})
        text=m.paragraph_text(p);old='(ref. 1)';start=text.index(old)
        plan=self.dump('plan.json',{'source_sha256':m.sha256(source),'replacements':[{'paragraph':0,'start':start,'end':start+len(old),'old':old,'references':['paper-a']}],
                                  'remove_paragraphs':[{'paragraph':2,'old':'1. Manual entry.'}]})
        before=m.sha256(source);out=self.base/'next.docx';audit=self.base/'audit.json'
        m.prepare_docx(source,out,mapping,plan,audit)
        _,parts,root=m.read_docx(out)
        self.assertIn('{Example, 2026 #35}',m.paragraph_text(list(root.iter(W+'p'))[0]))
        self.assertIsNotNone(root.find('.//'+W+'b'))
        self.assertEqual(parts['word/media/synthetic.png'],b'synthetic test bytes, not a real image')
        self.assertEqual(before,m.sha256(source))
        with self.assertRaises(ValueError):m.prepare_docx(source,out,mapping,plan,audit)

    def test_stale_plan_and_unmapped_reference_are_rejected(self):
        source=self.base/'source.docx';fixture_doc(source,['Evidence (ref. 1).'])
        mapping=self.dump('map.json',{'mappings':self.targets,'missing':[],'conflicts':[]})
        plan=self.dump('plan.json',{'source_sha256':'wrong','replacements':[]})
        with self.assertRaises(ValueError):m.prepare_docx(source,self.base/'next.docx',mapping,plan,self.base/'audit.json')
        with self.assertRaises(ValueError):m.temporary_citation(['unknown'],{r['id']:r for r in self.targets})

    def make_formatted(self,display='1',binary=False):
        p=ET.Element(W+'p');ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text='Evidence'
        xml=payload([self.targets[0]])
        field_runs(p,' ADDIN EN.CITE ' + ('' if binary else xml),display,(xml+'\x00') if binary else None)
        bibliography=ET.Element(W+'p');field_runs(bibliography,' ADDIN EN.REFLIST ','1. Example, A. '+self.rows[0]['title']+'. (2026).')
        path=self.base/('binary.docx' if binary else 'formatted.docx');fixture_doc(path,[p,'References',bibliography])
        return path

    def test_native_field_parser_and_numbering(self):
        for binary in [False,True]:
            path=self.make_formatted(binary=binary)
            result=m.verify_docx(path,expected_citations=1,expected_references=1)
            self.assertTrue(result['passed'],result)
            self.assertTrue(result['numbering_checked'])

    def test_wrong_display_number_is_rejected(self):
        result=m.verify_docx(self.make_formatted(display='2'))
        self.assertFalse(result['passed'])

    def test_swapped_citation_locations_fail_even_when_counts_match(self):
        paragraphs=[]
        for i,target in enumerate(self.targets,1):
            p=ET.Element(W+'p');ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text='Claim '
            field_runs(p,' ADDIN EN.CITE '+payload([target]),str(i))
            ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text='.';paragraphs.append(p)
        bibliography=ET.Element(W+'p')
        rows='\n'.join(f'{i}. Example. {r["title"]}. ({r["year"]}).' for i,r in enumerate(self.rows,1))
        field_runs(bibliography,' ADDIN EN.REFLIST ',rows)
        source=self.base/'source.docx';fixture_doc(source,paragraphs+['References',bibliography])
        mapping=self.dump('map.json',{'mappings':[],'missing':[],'conflicts':[]})
        plan=self.dump('plan.json',{'source_sha256':m.sha256(source),'replacements':[]})
        audit=self.base/'audit.json';out=self.base/'next.docx';m.prepare_docx(source,out,mapping,plan,audit)
        self.assertTrue(m.verify_docx(out,audit)['passed'])
        infos,parts,root=m.read_docx(out);body=root.find(W+'body');first,second=list(body)[:2]
        body.remove(second);body.remove(first);body.insert(0,second);body.insert(1,first)
        parts['word/document.xml']=ET.tostring(root,encoding='utf-8')
        with ZipFile(out,'w') as z:
            for info in infos:z.writestr(info,parts[info.filename])
        result=m.verify_docx(out,audit)
        self.assertFalse(result['passed']);self.assertIn('Citation locations differ from the edit plan',result['errors'])

    def test_empty_citation_result_is_rejected(self):
        self.assertFalse(m.verify_docx(self.make_formatted(display=''))['passed'])

    def test_existing_native_field_cannot_be_overwritten(self):
        source=self.make_formatted();mapping=self.dump('map.json',{'mappings':self.targets,'missing':[],'conflicts':[]})
        plan=self.dump('plan.json',{'source_sha256':m.sha256(source),'replacements':[{'paragraph':0,'start':8,'end':9,'old':'1','references':['paper-b']}]})
        with self.assertRaises(ValueError):m.prepare_docx(source,self.base/'next.docx',mapping,plan,self.base/'audit.json')

    def test_mc_ignorable_namespace_is_preserved(self):
        try:import lxml  # noqa: F401
        except ImportError:self.skipTest('lxml authoring dependency not installed')
        source=self.base/'modern.docx'
        xml=(f'<w:document xmlns:w="{W[1:-1]}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
             'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body><w:p><w:r><w:t>Text.</w:t></w:r></w:p></w:body></w:document>')
        with ZipFile(source,'w') as z:z.writestr('word/document.xml',xml)
        mapping=self.dump('map.json',{'mappings':self.targets,'missing':[],'conflicts':[]})
        plan=self.dump('plan.json',{'source_sha256':m.sha256(source),'replacements':[]})
        out=self.base/'next.docx';m.prepare_docx(source,out,mapping,plan,self.base/'audit.json')
        from lxml import etree as LE
        with ZipFile(out) as z:root=LE.fromstring(z.read('word/document.xml'))
        self.assertEqual(root.nsmap['w14'],'http://schemas.microsoft.com/office/word/2010/wordml')


class WordCommandTests(unittest.TestCase):
    def test_requires_windows_before_launching_a_process(self):
        with patch.object(m.sys, 'platform', 'linux'), patch.object(m.subprocess, 'run') as launch:
            with self.assertRaisesRegex(ValueError, 'requires Windows'):
                m.run_word('Probe')
            launch.assert_not_called()

    def test_rejects_invalid_parameters_before_launch(self):
        with patch.object(m.sys, 'platform', 'win32'), patch.object(m.subprocess, 'run') as launch:
            for arguments in [dict(timeout_seconds=0), dict(expected_citations=-1), dict(expected_references=-2)]:
                with self.assertRaises(ValueError):m.run_word('Probe', **arguments)
            with self.assertRaises(ValueError):m.run_word('Format')
            launch.assert_not_called()

    def native_environment(self):
        folder = tempfile.TemporaryDirectory();self.addCleanup(folder.cleanup)
        engine = Path(folder.name) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
        engine.parent.mkdir(parents=True);engine.touch()
        return folder.name, engine

    def test_format_quotes_paths_and_forwards_all_options(self):
        root, engine = self.native_environment()
        document = "C:/论文/O'Brien $draft `1.docx"
        style = "C:/Styles/O'Brien.ens"
        completed = m.subprocess.CompletedProcess([], 0, '{"action":"format","saved":true}', '')
        with patch.object(m.sys, 'platform', 'win32'), patch.dict(os.environ, {'SystemRoot':root}), \
                patch.object(m.subprocess, 'run', return_value=completed) as launch:
            self.assertTrue(m.run_word('Format', document, style, 25, 18, 90, 'work/final.pdf')['saved'])
        arguments = launch.call_args.args[0]
        self.assertEqual(arguments[:4], [str(engine), '-NoProfile', '-NonInteractive', '-EncodedCommand'])
        command = base64.b64decode(arguments[4]).decode('utf-16-le')
        self.assertTrue(command.endswith(" -Action 'Format' -Document 'C:/论文/O''Brien $draft `1.docx'"
                                         " -StylePath 'C:/Styles/O''Brien.ens' -ExpectedCitations '25'"
                                         " -ExpectedReferences '18' -TimeoutSeconds '90' -PdfPath 'work/final.pdf'"))
        self.assertNotIn('shell', launch.call_args.kwargs)

    def test_native_errors_are_reported_without_retry(self):
        root, _ = self.native_environment()
        completed = m.subprocess.CompletedProcess([], 1, '', 'Enable the EndNote add-in in Word.')
        with patch.object(m.sys, 'platform', 'win32'), patch.dict(os.environ, {'SystemRoot':root}), \
                patch.object(m.subprocess, 'run', return_value=completed) as launch:
            with self.assertRaisesRegex(ValueError, 'Enable the EndNote'):
                m.run_word('Probe')
            self.assertEqual(launch.call_count, 1)

    def test_unexpected_native_response_is_rejected(self):
        root, _ = self.native_environment()
        with patch.object(m.sys, 'platform', 'win32'), patch.dict(os.environ, {'SystemRoot':root}):
            for output in ['not JSON', '[]', '{"action":"format"}']:
                completed = m.subprocess.CompletedProcess([], 0, output, '')
                with patch.object(m.subprocess, 'run', return_value=completed), self.assertRaises(ValueError):
                    m.run_word('Probe')


class SkillInstallTests(unittest.TestCase):
    def test_upgrade_copies_uv_files_and_backs_up_legacy_requirements(self):
        spec=importlib.util.spec_from_file_location('install_skill',ROOT/'scripts/install_skill.py')
        installer=importlib.util.module_from_spec(spec);spec.loader.exec_module(installer)
        with tempfile.TemporaryDirectory() as folder:
            codex_home=Path(folder);destination=codex_home/'skills/endnote-cwyw'
            destination.mkdir(parents=True)
            (destination/'requirements.txt').write_text('lxml>=5.0\n',encoding='utf-8')
            (destination/'.venv').mkdir()
            (destination/'.venv/cache').write_bytes(b'test runtime cache')
            with patch.dict(os.environ, {'CODEX_HOME':folder}), \
                    patch.object(sys, 'argv', ['install_skill.py','--replace']), contextlib.redirect_stdout(io.StringIO()):
                installer.main()
            for name in ['pyproject.toml','uv.lock','.python-version']:
                self.assertEqual((destination/name).read_bytes(),(ROOT/name).read_bytes())
            self.assertFalse((destination/'requirements.txt').exists())
            backups=list((codex_home/'skill-backups').glob('*/requirements.txt'))
            self.assertEqual(len(backups),1)
            self.assertEqual(backups[0].read_text(encoding='utf-8'),'lxml>=5.0\n')
            self.assertFalse((backups[0].parent/'.venv').exists())


if __name__=='__main__':unittest.main()
