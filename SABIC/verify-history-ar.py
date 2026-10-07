"""Verify Arabic content, editable data, RTL layout, and ManFath run fonts."""
from io import BytesIO
from pathlib import Path
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent.parent
TEXT_FONT = 'Arsenica Arabic Antiqua Trial DemiBd'
DIGIT_FONT = 'thmanyah sans'
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}

with ZipFile(ROOT / 'SabicHistory.pptx') as original, ZipFile(ROOT / 'SabicHistoryAr.pptx') as archive:
    assert archive.testzip() is None
    names = set(archive.namelist())
    xml = {name: ET.fromstring(archive.read(name)) for name in names
           if name.endswith(('.xml', '.rels'))}
    slides = sorted(name for name in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', name))
    assert len(slides) == 12
    size = xml['ppt/presentation.xml'].find('p:sldSz', NS)
    width, height = int(size.get('cx')), int(size.get('cy'))
    digit_runs = 0
    for name, root in xml.items():
        data = archive.read(name).decode()
        assert 'Arial' not in data and 'Helvetica' not in data, f'Fallback font in {name}'
        runs = []
        if re.fullmatch(r'ppt/(slides/slide|notesSlides/notesSlide|charts/chart)\d+\.xml', name):
            runs = root.findall('.//a:r', NS) + root.findall('.//a:fld', NS)
        for run in runs:
            text = ''.join(node.text or '' for node in run.findall('a:t', NS))
            props = run.find('a:rPr', NS)
            if not text or props is None:
                continue
            face = props.find('a:latin', NS)
            assert face is not None, f'Unspecified run font in {name}'
            if re.search(r'[0-9]', text):
                assert re.fullmatch(r'[0-9]+', text), f'Mixed digit run in {name}: {text}'
                assert face.get('typeface') == DIGIT_FONT and props.get('b') == '1'
                digit_runs += 1
            else:
                assert face.get('typeface') == TEXT_FONT, f'Wrong text font in {name}'
        if name.endswith('.rels'):
            base = posixpath.dirname(posixpath.dirname(name))
            for relation in root:
                if relation.get('TargetMode') != 'External':
                    target = posixpath.normpath(posixpath.join(base, relation.get('Target'))).lstrip('/')
                    assert target in names, f'Missing relationship: {target}'
    for name in slides:
        root = xml[name]
        assert re.search(r'[\u0600-\u06ff]', ''.join(t.text or '' for t in root.findall('.//a:t', NS)))
        ids = [node.get('id') for node in root.findall('.//p:cNvPr', NS)]
        assert len(ids) == len(set(ids)), f'Duplicate object IDs in {name}'
        for paragraph in root.findall('.//a:p', NS):
            content = ''.join(t.text or '' for t in paragraph.findall('.//a:t', NS))
            if re.search(r'[\u0600-\u06ff]', content):
                assert paragraph.find('a:pPr', NS).get('rtl') == '1'
        for transform in root.findall('.//a:xfrm', NS) + root.findall('.//p:xfrm', NS):
            offset, extent = transform.find('a:off', NS), transform.find('a:ext', NS)
            if offset is not None and extent is not None:
                x, y = int(offset.get('x')), int(offset.get('y'))
                w, h = int(extent.get('cx')), int(extent.get('cy'))
                assert 0 <= x <= x + w <= width + 10, f'Horizontal overflow: {name}'
                assert 0 <= y <= y + h <= height + 10, f'Vertical overflow: {name}'
        number = re.search(r'slide(\d+)', name).group(1)
        notes = archive.read(f'ppt/notesSlides/notesSlide{number}.xml').decode()
        assert 'https://www.sabic.com/' in notes
    for name in [n for n in names if re.fullmatch(r'ppt/charts/chart\d+\.xml', n)]:
        source = ET.fromstring(original.read(name))
        expected = [n.text for n in source.findall('.//c:numCache/c:pt/c:v', NS)]
        actual = [n.text for n in xml[name].findall('.//c:numCache/c:pt/c:v', NS)]
        assert actual == expected, f'Chart values changed: {name}'
    for name in [n for n in names if n.startswith('ppt/media/') and not n.endswith('/')]:
        assert archive.read(name) == original.read(name), f'Artwork changed: {name}'
    for name in [n for n in names if n.endswith('.xlsx')]:
        with ZipFile(BytesIO(archive.read(name))) as workbook:
            assert workbook.testzip() is None
            for part in workbook.namelist():
                if part.endswith(('.xml', '.rels')):
                    ET.fromstring(workbook.read(part))
    assert digit_runs > 100
print(f'PASS: 12 Arabic RTL slides; {digit_runs} Thmanyah digit runs; Arsenica text; no Arial; unchanged chart values and photos; valid editable workbooks and slide bounds.')
