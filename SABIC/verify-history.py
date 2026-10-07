"""Check the generated history deck: python3 SABIC/verify-history.py."""
from pathlib import Path
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

file = Path(__file__).resolve().parent.parent / 'SabicHistory.pptx'
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
with ZipFile(file) as archive:
    assert archive.testzip() is None
    names = set(archive.namelist())
    xml = {name: ET.fromstring(archive.read(name)) for name in names
           if name.endswith(('.xml', '.rels'))}
    slides = [name for name in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', name)]
    charts = [name for name in names if re.fullmatch(r'ppt/charts/chart\d+\.xml', name)]
    assert len(slides) == 12
    assert len(charts) == 2
    values = {tuple(v.text for v in xml[name].findall('.//c:numCache/c:pt/c:v', ns)) for name in charts}
    assert ('70', '30') in values and ('6.3', '55.5') in values
    dimensions = xml['ppt/presentation.xml'].find('p:sldSz', ns)
    width, height = int(dimensions.get('cx')), int(dimensions.get('cy'))
    for name in slides:
        root = xml[name]
        assert root.findall('.//a:t', ns), f'No editable text: {name}'
        number = re.search(r'slide(\d+)\.xml', name).group(1)
        notes = archive.read(f'ppt/notesSlides/notesSlide{number}.xml').decode()
        assert 'https://www.sabic.com/' in notes, f'Missing source URL: {name}'
        for transform in root.findall('.//a:xfrm', ns):
            offset, extent = transform.find('a:off', ns), transform.find('a:ext', ns)
            if offset is not None and extent is not None:
                x, y = int(offset.get('x')), int(offset.get('y'))
                w, h = int(extent.get('cx')), int(extent.get('cy'))
                assert 0 <= x <= x + w <= width + 10, f'Horizontal overflow: {name}'
                assert 0 <= y <= y + h <= height + 10, f'Vertical overflow: {name}'
    for name, root in xml.items():
        if name.endswith('.rels'):
            base = posixpath.dirname(posixpath.dirname(name))
            for relation in root:
                if relation.get('TargetMode') != 'External':
                    target = posixpath.normpath(posixpath.join(base, relation.get('Target'))).lstrip('/')
                    assert target in names, f'Missing relationship target: {target}'
print('PASS: 12 editable slides, 2 charts with verified data, source URLs on every slide, valid XML/relationships and slide bounds.')
