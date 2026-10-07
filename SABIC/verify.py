"""Run after node SABIC/generate.cjs. Requires only Python's standard library."""
from pathlib import Path
import posixpath
import re
from zipfile import ZipFile
from xml.etree import ElementTree as ET

folder = Path(__file__).resolve().parent
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

for extension, kind in [('pptx', 'presentation'), ('potx', 'template')]:
    with ZipFile(folder / f'SABIC-Template.{extension}') as archive:
        assert archive.testzip() is None
        names = set(archive.namelist())
        xml = {name: ET.fromstring(archive.read(name)) for name in names
               if name.endswith(('.xml', '.rels'))}
        types = archive.read('[Content_Types].xml').decode()
        assert f'presentationml.{kind}.main+xml' in types
        slides = sorted(name for name in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', name))
        assert len(slides) == 8
        layouts = [root for name, root in xml.items() if re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', name)]
        titles = {root.find('p:cSld', ns).get('name') for root in layouts}
        assert {'SABIC_COVER', 'SABIC_CONTENT', 'SABIC_TWO_COLUMNS', 'SABIC_THREE_COLUMNS',
                'SABIC_FINANCIAL_TABLE', 'SABIC_TEXT_AND_TABLE'} <= titles
        assert sum(len(xml[name].findall('.//a:tbl', ns)) for name in slides) == 3
        size = xml['ppt/presentation.xml'].find('p:sldSz', ns)
        width, height = int(size.get('cx')), int(size.get('cy'))
        assert abs(width / height - 850.32 / 566.64) < 0.00001
        for name in slides:
            root = xml[name]
            assert root.findall('.//a:t', ns), f'No editable text in {name}'
            title = root.find('.//p:sp/p:txBody/a:p/a:pPr', ns)
            assert title is not None and title.get('algn') == 'l', f'Title is not left-aligned in {name}'
            for transform in root.findall('.//a:xfrm', ns):
                offset, extent = transform.find('a:off', ns), transform.find('a:ext', ns)
                if offset is not None and extent is not None:
                    x, y = int(offset.get('x')), int(offset.get('y'))
                    w, h = int(extent.get('cx')), int(extent.get('cy'))
                    assert 0 <= x <= x + w <= width + 10, f'Horizontal overflow in {name}'
                    assert 0 <= y <= y + h <= height + 10, f'Vertical overflow in {name}'
        for name, root in xml.items():
            if name.endswith('.rels'):
                base = posixpath.dirname(posixpath.dirname(name))
                for relation in root:
                    if relation.get('TargetMode') != 'External':
                        target = posixpath.normpath(posixpath.join(base, relation.get('Target'))).lstrip('/')
                        assert target in names, f'Missing relationship target: {target}'
    print(f'{extension}: valid XML and relationships; 8 editable slides, 6 SABIC layouts, 3 native tables; bounds checked.')
