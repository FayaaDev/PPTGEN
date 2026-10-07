"""Verify generated OPV Office packages; Python standard library only."""
from pathlib import Path
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

FOLDER = Path(__file__).resolve().parent
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
LABELS = ['Cover', 'Section', 'Content', 'Comparison', 'Three cards', 'Table',
          'Chart', 'Image', 'Timeline', 'Summary', 'Appendix', 'References']
LAYOUTS = {f'OPV {i:02d} - {label}' for i, label in enumerate(LABELS, 1)}

for filename, kind in [('OPV_Compact_Examples.pptx', 'presentation'),
                       ('OPV_Compact_Master.potx', 'template')]:
    with ZipFile(FOLDER / filename) as archive:
        assert archive.testzip() is None, 'Corrupt ZIP'
        names = set(archive.namelist())
        xml = {name: ET.fromstring(archive.read(name)) for name in names
               if name.endswith(('.xml', '.rels'))}
        assert f'presentationml.{kind}.main+xml' in archive.read('[Content_Types].xml').decode()
        slides = sorted(name for name in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', name))
        assert len(slides) == 12, 'Expected one example per layout'
        layouts = {root.find('p:cSld', NS).get('name'): root for name, root in xml.items()
                   if re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', name)}
        assert LAYOUTS <= layouts.keys(), 'Named layout missing'
        for name in LAYOUTS:
            assert layouts[name].findall('.//p:ph', NS), f'No placeholders: {name}'
            assert layouts[name].find('.//p:ph[@type="title"]', NS) is not None, f'No title: {name}'
        assert sum(len(xml[name].findall('.//a:tbl', NS)) for name in slides) == 1
        assert sum(len(xml[name].findall('.//c:chart', NS)) for name in slides) == 1
        assert any(name.startswith('ppt/embeddings/') and name.endswith('.xlsx') for name in names)
        size = xml['ppt/presentation.xml'].find('p:sldSz', NS)
        width, height = int(size.get('cx')), int(size.get('cy'))
        assert abs(width / height - 16 / 9) < 0.00001
        for name, root in xml.items():
            if re.fullmatch(r'ppt/(slides/slide|slideLayouts/slideLayout)\d+\.xml', name):
                if name in slides:
                    assert root.findall('.//a:t', NS), f'No editable text: {name}'
                for transform in root.findall('.//a:xfrm', NS) + root.findall('.//p:xfrm', NS):
                    offset, extent = transform.find('a:off', NS), transform.find('a:ext', NS)
                    if offset is None or extent is None:
                        continue
                    x, y, w, h = [int(node.get(key)) for node, key in
                                  [(offset, 'x'), (offset, 'y'), (extent, 'cx'), (extent, 'cy')]]
                    assert 0 <= x <= x + w <= width + 20, f'Horizontal overflow: {name}'
                    assert 0 <= y <= y + h <= height + 20, f'Vertical overflow: {name}'
                    if y < 6.5 * 914400 and not (x == 0 and w < 0.1 * 914400):
                        assert y + h <= 6.5 * 914400, f'Content intrudes into citations: {name}'
            if name.endswith('.rels'):
                base = posixpath.dirname(posixpath.dirname(name))
                for relation in root:
                    if relation.get('TargetMode') != 'External':
                        target = posixpath.normpath(posixpath.join(base, relation.get('Target'))).lstrip('/')
                        assert target in names, f'Missing relationship: {target}'
        theme = xml['ppt/theme/theme1.xml']
        assert theme.find('.//a:accent1/a:srgbClr', NS).get('val') == '136F7A'
        assert theme.find('.//a:majorFont/a:latin', NS).get('typeface') == 'Franklin Gothic Book'
    print(f'{filename}: PASS — 12 editable examples, 12 OPV layouts, native table/chart, valid relationships and bounds.')
