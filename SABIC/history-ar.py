"""Localize the existing history deck, retaining its editable charts and images.

Run: python3 SABIC/history-ar.py
Fonts: the exact ManFath assets, installed with their internal family names.
"""
from copy import deepcopy
from io import BytesIO
import json
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

FOLDER = Path(__file__).resolve().parent
TEXT_FONT = 'Arsenica Arabic Antiqua Trial DemiBd'
DIGIT_FONT = 'thmanyah sans'
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}


def tag(name):
    prefix, local = name.split(':')
    return f'{{{NS[prefix]}}}{local}'


def parse(data):
    for _, (prefix, uri) in ET.iterparse(BytesIO(data), events=['start-ns']):
        if not re.fullmatch(r'ns\d+', prefix):
            ET.register_namespace(prefix, uri)
    root = ET.fromstring(data)
    # ElementTree drops unused declarations referenced only by mc:Ignorable.
    for node in root.iter():
        node.attrib.pop('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable', None)
    return root


def serialize(root):
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def font(properties, family=TEXT_FONT):
    properties.set('lang', 'en-US' if family == DIGIT_FONT else 'ar-SA')
    properties.set('b', '1' if family == DIGIT_FONT else '0')
    properties.set('i', '0')
    for name in ['latin', 'ea', 'cs']:
        node = properties.find(f'a:{name}', NS)
        if node is None:
            node = ET.Element(tag(f'a:{name}'))
            # Font children precede hyperlinks in DrawingML's property order.
            index = next((i for i, child in enumerate(properties)
                          if child.tag in [tag('a:hlinkClick'), tag('a:hlinkMouseOver'), tag('a:extLst')]), len(properties))
            properties.insert(index, node)
        node.set('typeface', family)


def rewrite_text(body, content):
    first = body.find('a:p', NS)
    paragraph_style = first.find('a:pPr', NS) if first is not None else None
    style = body.find('.//a:rPr', NS)
    style = deepcopy(style) if style is not None else ET.Element(tag('a:rPr'))
    alignment = paragraph_style.get('algn') if paragraph_style is not None else None
    for paragraph in body.findall('a:p', NS):
        body.remove(paragraph)
    for line in content.split('\n'):
        rtl = bool(re.search(r'[\u0600-\u06ff]', line))
        paragraph = ET.SubElement(body, tag('a:p'))
        props = deepcopy(paragraph_style) if paragraph_style is not None else ET.Element(tag('a:pPr'))
        props.set('algn', 'ctr' if alignment == 'ctr' else 'r')
        props.set('rtl', '1' if rtl else '0')
        paragraph.append(props)
        default = props.find('a:defRPr', NS)
        if default is not None:
            font(default)
        for part in re.findall(r'[0-9]+|[^0-9]+', line):
            run = ET.SubElement(paragraph, tag('a:r'))
            properties = deepcopy(style)
            font(properties, DIGIT_FONT if part.isascii() and part.isdigit() else TEXT_FONT)
            run.append(properties)
            ET.SubElement(run, tag('a:t')).text = part
        end = ET.SubElement(paragraph, tag('a:endParaRPr'))
        font(end)


def set_default_fonts(root):
    for node in root.iter():
        if node.text in ['Arial', 'Helvetica Neue']:
            node.text = TEXT_FONT
        if 'typeface' in node.attrib:
            node.set('typeface', TEXT_FONT)
        if node.tag in [tag('a:rPr'), tag('a:defRPr'), tag('a:endParaRPr')]:
            font(node)


def mirror(root, width):
    for transform in root.findall('.//a:xfrm', NS) + root.findall('.//p:xfrm', NS):
        offset, extent = transform.find('a:off', NS), transform.find('a:ext', NS)
        if offset is not None and extent is not None and int(extent.get('cx', 0)):
            offset.set('x', str(width - int(offset.get('x')) - int(extent.get('cx'))))


def localize_workbook(data, translations):
    result = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(result, 'w') as output:
        for entry in source.infolist():
            payload = source.read(entry.filename)
            if entry.filename.endswith('.xml'):
                root = parse(payload)
                for node in root.iter():
                    if node.text in translations:
                        node.text = translations[node.text]
                    if node.get('name') in translations:
                        node.set('name', translations[node.get('name')])
                set_default_fonts(root)
                if entry.filename == 'xl/styles.xml':
                    for node in root.findall('.//s:font/s:name', NS):
                        node.set('val', TEXT_FONT)
                if entry.filename == 'xl/worksheets/sheet1.xml':
                    root.find('s:sheetViews/s:sheetView', NS).set('rightToLeft', '1')
                payload = serialize(root)
            output.writestr(entry, payload)
    return result.getvalue()


def chart_labels(root):
    # Axis labels contain only digits. Decimal data labels need separate runs
    # so their punctuation stays in Arsenica, matching the ManFath contract.
    for props in root.findall('.//a:defRPr', NS):
        font(props, DIGIT_FONT)
    values = [node.text for node in root.findall('.//c:ser/c:val/c:numRef/c:numCache/c:pt/c:v', NS)]
    labels = root.find('.//c:ser/c:dLbls', NS)
    for index, value in enumerate(values):
        label = ET.Element(tag('c:dLbl'))
        ET.SubElement(label, tag('c:idx'), {'val': str(index)})
        rich = ET.SubElement(ET.SubElement(label, tag('c:tx')), tag('c:rich'))
        ET.SubElement(rich, tag('a:bodyPr'))
        ET.SubElement(rich, tag('a:lstStyle'))
        paragraph = ET.SubElement(rich, tag('a:p'))
        run = ET.SubElement(paragraph, tag('a:r'))
        ET.SubElement(run, tag('a:rPr'), {'sz': '1600'})
        rewrite_text(rich, value)
        ET.SubElement(label, tag('c:showLegendKey'), {'val': '0'})
        ET.SubElement(label, tag('c:showVal'), {'val': '0'})
        ET.SubElement(label, tag('c:showCatName'), {'val': '0'})
        ET.SubElement(label, tag('c:showSerName'), {'val': '0'})
        labels.insert(index, label)
    orientation = root.find('.//c:catAx/c:scaling/c:orientation', NS)
    orientation.set('val', 'maxMin')


def main():
    copy = json.loads((FOLDER / 'history-ar-copy.json').read_text())
    destination = FOLDER.parent / 'SabicHistoryAr.pptx'
    with ZipFile(FOLDER.parent / 'SabicHistory.pptx') as source, ZipFile(destination, 'w') as output:
        size = parse(source.read('ppt/presentation.xml')).find('p:sldSz', NS)
        width = int(size.get('cx'))
        for entry in source.infolist():
            name = entry.filename
            data = source.read(name)
            if name.endswith('.xml'):
                root = parse(data)
                set_default_fonts(root)
                match = re.fullmatch(r'ppt/slides/slide(\d+)\.xml', name)
                if match:
                    number = match.group(1)
                    localized = copy['slides'][number]['shapes']
                    for shape in root.findall('.//p:sp', NS):
                        body = shape.find('p:txBody', NS)
                        if body is None or not body.findall('.//a:t', NS):
                            continue
                        field = body.find('.//a:fld', NS)
                        if field is not None:
                            props = field.find('a:rPr', NS)
                            font(props, DIGIT_FONT)
                            continue
                        identity = shape.find('p:nvSpPr/p:cNvPr', NS).get('id')
                        assert identity in localized, f'Missing Arabic text: slide {number}, shape {identity}'
                        rewrite_text(body, localized[identity])
                    mirror(root, width)
                    # The source's slide-number field duplicates a body id on
                    # slide 2. Renumber drawing objects without changing fields.
                    for index, node in enumerate(root.findall('.//p:cNvPr', NS), 1):
                        node.set('id', str(index))
                elif re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', name):
                    mirror(root, width)
                    for shape in root.findall('.//p:sp', NS):
                        body = shape.find('p:txBody', NS)
                        if body is not None:
                            text = '\n'.join(node.text or '' for node in body.findall('.//a:t', NS))
                            if text:
                                text = text.replace('CHEMISTRY THAT MATTERS™', 'كيمياء وتواصل').replace('SABIC |', 'سابك |')
                                rewrite_text(body, text)
                    # Reflect the cover ribbons, not the Arabic/Latin logo.
                    for image in root.findall('.//p:pic', NS):
                        transform = image.find('p:spPr/a:xfrm', NS)
                        if int(transform.find('a:ext', NS).get('cx')) > 2 * 914400:
                            transform.set('flipH', '1')
                elif re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml', name):
                    number = re.search(r'notesSlide(\d+)', name).group(1)
                    for shape in root.findall('.//p:sp', NS):
                        placeholder = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
                        body = shape.find('p:txBody', NS)
                        if placeholder is not None and placeholder.get('type') == 'body':
                            rewrite_text(body, copy['slides'][number]['notes'])
                        elif body is not None:
                            for props in body.findall('.//a:rPr', NS):
                                font(props, DIGIT_FONT)
                elif re.fullmatch(r'ppt/charts/chart\d+\.xml', name):
                    translations = copy['charts'][Path(name).name]
                    for node in root.iter():
                        if node.text in translations:
                            node.text = translations[node.text]
                    if name.endswith('chart2.xml'):
                        chart_labels(root)
                elif name == 'docProps/core.xml':
                    root.find('{http://purl.org/dc/elements/1.1/}title').text = 'تاريخ سابك: 1976–2026'
                    root.find('{http://purl.org/dc/elements/1.1/}subject').text = 'نسخة عربية بخطي منفذ: أرسينيكا للنص وثمانية للأرقام'
                data = serialize(root)
            elif name.endswith('.xlsx'):
                chart = 'chart1.xml' if name.endswith('Worksheet1.xlsx') else 'chart2.xml'
                data = localize_workbook(data, copy['charts'][chart])
            output.writestr(entry, data)
    print(f'Created {destination}: 12 Arabic slides with RTL layout and ManFath typography.')


if __name__ == '__main__':
    main()
