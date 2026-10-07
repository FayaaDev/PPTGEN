"""Recreate all 61 Polio slides with OPV Compact chrome, without flattening.

Run: python3 OPV/recreate.py [--verify-only]
Original slide order, text runs, relationships, notes and payloads are retained.
Template layouts remain reusable; inherited shape IDs are normalized on import.
Retained-content variants remove
template body decorations so the original native objects can occupy that space.
"""
from copy import deepcopy
from io import BytesIO
from itertools import count
import json
from pathlib import Path
import posixpath
import re
import sys
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

HERE = Path(__file__).resolve().parent
SOURCE = HERE / 'Polio_2026_7.pptx'
TEMPLATE = HERE / 'OPV_Compact_Master.potx'
OUTPUT = HERE / 'Polio_2026_7_OPV_Compact.pptx'
REPORT = HERE / 'Polio_2026_7_OPV_Compact_checks.json'
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
      'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)
EMU = 914400
WIDTH, HEIGHT = 12192000, 6858000
SECTIONS = {3, 8, 11, 18, 25, 36, 41, 47, 50}
LAYOUTS = {1: 'Cover', 55: 'Cover', **{i: 'Section' for i in SECTIONS},
           **{i: 'Table' for i in (6, 17, 29, 45, 57, 58)},
           **{i: 'Comparison' for i in (27, 51, 60)},
           26: 'Three cards', 39: 'Three cards', 54: 'Summary', 61: 'References',
           **{i: 'Image' for i in (5, 9, 12, 13, 15, 16, 20, 21, 22, 23, 24, 28, 37, 48, 49)},
           **{i: 'Timeline' for i in (4, 10, 43)}, 30: 'Chart'}
PALETTE = {'E4E8EA': 'F2F6F7', 'DBE8F3': 'DBE7F2', 'DAEBDF': 'D2E7E8',
           '2A6F97': '2B6F94', '2C7A5B': '2D7A62'}


def q(name):
    prefix, local = name.split(':')
    return f'{{{NS[prefix]}}}{local}'


def xml(data):
    # Preserve namespace bindings used in extension markup, including QName values.
    for _, (prefix, uri) in ET.iterparse(BytesIO(data), events=('start-ns',)):
        if prefix and not re.fullmatch(r'ns\d+', prefix):
            ET.register_namespace(prefix, uri)
    return ET.fromstring(data)


def dump(root):
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def relpath(part):
    folder, name = posixpath.split(part)
    return f'{folder}/_rels/{name}.rels'


def resolve(part, target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target)).lstrip('/')


def transform(shape):
    for path in ('p:spPr/a:xfrm', 'p:xfrm', 'p:grpSpPr/a:xfrm'):
        node = shape.find(path, NS)
        if node is not None:
            return node
    return None


def box(shape):
    tr = transform(shape)
    if tr is None:
        return None
    off, ext = tr.find('a:off', NS), tr.find('a:ext', NS)
    if off is None or ext is None:
        return None
    return tuple(int(node.get(key)) for node, key in
                 ((off, 'x'), (off, 'y'), (ext, 'cx'), (ext, 'cy')))


def place(shape, coords):
    tr = transform(shape)
    assert tr is not None, 'Expected explicit shape transform'
    off, ext = tr.find('a:off', NS), tr.find('a:ext', NS)
    x, y, w, h = (round(v) for v in coords)
    off.set('x', str(x)); off.set('y', str(y))
    ext.set('cx', str(w)); ext.set('cy', str(h))


def text(shape):
    return ''.join(t.text or '' for t in shape.findall('.//a:t', NS))


def ph(shape):
    return shape.find('.//p:ph', NS)


def remove_ph(shape):
    for nv in shape.findall('.//p:nvPr', NS):
        for child in list(nv):
            if child.tag == q('p:ph'):
                nv.remove(child)


def font_style(shape, factor=1, size=None, color=None, bold=None):
    for node in shape.iter():
        if node.tag in (q('a:rPr'), q('a:defRPr'), q('a:endParaRPr')):
            if size is not None:
                node.set('sz', str(round(size * 100)))
            elif 'sz' in node.attrib:
                node.set('sz', str(max(600, round(int(node.get('sz')) * factor))))
            if bold is not None:
                node.set('b', '1' if bold else '0')
            if color:
                for child in list(node):
                    if child.tag in (q('a:solidFill'), q('a:noFill'), q('a:gradFill')):
                        node.remove(child)
                fill = ET.Element(q('a:solidFill'))
                ET.SubElement(fill, q('a:srgbClr'), val=color)
                node.insert(0, fill)
        if node.tag in (q('a:latin'), q('a:ea'), q('a:cs')):
            node.set('typeface', 'Franklin Gothic Book')
        if node.tag == q('a:spcPts'):
            node.set('val', str(round(int(node.get('val')) * factor)))
        if node.tag in (q('a:bodyPr'), q('a:tcPr'), q('a:pPr')):
            for key in ('lIns', 'rIns', 'tIns', 'bIns', 'marL', 'marR', 'indent'):
                if key in node.attrib:
                    node.set(key, str(round(int(node.get(key)) * factor)))


def bind(shape, template_shape):
    """Use template geometry/placeholder while retaining original runs and links."""
    place(shape, box(template_shape))
    remove_ph(shape)
    nv = shape.find('p:nvSpPr/p:nvPr', NS)
    if nv is not None:
        nv.append(deepcopy(ph(template_shape)))


def new_text(template_shape, value, shape_id, name):
    shape = deepcopy(template_shape)
    prop = shape.find('p:nvSpPr/p:cNvPr', NS)
    prop.set('id', str(shape_id)); prop.set('name', f'OPV added: {name}')
    body = shape.find('p:txBody', NS)
    for paragraph in body.findall('a:p', NS):
        body.remove(paragraph)
    paragraph = ET.SubElement(body, q('a:p'))
    ET.SubElement(paragraph, q('a:pPr'), marL='0', indent='0')
    run = ET.SubElement(paragraph, q('a:r'))
    default = body.find('.//a:defRPr', NS)
    if default is not None:
        props = deepcopy(default); props.tag = q('a:rPr'); run.append(props)
    ET.SubElement(run, q('a:t')).text = value
    return shape


def placeholders(layout):
    result = {}
    for shape in layout.findall('p:cSld/p:spTree/p:sp', NS):
        placeholder, coords = ph(shape), box(shape)
        if placeholder is None or coords is None:
            continue
        y = coords[1] / EMU
        kind = placeholder.get('type')
        if kind == 'title':
            result['title'] = shape
        elif kind == 'sldNum':
            result['number'] = shape
        elif y > 7:
            result['footer'] = shape
        elif 6.5 < y < 6.9:
            result['source'] = shape
        elif y < 0.7:
            result['section'] = shape
        elif abs(y - 1.55) < 0.02 or abs(y - 3.76) < 0.02:
            result['subtitle'] = shape
    return result


def verify_office_ids(roots):
    """Office shares one ID registry across slide masters and their layouts."""
    pres = roots['ppt/presentation.xml']
    ids = [n.get('id') for n in pres.findall('p:sldMasterIdLst/p:sldMasterId', NS)]
    for name, root in roots.items():
        if name.startswith('ppt/slideMasters/') and name.endswith('.xml'):
            ids.extend(n.get('id') for n in root.findall('p:sldLayoutIdLst/p:sldLayoutId', NS))
        shape_ids = [n.get('id') for n in root.findall('.//p:cNvPr', NS)]
        assert len(shape_ids) == len(set(shape_ids)), f'Duplicate shape ID: {name}'
        assert all(0 <= int(n) <= 4294967295 for n in shape_ids), f'Invalid shape ID: {name}'
    assert len(ids) == len(set(ids)), 'Presentation-wide master/layout ID collision'
    assert all(2147483648 <= int(n) <= 4294967295 for n in ids), 'Invalid master/layout ID'
    return {'global_master_layout_ids': len(ids), 'shape_id_scopes': sum(
        bool(r.findall('.//p:cNvPr', NS)) for r in roots.values())}


def recreate():
    with ZipFile(SOURCE) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    with ZipFile(TEMPLATE) as z:
        theme = {n: z.read(n) for n in z.namelist()}
    types, theme_types = xml(parts['[Content_Types].xml']), xml(theme['[Content_Types].xml'])
    imported = {}
    inherited_id_fixes = []

    def import_part(old):
        if old in imported:
            return imported[old]
        folder, filename = posixpath.split(old)
        new = f'{folder}/opv_{filename}'
        assert new not in parts, f'Import collision: {new}'
        imported[old] = new
        parts[new] = theme[old]
        if old.endswith('.xml'):
            root = xml(parts[new])
            props = root.findall('.//p:cNvPr', NS)
            seen = set()
            changed = False
            next_shape_id = max((int(p.get('id')) for p in props), default=0) + 1
            for prop in props:
                old_id = prop.get('id')
                if old_id in seen:
                    changed = True
                    prop.set('id', str(next_shape_id))
                    inherited_id_fixes.append({'part': new, 'shape': prop.get('name'),
                                              'old_id': old_id, 'new_id': str(next_shape_id)})
                    next_shape_id += 1
                seen.add(prop.get('id'))
            if changed:
                parts[new] = dump(root)
        if relpath(old) in theme:
            relationships = xml(theme[relpath(old)])
            for relation in relationships:
                if relation.get('TargetMode') != 'External':
                    target = import_part(resolve(old, relation.get('Target')))
                    relation.set('Target', posixpath.relpath(target, posixpath.dirname(new)))
            parts[relpath(new)] = dump(relationships)
        return new

    template_pres = xml(theme['ppt/presentation.xml'])
    template_rels = xml(theme['ppt/_rels/presentation.xml.rels'])
    template_relmap = {r.get('Id'): r for r in template_rels}
    pres, pres_rels = xml(parts['ppt/presentation.xml']), xml(parts['ppt/_rels/presentation.xml.rels'])
    master_list = pres.find('p:sldMasterIdLst', NS)
    occupied_ids = [int(n.get('id')) for n in master_list]
    for name, data in parts.items():
        if name.startswith('ppt/slideMasters/') and name.endswith('.xml'):
            occupied_ids.extend(int(n.get('id')) for n in xml(data).findall('p:sldLayoutIdLst/p:sldLayoutId', NS))
    office_ids = count(max(occupied_ids) + 1)
    for n in template_pres.find('p:sldMasterIdLst', NS):
        relation = template_relmap[n.get(q('r:id'))]
        new = import_part(resolve('ppt/presentation.xml', relation.get('Target')))
        next_master_id = next(office_ids)
        rid = f'rIdOPVMaster{next_master_id}'
        ET.SubElement(pres_rels, q('rel:Relationship'), Id=rid, Type=relation.get('Type'),
                      Target=posixpath.relpath(new, 'ppt'))
        ET.SubElement(master_list, q('p:sldMasterId'), id=str(next_master_id), **{q('r:id'): rid})
        master_root = xml(parts[new])
        for layout_id in master_root.findall('p:sldLayoutIdLst/p:sldLayoutId', NS):
            layout_id.set('id', str(next(office_ids)))
        parts[new] = dump(master_root)
    # New objects and Office's theme picker use the imported theme, too.
    for relation in pres_rels:
        if relation.get('Type').endswith('/theme'):
            relation.set('Target', posixpath.relpath(import_part('ppt/theme/theme1.xml'), 'ppt'))

    layouts, adapted = {}, {}
    for old, new in list(imported.items()):
        if not re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', old):
            continue
        root = xml(parts[new])
        name = root.find('p:cSld', NS).get('name', '')
        if not name.startswith('OPV '):
            continue
        label = name.split(' - ', 1)[1]
        layouts[label] = placeholders(root)
        variant = deepcopy(root)
        variant.find('p:cSld', NS).set('name', name + ' — Retained content')
        tree = variant.find('p:cSld/p:spTree', NS)
        dark = label in ('Cover', 'Section')
        # Keep shared header/footer chrome. Body objects remain native slide objects.
        for shape in list(tree)[2:]:
            coords = box(shape)
            if coords is None:
                continue
            y = coords[1] / EMU
            keep = y >= 6.5 or (y < 1.9 if not dark else ph(shape) is not None or abs(y - 1.91) < 0.02)
            if not keep:
                tree.remove(shape)
        # Cover/section titles are slide instances; template artwork must not double
        # the cover image or occupy the original divider's wide text area.
        for shape in list(tree):
            if dark and shape.tag == q('p:pic'):
                tree.remove(shape)
        variant_name = new.replace('opv_', 'opv_retained_', 1)
        adapted[label] = variant_name
        parts[variant_name] = dump(variant)
        parts[relpath(variant_name)] = parts[relpath(new)]
        master_relation = next(r for r in xml(parts[relpath(new)]) if r.get('Type').endswith('/slideMaster'))
        master = resolve(new, master_relation.get('Target'))
        master_xml = xml(parts[master]); master_rels = xml(parts[relpath(master)])
        rid = f'rIdOPVRetained{len(adapted)}'
        ET.SubElement(master_rels, q('rel:Relationship'), Id=rid,
                      Type=NS['r'] + '/slideLayout', Target=posixpath.relpath(variant_name, posixpath.dirname(master)))
        lst = master_xml.find('p:sldLayoutIdLst', NS)
        ET.SubElement(lst, q('p:sldLayoutId'), id=str(next(office_ids)), **{q('r:id'): rid})
        parts[master] = dump(master_xml); parts[relpath(master)] = dump(master_rels)

    # Add content types for every imported dependency; never overwrite source types.
    for node in theme_types:
        old = node.get('PartName', '').lstrip('/')
        if old in imported:
            copy = deepcopy(node); copy.set('PartName', '/' + imported[old]); types.append(copy)
        elif node.tag == q('ct:Default') and not any(
                n.get('Extension') == node.get('Extension') for n in types):
            types.append(deepcopy(node))
    for variant in adapted.values():
        ET.SubElement(types, q('ct:Override'), PartName='/' + variant,
                      ContentType='application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml')

    slide_reports = []
    current_section = 'SESSION OVERVIEW'
    for i in range(1, 62):
        part = f'ppt/slides/slide{i}.xml'
        root = xml(parts[part]); tree = root.find('p:cSld/p:spTree', NS)
        shapes = list(tree)[2:]
        label = LAYOUTS.get(i, 'Appendix' if i > 55 else 'Content')
        slots = layouts[label]
        dark = label in ('Cover', 'Section')
        title = next(s for s in shapes if ph(s) is not None and ph(s).get('type') in ('title', 'ctrTitle'))
        if i in SECTIONS:
            current_section = text(title).upper()
        elif i == 56:
            current_section = 'APPENDIX / SUPPORTING DETAIL'
        section = ('VACCINE PREVENTABLE DISEASES' if i == 1 else
                   'QUESTIONS / DISCUSSION' if i == 55 else
                   'APPENDIX / REFERENCES' if i == 61 else current_section)
        number = next((s for s in shapes if ph(s) is not None and ph(s).get('type') == 'sldNum'), None)
        citations, subtitles, body = [], [], []
        for shape in shapes:
            coords = box(shape)
            if shape is title or shape is number or coords is None:
                continue
            x, y, w, h = (v / EMU for v in coords)
            if not text(shape) and shape.tag == q('p:sp') and y < 0.7 and h < 0.1:
                tree.remove(shape)  # Original header accent replaced by layout accent.
            elif not dark and text(shape) and y >= 6.5 and h <= 0.5:
                citations.append(shape)
            elif not dark and text(shape) and 1.15 <= y <= 1.21 and h <= 0.35 and w > 10:
                subtitles.append(shape)
            else:
                body.append(shape)
        bind(title, slots['title'])
        title_size = (38 if i in SECTIONS else 42) if dark else (30 if len(text(title)) < 65 else 26)
        if dark and len(text(title)) > 38:
            title_size = 32
        font_style(title, size=title_size, color='FFFFFF' if dark else '0C2A33', bold=True)
        title.find('p:txBody/a:bodyPr', NS).set('anchor', 'ctr')
        if number is not None:
            bind(number, slots['number'])
            font_style(number, size=10, color='A8C6CC' if dark else '62767E')
        for shape in citations:
            bind(shape, slots['source'])
            font_style(shape, size=10, color='62767E')
        for shape in subtitles:
            bind(shape, slots['subtitle'])
            font_style(shape, size=14, color='62767E')
        scale = 1
        dense = False
        if dark:
            for shape in body:
                remove_ph(shape)
                if shape.tag == q('p:pic'):
                    place(shape, tuple(v * EMU for v in (9.05, 1.88, 3.7, 3.7 * box(shape)[3] / box(shape)[2])))
                elif text(shape):
                    place(shape, tuple(v * EMU for v in (0.6, 3.76, 8.1, 2.5 if i == 1 else 1.25)))
                    font_style(shape, size=20 if i == 1 else 22, color='BED2D7')
                else:
                    tree.remove(shape)  # Cover's former accent; Compact variant supplies it below.
            if i in SECTIONS:
                # Original divider text spans the slide; no inherited artwork to clash.
                place(title, tuple(v * EMU for v in (0.6, 2.18, 12.1, 1.4)))
                for shape in body:
                    if text(shape):
                        place(shape, tuple(v * EMU for v in (0.6, 3.76, 12.1, 1.25)))
        elif i in (20, 21, 22, 23, 48):
            # Full-height source screenshots have a separate commentary column.
            # Fit visual objects together, preserving their relative placement;
            # do not shrink commentary because the screenshot spans 7.5 inches.
            dense = True
            commentary = [s for s in body if text(s) and box(s)[0] < 3 * EMU]
            visuals = [s for s in body if s not in commentary]
            vb = [box(s) for s in visuals]
            left, top = min(b[0] for b in vb), min(b[1] for b in vb)
            right, bottom = max(b[0] + b[2] for b in vb), max(b[1] + b[3] for b in vb)
            ratio = min(7.88 * EMU / (right - left), 4.4 * EMU / (bottom - top))
            ox = 0.45 * EMU + (7.88 * EMU - (right - left) * ratio) / 2
            oy = 2.0 * EMU + (4.4 * EMU - (bottom - top) * ratio) / 2
            for shape in visuals:
                remove_ph(shape)
                x, y, w, h = box(shape)
                place(shape, (ox + (x - left) * ratio, oy + (y - top) * ratio, w * ratio, h * ratio))
                font_style(shape, factor=ratio)
            for shape in commentary:
                remove_ph(shape)
                place(shape, tuple(v * EMU for v in (8.75, 2.08, 4.13, 4.15)))
                font_style(shape, size=20)
        elif body:
            coords = [box(s) for s in body if box(s) is not None]
            left = min(b[0] for b in coords); top = min(b[1] for b in coords)
            right = max(b[0] + b[2] for b in coords); bottom = max(b[1] + b[3] for b in coords)
            dense = (bottom - top) / EMU > 4.7 or sum(len(text(s)) for s in body) > 1700
            start = (1.98 if subtitles else 1.7) if dense else 2.0
            sx = 12.433333 * EMU / (right - left)
            sy = min(1.0, (6.4 - start) * EMU / (bottom - top))
            scale = min(1.0, sx, sy)
            for shape in body:
                remove_ph(shape)
                old = box(shape)
                if old is None:
                    continue
                x, y, w, h = old
                nx, ny = 0.45 * EMU + (x - left) * sx, start * EMU + (y - top) * sy
                nw, nh = w * sx, h * sy
                if shape.tag == q('p:pic'):
                    ratio = min(sx, sy)
                    nx += (nw - w * ratio) / 2; ny += (nh - h * ratio) / 2
                    nw, nh = w * ratio, h * ratio
                place(shape, (nx, ny, nw, nh))
                font_style(shape, factor=scale)
                for col in shape.findall('.//a:gridCol', NS):
                    col.set('w', str(round(int(col.get('w')) * sx)))
                for row in shape.findall('.//a:tr', NS):
                    row.set('h', str(round(int(row.get('h')) * sy)))
        for node in root.findall('.//a:srgbClr', NS):
            node.set('val', PALETTE.get(node.get('val'), node.get('val')))
        # Slide-level background overrides from the source must not hide Compact.
        csld = root.find('p:cSld', NS)
        bg = csld.find('p:bg', NS)
        if bg is not None:
            csld.remove(bg)
        root.set('showMasterSp', '1')
        next_id = max(int(n.get('id')) for n in tree.findall('.//p:cNvPr', NS)) + 1
        # Restyle original pale cards in place; retain their native text overlays.
        for shape in body:
            geom = shape.find('p:spPr/a:prstGeom', NS)
            fill = shape.find('p:spPr/a:solidFill/a:srgbClr', NS)
            coords = box(shape)
            if (dark or text(shape) or geom is None or fill is None or coords is None or
                    geom.get('prst') not in ('rect', 'roundRect') or
                    fill.get('val') not in ('F2F6F7', 'D2E7E8', 'DBE7F2') or
                    coords[2] < 1.5 * EMU or coords[3] < 0.55 * EMU):
                continue
            geom.set('prst', 'rect')
            rule = deepcopy(shape)
            prop = rule.find('p:nvSpPr/p:cNvPr', NS)
            prop.set('id', str(next_id)); prop.set('name', 'OPV added: card side rule')
            next_id += 1
            place(rule, (coords[0], coords[1], 0.035 * EMU, coords[3]))
            rule.find('p:spPr/a:solidFill/a:srgbClr', NS).set('val', '136F7A')
            sppr = rule.find('p:spPr', NS)
            line = sppr.find('a:ln', NS)
            if line is not None:
                sppr.remove(line)
            ET.SubElement(ET.SubElement(sppr, q('a:ln')), q('a:noFill'))
            tree.insert(list(tree).index(shape) + 1, rule)
        for slot, value in (('section', section), ('footer', 'VPD Department · Public Health Authority')):
            tree.append(new_text(slots[slot], value, next_id, slot)); next_id += 1
        if number is None:
            tree.append(new_text(slots['number'], str(i), next_id, 'number'))
            next_id += 1
        frame = ET.SubElement(tree, q('p:sp'))
        nv = ET.SubElement(frame, q('p:nvSpPr'))
        ET.SubElement(nv, q('p:cNvPr'), id=str(next_id), name='OPV added: slide frame')
        ET.SubElement(nv, q('p:cNvSpPr'))
        ET.SubElement(nv, q('p:nvPr'))
        props = ET.SubElement(frame, q('p:spPr'))
        tr = ET.SubElement(props, q('a:xfrm'))
        inset = round(0.08 * EMU)
        ET.SubElement(tr, q('a:off'), x=str(inset), y=str(inset))
        ET.SubElement(tr, q('a:ext'), cx=str(WIDTH - 2 * inset), cy=str(HEIGHT - 2 * inset))
        ET.SubElement(ET.SubElement(props, q('a:prstGeom'), prst='rect'), q('a:avLst'))
        ET.SubElement(props, q('a:noFill'))
        line = ET.SubElement(props, q('a:ln'), w=str(2 * inset))
        ET.SubElement(ET.SubElement(line, q('a:solidFill')), q('a:srgbClr'), val='FFFFFF' if dark else '0C2A33')
        relationships = xml(parts[relpath(part)])
        for relation in relationships:
            if relation.get('Type').endswith('/slideLayout'):
                relation.set('Target', posixpath.relpath(adapted[label], 'ppt/slides'))
        parts[part] = dump(root); parts[relpath(part)] = dump(relationships)
        slide_reports.append({'slide': i, 'layout': label, 'dense': dense,
                              'body_font_scale': round(scale, 4)})
    parts['ppt/presentation.xml'] = dump(pres)
    parts['ppt/_rels/presentation.xml.rels'] = dump(pres_rels)
    parts['[Content_Types].xml'] = dump(types)
    temp = OUTPUT.with_suffix('.pptx.tmp')
    try:
        with ZipFile(temp, 'w', ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
        verify(temp, slide_reports, inherited_id_fixes)
        temp.replace(OUTPUT)
    finally:
        temp.unlink(missing_ok=True)


def verify(path=OUTPUT, slide_reports=None, inherited_id_fixes=None):
    with ZipFile(SOURCE) as z:
        source = {n: z.read(n) for n in z.namelist()}
    with ZipFile(path) as z:
        assert z.testzip() is None, 'Corrupt ZIP'
        result = {n: z.read(n) for n in z.namelist()}
    roots = {n: xml(b) for n, b in result.items() if n.endswith(('.xml', '.rels'))}
    office_id_checks = verify_office_ids(roots)
    slides = [n for n in result if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)]
    assert len(slides) == 61
    assert set(source) <= set(result), 'Source part removed'
    bounds_checked = 0
    hyperlinks = 0
    for i in range(1, 62):
        part = f'ppt/slides/slide{i}.xml'
        old, new = xml(source[part]), roots[part]
        # Only deliberately added section/footer/number shapes are excluded.
        kept = deepcopy(new)
        tree = kept.find('p:cSld/p:spTree', NS)
        for s in list(tree):
            props = s.find('p:nvSpPr/p:cNvPr', NS)
            if props is not None and props.get('name', '').startswith('OPV added: '):
                tree.remove(s)
        assert [n.text for n in old.findall('.//a:t', NS)] == [n.text for n in kept.findall('.//a:t', NS)], f'Text changed: slide {i}'
        assert len(old.findall('.//a:tbl', NS)) == len(new.findall('.//a:tbl', NS)), f'Table lost: {i}'
        assert len(old.findall('.//c:chart', NS)) == len(new.findall('.//c:chart', NS)), f'Chart lost: {i}'
        assert len(old.findall('.//p:pic', NS)) == len(new.findall('.//p:pic', NS)), f'Picture lost: {i}'
        for before, after in zip(old.findall('.//p:pic', NS), new.findall('.//p:pic', NS)):
            ob, nb = box(before), box(after)
            if ob and nb and ob[2] and ob[3]:
                assert abs((nb[2] / nb[3]) / (ob[2] / ob[3]) - 1) < 0.00001, f'Image aspect changed: {i}'
        for tag in ('a:hlinkClick', 'a:hlinkMouseOver'):
            old_links = [ET.tostring(n) for n in old.findall('.//' + tag, NS)]
            new_links = [ET.tostring(n) for n in new.findall('.//' + tag, NS)]
            assert old_links == new_links, f'Hyperlink changed: {i}'
            hyperlinks += len(old_links)
        for s in new.find('p:cSld/p:spTree', NS):
            coords = box(s)
            if coords is None:
                continue
            x, y, w, h = coords
            assert 0 <= x <= x + w <= WIDTH + 20 and 0 <= y <= y + h <= HEIGHT + 20, f'Bounds: slide {i}, {coords}'
            bounds_checked += 1
        ids = [n.get('id') for n in new.findall('.//p:cNvPr', NS)]
        assert len(ids) == len(set(ids)), f'Duplicate shape id: {i}'
        old_rels = xml(source[relpath(part)])
        new_rels = {r.get('Id'): r for r in roots[relpath(part)]}
        for rel in old_rels:
            if not rel.get('Type').endswith('/slideLayout'):
                assert rel.attrib == new_rels[rel.get('Id')].attrib, f'Relationship changed: {i}'
    payloads = [n for n in source if n.startswith(('ppt/media/', 'ppt/charts/', 'ppt/embeddings/', 'ppt/notesSlides/', 'ppt/notesMasters/'))]
    assert all(source[n] == result[n] for n in payloads), 'Original payload/notes bytes changed'
    mutable = {'[Content_Types].xml', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels'} | set(slides) | {relpath(n) for n in slides}
    assert all(source[n] == result[n] for n in source if n not in mutable), 'Unexpected original part edit'
    rel_count = 0
    for name, root in roots.items():
        if not name.endswith('.rels'):
            continue
        part = posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
        if name == '_rels/.rels':
            part = ''
        ids = [r.get('Id') for r in root]
        assert len(ids) == len(set(ids)), f'Duplicate relationship id: {name}'
        for rel in root:
            if rel.get('TargetMode') != 'External':
                assert resolve(part, rel.get('Target')) in result, f'Missing target: {name}: {rel.get("Target")}'
            rel_count += 1
        if part in roots:
            for node in roots[part].iter():
                for key, value in node.attrib.items():
                    if key.startswith('{' + NS['r'] + '}'):
                        assert value in ids, f'Unresolved relationship id: {part}: {value}'
    types = roots['[Content_Types].xml']
    overrides = {n.get('PartName').lstrip('/'): n.get('ContentType') for n in types if n.tag == q('ct:Override')}
    defaults = {n.get('Extension'): n.get('ContentType') for n in types if n.tag == q('ct:Default')}
    assert overrides['ppt/presentation.xml'].endswith('presentation.main+xml')
    assert all(n in result for n in overrides), 'Dangling content type'
    for name in result:
        if name == '[Content_Types].xml' or name.endswith('/'):
            continue
        assert name in overrides or name.rsplit('.', 1)[-1] in defaults, f'Missing content type: {name}'
    layout_names = [r.find('p:cSld', NS).get('name', '') for n, r in roots.items() if n.startswith('ppt/slideLayouts/opv_') and n.endswith('.xml')]
    assert sum(n.endswith('— Retained content') for n in layout_names) == 12
    pres = roots['ppt/presentation.xml']
    assert len(pres.find('p:sldIdLst', NS)) == 61
    assert ET.tostring(pres.find('p:sldIdLst', NS)) == ET.tostring(xml(source['ppt/presentation.xml']).find('p:sldIdLst', NS)), 'Slide order/IDs changed'
    assert pres.find('p:sldSz', NS).attrib == {'cx': str(WIDTH), 'cy': str(HEIGHT)}
    chart_count = sum(len(roots[n].findall('.//c:chart', NS)) for n in slides)
    table_count = sum(len(roots[n].findall('.//a:tbl', NS)) for n in slides)
    report = {'status': 'PASS', 'slides': 61, 'original_text_order': 'exact per-slide a:t sequence',
              'original_picture_objects': sum(len(roots[n].findall('.//p:pic', NS)) for n in slides),
              'native_tables': table_count, 'native_charts': chart_count,
              'unchanged_payload_and_notes_parts': len(payloads),
              'original_media_parts': sum(n.startswith('ppt/media/') for n in source),
              'original_workbooks': sum(n.startswith('ppt/embeddings/') for n in source),
              'hyperlink_elements_preserved': hyperlinks, 'relationships_checked': rel_count,
              'shape_bounds_checked': bounds_checked, 'normalized_named_layouts': 12,
              'retained_content_layouts': 12,
              'office_id_checks': office_id_checks,
              'concerns': ['Package checks cannot establish rendered text fit or visual overlap.',
                           'Original screenshot/news-image placeholders remain where source content was missing.']}
    if slide_reports is None and REPORT.exists():
        previous = json.loads(REPORT.read_text())
        slide_reports = previous.get('slide_mapping', [])
        if inherited_id_fixes is None:
            inherited_id_fixes = previous.get('inherited_shape_id_fixes', [])
    report['inherited_shape_id_fixes'] = inherited_id_fixes or []
    report['slide_mapping'] = slide_reports or []
    REPORT.write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps({k: v for k, v in report.items() if k != 'slide_mapping'}, indent=2))
    return report


if __name__ == '__main__':
    if sys.argv[1:] == ['--verify-only']:
        verify()
    elif not sys.argv[1:]:
        recreate()
        print(f'Created {OUTPUT}')
    else:
        raise SystemExit('Usage: python3 OPV/recreate.py [--verify-only]')
