#!/usr/bin/env python3
"""按 PPTX 关系表读取雨课堂题目结构；只读，不启动 PowerPoint。"""
import json
import posixpath
import sys
import zipfile
from xml.etree import ElementTree as ET
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def inspect(path):
    """Follow actual slide/shape relationships, not orphaned tag parts."""
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None, 'Damaged ZIP'

        def relationships(part):
            folder, name = posixpath.split(part)
            relpath = posixpath.join(folder, '_rels', name + '.rels')
            return {e.get('Id'): posixpath.normpath(posixpath.join(folder, e.get('Target')))
                    for e in ET.fromstring(z.read(relpath))
                    if e.get('TargetMode') != 'External'}

        pres = ET.fromstring(z.read('ppt/presentation.xml'))
        prels = relationships('ppt/presentation.xml')
        slides = []
        for sid in pres.findall('p:sldIdLst/p:sldId', NS):
            part = prels[sid.get('{%s}id' % NS['r'])]
            root = ET.fromstring(z.read(part))
            rels = relationships(part)

            def tags(element):
                result = {}
                for tagref in element.findall('p:custDataLst/p:tags', NS):
                    target = rels[tagref.get('{%s}id' % NS['r'])]
                    result.update({e.get('name'): e.get('val') for e in ET.fromstring(z.read(target))})
                return result

            slide = {'tags': tags(root.find('p:cSld', NS)), 'shapes': []}
            for shape in root.findall('.//p:sp', NS):
                nv = shape.find('p:nvSpPr/p:nvPr', NS)
                slide['shapes'].append({
                    'text': ''.join(e.text or '' for e in shape.findall('.//a:t', NS)),
                    'tags': tags(nv) if nv is not None else {}})
            slides.append(slide)
        return {'slide_count': len(slides), 'has_embedded_vba': 'ppt/vbaProject.bin' in z.namelist(), 'slides': slides}

if __name__ == "__main__":
    print(json.dumps(inspect(sys.argv[1]), ensure_ascii=False, indent=2))
