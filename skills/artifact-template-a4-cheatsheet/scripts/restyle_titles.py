"""Restyle generated Q&A titles in place, preserving all original text."""
from pathlib import Path
from zipfile import ZipFile
import argparse
import json
from lxml import etree
from build_cheatsheet import NS, qn

def restyle(path):
    path = Path(path)
    with ZipFile(path) as z:
        files = [(entry, z.read(entry.filename)) for entry in z.infolist()]
    root = etree.fromstring(next(data for info, data in files if info.filename == 'word/document.xml'))
    old = root.xpath('.//w:t/text()|.//m:t/text()', namespaces=NS)
    questions = headings = 0
    for p in root.xpath('.//w:body/w:p', namespaces=NS):
        runs = p.findall(qn('w:r'))
        labels = [i for i, r in enumerate(runs) if ''.join(r.xpath('.//w:t/text()', namespaces=NS)) == ' A: ']
        assert len(labels) <= 1, 'Ambiguous answer boundary; inspect manually'
        if labels:
            selected, color = runs[:labels[0]], 'B00020'
            questions += 1
        else:
            style = p.find('./' + qn('w:pPr') + '/' + qn('w:pStyle'))
            is_title = style is not None and style.get(qn('w:val')) == 'Title'
            is_section = any(r.xpath('./w:rPr/w:color[@w:val="1F4E79"]', namespaces=NS) for r in runs)
            if not (is_title or is_section):
                continue
            selected, color = runs, '1F4E79'
            headings += 1
        for run in selected:
            rp = run.find(qn('w:rPr'))
            if rp is None:
                rp = etree.Element(qn('w:rPr'))
                run.insert(0, rp)
            for key, value in [('b', '0'), ('bCs', '0'), ('color', color)]:
                el = rp.find(qn('w:' + key))
                if el is None:
                    el = etree.SubElement(rp, qn('w:' + key))
                el.attrib.clear()
                el.set(qn('w:val'), value)
    assert questions, 'No supported Q&A paragraphs; inspect source instead of applying this helper'
    assert root.xpath('.//w:t/text()|.//m:t/text()', namespaces=NS) == old
    xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    tmp = path.with_name(path.stem + '.styling-tmp.docx')
    assert not tmp.exists(), 'Temporary path already exists'
    with ZipFile(tmp, 'w') as z:
        for entry, data in files:
            z.writestr(entry, xml if entry.filename == 'word/document.xml' else data)
    tmp.replace(path)
    return {'questions': questions, 'headings': headings, 'text_unchanged': True, 'output': str(path.resolve()), 'render_required': True}

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    print(json.dumps(restyle(parser.parse_args().docx), ensure_ascii=False))
