"""Check exact question content, style, and a separately rendered PDF."""
from pathlib import Path
from zipfile import ZipFile
import argparse
import hashlib
import json
import re
from copy import deepcopy
from lxml import etree
from pypdf import PdfReader
from build_cheatsheet import read_data, answer_text, FORMULA, NS, qn

def norm(text):
    return re.sub(r'\s+', '', text).translate(str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹', '0123456789'))

def check(data, docx, pdf=None, expected_pages=2):
    with ZipFile(docx) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    paragraphs = root.xpath('.//w:body/w:p', namespaces=NS)
    def flat(el):
        parts = []
        for node in el.iter():
            if node.tag in (qn('w:t'), qn('m:t')):
                parts.append(node.text or '')
            elif node.tag in (qn('w:br'), qn('w:cr')):
                parts.append('\n')
            elif node.tag == qn('w:tab'):
                parts.append('\t')
        return ''.join(parts)
    items = data['items']
    seen = []
    by_id = {x['id']: x for x in items}
    for p in paragraphs:
        text = flat(p)
        key = text.split(' ', 1)[0]
        runs = p.findall(qn('w:r'))
        if key in by_id:
            x = by_id[key]
            compared = deepcopy(p)
            native = compared.xpath('.//m:oMathPara|.//m:oMath[not(ancestor::m:oMathPara)]', namespaces=NS)
            formulas = [data['formulas'][name] for name in FORMULA.findall(x['answer']) if 'omml' in data['formulas'][name]]
            assert len(native) == len(formulas), f'Native formula count mismatch: {key}'
            def signature(node):
                content = node.text if node.tag in (qn('m:t'), qn('w:t')) else None
                return (node.tag, sorted(node.attrib.items()), content, [signature(c) for c in node])
            for node, formula in zip(native, formulas):
                source = etree.fromstring(formula['omml'].encode(), parser=etree.XMLParser(resolve_entities=False, no_network=True))
                assert signature(node) == signature(source), f'Native formula structure mismatch: {key}'
                placeholder = etree.Element(qn('w:r'))
                etree.SubElement(placeholder, qn('w:t')).text = formula['plain_text']
                node.getparent().replace(node, placeholder)
            assert flat(compared) == x['id'] + ' ' + x['question'] + ' A: ' + answer_text(x, data), f'Word text mismatch: {key}'
            boundary = next(i for i, r in enumerate(runs) if flat(r) == ' A: ')
            selected, color = runs[:boundary], 'B00020'
            seen.append(key)
        else:
            assert text in [data['title'], *data['sections'].values()], f'Unexpected paragraph: {text[:60]}'
            selected, color = runs, '1F4E79'
        for run in selected:
            rp = run.find(qn('w:rPr'))
            assert rp.find(qn('w:b')).get(qn('w:val')) == '0', 'Bold title'
            assert rp.find(qn('w:color')).get(qn('w:val')) == color, 'Wrong title color'
    assert seen == [x['id'] for x in items], 'Question IDs missing, repeated or reordered'
    assert not root.xpath('.//w:headerReference|.//w:footerReference|.//w:footnoteReference|.//w:endnoteReference|.//w:instrText|.//w:fldSimple', namespaces=NS)
    pg = root.find('.//' + qn('w:pgSz'))
    assert abs(int(pg.get(qn('w:w'))) / 1440 * 25.4 - 297) < 0.1
    assert abs(int(pg.get(qn('w:h'))) / 1440 * 25.4 - 210) < 0.1
    result = {'questions': len(items), 'word_content_and_styles_verified': True,
              'docx_sha256': hashlib.sha256(Path(docx).read_bytes()).hexdigest(), 'visual_inspection': 'required separately'}
    if pdf:
        reader = PdfReader(pdf)
        assert len(reader.pages) == expected_pages, f'Expected {expected_pages} pages, got {len(reader.pages)}'
        rendered = norm('\n'.join(p.extract_text() for p in reader.pages))
        position = 0
        for x in items:
            question = norm(x['id'] + ' ' + x['question'] + ' A: ')
            start = rendered.find(question, position)
            assert start >= 0, f'Question absent/reordered in PDF: {x["id"]}'
            answer = norm(answer_text(x, data))
            assert rendered.startswith(answer, start + len(question)), f'Answer mismatch in PDF: {x["id"]}'
            position = start + len(question) + len(answer)
        for p in reader.pages:
            assert abs(float(p.mediabox.width) - 841.89) < 1 and abs(float(p.mediabox.height) - 595.28) < 1
        result.update({'rendered_pages': len(reader.pages), 'pdf_text_complete': True})
    return result

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    parser.add_argument('docx')
    parser.add_argument('--pdf')
    parser.add_argument('--expected-pages', type=int, default=2)
    parser.add_argument('--report')
    args = parser.parse_args()
    result = check(read_data(args.input), args.docx, args.pdf, args.expected_pages)
    if args.report:
        Path(args.report).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))
