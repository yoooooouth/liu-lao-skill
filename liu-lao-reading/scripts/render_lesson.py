"""Fill the bundled school template and optionally render a two-part reading guide."""
import argparse
from copy import deepcopy
import json
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
TEMPLATE = Path(__file__).resolve().parents[1] / 'assets' / '模板-备课记录本.docx'
PART = 'word/document.xml'
FIELDS = {
    (0, 1): 'teacher', (0, 5): 'class', (1, 1): 'date', (1, 3): 'type',
    (2, 1): 'title', (3, 1): 'standards', (4, 1): 'text', (5, 1): 'learners',
    (6, 1): 'goals', (7, 1): 'focus', (8, 1): 'difficulty', (9, 1): 'methods',
    (11, 1): 'process', (12, 1): 'board', (13, 1): 'reflection',
}
GRID = [[1] * 6, [1, 3, 1, 1]] + [[1, 5]] * 8 + [[5, 1]] + [[1, 5]] * 3
GUIDE_FIELDS = ('读法提示', '工具', '任务', '自评')


def line_value(data, key):
    value = data.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f'请填写 {key}')
    return value.strip()


def fill_cell(cell, content):
    paragraphs = cell.findall(W + 'p')
    base = paragraphs[0]
    for paragraph in paragraphs:
        cell.remove(paragraph)
    prop = base.find(W + 'pPr')
    for line in content.splitlines():
        paragraph = etree.SubElement(cell, W + 'p')
        if prop is not None:
            paragraph.append(deepcopy(prop))
        run = etree.SubElement(paragraph, W + 'r')
        props = etree.SubElement(run, W + 'rPr')
        etree.SubElement(props, W + 'rFonts', {W + x: '仿宋' for x in ('ascii', 'hAnsi', 'eastAsia')})
        etree.SubElement(props, W + 'sz', {W + 'val': '24'})
        text = etree.SubElement(run, W + 't', {XML_SPACE: 'preserve'})
        text.text = line


def grid(rows):
    return [[int(cell.find(W + 'tcPr/' + W + 'gridSpan').get(W + 'val'))
             if cell.find(W + 'tcPr/' + W + 'gridSpan') is not None else 1
             for cell in row.findall(W + 'tc')] for row in rows]


def serialize(root):
    return etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def build_lesson(data, output):
    values = {key: line_value(data, key) for key in FIELDS.values() if key != 'date'}
    values['date'] = line_value(data, 'date') if data.get('date') else '待定'
    if len(values['board'].splitlines()) > 3:
        raise ValueError('板书最多三行')
    if len(values['goals'].splitlines()) != 4:
        raise ValueError('教学目标须恰好四条，每行一条')
    total = data.get('total_period')
    if total is not None and (type(total) is not int or total < 1):
        raise ValueError('总课时序号须为正整数')
    with ZipFile(TEMPLATE) as source:
        root = etree.fromstring(source.read(PART))
        body = root.find(W + 'body')
        table = body.find(W + 'tbl')
        rows = table.findall(W + 'tr')
        if grid(rows) != GRID:
            raise ValueError('模板的行列或合并关系不匹配')
        for (r, cell), key in FIELDS.items():
            fill_cell(rows[r].findall(W + 'tc')[cell], values[key])
        if total is not None:
            nodes = body.findall(W + 'p/' + W + 'r/' + W + 't')
            target = next((node for node in nodes if '总____课时' in (node.text or '')), None)
            if target is not None:
                target.text = target.text.replace('总____课时', f'总{total}课时')
            else:
                for p in body.findall(W+'p'):
                    nodes = p.findall(W+'r/'+W+'t')
                    original=''.join(n.text or '' for n in nodes)
                    if '总____课时' in original:
                        nodes[0].text=original.replace('总____课时', f'总{total}课时')
                        for n in nodes[1:]:
                            n.text=''
                        break
                else:
                    raise ValueError('找不到总课时位置')
        if grid(rows) != GRID:
            raise ValueError('填表后合并关系异常')
        updated = serialize(root)
        with ZipFile(output, 'x', ZIP_DEFLATED) as out:
            for part in source.infolist():
                out.writestr(part, updated if part.filename == PART else source.read(part))


def build_guide(data, output):
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor

    title = line_value(data, 'title')
    student = data.get('student', {})
    teacher = data.get('teacher_notes', {})
    if set(student) != set(GUIDE_FIELDS) or set(teacher) != set(GUIDE_FIELDS):
        raise ValueError('两卷须有相同的读法提示、工具、任务、自评')
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.left_margin = section.right_margin = Mm(23)
    section.top_margin = section.bottom_margin = Mm(20)
    for style_name, font_size in (('Normal', 12), ('Heading 1', 17), ('Heading 2', 13)):
        style = doc.styles[style_name]
        style.font.name = '仿宋'
        style.font.size = Pt(font_size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), '仿宋')
    for version in ('学生卷', '教师卷'):
        heading = doc.add_heading(f'{title} 导学单 · {version}', level=1)
        if version == '教师卷':
            heading.paragraph_format.page_break_before = True
        for field in GUIDE_FIELDS:
            doc.add_heading(field, level=2)
            for line in line_value(student, field).splitlines():
                doc.add_paragraph(line)
            if version == '教师卷':
                for line in line_value(teacher, field).splitlines():
                    doc.add_paragraph(line)
    with output.open('xb') as stream:
        doc.save(stream)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('kind', choices=('lesson', 'guide'))
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    if args.output.suffix.lower() != '.docx':
        parser.error('输出需为 .docx')
    data = json.loads(args.source.read_text(encoding='utf-8'))
    (build_lesson if args.kind == 'lesson' else build_guide)(data, args.output)
    print(args.output)
