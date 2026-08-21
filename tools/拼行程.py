#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按路线编号，把模板里那张表**连同图片和排版原样搬进**新的交付文档，再批量替换变量。

不重打内容、不丢图片——这是做逐日行程的正确姿势。

用法：
    python3 tools/拼行程.py 配方.json 输出.docx

配方.json 格式：
{
  "标题": "西意法15天行程　文档1",
  "概述": ["✓ 9.09 周三 巴塞罗那：...", "..."],
  "全局替换": {"8月": "9月"},
  "天": [
    {"标题": "D1｜9.09（周三）巴塞罗那", "编号": "BCN-R02",
     "说明": "抵达日，上午落地可当完整一天用",
     "替换": {"5 月": "9 月", "09:30 - 12:00": "10:30 - 12:15"}}
  ]
}
"""
import os, sys, csv, json, copy, argparse, io, hashlib
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCATOR = os.path.join(HERE, 'catalog', '路线定位表.csv')
FONT, SIZE = 'HYJunHei-EEJ', Pt(9)
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def find_root(arg=None):
    if arg:
        return os.path.expanduser(arg)
    cfg = os.path.join(HERE, 'config', 'local-paths.yaml')
    if os.path.exists(cfg):
        for line in open(cfg, encoding='utf-8'):
            if line.strip().startswith('template_root:'):
                v = line.split(':', 1)[1].strip().strip('"\'')
                if v and not v.startswith('/path/to'):
                    return os.path.expanduser(v)
    return os.path.expanduser('~/文件夹/旅游通用模板')


def locate(code):
    with open(LOCATOR, encoding='utf-8') as f:
        rows = [r for r in csv.DictReader(f) if r['路线编号'].upper() == code.upper()]
    rows = [r for r in rows if not r.get('备注')] or rows
    if not rows:
        raise SystemExit(f'找不到路线编号 {code}')
    return rows[0]


def get_table(src_doc, idx):
    """定位表：伦敦模板首张是总索引表，需跳过。"""
    tables = src_doc.tables
    off = 1 if (tables and tables[0].rows and '路线编号' in tables[0].rows[0].cells[0].text) else 0
    return tables[idx - 1 + off]


def copy_images(tbl_el, src_part, dst_part, seen):
    """把表格里引用的图片搬到新文档，并重写 rId。

    不能直接 relate_to 源文档的 part：不同模板里都有 image1.png，
    partname 会在新包里撞车，产出重复 zip 条目、文件损坏。
    这里改为取出图片二进制，交给 get_or_add_image 重新登记，
    它会按 sha1 去重并分配唯一 partname。
    """
    n = 0
    for el in tbl_el.iter():
        for attr in (R_NS + 'embed', R_NS + 'link', R_NS + 'id'):
            rid = el.get(attr)
            if not rid:
                continue
            try:
                part = src_part.related_parts[rid]
            except KeyError:
                continue
            blob = part.blob
            key = hashlib.sha1(blob).hexdigest()
            if key not in seen:
                new_rid, _ = dst_part.get_or_add_image(io.BytesIO(blob))
                seen[key] = new_rid
            el.set(attr, seen[key])
            n += 1
    return n


def replace_in_table(tbl_el, mapping):
    """在表格的文字节点上做替换，保留原有格式与图片。"""
    hits = 0
    for t in tbl_el.iter(qn('w:t')):
        if not t.text:
            continue
        new = t.text
        for a, b in mapping.items():
            if a in new:
                new = new.replace(a, b)
        if new != t.text:
            t.text = new
            hits += 1
    return hits


def styled(p, text, bold=False, size=SIZE, color=None):
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = size
    r.bold = bold
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    if color:
        r.font.color.rgb = color
    return p


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('recipe'); ap.add_argument('out'); ap.add_argument('--root')
    a = ap.parse_args()
    cfg = json.load(open(a.recipe, encoding='utf-8'))
    root = find_root(a.root)

    doc = Document()
    st = doc.styles['Normal']
    st.font.name = FONT; st.font.size = SIZE
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    for s in doc.sections:
        s.left_margin = s.right_margin = Pt(36)

    styled(doc.add_paragraph(), cfg.get('标题', ''), bold=True, size=Pt(16))
    for line in cfg.get('概述', []):
        styled(doc.add_paragraph(), line)
    doc.add_paragraph()

    total_img = 0
    cache = {}
    seen_imgs = {}
    for day in cfg['天']:
        code = day['编号']
        row = locate(code)
        path = os.path.join(root, row['模板文件（相对旅游通用模板/）'])
        if path not in cache:
            cache[path] = Document(path)
        src = cache[path]
        tbl = get_table(src, int(row['文件内第几张表']))

        styled(doc.add_paragraph(), day['标题'], bold=True, size=Pt(13))
        if day.get('说明'):
            styled(doc.add_paragraph(), day['说明'], size=Pt(8.5),
                   color=RGBColor(0x88, 0x44, 0x00))

        new_el = copy.deepcopy(tbl._tbl)
        n = copy_images(new_el, src.part, doc.part, seen_imgs)
        total_img += n
        mapping = dict(cfg.get('全局替换', {})); mapping.update(day.get('替换', {}))
        hits = replace_in_table(new_el, mapping)
        doc.element.body.append(new_el)
        doc.add_paragraph()
        print(f"  {day['标题'][:22]:24} ← 【{code}】{row['路线名称'][:18]}　图{n} 替换{hits}处")

    doc.save(a.out)
    print(f'\n已生成：{a.out}　（共 {len(cfg["天"])} 天，搬运图片 {total_img} 处）')


if __name__ == '__main__':
    main()
