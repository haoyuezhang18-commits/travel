#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按路线编号从本地模板 DOCX 里只取出那一张表，不通读整份文件。

用法：
    python3 tools/取路线.py LAK-R03
    python3 tools/取路线.py ROM-R02 --root /path/to/旅游通用模板
    python3 tools/取路线.py --list 罗马          # 按关键词列出编号

模板根目录按以下顺序查找：
    1) --root 参数
    2) config/local-paths.yaml 里的 template_root
    3) ~/文件夹/旅游通用模板

依赖：python-docx（pip install python-docx）
"""
import os, sys, csv, argparse

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCATOR = os.path.join(HERE, 'catalog', '路线定位表.csv')


def find_root(arg_root):
    if arg_root:
        return os.path.expanduser(arg_root)
    cfg = os.path.join(HERE, 'config', 'local-paths.yaml')
    if os.path.exists(cfg):
        for line in open(cfg, encoding='utf-8'):
            if line.strip().startswith('template_root:'):
                v = line.split(':', 1)[1].strip().strip('"\'')
                if v and not v.startswith('/path/to'):
                    return os.path.expanduser(v)
    return os.path.expanduser('~/文件夹/旅游通用模板')


def load_rows():
    with open(LOCATOR, encoding='utf-8') as f:
        return list(csv.DictReader(f))


def dump_table(tbl):
    rows = []
    for tr in tbl.rows:
        rows.append([c.text.strip().replace('\n', ' ') for c in tr.cells])
    if not rows:
        return '(空表)'
    out = ['| ' + ' | '.join(rows[0]) + ' |',
           '|' + '---|' * len(rows[0])]
    for r in rows[1:]:
        out.append('| ' + ' | '.join(x.replace('|', '\\|') for x in r) + ' |')
    return '\n'.join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('code', nargs='?', help='路线编号，如 LAK-R03')
    ap.add_argument('--root', help='模板库根目录')
    ap.add_argument('--list', dest='kw', help='按关键词列出编号')
    a = ap.parse_args()
    rows = load_rows()

    if a.kw:
        hit = [r for r in rows if a.kw in r['路线编号'] or a.kw in r['路线名称'] or a.kw in r['模板文件（相对旅游通用模板/）']]
        if not hit:
            print(f'没有匹配 "{a.kw}" 的路线'); return 1
        for r in hit:
            print(f"{r['路线编号']:16} {r['路线名称']}")
        return 0

    if not a.code:
        ap.print_help(); return 1

    code = a.code.upper() if a.code[:3].isalpha() else a.code
    hit = [r for r in rows if r['路线编号'].upper() == code]
    primary = [r for r in hit if not r.get('备注')]
    if primary and len(primary) < len(hit):
        skipped = len(hit) - len(primary)
        hit = primary
        print(f'（另有 {skipped} 份同源副本，已跳过）\n')
    if not hit:
        near = [r['路线编号'] for r in rows if code.split('-')[0] in r['路线编号'].upper()][:8]
        print(f'找不到编号 {a.code}' + (f'，同城的有：{", ".join(near)}' if near else ''))
        return 1

    from docx import Document
    root = find_root(a.root)
    for r in hit:
        path = os.path.join(root, r['模板文件（相对旅游通用模板/）'])
        if not os.path.exists(path):
            print(f'模板文件不存在：{path}'); return 1
        doc = Document(path)
        idx = int(r['文件内第几张表']) - 1
        # 有总索引表的文件（伦敦），实际表序要跳过那张索引表
        tables = doc.tables
        offset = 1 if (tables and tables[0].rows and '路线编号' in tables[0].rows[0].cells[0].text) else 0
        tbl = tables[idx + offset]
        print(f"# 【{r['路线编号']}】{r['路线名称']}")
        print(f"来源：{r['模板文件（相对旅游通用模板/）']}（第 {r['文件内第几张表']} 张表）\n")
        print(dump_table(tbl))
    return 0


if __name__ == '__main__':
    sys.exit(main())
