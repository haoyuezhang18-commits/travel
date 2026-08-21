#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
让 AI 直接看已交付成品的实际效果，而不是只看空骨架。

用法：
    python3 tools/看成品.py --list                 # 列出可参考的成品
    python3 tools/看成品.py 英国15天                # 看结构大纲（概述+每天标题+表头）
    python3 tools/看成品.py 英国15天 --day 2        # 完整打印第2天那张表，看备注写到什么颗粒度
    python3 tools/看成品.py 票务与预约 --day 1      # 文档2 同理，按表序号

注意：这些是**真实客户成品**，只看格式和颗粒度，
不要把里面的客户日期、人数、酒店、航班抄进新订单。
"""
import os, sys, csv, re, zipfile, argparse
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(HERE, 'catalog', '成品参考.csv')
HOME = os.path.expanduser('~')


def load():
    with open(REF, encoding='utf-8') as f:
        return list(csv.DictReader(f))


def txt(el):
    return re.sub(r'\s+', ' ', ''.join(x.text or '' for x in el.iter(W + 't'))).strip()


def body_of(path):
    z = zipfile.ZipFile(path)
    return ET.fromstring(z.read('word/document.xml')).find(W + 'body'), z


def outline(path):
    body, z = body_of(path)
    imgs = len([n for n in z.namelist() if n.startswith('word/media/')])
    ti = 0
    overview, out = [], []
    for ch in body:
        if ch.tag == W + 'p':
            t = txt(ch)
            if not t:
                continue
            if t.startswith('✓') or t.startswith('✔'):
                overview.append(t)
            else:
                out.append(('§', t))
        elif ch.tag == W + 'tbl':
            ti += 1
            trs = ch.findall(W + 'tr')
            head = [txt(tc) for tc in trs[0].findall(W + 'tc')] if trs else []
            out.append(('T', ti, ' | '.join(head), len(trs)))
    print(f'（共 {ti} 张表、{imgs} 张图）\n')
    if overview:
        print('—— 开头的每日一行概述 ——')
        for o in overview:
            print(' ', o[:110])
        print()
    print('—— 正文结构 ——')
    for item in out:
        if item[0] == '§':
            print(' §', item[1][:90])
        else:
            print(f'   [表{item[1]}] {item[2]}   （{item[3]}行）')
    print(f'\n看某一张表的完整写法：--day <表序号>')


def one_table(path, n):
    body, _ = body_of(path)
    tbls = body.findall(W + 'tbl')
    if not 1 <= n <= len(tbls):
        print(f'只有 {len(tbls)} 张表'); return 1
    # 顺序遍历，记住第 n 张表前面最近的一段非空文字
    ti, title = 0, ''
    for ch in body:
        if ch.tag == W + 'p':
            t = txt(ch)
            if t:
                title = t
        elif ch.tag == W + 'tbl':
            ti += 1
            if ti == n:
                break
    if title:
        print(f'§ {title}\n')
    for tr in tbls[n - 1].findall(W + 'tr'):
        cells = [txt(tc) for tc in tr.findall(W + 'tc')]
        print(' ‖ '.join(cells))
        print('-' * 100)
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('name', nargs='?', help='参考名称关键词')
    ap.add_argument('--day', type=int, help='完整打印第几张表')
    ap.add_argument('--list', action='store_true', dest='ls')
    a = ap.parse_args()
    rows = load()

    if a.ls or not a.name:
        cur = None
        for r in rows:
            if r['类型'] != cur:
                cur = r['类型']; print(f'\n【{cur}】')
            print(f"  {r['参考名称']:<24} {r['展示了什么']}")
        print('\n用法：python3 tools/看成品.py <参考名称关键词> [--day N]')
        return 0

    hit = [r for r in rows if a.name in r['参考名称'] or a.name in r['相对家目录路径']]
    if not hit:
        print(f'没有匹配 "{a.name}" 的参考成品，用 --list 看全部'); return 1
    r = hit[0]
    path = os.path.join(HOME, r['相对家目录路径'])
    if not os.path.exists(path):
        print(f'文件不存在：{path}'); return 1
    print(f"# {r['参考名称']}（{r['类型']}）")
    print(f"来源：~/{r['相对家目录路径']}")
    print(f"看点：{r['展示了什么']}")
    print('⚠ 真实客户成品，只看格式和颗粒度，客户日期/人数/酒店/航班不要抄进新订单\n')
    return one_table(path, a.day) if a.day else (outline(path) or 0)


if __name__ == '__main__':
    sys.exit(main())
