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
import os, sys, csv, json, copy, argparse, io, hashlib, re
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


def table_images(tbl):
    return len(list(tbl._tbl.iter(qn('a:blip'))))


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


TR = qn('w:tr'); TC = qn('w:tc'); WT = qn('w:t')


def row_text(tr, col=None):
    tcs = tr.findall(TC)
    if col is None:
        return ' '.join(''.join(t.text or '' for t in tc.iter(WT)) for tc in tcs)
    if col >= len(tcs):
        return ''
    return ''.join(t.text or '' for t in tcs[col].iter(WT))


def _has_img(el):
    return el.find('.//' + qn('w:drawing')) is not None or el.find('.//' + qn('w:pict')) is not None


def _proto_run(p):
    """从段落里挑一个「只带格式」的 run 当模子：只保留 w:rPr，其余子元素全部剔掉。

    必须剔干净，不能只删 w:t。模板里的 run 可能带：
      · w:fldChar —— 域代码的开头／结尾（超链接、页码这些都是域）
      · w:instrText —— 域的指令文字
      · w:br / w:tab / w:drawing / w:pict / w:footnoteReference …

    踩过的坑（2026-09 伦敦5天单）：模子里留下了一个 `w:fldChar begin`，
    于是追加的每一行都开了一个**永不闭合的域**。后果是
    **WPS 从那里开始整篇停止渲染（5 页塌成 1 页、字数 1546 → 217），
    Word 直接报「文件已损坏」**——而文字其实都在 XML 里，
    python-docx 读得出来，所以原来的自检一条都抓不到。
    """
    for r in p.findall(qn('w:r')):
        if _has_img(r):
            continue
        proto = copy.deepcopy(r)
        for ch in list(proto):
            if ch.tag != qn('w:rPr'):
                proto.remove(ch)
        return proto
    return None


def append_cell(tr, col, text):
    """在格子末尾追加几行，原有文字、格式、图片全部保留。
    用于「模板写得好、只想补一两句」的情况——别用 set_cell 整格覆盖。"""
    tcs = tr.findall(TC)
    if col >= len(tcs):
        return
    tc = tcs[col]
    ps = tc.findall(qn('w:p'))
    proto_p = None
    for p_ in ps:
        if not _has_img(p_):
            proto_p = p_
    if proto_p is None:
        proto_p = ps[-1] if ps else None
    for line in str(text).split('\n'):
        newp = copy.deepcopy(proto_p) if proto_p is not None else tc.makeelement(qn('w:p'), {})
        for ch in list(newp):
            if ch.tag != qn('w:pPr'):
                newp.remove(ch)
        proto_r = _proto_run(proto_p) if proto_p is not None else None
        r = proto_r if proto_r is not None else newp.makeelement(qn('w:r'), {})
        t = r.makeelement(WT, {})
        t.text = line
        t.set(qn('xml:space'), 'preserve')
        r.append(t)
        newp.append(r)
        tc.append(newp)


def set_cell(tr, col, text):
    """把某格文字整体换掉，保留该格第一个 run 的格式；● 分点用真换行。
    带图片的段落原样保留并挪到文字后面——改备注不能把模板攒的图删掉。"""
    tcs = tr.findall(TC)
    if col >= len(tcs):
        return
    tc = tcs[col]
    ps = tc.findall(qn('w:p'))
    imgs = [p for p in ps if _has_img(p)]      # 图片段落，原样留着
    texts = [p for p in ps if p not in imgs]   # 文字段落，才是要换掉的
    if texts:
        keep = texts[0]
        for extra in texts[1:]:
            tc.remove(extra)
    else:
        keep = tc.makeelement(qn('w:p'), {})
        (imgs[0].addprevious(keep) if imgs else tc.append(keep))
    for p in imgs:                              # 图片统一挪到文字后面
        tc.remove(p)
        tc.append(p)
    proto = _proto_run(keep)                    # 只带格式的模子，见 _proto_run 的说明
    for r in keep.findall(qn('w:r')):
        keep.remove(r)
    for i, line in enumerate(str(text).split('\n')):
        r = copy.deepcopy(proto) if proto is not None else keep.makeelement(qn('w:r'), {})
        if i:
            r.append(r.makeelement(qn('w:br'), {}))
        t = r.makeelement(WT, {})
        t.text = line
        t.set(qn('xml:space'), 'preserve')
        r.append(t)
        keep.append(r)


COLS = {'时间段': 0, '行程内容': 1, '交通': 2, '备注': 3}


def find_rows(rows, kw, whole=False, col=None, exact=False):
    """默认只在「行程内容」列匹配。

    备注列里经常顺带提到别的地名（例如古罗马广场的备注写着
    "从后门6出口出去就是真理之口"），整行匹配会改错行。
    """
    if exact:
        hit = [tr for tr in rows if row_text(tr, 1).strip() == kw]
        if hit:
            return hit
    if col is not None:
        return [tr for tr in rows if kw in row_text(tr, COLS.get(col, col))]
    if whole:
        return [tr for tr in rows if kw in row_text(tr)]
    hit = [tr for tr in rows if kw in row_text(tr, 1)]
    return hit if hit else [tr for tr in rows if kw in row_text(tr, 0)]


def apply_row_ops(tbl_el, day):
    """删 / 改 / 增 / 重排——字符串替换做不到的都在这里。"""
    log = []
    def data_rows():
        return tbl_el.findall(TR)[1:]

    for item in day.get('删', []):
        kw, col = (item, None) if isinstance(item, str) else (item['配'], item.get('配列'))
        hit = find_rows(data_rows(), kw, whole=day.get('整行匹配', False), col=col)
        if not hit:
            log.append(f'⚠ 删除失败，找不到「{kw}」')
        for tr in hit:
            tbl_el.remove(tr); log.append(f'删除「{kw}」')

    for op in day.get('改', []):
        kw = op['配']
        hit = find_rows(data_rows(), kw, whole=op.get('整行匹配', False),
                        col=op.get('配列'), exact=op.get('精确', False))
        if not hit:
            log.append(f'⚠ 修改失败，找不到「{kw}」'); continue
        if len(hit) > 1 and not op.get('允许多行'):
            log.append(f'⚠ 「{kw}」匹配到 {len(hit)} 行，只改第一行（如需全改加 允许多行:true）')
        for tr in (hit if op.get('允许多行') else hit[:1]):
            for ci, key in enumerate(['时间段', '行程内容', '交通', '备注']):
                if key in op:
                    set_cell(tr, ci, op[key])
            for ci, key in enumerate(['追加时间段', '追加行程内容', '追加交通', '追加备注']):
                if key in op:
                    append_cell(tr, ci, op[key])
            log.append(f'修改「{kw}」')

    for op in day.get('增', []):
        rows = tbl_el.findall(TR)
        proto = copy.deepcopy(rows[-1] if len(rows) > 1 else rows[0])
        for dr in proto.iter(qn('w:drawing')):     # 新行不带图
            dr.getparent().remove(dr)
        for ci, key in enumerate(['时间段', '行程内容', '交通', '备注']):
            set_cell(proto, ci, op.get(key, ''))
        anchor = op.get('位置', '尾')
        if anchor == '首':
            rows[0].addnext(proto)
        elif anchor == '尾':
            rows[-1].addnext(proto)
        else:
            hit = find_rows(tbl_el.findall(TR)[1:], anchor)
            if not hit:
                log.append(f'⚠ 插入位置找不到「{anchor}」，改放到表尾')
                tbl_el.findall(TR)[-1].addnext(proto)
            else:
                hit[-1].addnext(proto)
        log.append(f"新增「{str(op.get('行程内容',''))[:16]}」")

    # 搬图：把带图片的段落从一行挪到另一行的同一列。
    # 为什么需要：模板里的照片是跟着具体那家店／那个机位的。换掉一格的餐厅推荐，
    # 照片就对不上人了（踩过：D1 晚餐改成考文特花园，可那格的照片是孔雀临江宴——
    # 窗外大本钟的中餐厅，属于西敏区）。正确做法是把店和它的照片一起挪到对得上的那一行，
    # 而不是删照片（铁律第 13 条）。
    for op in day.get('搬图', []):
        src = find_rows(tbl_el.findall(TR)[1:], op['从'], col=op.get('从列'))
        dst = find_rows(tbl_el.findall(TR)[1:], op['到'], col=op.get('到列'))
        if not src or not dst:
            log.append(f"⚠ 搬图找不到「{op['从'] if not src else op['到']}」")
            continue
        ci = COLS.get(op.get('列', '备注'), 3)
        stc, dtc = src[0].findall(TC), dst[0].findall(TC)
        if ci >= len(stc) or ci >= len(dtc):
            log.append(f"⚠ 搬图：第 {ci + 1} 列不存在")
            continue
        moved = 0
        for p in [p for p in stc[ci].findall(qn('w:p')) if _has_img(p)]:
            stc[ci].remove(p)
            dtc[ci].append(p)
            moved += 1
        if moved:
            log.append(f"搬图「{op['从'][:10]}」→「{op['到'][:10]}」{moved} 张")
        else:
            log.append(f"⚠ 搬图：「{op['从'][:10]}」里没有带图的段落")

    if day.get('序'):
        rows = tbl_el.findall(TR)
        head, data = rows[0], rows[1:]
        ordered, rest = [], list(data)
        for kw in day['序']:
            hit = [tr for tr in rest if kw in row_text(tr, 1)] or [tr for tr in rest if kw in row_text(tr)]
            if hit:
                # 只取「行程内容」与首个命中完全相同的那一组：
                # 既能让交通备选这类重复行整组移动，又不会把
                # 「中央市场」误吞掉「中央市场 → 学院美术馆」。
                same = row_text(hit[0], 1).strip()
                group = [tr for tr in hit if row_text(tr, 1).strip() == same] or hit[:1]
                for tr in group:
                    ordered.append(tr); rest.remove(tr)
            else:
                log.append(f'⚠ 排序找不到「{kw}」')
        ordered += rest
        for tr in data:
            tbl_el.remove(tr)
        prev = head
        for tr in ordered:
            prev.addnext(tr); prev = tr
        log.append('已按指定顺序重排')
    return log


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


TIME_RE = re.compile(r'(\d{1,2})\s*[:：]\s*(\d{2})')

# 交付文件面向客户，不得出现内部用语（铁律第20条）
INTERNAL = ['客户', '负责人', '模板', '沿用', '新增内容', '本单', '上一位',
            '复用', '匹配度', '提醒客户', '这一单', '原表']


def audit(tbl_el, day):
    """生成后自检：时间倒流、旧客户信息残留、重复时段。"""
    warn = []
    rows = tbl_el.findall(TR)[1:]
    mins, labels = [], []
    for tr in rows:
        m = TIME_RE.search(row_text(tr, 0))
        if m:
            mins.append(int(m.group(1)) * 60 + int(m.group(2)))
            labels.append(row_text(tr, 0).strip()[:14] + ' ' + row_text(tr, 1).strip()[:16])
    for i in range(1, len(mins)):
        if mins[i] < mins[i - 1] - 5:
            warn.append(f'时间倒流：{labels[i-1]} → {labels[i]}')
    exact = {}
    for tr in rows:
        span = row_text(tr, 0).strip()
        body = row_text(tr, 1).strip()
        if not span or '备选' in body or '备选' in row_text(tr, 2):
            continue
        exact.setdefault(span, []).append(body[:20])
    for span, bodies in exact.items():
        if len(bodies) > 1:
            warn.append(f'时段完全重复（{span}）：' + ' / '.join(bodies))
    # 只查「行程内容」列：备注里顺带提一句别的景点是正常的
    body = ' '.join(row_text(tr, 1) for tr in rows)
    for bad in day.get('不应出现', []):
        if bad in body:
            warn.append(f'残留旧信息：「{bad}」')

    # 内部用语：整行（含备注）都要查
    for tr in rows:
        whole = row_text(tr)
        for w in INTERNAL:
            if w in whole:
                frag = next((ln.strip() for ln in whole.split('●') if w in ln), whole)
                warn.append(f'出现内部用语「{w}」：{frag[:46]}')
                break
    return warn


# ---------------------------------------------------------------- 样式补齐
# 为什么需要这一步：输出文档是 python-docx 的空白模板（styles.xml 用
# Normal / Heading1 这类英文名 ID），而 WPS 做的城市模板用的是 "1"/"2"/"5"
# 这种数字 ID。整表搬过来后，表格里的 pStyle/rStyle/tblStyle 指向的 ID 在
# 输出文档里不存在——Word 遇到悬空样式引用直接报「文件已损坏」，WPS 宽容
# 一些但会渲染不全（实测伦敦5天单：4 个悬空 ID，Word 打不开）。
# 多模板拼表时尤其容易踩到，因为只可能带走其中一份 styles.xml。
STYLE_REF = {'pStyle': 'paragraph', 'rStyle': 'character', 'tblStyle': 'table'}


def _style_index(styles_el):
    """styleId -> (type, name)　和　(type, name) -> styleId"""
    by_id, by_name = {}, {}
    for st in styles_el.findall(qn('w:style')):
        sid = st.get(qn('w:styleId')); typ = st.get(qn('w:type'))
        nm = st.find(qn('w:name'))
        name = nm.get(qn('w:val')) if nm is not None else ''
        by_id[sid] = (typ, name, st)
        by_name[(typ, name)] = sid
    return by_id, by_name


def fix_styles(doc, src_docs):
    """把搬进来的表格里悬空的样式引用补齐或重映射。返回日志。"""
    log = []
    dst_styles = doc.styles.element
    d_by_id, d_by_name = _style_index(dst_styles)
    src_idx = [_style_index(s.styles.element)[0] for s in src_docs]

    imported, remapped, dropped = {}, 0, 0
    for tag, typ in STYLE_REF.items():
        for ref in list(doc.element.body.iter(qn('w:' + tag))):
            sid = ref.get(qn('w:val'))
            if sid is None or sid in d_by_id:
                continue
            found = None
            for idx in src_idx:                       # 在各来源模板里按 ID 找
                if sid in idx and idx[sid][0] == typ:
                    found = idx[sid]; break
            if found is None:
                ref.getparent().remove(ref); dropped += 1
                continue
            _, name, st_el = found
            if (typ, name) in d_by_name:              # 输出里有同名同类样式 → 改指向
                ref.set(qn('w:val'), d_by_name[(typ, name)])
                remapped += 1
                continue
            if sid not in imported:                   # 没有 → 把定义搬过来
                new = copy.deepcopy(st_el)
                for child in ('basedOn', 'link', 'next'):
                    el = new.find(qn('w:' + child))
                    if el is None:
                        continue
                    tgt = el.get(qn('w:val'))
                    if tgt in d_by_id:
                        continue
                    # 套娃引用也要解：按名字找输出里的对应样式，找不到就去掉
                    r = None
                    for idx in src_idx:
                        if tgt in idx:
                            r = d_by_name.get((idx[tgt][0], idx[tgt][1])); break
                    if r:
                        el.set(qn('w:val'), r)
                    else:
                        new.remove(el)
                dst_styles.append(new)
                d_by_id[sid] = (typ, name, new)
                d_by_name[(typ, name)] = sid
                imported[sid] = name
    if imported:
        log.append('补入样式定义：' + '、'.join(f'{k}({v})' for k, v in imported.items()))
    if remapped:
        log.append(f'重映射同名样式引用 {remapped} 处')
    if dropped:
        log.append(f'删除无法解析的样式引用 {dropped} 处')
    return log


def unbalanced_fields(doc):
    """自检：域（w:fldChar）有没有 begin/end 不配对。

    一个没闭合的 `fldChar begin` 会让 WPS 从那里开始整篇停止渲染、
    Word 报「文件已损坏」，而文字在 XML 里都还在——内容类的自检抓不到，
    必须单独查。逐个单元格查，不只看全篇总数，避免一个多一个少互相抵消。
    """
    bad = []
    for ti, tbl in enumerate(doc.tables, 1):
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                kinds = [fc.get(qn('w:fldCharType'))
                         for fc in cell._tc.iter(qn('w:fldChar'))]
                if kinds.count('begin') != kinds.count('end'):
                    bad.append(f'第{ti}张表 第{ri}行 第{ci + 1}列 域未闭合'
                               f'（begin {kinds.count("begin")} / end {kinds.count("end")}）')
    return bad


def dangling_styles(doc):
    """交付前自检：还有没有指向不存在样式的引用（Word 会判文件损坏）。"""
    have = {st.get(qn('w:styleId')) for st in doc.styles.element.findall(qn('w:style'))}
    bad = set()
    for tag in STYLE_REF:
        for ref in doc.element.body.iter(qn('w:' + tag)):
            v = ref.get(qn('w:val'))
            if v and v not in have:
                bad.add(f'{tag}={v}')
    return sorted(bad)


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

    # 标题与概述同样面向客户，一并查内部用语
    head_warn = []
    for txt in [cfg.get('标题', '')] + list(cfg.get('概述', [])):
        for w in INTERNAL:
            if w in txt:
                head_warn.append(f'标题/概述出现内部用语「{w}」：{txt.strip()[:46]}')
                break
    styled(doc.add_paragraph(), cfg.get('标题', ''), bold=True, size=Pt(16))
    for line in cfg.get('概述', []):
        styled(doc.add_paragraph(), line)
    doc.add_paragraph()

    total_img = 0
    cache = {}
    used_srcs = set()
    seen_imgs = {}
    problems = list(head_warn)
    for w in head_warn:
        print('  ⚠ ' + w)

    for day in cfg['天']:
        sources = day.get('取') or [{'编号': day['编号']}]

        styled(doc.add_paragraph(), day['标题'], bold=True, size=Pt(13))
        # 「说明」是给负责人的内部信息，只打印到终端，不写进客户文件（铁律第20条）

        new_el, n, srcnote = None, 0, []
        for si, sc in enumerate(sources):
            row = locate(sc['编号'])
            path = os.path.join(root, row['模板文件（相对旅游通用模板/）'])
            if path not in cache:
                cache[path] = Document(path)
            src = cache[path]
            used_srcs.add(path)
            tbl = get_table(src, int(row['文件内第几张表']))
            srcnote.append(f"{sc['编号']}(图{table_images(tbl)})")
            part = copy.deepcopy(tbl._tbl)
            n += copy_images(part, src.part, doc.part, seen_imgs)
            keep = sc.get('要')
            if keep is not None:                       # 只取指定的行
                for tr in part.findall(TR)[1:]:
                    if not any(k in row_text(tr, 1) or k in row_text(tr, 0) for k in keep):
                        part.remove(tr)
            if new_el is None:
                new_el = part
            else:                                       # 追加到主表后面
                if keep is None:
                    rows_to_add = part.findall(TR)[1:]
                else:
                    rows_to_add = part.findall(TR)[1:]
                for tr in rows_to_add:
                    new_el.append(tr)
        code = ' + '.join(srcnote)
        total_img += n
        mapping = dict(cfg.get('全局替换', {})); mapping.update(day.get('替换', {}))
        hits = replace_in_table(new_el, mapping)
        oplog = apply_row_ops(new_el, day)
        # 必须插在 sectPr 之前、且紧跟本天标题；
        # 直接 body.append 会把所有表格堆到文档末尾、彼此黏连。
        anchor_p = doc.add_paragraph()
        anchor_p._p.addprevious(new_el)
        gap = anchor_p
        gap.paragraph_format.space_after = Pt(16)
        gap.paragraph_format.space_before = Pt(8)
        styled(gap, '')
        print(f"  {day['标题'][:20]:22} ← {code} 图{n} 替换{hits}处 " +
              (f"行操作{len(oplog)}项" if oplog else ""))
        if day.get('说明'):
            print(f"      〔内部〕{day['说明']}")
        for L in oplog:
            if L.startswith('⚠'):
                print('      ' + L); problems.append(f"{day['标题'][:12]}｜{L}")
        for w in audit(new_el, day):
            print('      ⚠ ' + w); problems.append(f"{day['标题'][:12]}｜{w}")

    for L in fix_styles(doc, [cache[p] for p in sorted(used_srcs)]):
        print('  · ' + L)
    bad = dangling_styles(doc)
    if bad:
        problems.append('样式引用悬空（Word 会判文件损坏）：' + '、'.join(bad))
        print('  ⚠ 样式引用悬空：' + '、'.join(bad))
    for w in unbalanced_fields(doc):
        problems.append('域未闭合（WPS 会整篇停止渲染、Word 判损坏）：' + w)
        print('  ⚠ ' + w)

    doc.save(a.out)
    print(f'\n已生成：{a.out}　（共 {len(cfg["天"])} 天，搬运图片 {total_img} 处）')
    if problems:
        print(f'\n❌ 自检发现 {len(problems)} 个问题，必须修完再交付：')
        for p in problems:
            print('   ' + p)
    else:
        print('\n✅ 自检通过：时间顺序正常、无重复时段、无旧客户信息残留')


if __name__ == '__main__':
    main()
