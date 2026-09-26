# -*- coding: utf-8 -*-
"""把 JPS_manuscript_draft_v1.md 转成 .docx（含 5 张表）。
针对本稿的 markdown 子集：# ## ### 标题、**粗体**、`代码`、| 表格 |、- 列表、> 引用、--- 分隔。
排版：Times New Roman 12pt，正文双倍行距，表格单倍行距，标题黑色。

2026-07-29 新增（审稿指出表格版式不合格：列宽未设置导致长文本逐字断行，
跨页无重复表头，插图无 alt text）：
  · 表格改 9pt，按各列内容的最大字符数比例分配列宽，窄列有下限，不再被压成逐字断行。
  · 单表所需自然宽度超过顶版可用宽度时，整张表连同其标题一起放进单独的横版（landscape）
    分节，表格结束后立刻切回竖版，不影响前后正文的分节/页边距。
  · 每张表表头行标记 <w:tblHeader/>，跨页续表时表头随页重复。
  · 插图设置 wp:docPr 的 descr（Word“替换文字”/无障碍朗读用）。
"""
import os
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import sys
SRC = sys.argv[1] if len(sys.argv) > 1 else "JPS_manuscript_draft_v2.md"
OUT = sys.argv[2] if len(sys.argv) > 2 else SRC.rsplit(".", 1)[0] + ".docx"
STAR, PIPE = "\x01", "\x02"  # 转义占位符

doc = Document()
st = doc.styles["Normal"]
st.font.name = "Times New Roman"          # 西文与数字
st.font.size = Pt(12)                      # 小四
_rpr = st.element.get_or_add_rPr()
_rpr.get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")   # 中文
# 伦理送审文件为中文公文体裁，单倍行距、段后留白。
st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
st.paragraph_format.space_after = Pt(6)

# ---------------------------------------------------------------- 页面/表格几何
PORTRAIT_W, PORTRAIT_H = Inches(8.5), Inches(11.0)
LANDSCAPE_W, LANDSCAPE_H = Inches(11.0), Inches(8.5)
MARGIN_LR, MARGIN_TB = Inches(1.25), Inches(1.0)
PORTRAIT_USABLE_IN = 8.5 - 2 * 1.25   # 6.0
LANDSCAPE_USABLE_IN = 11.0 - 2 * 1.25  # 8.5
TABLE_FONT_PT = 9
CHARS_PER_INCH = 11.0   # 9pt Times New Roman 粗略估计，宁可偏保守（列偏宽）
MIN_COL_IN = 0.55       # 数字列的下限——曾因下限太小、又被整体等比例压缩，
                        # 把 "0.0016" 这样 6 字符的 p 值从中间挤断成两行
CELL_PAD_IN = 0.16      # 单元格左右内边距合计的粗略估计

sec0 = doc.sections[0]
sec0.page_width, sec0.page_height = PORTRAIT_W, PORTRAIT_H
sec0.left_margin = sec0.right_margin = MARGIN_LR
sec0.top_margin = sec0.bottom_margin = MARGIN_TB

def restore(s):
    return s.replace(STAR, "*").replace(PIPE, "|")

def _clean_len(s):
    """去粗体/斜体标记后估算显示字符数，供列宽估算——不影响实际渲染文本。"""
    s2 = re.sub(r"\*\*([^*]+)\*\*", r"\1", s)
    s2 = re.sub(r"\*([^*]+)\*", r"\1", s2)
    return sum(2 if "\u4e00" <= ch <= "\u9fff" else 1 for ch in s2)

def add_runs(p, text, single=False):
    if single:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    text = text.replace("\\*", STAR).replace("\\|", PIPE)
    # 先切粗体/代码/上标，再在剩余片段里切单星号斜体（期刊名、强调词）
    for part in re.split(r"(\*\*[^*]+\*\*|`[^`]+`|\^[^\s^]+\^)", text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            r = p.add_run(restore(part[2:-2])); r.bold = True
        elif part.startswith("`") and part.endswith("`"):
            r = p.add_run(restore(part[1:-1])); r.font.name = "Consolas"; r.font.size = Pt(10)
        elif part.startswith("^") and part.endswith("^") and len(part) > 2:
            r = p.add_run(restore(part[1:-1])); r.font.superscript = True
        else:
            for sub in re.split(r"(\*[^*]+\*)", part):
                if not sub:
                    continue
                if sub.startswith("*") and sub.endswith("*") and len(sub) > 2:
                    r = p.add_run(restore(sub[1:-1])); r.italic = True
                else:
                    p.add_run(restore(sub))

def add_pagebreak():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

_ALT_COUNTER = [0]
def add_image(path, width_in=6.3, alt=None):
    """整幅居中插图；宽度按页宽收缩，超高的图再按比例限高；设置无障碍替换文字。"""
    if not os.path.exists(path):
        print("  [warn] 图片不存在，跳过:", path); return
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.add_run().add_picture(path, width=Inches(width_in))
    pic = doc.inline_shapes[-1]
    max_h = Inches(8.2)
    if pic.height > max_h:                     # 竖长图按高度重新缩放，避免跨页
        ratio = max_h / pic.height
        pic.height = int(pic.height * ratio); pic.width = int(pic.width * ratio)
    _ALT_COUNTER[0] += 1
    descr = alt or (os.path.splitext(os.path.basename(path))[0].replace("_", " "))
    docPr = pic._inline.find(qn("wp:docPr"))
    if docPr is not None:
        docPr.set("descr", descr)
        docPr.set("title", descr)

def _set_cell_border(cell, edge, val="nil", sz=4, color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders"); tcPr.append(borders)
    el = borders.find(qn("w:%s" % edge))
    if el is None:
        el = OxmlElement("w:%s" % edge); borders.append(el)
    el.set(qn("w:val"), val); el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), "0"); el.set(qn("w:color"), color)

def three_line_table(tbl):
    """三线表：仅表顶粗线、表头下细线、表底粗线；无竖线，正文行间无横线。"""
    tblPr = tbl._tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:%s" % edge); el.set(qn("w:val"), "nil"); tblBorders.append(el)
    tblPr.append(tblBorders)
    for row in tbl.rows:
        for cell in row.cells:
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                _set_cell_border(cell, edge, val="nil")
    THICK, THIN = 12, 6   # sz 单位=1/8pt：12=1.5pt（表顶/表底），6=0.75pt（表头下）
    for cell in tbl.rows[0].cells:
        _set_cell_border(cell, "top", val="single", sz=THICK)
        _set_cell_border(cell, "bottom", val="single", sz=THIN)
    for cell in tbl.rows[-1].cells:
        _set_cell_border(cell, "bottom", val="single", sz=THICK)

def mark_repeating_header(tbl):
    """表头行跨页重复：给首行 trPr 加 <w:tblHeader/>。"""
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true")
    trPr.append(el)

def estimate_col_widths(header, body, usable_in):
    """按各列（表头+全部单元格）最大显示字符数分配列宽。
    两阶段分配，而不是整体等比例缩放——等比例缩放会把窄的数字列也按比例压小，
    曾把 "0.0016" 这样 6 字符的 p 值从中间挤断成两行：
      1) 每列先保证 MIN_COL_IN 下限；
      2) 剩余空间按各列"超出下限的自然需求"比例分配给长文本列，短列不再被压缩。
    返回 (每列宽度, 未截断的自然总宽——供横版判定用)。"""
    ncol = len(header)
    maxlen = []
    for j in range(ncol):
        vals = [_clean_len(header[j])]
        for r in body:
            if j < len(r): vals.append(_clean_len(r[j]))
        maxlen.append(max(vals) if vals else 1)
    natural = [max(MIN_COL_IN, ln / CHARS_PER_INCH + CELL_PAD_IN) for ln in maxlen]
    natural_total = sum(natural)
    floor_total = MIN_COL_IN * ncol
    remaining = usable_in - floor_total
    extra = [n - MIN_COL_IN for n in natural]
    extra_total = sum(extra)
    if remaining <= 0:
        widths = [usable_in / ncol] * ncol
    elif extra_total <= remaining:
        # 底线 + 自然多出部分都放得下：多余空间按现有比例再摊一次，刚好铺满
        leftover = remaining - extra_total
        widths = [MIN_COL_IN + extra[j] + (leftover * extra[j] / extra_total if extra_total > 0 else leftover / ncol)
                  for j in range(ncol)]
    else:
        # 放不下：按各列超出下限的需求比例分配剩余空间，短列始终保住下限
        widths = [MIN_COL_IN + remaining * (extra[j] / extra_total if extra_total > 0 else 1 / ncol)
                  for j in range(ncol)]
    return widths, natural_total

def apply_col_widths(tbl, widths_in):
    tbl.autofit = False
    grid = tbl._tbl.tblGrid
    gridcols = grid.findall(qn("w:gridCol"))
    for gc, w in zip(gridcols, widths_in):
        gc.set(qn("w:w"), str(int(Inches(w).twips)))
    for row in tbl.rows:
        for cell, w in zip(row.cells, widths_in):
            cell.width = Inches(w)

def start_landscape():
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = LANDSCAPE_W, LANDSCAPE_H
    sec.left_margin = sec.right_margin = MARGIN_LR
    sec.top_margin = sec.bottom_margin = MARGIN_TB

def end_landscape():
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = PORTRAIT_W, PORTRAIT_H
    sec.left_margin = sec.right_margin = MARGIN_LR
    sec.top_margin = sec.bottom_margin = MARGIN_TB

def add_h(text, level):
    h = doc.add_heading("", level=level)
    r = h.add_run(text)
    r.font.color.rgb = RGBColor(0, 0, 0)
    r.font.name = "Times New Roman"
    r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")

def build_table(header, body):
    """建表 + 三线边框 + 列宽 + 表头重复 + 9pt 字号。返回是否横版。"""
    widths, natural_total = estimate_col_widths(header, body, PORTRAIT_USABLE_IN)
    landscape = natural_total > PORTRAIT_USABLE_IN
    usable = LANDSCAPE_USABLE_IN if landscape else PORTRAIT_USABLE_IN
    if landscape:
        widths, _ = estimate_col_widths(header, body, usable)
    tbl = doc.add_table(rows=1, cols=len(header))
    for j, htext in enumerate(header):
        cell = tbl.rows[0].cells[j]; cell.paragraphs[0].text = ""
        add_runs(cell.paragraphs[0], htext, single=True)
        for rr in cell.paragraphs[0].runs: rr.bold = True
    for br in body:
        cells = tbl.add_row().cells
        for j, val in enumerate(br):
            if j < len(cells):
                cells[j].paragraphs[0].text = ""
                add_runs(cells[j].paragraphs[0], val, single=True)
    for row in tbl.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(TABLE_FONT_PT)
    three_line_table(tbl)
    apply_col_widths(tbl, widths)
    mark_repeating_header(tbl)
    return tbl, landscape

lines = open(SRC, encoding="utf-8").read().split("\n")

def _peek_table(start):
    """从 start 行起跳过空行，若紧接着是表格则返回 (header, body, 表格结束后的行号)，否则 None。"""
    k = start
    while k < len(lines) and lines[k].strip() == "":
        k += 1
    if k < len(lines) and lines[k].startswith("|") and k + 1 < len(lines) \
            and re.match(r"^\|[\s:|\-]+\|$", lines[k + 1]):
        tl = []
        j = k
        while j < len(lines) and lines[j].startswith("|"):
            tl.append(lines[j]); j += 1
        rows = [[c.strip() for c in r.strip().strip("|").split("|")] for r in tl]
        return rows[0], rows[2:], j
    return None

i = 0
in_landscape = False
while i < len(lines):
    line = lines[i]
    # 表格：可能是紧跟在 "**Table N.**" / "**Supplementary Table SN.**" 标题之后，
    # 也可能是校正模型这类无独立标题的第二张表——两种都在这里统一处理列宽/横版判定。
    if line.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s:|\-]+\|$", lines[i + 1]):
        tl = []
        while i < len(lines) and lines[i].startswith("|"):
            tl.append(lines[i]); i += 1
        rows = [[c.strip() for c in r.strip().strip("|").split("|")] for r in tl]
        header, body = rows[0], rows[2:]
        _, natural_total = estimate_col_widths(header, body, PORTRAIT_USABLE_IN)
        need_landscape = natural_total > PORTRAIT_USABLE_IN
        if need_landscape and not in_landscape:
            start_landscape(); in_landscape = True
        build_table(header, body)
        if in_landscape and not need_landscape:
            # 理论上不会出现（need_landscape 已据同一次估算得出），保留作为防御
            pass
        doc.add_paragraph()
        # 表格后紧跟的若不是另一张表，则关闭横版，回到竖版继续正文/脚注
        nxt = _peek_table(i)
        if in_landscape and nxt is None:
            end_landscape(); in_landscape = False
        continue
    m_img = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$", line.strip())
    if m_img:
        add_image(m_img.group(2), alt=m_img.group(1) or None); i += 1; continue
    if line.strip() in ("<!--pagebreak-->", "\\pagebreak"):
        add_pagebreak(); i += 1; continue
    if line.startswith("### "):
        add_h(line[4:], 2)
    elif line.startswith("## "):
        add_h(line[3:], 1)
    elif line.startswith("# "):
        add_h(line[2:], 0)
    elif line.strip() == "---":
        pass
    elif line.startswith("> "):
        p = doc.add_paragraph(); p.paragraph_format.left_indent = Pt(18)
        add_runs(p, line[2:])
        for r in p.runs: r.italic = True
    elif line.startswith("- "):
        add_runs(doc.add_paragraph(style="List Bullet"), line[2:])
    elif line.strip() == "":
        pass
    else:
        # "**Table N.**" / "**Supplementary Table SN.**" 标题：若紧跟的表需要横版，
        # 提前在标题之前开横版分节，让标题与表格同页，不被分节断开。
        is_table_caption = bool(re.match(r"^\*\*(Table \d|Supplementary Table S\d)\.\*\*", line))
        if is_table_caption and not in_landscape:
            peeked = _peek_table(i + 1)
            if peeked:
                hdr, bdy, _ = peeked
                _, nat = estimate_col_widths(hdr, bdy, PORTRAIT_USABLE_IN)
                if nat > PORTRAIT_USABLE_IN:
                    start_landscape(); in_landscape = True
        add_runs(doc.add_paragraph(), line)
    i += 1

if in_landscape:
    end_landscape()

doc.save(OUT)
print("saved", OUT, "| tables:", len(doc.tables), "| paragraphs:", len(doc.paragraphs))
