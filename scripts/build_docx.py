#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_docx.py — 把 transcript.json 渲染成排版完整的 Word 对话记录。

用法：
  python build_docx.py transcript.json -o 输出.docx
  python build_docx.py transcript.json -o 输出.docx --title "第一个 skill" --subtitle "副标题"
  python build_docx.py transcript.json -o 输出.docx --narrative narrative.json

narrative.json 可选字段（缺省时自动生成中性版本，绝不编造事实）：
  {
    "kicker": "WORKBUDDY ｜ SKILL 构建实录",
    "title": "第一个 skill",
    "subtitle": "一句话说清这段对话在做什么",
    "cover_rows": [["字段","值"], ...],
    "lead": ["文档说明段落", ...],
    "stages": [["阶段一","主题","时间","当时要解决什么","结果"], ...],
    "after_stages": "表格后的补充段落",
    "sections": [
      {"h": "四、交付物清单",
       "blocks": [
         {"type": "body", "text": "..."},
         {"type": "h2", "text": "4.1 小节"},
         {"type": "kv", "k": "路径", "v": "..."},
         {"type": "code", "text": "git clone ..."},
         {"type": "table", "cols": ["列名","列名"], "rows": [["",""], ...], "widths": [4,12]}
       ]}
    ],
    "appendix": [["标题","内容"], ...],
    "disclaimer": "..."
  }

依赖：python-docx（pip install python-docx）
"""
import argparse
import datetime
import json
import os
import re

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DARK = RGBColor(0x1F, 0x2A, 0x37)
ACCENT = RGBColor(0xB0, 0x3A, 0x2E)
GRAY = RGBColor(0x6B, 0x72, 0x80)
BLUE = RGBColor(0x2C, 0x5F, 0x8A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT_CN = "微软雅黑"
FONT_HEI = "黑体"
FONT_MONO = "Consolas"

USABLE_CM = 16.2   # A4 减去左右各 2.4cm 页边距


# ---------- 底层工具 ----------

def set_font(run, size=10.5, bold=False, color=None, cn=FONT_CN, ascii_font="Segoe UI"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = ascii_font
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    rf.set(qn("w:eastAsia"), cn)
    rf.set(qn("w:ascii"), ascii_font)
    rf.set(qn("w:hAnsi"), ascii_font)


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def fix_layout(table, widths_cm):
    """固定列宽：必须同时关掉 autofit、写 tblLayout=fixed、重写 tblGrid，否则 Word 打开时重算列宽。"""
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    tblW.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths_cm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w * 567)))
        grid.append(gc)
    tblPr.addnext(grid)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])


def cell_margins(table, top=40, bottom=40, left=80, right=80):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement("w:" + k)
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.append(mar)


def para_border_bottom(p, color="D5D9E0", size=6):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    pPr.append(pbdr)


def para_shade(p, hexcolor):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    pPr.append(shd)


def add_page_field(paragraph, instr="PAGE"):
    r = paragraph.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r._r.addnext(fld)


def keep_next(p):
    p.paragraph_format.keep_with_next = True


# ---------- 段落构件 ----------

def h1(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.25
    keep_next(p)
    set_font(p.add_run(text), 17, bold=True, color=DARK, cn=FONT_HEI)
    para_border_bottom(p, "B03A2E", 10)
    return p


def h2(doc, text):
    p = doc.add_paragraph(style="Heading 2")
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    keep_next(p)
    set_font(p.add_run(text), 12, bold=True, color=BLUE, cn=FONT_HEI)
    return p


def body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    add_rich(p, text, 10.5, DARK)
    return p


def kv(doc, k, v):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.4)
    set_font(p.add_run(k + "　"), 9.5, bold=True, color=GRAY)
    add_rich(p, v, 9.5, DARK)
    return p


def code_block(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = Cm(0.4)
    para_shade(p, "F4F6F9")
    set_font(p.add_run(text), 8.5, color=RGBColor(0x24, 0x2C, 0x38), ascii_font=FONT_MONO)
    return p


def add_rich(p, text, size, color):
    """渲染 **粗体** / `代码` / [文字](链接)。"""
    for tk in re.split(r"(\*\*.+?\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))", text):
        if not tk:
            continue
        if tk.startswith("**") and tk.endswith("**"):
            set_font(p.add_run(tk[2:-2]), size, bold=True, color=color)
        elif tk.startswith("`") and tk.endswith("`"):
            set_font(p.add_run(tk[1:-1]), size - 0.5, color=ACCENT, ascii_font=FONT_MONO)
        elif tk.startswith("[") and "](" in tk:
            m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", tk)
            set_font(p.add_run(m.group(1)), size, color=BLUE)
        else:
            set_font(p.add_run(tk), size, color=color)


# ---------- markdown 轻量渲染 ----------

def md_lines(doc, text):
    lines = text.split("\n")
    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue

        if s.startswith("|") and s.count("|") >= 2:
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1
            render_md_table(doc, block)
            continue

        if s.startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            code_block(doc, "\n".join(buf))
            continue

        m = re.match(r"^(#{1,4})\s+(.*)$", s)
        if m:
            lvl = len(m.group(1))
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
            keep_next(p)
            set_font(p.add_run(re.sub(r"\*\*(.+?)\*\*", r"\1", m.group(2))),
                     {1: 13, 2: 11.5, 3: 10.5, 4: 10.5}[lvl], bold=True,
                     color=(DARK if lvl <= 2 else BLUE), cn=FONT_HEI)
            i += 1
            continue

        if s.startswith(">"):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6)
            p.paragraph_format.space_after = Pt(5)
            para_shade(p, "F7F8FA")
            add_rich(p, s.lstrip("> ").strip(), 9.5, GRAY)
            i += 1
            continue

        m = re.match(r"^([-*+]|\d+\.)\s+(.*)$", s)
        if m:
            marker = "・" if not m.group(1)[0].isdigit() else m.group(1)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.75)
            p.paragraph_format.space_after = Pt(3)
            set_font(p.add_run(marker + " "), 10.5, bold=True, color=ACCENT)
            add_rich(p, m.group(2), 10.5, DARK)
            i += 1
            continue

        if re.match(r"^-{3,}$", s):
            p = doc.add_paragraph()
            para_border_bottom(p, "E3E6EB", 4)
            p.paragraph_format.space_after = Pt(6)
            i += 1
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        add_rich(p, s, 10.5, DARK)
        i += 1


def render_md_table(doc, block):
    rows = []
    for line in block:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.match(r"^:?-{2,}:?$", c) for c in cells if c):
            continue
        rows.append(cells)
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]

    t = doc.add_table(rows=0, cols=ncol)
    t.style = "Table Grid"
    cell_margins(t, 30, 30, 60, 60)
    # 按各列最大内容长度分配列宽，限制在 1.5–6.5cm
    maxlens = [max(len(r[ci]) for r in rows) for ci in range(ncol)]
    s = sum(maxlens) or 1
    widths = [min(6.5, max(1.5, USABLE_CM * m / s)) for m in maxlens]
    k = (USABLE_CM - 0.4) / sum(widths)
    widths = [w * k for w in widths]
    for ri, rowdata in enumerate(rows):
        cells = t.add_row().cells
        for ci, v in enumerate(rowdata):
            p = cells[ci].paragraphs[0]
            p.paragraph_format.line_spacing = 1.1
            p.paragraph_format.space_after = Pt(0)
            set_font(p.add_run(re.sub(r"\*\*(.+?)\*\*", r"\1", v)), 8, bold=(ri == 0),
                     color=(WHITE if ri == 0 else DARK))
            if ri == 0:
                shade(cells[ci], "2C5F8A")
    fix_layout(t, widths)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(2)
    sp.paragraph_format.line_spacing = 0.6


# ---------- 数据表渲染 ----------

def grid_table(doc, cols, rows, widths, head_fill="1F2A37", font_size=8.5,
               first_col_accent=False, zebra=False, bold_first=False):
    t = doc.add_table(rows=1, cols=len(cols))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_margins(t, 45, 45, 80, 80)
    for i, c in enumerate(cols):
        cell = t.rows[0].cells[i]
        shade(cell, head_fill)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        set_font(p.add_run(c), font_size, bold=True, color=WHITE)
    for ri, rowdata in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(rowdata):
            if i >= len(cols):
                break
            p = cells[i].paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            col = ACCENT if (first_col_accent and i == 0) else DARK
            set_font(p.add_run(str(v)), font_size, bold=(bold_first and i == 0) or (first_col_accent and i == 0),
                     color=col)
            if zebra and ri % 2 == 1:
                shade(cells[i], "FAFBFC")
    fix_layout(t, widths)
    return t


# ---------- 主构建 ----------

def build(data, out, narrative=None):
    narrative = narrative or {}
    meta = data["meta"]
    turns = data["turns"]
    stats = data.get("stats") or {}

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.4)
    sec.right_margin = Cm(2.4)

    normal = doc.styles["Normal"]
    normal.font.size = Pt(10.5)
    normal.font.name = "Segoe UI"
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
    normal.paragraph_format.line_spacing = 1.45
    normal.paragraph_format.space_after = Pt(4)

    for name, size in (("Heading 1", 17), ("Heading 2", 12), ("Heading 3", 10.5)):
        try:
            st = doc.styles[name]
            st.font.name = "Segoe UI"
            st.font.size = Pt(size)
            st.font.color.rgb = DARK
            st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_HEI)
        except KeyError:
            pass

    # 页眉页脚
    hdr = sec.header.paragraphs[0]
    hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_font(hdr.add_run(narrative.get("title", "完整对话记录") + "｜对话记录"), 8.5, color=GRAY)
    para_border_bottom(hdr, "E3E6EB", 4)

    ftr = sec.footer.paragraphs[0]
    ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(ftr.add_run("— "), 9, color=GRAY)
    add_page_field(ftr)
    set_font(ftr.add_run(" —"), 9, color=GRAY)

    # ===== 封面 =====
    for _ in range(3):
        doc.add_paragraph()

    set_font(doc.add_paragraph().add_run(narrative.get("kicker", "WORKBUDDY ｜ 会话记录导出")),
             9.5, bold=True, color=ACCENT)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    set_font(p.add_run(narrative.get("title", "完整对话记录")), 40, bold=True, color=DARK, cn=FONT_HEI)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(18)
    sub = narrative.get("subtitle") or ("会话 %s" % meta["session_id"][:8])
    set_font(p.add_run(sub), 14, color=GRAY)

    p = doc.add_paragraph()
    para_border_bottom(p, "B03A2E", 12)
    p.paragraph_format.space_after = Pt(24)

    cover_rows = narrative.get("cover_rows") or [
        ["对话轮次", "%d 轮" % meta["turns"]],
        ["工具调用", "%d 次（失败/中断 %d 次）" % (meta["tools"], meta["failed"])],
        ["时间跨度", "%s — %s" % (meta["start"][:16], meta["end"][:16])],
        ["助手正文", "%d 字" % meta["assistant_chars"]],
        ["原始记录", "%d 条，%s" % (meta["records"], os.path.basename(meta["source"]))],
        ["导出时间", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")],
    ]
    tb = doc.add_table(rows=0, cols=2)
    cell_margins(tb, 50, 50, 0, 100)
    for k, v in cover_rows:
        row = tb.add_row()
        p0 = row.cells[0].paragraphs[0]
        set_font(p0.add_run(k), 9.5, color=GRAY)
        p1 = row.cells[1].paragraphs[0]
        set_font(p1.add_run(v), 9.5, bold=True, color=DARK)
    fix_layout(tb, [3.2, 12.0])

    doc.add_page_break()

    # ===== 文档说明 =====
    h1(doc, "文档说明")
    lead = narrative.get("lead") or [
        "本文档由会话记录直接导出，对话内容未做删改，目的是留下一条可复核的完整链路："
        "从最初的一句需求，到最终产出的结果。",
        "结构上分为任务概览、完整对话记录、工具执行明细三部分。对话记录保留用户原话与助手回复全文；"
        "工具执行明细记录每一次调用的动作与结果摘要，参数与返回值做了长度截断，原始值仍可在会话文件中检索。",
        "需要说明的是：会话文件里存在历史记录的重放副本，同一句提问会在后续轮次被整段重发。"
        "本文档已按内容去重，只保留首次出现，因此轮次数少于原始记录条数，属正常现象。",
    ]
    for t in lead:
        body(doc, t)

    doc.add_paragraph()

    # ===== 一、任务概览 =====
    h1(doc, "一、任务概览")
    if narrative.get("stages"):
        body(doc, narrative.get("stages_lead",
                                "整个任务按阶段推进，下表归纳每一步的动机与结果。"))
        grid_table(doc, ["阶段", "主题", "时间", "当时要解决什么", "结果"],
                   narrative["stages"], [1.5, 2.8, 3.1, 4.1, 4.7],
                   font_size=8.5, bold_first=True, first_col_accent=True)
        if narrative.get("after_stages"):
            doc.add_paragraph()
            body(doc, narrative["after_stages"])
    else:
        body(doc, "本次共 %d 轮提问、%d 次操作，时间跨度 %s 至 %s。"
                  % (meta["turns"], meta["tools"], meta["start"], meta["end"]))
        rows = [["对话轮次", "%d 轮" % meta["turns"]],
                ["用户提问字数", "%d 字" % meta["user_chars"]],
                ["助手回复字数", "%d 字" % meta["assistant_chars"]],
                ["工具调用次数", "%d 次" % meta["tools"]],
                ["其中失败或中断", "%d 次" % meta["failed"]],
                ["原始记录条数", "%d 条" % meta["records"]]]
        grid_table(doc, ["项目", "数值"], rows, [4.5, 10.0], font_size=9)

        if stats.get("tool_usage"):
            doc.add_paragraph()
            h2(doc, "操作类型分布")
            total = sum(v for _, v in stats["tool_usage"]) or 1
            rows = [[k, "%d 次" % v, "%.0f%%" % (v * 100.0 / total)] for k, v in stats["tool_usage"]]
            grid_table(doc, ["动作", "次数", "占比"], rows, [6.0, 4.0, 4.0], font_size=8.5)

    doc.add_page_break()

    # ===== 二、完整对话记录 =====
    h1(doc, "二、完整对话记录")
    body(doc, "以下为逐轮原文。用户消息为原话（含错别字，未作修正）；助手消息保留当时的答复全文。"
              "回复中的表格、代码块与列表已转为 Word 格式。")

    for t in turns:
        p = doc.add_paragraph(style="Heading 3")
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(2)
        keep_next(p)
        set_font(p.add_run("第 %d 轮" % t["no"]), 9, bold=True, color=ACCENT)
        set_font(p.add_run("　" + (t["time"] or "")), 8.5, color=GRAY)

        if t.get("q"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(6)
            para_shade(p, "F4F6F9")
            set_font(p.add_run("问　"), 9, bold=True, color=BLUE)
            set_font(p.add_run(t["q"]["text"].strip()), 10.5, bold=True, color=DARK)

        for a in t["a"]:
            md_lines(doc, a["text"])

        if t["tools"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            set_font(p.add_run("（本轮执行 %d 次操作%s，明细见第三节 E%02d）"
                               % (len(t["tools"]),
                                  "，其中 %d 次失败" % sum(1 for x in t["tools"] if x.get("status") == "失败")
                                  if any(x.get("status") == "失败" for x in t["tools"]) else "",
                                  t["no"])), 8.5, color=GRAY)

    doc.add_page_break()

    # ===== 三、工具执行明细 =====
    h1(doc, "三、工具执行明细")
    body(doc, "下表按时间顺序记录每一次操作。参数与结果均做了截断，仅保留可辨识的要点；"
              "失败或被拒绝的操作同样保留，不做美化。")

    rows = []
    for t in turns:
        first = True
        for tl in t["tools"]:
            rows.append(["E%02d" % t["no"] if first else "", tl["time"], tl["label"],
                         tl["params"], tl["result"]])
            first = False
    if rows:
        grid_table(doc, ["轮次", "时间", "动作", "关键参数", "结果摘要"], rows,
                   [1.4, 1.9, 2.2, 5.1, 5.6], font_size=7.5, zebra=True, first_col_accent=True)
    else:
        body(doc, "本次会话没有工具调用记录。")

    failed = stats.get("failed_ops") or []
    if failed:
        doc.add_paragraph()
        h2(doc, "失败与中断的操作")
        body(doc, "以下操作未成功完成，保留原因以便回溯。")
        frows = [["E%02d" % f["turn"], f["time"], f["label"], f["status"], f["params"], f["result"]]
                 for f in failed]
        grid_table(doc, ["轮次", "时间", "动作", "状态", "参数", "结果"], frows,
                   [1.3, 1.7, 1.8, 1.2, 4.2, 6.0], font_size=7.5)

    # ===== 自定义章节 =====
    for s in narrative.get("sections") or []:
        doc.add_page_break()
        h1(doc, s.get("h", ""))
        if s.get("lead"):
            body(doc, s["lead"])
        for b in s.get("blocks") or []:
            typ = b.get("type", "body")
            if typ == "h2":
                h2(doc, b.get("text", ""))
            elif typ == "kv":
                kv(doc, b.get("k", ""), b.get("v", ""))
            elif typ == "code":
                code_block(doc, b.get("text", ""))
            elif typ == "table":
                cols = b.get("cols") or []
                widths = b.get("widths")
                if not widths:
                    n = max(1, len(cols))
                    widths = [USABLE_CM / n] * n
                grid_table(doc, cols, b.get("rows") or [], widths, font_size=8.5)
                doc.add_paragraph()
            else:
                body(doc, b.get("text", ""))

    # ===== 附录 =====
    appendix = narrative.get("appendix") or []
    if appendix:
        doc.add_page_break()
        h1(doc, "附录")
        for i, item in enumerate(appendix, 1):
            title, content = (item[0], item[1]) if isinstance(item, (list, tuple)) else (item.get("h"), item.get("text"))
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            set_font(p.add_run("%d. %s" % (i, title)), 10.5, bold=True, color=ACCENT)
            body(doc, content)

    # ===== 免责 =====
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    para_border_bottom(p, "E3E6EB", 4)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    set_font(p.add_run("文档说明与责任"), 10, bold=True, color=GRAY)
    body(doc, narrative.get(
        "disclaimer",
        "本文档为会话过程记录，内容忠实于当时的对话与工具输出，不代表对其中任何结论的再次确认或背书。"
        "文中涉及的数据、判断与建议仅供回溯参考；若涉及投资、法律、医疗等专业领域，请以持牌专业机构的意见为准。"))

    doc.save(out)
    return out


def main():
    ap = argparse.ArgumentParser(description="把 transcript.json 渲染成 Word 文档")
    ap.add_argument("transcript", help="export_session.py 产出的 JSON")
    ap.add_argument("-o", "--out", required=True, help="输出 docx 路径")
    ap.add_argument("--title", help="封面标题")
    ap.add_argument("--subtitle", help="封面副标题")
    ap.add_argument("--narrative", help="narrative.json 路径（可选）")
    args = ap.parse_args()

    with open(args.transcript, encoding="utf-8") as f:
        data = json.load(f)
    narrative = {}
    if args.narrative:
        with open(args.narrative, encoding="utf-8") as f:
            narrative = json.load(f)
    if args.title:
        narrative["title"] = args.title
    if args.subtitle:
        narrative["subtitle"] = args.subtitle

    out = build(data, args.out, narrative)
    print("已生成：%s（%d KB，%d 轮，%d 次操作）"
          % (out, os.path.getsize(out) // 1024, data["meta"]["turns"], data["meta"]["tools"]))


if __name__ == "__main__":
    main()
