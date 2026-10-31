# python-docx 排版要点与已验证的雷区

以下每一条都在实际生成 60–76 页对话记录文档时验证过。

## 雷区 1：列宽被 Word 重算（最常见）

只写 `cell.width = Cm(x)` 不够——Word 打开时会按 autofit 重算，窄列（如「轮次」「动作」）会被挤成**单字竖排**。必须三件套同时做（见 `build_docx.py` 的 `fix_layout()`）：

```python
table.autofit = False
# 1) tblPr 里加 <w:tblLayout w:type="fixed"/>
# 2) tblW 设固定总宽（dxa = cm * 567）
# 3) 删除旧 tblGrid，按列宽重写 <w:gridCol/>
# 4) 最后逐行 cell.width 兜底
```

## 雷区 2：表格后紧贴下一段

表格和后续段落之间要插一个小空段（`line_spacing=0.6`），否则连续两张表会粘连。

## 雷区 3：中文字体

`run.font.name` 只作用于西文；中文必须设 `w:eastAsia`：

```python
rf = run._element.get_or_add_rPr()... # rFonts set(qn("w:eastAsia"), "微软雅黑")
```

标题想要黑体观感就单独传 `cn=FONT_HEI`。Normal 样式也要同样处理，否则正文回退宋体。

## Heading 样式的取舍

用内置 `Heading 1/2/3` 样式（而非手写加粗段落）的好处：Word 导航窗格可用、可自动生成目录。但要在 styles 里重设字体颜色，否则默认蓝色 Calibri 很难看。轮次标题用 Heading 3，一页一个层级不喧宾夺主。

## 页码与页眉

页脚页码用 `fldSimple` 域：

```python
fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE")
```

页眉下边框用段落边框（`w:pBdr/w:bottom`），比插入横线图片干净。

## 表格宽度预算

A4（21cm）左右页边距各 2.4cm，可用 16.2cm。多列表格的 widths 之和不要超过 16.2，否则溢出页边。对话里的 markdown 表格按**各列最大内容长度**比例分配、夹在 1.5–6.5cm 之间，实测效果稳定。

## 对话 markdown 的轻量还原规则

助手回复是 markdown，逐行状态机转换（见 `md_lines()`）：

| 语法 | 处理 |
|---|---|
| `| a | b |` 连续行 | 跳过分隔行，转 Word 表格，蓝底白字表头 |
| 代码块 | 灰底 Consolas 段落 |
| `#`–`####` | 按层级映射 13/11.5/10.5pt 黑体 |
| `>` 引用 | 缩进 + 浅灰底 |
| `- ` / `1. ` | 「・」或原序号 + 悬挂缩进 |
| `**x**`、`` `x` ``、`[x](y)` | 行内富文本（粗体/赭红等宽/链接蓝） |

`**粗体**` 要先于普通文本切分，用 `re.split` 带捕获组。

## 核验流程（不可省）

1. 结构自检：`python -c` 重新打开 docx，数段落/表格、查 `**`、行首 `|`、`###` 残留。
2. 视觉核验：Word COM 导出 PDF（`ExportAsFixedFormat(path, 17)`），pypdfium2 渲染成 PNG 目检封面、概览表、明细表、末页。
3. 已知渲染环境差异：PowerShell 控制台打印中文会乱码，但文件内容正确——用 UTF-8 通道（Python/文本工具）核验，别信终端显示。

## 体积参考

29 轮 / 281 次操作 / 3.5 万字回复 → 约 76 页 / 114KB。超过 100 页时考虑建议用户用 `--cutoff` 或 `--max-turns` 拆分。
