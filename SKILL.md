---
name: conversation-export
description: 把 WorkBuddy 会话记录完整导出为排版 Word 文档。当用户要求导出对话、生成对话记录、把聊天过程归档成 Word/docx、做任务复盘文档时使用。自动定位会话 jsonl、按内容去重历史重放、还原对话中的表格与代码块、逐次列出工具调用明细（含失败操作），支持自定义概览与交付物章节。也可产出 Markdown 版本供其他平台导入。
agent_created: true
---

# 会话记录导出（Word）

## 用途

把「会话 jsonl → 结构化记录 → 排版完整的 Word 文档」固化为一条命令链。核心价值不是拿到原始文件，而是三件事：**去重**（jsonl 里的历史重放副本会让提问凭空翻倍）、**可读**（对话里的 markdown 表格/代码块要转成 Word 原生格式）、**可回溯**（每一次工具调用含失败操作都保留摘要）。

## 能力边界

- 只做**记录导出**，不改写对话内容；用户消息中的错别字保留原样。
- 工具调用的参数与返回值做**长度截断**（全量会有数百万字符），文档中必须写明这一点。
- narrative 中的阶段划分、交付物清单等内容**只能来自 transcript 里的真实事实**，禁止编造未发生的过程。
- 不承诺还原原始排版细节（如聊天里的可视化图表只留文字记录）。

## 执行流程

### 步骤 1 · 定位会话并抽取

```bash
python scripts/export_session.py --list                 # 查看所有会话文件
python scripts/export_session.py --latest -o t.json --md t.md
```

`--latest` 按当前工作目录匹配项目目录取最新会话；跨项目导出时显式传 jsonl 路径。
常用参数：`--cutoff "2026-10-31 12:20"` 截断到某时刻（用于排除导出动作本身）；`--max-turns N` 只取前 N 轮。

先看 stdout 的统计行（轮次/工具/失败数/字数/时间范围），确认体量合理再继续。

### 步骤 2 · 写 narrative（可选但推荐）

narrative.json 决定文档的「概览、交付物、附录」等叙事章节，字段见 `scripts/build_docx.py` 头部注释。写之前**必须通读 transcript 的提问清单与关键回复**，确保每一条都有出处：

- `stages`：按对话实际推进划分阶段，时间用对话里的真实时间。
- `sections`：交付物清单、关键数据等，只列对话中确实产生的东西。
- `appendix`：过程中真实的判断修正、踩坑复盘。

用户没有要求叙事章节时可以跳过，脚本会用统计表自动补一版中性概览。

### 步骤 3 · 生成 docx

```bash
python scripts/build_docx.py t.json -o 输出.docx --title "标题" --narrative narrative.json
```

无 narrative 时省略 `--narrative`；标题可用 `--title` 直接给。

### 步骤 4 · 版面核验（不可跳过）

python-docx 生成的文件必须实际打开验证过才算完成：

1. 结构自检：段落/表格数量、无 markdown 残留符号（`**`、行首 `|`、`###`）。
2. 渲染核验：有 Word 时用 COM 导出 PDF 再渲染成图目检（重点看表格列宽是否被 Word 重算挤成竖排）；无 Word 时至少做结构自检并向用户说明未做视觉核对。
3. 常见雷区见 `references/docx-layout.md`——尤其是**固定列宽三件套**，漏一步 Word 就会把窄列挤成单字竖排。

### 步骤 5 · 交付

`present_files` 打开 docx，回复中给出：文档结构、轮次/操作统计、去重说明、截断说明。用户提到飞书/腾讯文档等平台时，按可用连接器决定是直接上传还是给手动导入路径。

## 硬性要求

1. **去重必须做**：jsonl 中同一提问会被历史重放多次（时间戳成批相同），必须按规范化文本去重只留首次，否则轮次数虚高一倍。
2. **截断必须声明**：文档说明里写明参数与返回值做了截断，全量在会话文件中。
3. **失败操作必须保留**：状态判定不能只看 status 字段（几乎全是 completed），要叠加结果文本的失败信号（`Exit Code: [1-9]`、`Traceback`、`Error:`、`Command rejected` 等）。
4. **narrative 不编造**：写进概览表的每个阶段、每个交付物都要能在 transcript 里指出来源。
5. 依赖仅 python-docx（`pip install python-docx`）；PDF 渲染核验用 pypdfium2，属可选增强。

## 资源

- `scripts/export_session.py` — jsonl → transcript.json / markdown（去重、截断、状态判定、自动定位）
- `scripts/build_docx.py` — transcript.json → 排版 Word（封面/概览/对话/明细/自定义章节/附录）
- `references/jsonl-schema.md` — 会话文件结构、重放规律、去重与状态判定规则
- `references/docx-layout.md` — python-docx 排版要点与已验证的雷区清单
