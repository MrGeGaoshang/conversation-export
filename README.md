# conversation-export

把 WorkBuddy 会话记录一键导出为**排版完整的 Word 文档**的 skill。

适用于：导出聊天记录归档、做任务复盘文档、把对话过程交给别人复核。

---

## 安装

```bash
git clone https://gitee.com/Mr-GeGaoshang/conversation-export.git ~/.workbuddy/skills/conversation-export
```

Windows（PowerShell）：

```powershell
git clone https://gitee.com/Mr-GeGaoshang/conversation-export.git "$env:USERPROFILE\.workbuddy\skills\conversation-export"
```

GitHub 用户把地址换成 `https://github.com/MrGeGaoshang/conversation-export.git` 即可，内容完全一致。

装好后**新开一个对话**，说「把这次对话导出成 Word」就会自动触发（skill 在会话启动时加载，当前已开的对话不会生效）。

## 前置依赖

```bash
pip install python-docx
```

仅此一项必需。`pypdfium2` 为可选，用于生成后把 PDF 渲染成图做版面目检。

## 它解决什么问题

直接用脚本读会话文件再拼 Word，会踩三个坑。这个 skill 把三个坑都固化了：

**① 历史重放导致轮次虚高。** 会话 jsonl 里同一批提问会被整段重发多次（时间戳成批相同），表面上 50 多条提问，实际可能只有 25 轮。脚本按规范化文本去重，只保留首次出现。

**② 失败操作看不出来。** 工具调用的 `status` 字段几乎全是 `completed`，真实失败藏在返回文本里。脚本叠加结果信号判定（`Exit Code` 非零、`Traceback`、`Error:`、`Command rejected`、超时等），把失败与中断的操作单独成表——这是复盘时最有价值的部分。

**③ Word 表格列宽会被重算。** python-docx 设了列宽，Word 打开时仍会按内容重算，把窄列挤成单字竖排。脚本用「固定列宽三件套」锁死：`autofit=False` + `<w:tblLayout w:type="fixed"/>` + 重写 `<w:tblGrid>`。漏任何一步都会出问题，详见 `references/docx-layout.md`。

## 用法

```bash
# 1. 定位会话并抽取（--latest 自动匹配当前项目的最新会话）
python scripts/export_session.py --list
python scripts/export_session.py --latest -o transcript.json --md transcript.md

# 2. 生成 Word
python scripts/build_docx.py transcript.json -o 对话记录.docx --title "任务复盘"

# 3. 可选：带自定义概览、交付物清单、附录
python scripts/build_docx.py transcript.json -o 对话记录.docx \
    --title "任务复盘" --narrative narrative.json
```

不写 `narrative.json` 也能出文档，脚本会用统计表自动补一版中性概览。

常用参数：

| 参数 | 作用 |
|---|---|
| `--latest` | 按当前工作目录匹配项目目录，取最新会话 |
| `--cutoff "2026-10-31 12:20"` | 截断到某时刻，排除导出动作本身 |
| `--max-turns N` | 只取前 N 轮 |
| `--md out.md` | 同时输出 Markdown，便于导入飞书/腾讯文档等平台 |

## 生成的文档结构

封面 → 文档说明 → 一、任务概览 → 二、完整对话记录（原文，含表格与代码块还原）→ 三、工具执行明细（逐次列出，含失败操作）→ 四、交付物清单 → 五、关键数据留档 → 附录 → 免责声明。

章节按需生成，没有内容的章节不会硬凑。

## 目录结构

```
conversation-export/
├── SKILL.md                     5 步流程 + 5 条硬性要求
├── scripts/
│   ├── export_session.py        jsonl → transcript.json / markdown
│   └── build_docx.py            transcript.json → 排版 Word
└── references/
    ├── jsonl-schema.md          会话文件结构、重放规律、状态判定规则
    └── docx-layout.md           python-docx 排版雷区清单
```

## 边界

- **只做记录导出**，不改写对话内容，用户原文中的错别字保留。
- 工具调用的参数与返回值做**长度截断**（全量会是数百万字符、上千页），文档中会明确标注，原始数据仍在会话文件里。
- 不还原可视化图表等非文本内容，只保留其文字记录。
- 概览、交付物等叙事章节只能来自会话中的真实事实，不编造未发生的过程。

## 不包含：云文档上传

**这个 skill 只在本地生成 `.docx`，不会自动上传到飞书、腾讯文档等云平台。**

原因是有意为之，不是遗漏：

- **上传是投递动作，不是文档生成的一部分。** 这次传飞书、下次传腾讯文档、后天直接发微信，目标每次不同，没必要固化进 skill。
- **硬编码会造成隐性依赖。** 如果把飞书调用写死在脚本里，没装飞书连接器的人执行时就会报错，而这个 skill 的核心功能其实只需要 python-docx。

那怎么传到云文档？两条路：

1. **让当前对话代传。** 只要那个对话环境里装了对应的连接器（如飞书 `feishu`），直接说「把这份 docx 传到飞书云文档」即可，AI 会用连接器完成上传。这依赖对话环境，不依赖本 skill。
2. **手动导入（最稳）。** 飞书：云文档 → 新建 → 导入文件 → 选 docx，平台会自动转成在线文档并保留标题层级与表格。

## 许可

MIT，见 `LICENSE`。可自由使用、修改与分发。

---

作者：Mr. Ge Gaoshang
