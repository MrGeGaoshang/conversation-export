# 会话 jsonl 结构与解析规则

文件位置：`~/.workbuddy/projects/<项目slug>/<sessionId>.jsonl`。项目 slug 是把工作目录绝对路径里的非字母数字字符替换成 `-`（如 `C:/Users/Administrator/WorkBuddy/2026-10-30-22-27-21` → `c-Users-Administrator-WorkBuddy-2026-10-30-22-27-21`）。

## 记录类型

| type | 关键字段 | 说明 |
|---|---|---|
| `session-meta` | id, meta | 会话元信息，8 条左右 |
| `message` | role, content | role=user/assistant；content 是块数组 |
| `reasoning` | content | 内部思考，导出对话时**跳过** |
| `function_call` | name, arguments, callId | 工具调用；arguments 是 JSON 字符串 |
| `function_call_result` | callId, status, output | 与 callId 配对；output 可能是 JSON 字符串 |
| `file-history-snapshot` | snapshot | 文件快照，跳过 |
| `ai-title` | aiTitle | 会话标题 |

`message.content` 的块：用户是 `{"type":"input_text","text":...}`，助手是 `{"type":"output_text","text":...}`。真正的用户提问藏在文本的 `<user_query>...</user_query>` 标签里，其余是 system-reminder（跳过）。

## 历史重放（去重的根本原因）

用户消息会被**整段重发**：会话进行到后面的轮次时，早先的 `<user_query>` 会在新的 message 记录里再次出现，且时间戳成批相同（同一秒内十几条）。实测一次会话 53 条提问记录只有 25 条是真实的。

去重规则：`key = re.sub(r"\s+"," ",text).strip()`，按 key 去重只保留**首次出现**。助手消息与 function_call（按 callId）实测无重复，但去重逻辑对它们也无害。

## 状态判定

`function_call_result.status` 实测只有 `completed` / `incomplete` 两种，**不能反映真实成败**。真实失败藏在 output 文本里，需叠加关键词判定：

```
Command rejected / Error: / Traceback (most recent call last) /
ModuleNotFoundError / No such file or directory / Exit Code: [1-9] /
timed out / SIGTERM / Permission denied / command not found
```

注意 `Exit Code: 0` 是成功，正则必须限定非零。连接器被拒绝执行（如 `Command rejected: ...`）也在此列。

## 截断基准（实测）

一次约 4MB 的会话：助手正文约 3.5 万字符，工具参数约 19 万，工具结果约 57 万。全量入文档会有上千页，因此：

- 参数摘要：取前 2 个有意义的键（message/code/skill/file_path/command/entry/toolName 等），各截 60 字符
- 结果摘要：压平空白后截 150 字符
- 文档中必须声明截断，并给出原始 jsonl 路径

## 其他坑

- `assistant` 消息里可能出现 `[omitted]` 或 `<omitted />` 占位（压缩摘要），导出时跳过。
- 会话仍在进行时导出，最后一轮可能只有提问没有回复，属正常。
- `--cutoff` 截断用北京时间毫秒时间戳；脚本已内置 `to_ms`。
- jsonl 用 UTF-8 无 BOM；Windows 下读写都要显式 `encoding="utf-8"`。
