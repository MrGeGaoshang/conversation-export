#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_session.py — 把 WorkBuddy 会话记录（jsonl）解析成结构化 transcript.json / markdown。

用法：
  python export_session.py --latest                     # 自动定位当前项目最近的会话
  python export_session.py <session.jsonl>              # 指定文件
  python export_session.py --latest -o t.json --md t.md # 同时产出 markdown
  python export_session.py --latest --cutoff "2026-10-31 12:20"
  python export_session.py --list                        # 列出所有会话文件

输出 JSON 结构：
  meta  : 来源路径、会话 ID、起止时间、轮次/工具/正文字数、失败操作数
  turns : [{no, time, q:{time,text}, a:[{time,text}], tools:[{time,name,label,params,status,result}]}]
  stats : 工具使用分布、失败操作清单、用户提问清单
"""
import argparse
import datetime
import json
import os
import re
import sys
from collections import Counter

TZ = datetime.timezone(datetime.timedelta(hours=8))

TOOL_LABEL = {
    "Skill": "加载技能", "Bash": "执行命令", "PowerShell": "执行命令", "Read": "读取文件",
    "Write": "写入文件", "Edit": "修改文件", "Glob": "查找文件", "Grep": "检索内容",
    "ToolSearch": "加载工具", "DeferExecuteTool": "调用连接器", "WebSearch": "联网搜索",
    "WebFetch": "抓取网页", "present_files": "展示产物", "show_widget": "生成图表",
    "widget_guidelines": "读取规范", "AskUserQuestion": "征询选择", "Agent": "派发子任务",
    "TaskCreate": "建任务", "TaskUpdate": "更新任务", "TaskList": "查任务",
}

PARAM_KEYS = ["message", "code", "skill", "file_path", "command", "entry", "toolName",
              "query", "files", "rang", "content", "title", "description", "instructions",
              "new_string", "old_string", "connectorName", "path"]

FAIL_STATUS = {"failed", "error", "rejected", "denied", "cancelled", "aborted"}

# 结果文本层面的失败信号：工具可能返回 completed 但 shell/脚本其实失败了
FAIL_PATTERNS = [
    r"Command rejected", r"^Error", r"Error:", r"Traceback \(most recent call last\)",
    r"ModuleNotFoundError", r"No such file or directory", r"Exit Code: [1-9]",
    r"timed out", r"SIGTERM", r"Permission denied", r"command not found",
]


def judge(status, result_text):
    """把 status 字段与结果文本合并成一个可读状态。"""
    st = (status or "").lower()
    if st in FAIL_STATUS:
        return "失败"
    for pat in FAIL_PATTERNS:
        if re.search(pat, result_text or "", re.M | re.I):
            return "失败"
    if st == "incomplete":
        return "中断"
    return "成功"


# ---------- 定位会话文件 ----------

def projects_root():
    return os.path.join(os.path.expanduser("~"), ".workbuddy", "projects")


def list_sessions(root=None):
    root = root or projects_root()
    out = []
    for dirpath, _, files in os.walk(root):
        for fn in files:
            if fn.endswith(".jsonl"):
                p = os.path.join(dirpath, fn)
                out.append((os.path.getmtime(p), p))
    out.sort(reverse=True)
    return out


def slug_variants(cwd):
    cwd = os.path.abspath(cwd).replace("\\", "/").strip("/")
    base = re.sub(r"[^A-Za-z0-9]+", "-", cwd).strip("-")
    return {base.lower(), base.lower().replace("--", "-")}


def find_latest(cwd=None):
    """优先返回与当前工作目录同名的项目目录下最新的会话，找不到则退化为全局最新。"""
    cwd = cwd or os.getcwd()
    root = projects_root()
    want = slug_variants(cwd)
    same = []
    for mtime, p in list_sessions(root):
        proj = os.path.basename(os.path.dirname(p))
        if proj.lower() in want or proj.lower().endswith("-" + os.path.basename(os.path.abspath(cwd)).lower()):
            same.append((mtime, p))
    if same:
        return same[0][1]
    alls = list_sessions(root)
    return alls[0][1] if alls else None


# ---------- 解析 ----------

def bj(ms, fmt="%Y-%m-%d %H:%M:%S"):
    if not ms:
        return ""
    return datetime.datetime.fromtimestamp(ms / 1000, TZ).strftime(fmt)


def to_ms(text):
    """'2026-10-31 12:20' / '2026-10-31' → 毫秒时间戳（北京时间）。"""
    text = text.strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            dt = datetime.datetime.strptime(text, fmt).replace(tzinfo=TZ)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    raise SystemExit("无法识别的时间：%s（用 'YYYY-MM-DD HH:MM' 格式）" % text)


def blocks_text(c):
    if isinstance(c, str):
        return c
    return "\n".join(b.get("text") or "" for b in (c or [])
                     if isinstance(b, dict) and b.get("type") in ("input_text", "text", "output_text"))


def norm(s):
    return re.sub(r"\s+", " ", s or "").strip()


def short(s, n):
    s = norm(s)
    return s if len(s) <= n else s[: n - 1] + "…"


def param_brief(args, n=130):
    if isinstance(args, str):
        try:
            args = json.loads(args)
        except Exception:
            return short(args, n)
    if not isinstance(args, dict):
        return short(str(args), n)
    parts = []
    for k in PARAM_KEYS:
        v = args.get(k)
        if v in (None, "", [], {}):
            continue
        parts.append("%s=%s" % (k, short(str(v), 60)))
        if len(parts) >= 2:
            break
    if not parts:
        for k, v in list(args.items())[:2]:
            parts.append("%s=%s" % (k, short(str(v), 50)))
    return short("; ".join(parts), n)


def result_brief(res, n=150):
    if res is None:
        return ""
    if isinstance(res, dict):
        txt = res.get("text") or res.get("output") or res.get("message") or ""
        if not txt:
            txt = json.dumps(res, ensure_ascii=False)
    else:
        txt = str(res)
    return short(txt, n)


def parse(path, cutoff_ms=None, max_turns=None):
    rows = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue

    if cutoff_ms:
        rows = [d for d in rows if (d.get("timestamp") or 0) <= cutoff_ms]

    events, seen_q, pending = [], set(), set()

    for d in rows:
        t = d.get("type")
        if t == "message":
            txt = blocks_text(d.get("content"))
            role = d.get("role")
            if role == "user":
                for q in re.findall(r"<user_query>(.*?)</user_query>", txt, re.S):
                    key = norm(q)
                    # 同一提问会被历史重放多次，只保留首次
                    if not key or key in seen_q:
                        continue
                    seen_q.add(key)
                    events.append({"k": "user", "ts": d.get("timestamp") or 0,
                                   "time": bj(d.get("timestamp")), "text": q.strip()})
            elif role == "assistant":
                body = txt.strip()
                if body and body not in ("[omitted]",) and "<omitted />" not in body[:40]:
                    events.append({"k": "assistant", "ts": d.get("timestamp") or 0,
                                   "time": bj(d.get("timestamp")), "text": body})
        elif t == "function_call":
            cid = d.get("callId")
            if cid and cid in seen_q:
                continue
            if cid:
                seen_q.add(cid)
            ev = {"k": "tool", "ts": d.get("timestamp") or 0, "time": bj(d.get("timestamp"), "%m-%d %H:%M"),
                  "name": d.get("name"), "label": TOOL_LABEL.get(d.get("name"), d.get("name")),
                  "params": param_brief(d.get("arguments")), "status": "", "result": ""}
            events.append(ev)
            if cid:
                pending.add(cid)
                ev["_cid"] = cid
        elif t == "function_call_result":
            cid = d.get("callId")
            for ev in events:
                if ev.get("_cid") == cid and not ev.get("status"):
                    body = result_brief(d.get("output"), 400)
                    ev["result"] = result_brief(d.get("output"))
                    ev["status"] = judge(d.get("status"), body)
                    break

    events.sort(key=lambda e: e["ts"])

    turns, cur = [], None
    for e in events:
        if e["k"] == "user":
            cur = {"no": 0, "time": e["time"], "q": {"time": e["time"], "text": e["text"]},
                   "a": [], "tools": []}
            turns.append(cur)
        elif cur is None:
            cur = {"no": 0, "time": e["time"], "q": None, "a": [], "tools": []}
            turns.append(cur)
            (cur["a"] if e["k"] == "assistant" else cur["tools"]).append(e)
        else:
            (cur["a"] if e["k"] == "assistant" else cur["tools"]).append(e)

    turns = [t for t in turns if t["q"] or t["a"] or t["tools"]]
    if max_turns:
        turns = [t for t in turns if t["no"] != 0][:max_turns] if turns and turns[0]["q"] else turns[:max_turns]
    for i, t in enumerate(turns, 1):
        t["no"] = i
        for tl in t["tools"]:
            tl.pop("_cid", None)
            if not tl.get("status"):
                tl["status"] = "无返回"

    stamps = [d.get("timestamp") for d in rows if d.get("timestamp")]
    meta = {
        "source": os.path.abspath(path),
        "session_id": os.path.splitext(os.path.basename(path))[0],
        "start": bj(min(stamps)) if stamps else "",
        "end": bj(max(stamps)) if stamps else "",
        "records": len(rows),
        "turns": len(turns),
        "tools": sum(len(t["tools"]) for t in turns),
        "answers": sum(len(t["a"]) for t in turns),
        "assistant_chars": sum(len(a["text"]) for t in turns for a in t["a"]),
        "user_chars": sum(len(t["q"]["text"]) for t in turns if t["q"]),
        "failed": sum(1 for t in turns for tl in t["tools"] if tl.get("status") in ("失败", "中断")),
        "cutoff": bj(cutoff_ms) if cutoff_ms else "",
    }

    usage = Counter(tl["label"] for t in turns for tl in t["tools"])
    stats = {
        "tool_usage": [[k, v] for k, v in usage.most_common()],
        "failed_ops": [{"turn": t["no"], "time": tl["time"], "label": tl["label"],
                        "params": tl["params"], "result": tl["result"], "status": tl.get("status")}
                       for t in turns for tl in t["tools"] if tl.get("status") in ("失败", "中断")],
        "questions": [{"turn": t["no"], "time": t["time"], "text": t["q"]["text"]} for t in turns if t["q"]],
    }
    return {"meta": meta, "turns": turns, "stats": stats}


# ---------- markdown 输出 ----------

def to_markdown(data, title="完整对话记录"):
    m = data["meta"]
    out = ["# %s" % title, "",
           "> 来源：`%s`  " % m["source"],
           "> 时间：%s — %s  " % (m["start"], m["end"]),
           "> 轮次：%d　工具调用：%d 次　助手正文：%d 字" % (m["turns"], m["tools"], m["assistant_chars"]),
           ""]
    for t in data["turns"]:
        out.append("## 第 %d 轮　%s" % (t["no"], t["time"]))
        if t["q"]:
            out.append("")
            out.append("**问：** %s" % t["q"]["text"].replace("\n", "  \n"))
        for a in t["a"]:
            out.append("")
            out.append(a["text"])
        if t["tools"]:
            out.append("")
            out.append("<sub>本轮执行 %d 次操作</sub>" % len(t["tools"]))
        out.append("")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description="导出 WorkBuddy 会话记录为结构化 JSON / Markdown")
    ap.add_argument("session", nargs="?", help="会话 jsonl 路径；省略时配合 --latest")
    ap.add_argument("--latest", action="store_true", help="自动定位当前项目最近的会话")
    ap.add_argument("--list", action="store_true", help="列出所有会话文件")
    ap.add_argument("-o", "--out", default="transcript.json", help="输出 JSON 路径")
    ap.add_argument("--md", help="同时输出 markdown 到该路径")
    ap.add_argument("--title", default="完整对话记录", help="markdown 标题")
    ap.add_argument("--cutoff", help="只保留此时间之前的记录，如 '2026-10-31 12:20'")
    ap.add_argument("--max-turns", type=int, help="最多保留前 N 轮")
    args = ap.parse_args()

    if args.list:
        for mtime, p in list_sessions():
            print("%s  %6.1f KB  %s" % (
                datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M"),
                os.path.getsize(p) / 1024, p))
        return

    path = args.session
    if not path and args.latest:
        path = find_latest()
        if not path:
            raise SystemExit("没有找到任何会话文件，请手动指定 jsonl 路径")
        print("自动定位到会话：%s" % path)
    if not path:
        raise SystemExit("请提供会话文件路径，或加 --latest")

    data = parse(path, to_ms(args.cutoff) if args.cutoff else None, args.max_turns)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    if args.md:
        with open(args.md, "w", encoding="utf-8") as f:
            f.write(to_markdown(data, args.title))

    m = data["meta"]
    print("轮次 %d ｜ 工具 %d 次（失败 %d）｜ 助手正文 %d 字 ｜ 用户提问 %d 字"
          % (m["turns"], m["tools"], m["failed"], m["assistant_chars"], m["user_chars"]))
    print("时间范围 %s — %s" % (m["start"], m["end"]))
    print("已写出：%s%s" % (args.out, ("　和　" + args.md) if args.md else ""))


if __name__ == "__main__":
    main()
