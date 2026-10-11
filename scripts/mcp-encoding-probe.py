#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""编码探针：直接以字节流读取 mcp_server.py 的 stdout，判断它到底输出什么编码。

背景：MCP 规范要求 stdio 传输的消息必须是 UTF-8。Windows 上 Python 的
sys.stdout 在重定向/管道场景默认用 locale 编码（中文机器 = cp936/GBK），
于是 server 写 GBK 字节、客户端按 UTF-8 解码 → 中文变乱码。

用法： python scripts/mcp-encoding-probe.py [server.py 路径]
"""
import json
import os
import subprocess
import sys

try:                       # 探针自身也要 UTF-8 输出，否则本报告的汉字会被二次编码
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
SERVER = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "tools", "vba", "mcp_server.py")
SERVER = os.path.abspath(SERVER)

REQUESTS = [
    {"jsonrpc": "2.0", "id": 1, "method": "initialize",
     "params": {"protocolVersion": "2025-06-18", "capabilities": {},
                "clientInfo": {"name": "enc-probe", "version": "1"}}},
    {"jsonrpc": "2.0", "method": "notifications/initialized"},
    {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
]
# 请求本身用 UTF-8 写（含中文的 request 也要能进去）
payload = "".join(json.dumps(r, ensure_ascii=False) + "\n" for r in REQUESTS).encode("utf-8")

p = subprocess.Popen([sys.executable, SERVER], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                     stderr=subprocess.PIPE)
out, err = p.communicate(payload)

print("server      : %s" % SERVER)
print("exit        : %s" % p.returncode)
print("stdout 字节 : %d" % len(out))
print("stderr      : %s" % (err.decode("utf-8", "replace").strip() or "(空)"))
print("stdout 纯 ASCII? : %s" % out.isascii())
print()

for enc in ("utf-8", "cp936", "cp1252", "cp932"):
    try:
        txt = out.decode(enc)
        # 中文字符 + 替换符计数，够判断这次解码是不是“看起来成功但其实是错的”
        has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in txt)
        bad = txt.count("\ufffd")
        print("  %-8s 解码成功   中文=%s  U+FFFD=%d" % (enc, has_cjk, bad))
    except UnicodeDecodeError as e:
        print("  %-8s 解码失败   %s" % (enc, e))

print()
print("---- 按 UTF-8 解码后的 tools/list 描述（前 120 字）----")
try:
    txt = out.decode("utf-8")
    for line in txt.splitlines():
        m = json.loads(line)
        if m.get("id") == 2:
            for t in m["result"]["tools"]:
                d = t["description"]
                print("  %-16s %s" % (t["name"], d[:120]))
except Exception as e:
    print("  (无法按 UTF-8 解析: %r)" % (e,))
