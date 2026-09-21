#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""本地自检 / CI 用：脚本语法 + 技能与指令文件的 frontmatter 规范 + 密钥/个人路径 + 文风。

用法： python tools/validate.py
退出码：0=全部通过；1=有失败

为什么需要它：技能/指令文件的 frontmatter 写错（缺 name/description、目录名与 name 不一致、
用 `org/` 前缀）会导致**宿主静默不加载**，不报错、很难查。这个检查能提前抓到。
另外两层护栏同样重要：① 提交前拦个人路径/密钥特征；② `.github/**` 是**模型可读**文本，
比喻与口语标记（"面孔""地雷""兜住"）只增加歧义，由 check_prose() 拦住。
"""

import os
import re
import sys
import py_compile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
fails = []
oks = []


def check_python(paths):
    for p in paths:
        try:
            py_compile.compile(p, doraise=True)
            oks.append("语法 OK  %s" % os.path.relpath(p, ROOT))
        except py_compile.PyCompileError as e:
            fails.append("语法错误 %s: %s" % (os.path.relpath(p, ROOT), e))


def frontmatter(path):
    """极简 frontmatter 解析：返回 dict（只支持 key: value 单行）。"""
    with open(path, encoding="utf-8") as f:
        txt = f.read()
    if not txt.startswith("---"):
        return None, txt
    end = txt.find("\n---", 3)
    if end < 0:
        return None, txt
    head = txt[3:end]
    data = {}
    for line in head.splitlines():
        m = re.match(r"^([A-Za-z_][\w-]*)\s*:\s*(.+?)\s*$", line)
        if m and not line.startswith(" "):
            data[m.group(1)] = m.group(2).strip("'\"")
    return data, txt[end + 4:]


def check_skills():
    skills_dir = os.path.join(ROOT, ".github", "skills")
    if not os.path.isdir(skills_dir):
        return
    for name in sorted(os.listdir(skills_dir)):
        sk = os.path.join(skills_dir, name, "SKILL.md")
        if not os.path.isfile(sk):
            continue
        fm, _ = frontmatter(sk)
        rel = os.path.relpath(sk, ROOT)
        if fm is None:
            fails.append("%s 缺少 YAML frontmatter（以 --- 开头/结尾）" % rel)
            continue
        if not fm.get("name"):
            fails.append("%s 缺少 name" % rel)
        elif fm["name"] != name:
            fails.append("%s 的 name=%r 与目录名 %r 不一致（宿主会静默不加载）" % (rel, fm["name"], name))
        elif "/" in fm["name"] or ":" in fm["name"]:
            fails.append("%s 的 name 含 / 或 : （会导致静默加载失败）" % rel)
        if not fm.get("description"):
            fails.append("%s 缺少 description（模型靠它决定何时加载）" % rel)
        if fm.get("name") == name and fm.get("description"):
            oks.append("技能 OK  %s (name=%s)" % (rel, name))


def check_md_dir(sub, suffix):
    d = os.path.join(ROOT, ".github", sub)
    if not os.path.isdir(d):
        return
    for f in sorted(os.listdir(d)):
        if not f.endswith(suffix):
            continue
        p = os.path.join(d, f)
        fm, _ = frontmatter(p)
        rel = os.path.relpath(p, ROOT)
        if fm is None:
            fails.append("%s 缺少 frontmatter" % rel)
        elif not fm.get("description"):
            fails.append("%s 缺少 description" % rel)
        else:
            oks.append("文件 OK  %s" % rel)


def check_secrets():
    """防误提交：扫绝对个人路径 / 常见密钥前缀；可另用环境变量 VBA_KIT_DENYLIST 补充私有关键词（逗号分隔，不写进仓库）。"""
    patterns = [
        (re.compile(r"C:\\+Users\\+[^\\\s\"']+"), "Windows 绝对个人路径"),
        (re.compile(r"/Users/[A-Za-z0-9._-]+/"), "macOS 绝对个人路径"),
        (re.compile(r"\bghp_[A-Za-z0-9]{20,}"), "GitHub PAT"),
        (re.compile(r"\bsk-[A-Za-z0-9]{20,}"), "OpenAI 风格密钥"),
        (re.compile(r"\bAKIA[0-9A-Z]{16}\b"), "AWS Access Key"),
    ]
    extra = [t.strip() for t in os.environ.get("VBA_KIT_DENYLIST", "").split(",") if t.strip()]
    patterns += [(re.compile(re.escape(t), re.I), "自定义禁用词 %r" % t) for t in extra]

    exts = {".py", ".md", ".json", ".yml", ".yaml", ".txt", ".bas", ".vbs", ".ps1", ".sh"}
    hits = 0
    for root, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in (".git", "__pycache__", ".venv", "venv")]
        for f in files:
            if os.path.splitext(f)[1].lower() not in exts:
                continue
            p = os.path.join(root, f)
            try:
                txt = open(p, encoding="utf-8", errors="ignore").read()
            except Exception:
                continue
            for rx, why in patterns:
                if rx.search(txt):
                    fails.append("%s 命中 %s（公开仓库前请脱敏）" % (os.path.relpath(p, ROOT), why))
                    hits += 1
    if hits == 0:
        oks.append("内容 OK  无个人路径/密钥特征")


def check_prose():
    """文风检查：`.github/**` 下的技能 / 指令 / 提示词是**模型可读**的上下文。

    比喻（"第二张面孔""地雷""被咬""兜住"）与口语（"跑得好好的""白折腾"）只增加
    解码成本与歧义，不增加信息；应写成字面表述（"同一成因的第二种表现""失败""On Error
    包裹"）。人读的文档（README.md、examples/*/README.md，与本目录无关）不受此限制。
    确需保留的单个句子可在该行加 `lang-ok` 豁免。
    """
    patterns = [
        (re.compile(r"面孔|地雷|被咬|兜住|踩坑|白折腾|跑得好好的|头号|偷偷|全毁|会炸"
                    r"|一眼看出|一记|钉死|挂死|最隐蔽"), "比喻/口语"),
        (re.compile(r"\b(landmine|landmines|sneakiest)\b|second face|third face|bites\b",
                    re.I), "metaphor"),
    ]
    targets = []
    for sub in ("skills", "instructions", "prompts"):
        for root, dirs, files in os.walk(os.path.join(ROOT, ".github", sub)):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            targets += [os.path.join(root, f) for f in files if f.endswith(".md")]
    hits = 0
    for p in targets:
        rel = os.path.relpath(p, ROOT)
        try:
            lines = open(p, encoding="utf-8", errors="ignore").read().splitlines()
        except Exception:
            continue
        for i, line in enumerate(lines, 1):
            if "lang-ok" in line:
                continue
            for rx, why in patterns:
                m = rx.search(line)
                if m:
                    fails.append("%s:%d 命中%s标记 %r（改成字面表述，或该行加 lang-ok 豁免）"
                                 % (rel, i, why, m.group(0)))
                    hits += 1
    if hits == 0:
        oks.append("文风 OK  %d 个模型可读文件无比喻/口语标记" % len(targets))


def main():
    py_files = []
    for root, dirs, files in os.walk(os.path.join(ROOT, "tools")):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        py_files += [os.path.join(root, f) for f in files if f.endswith(".py")]
    check_python(py_files)
    check_skills()
    check_md_dir("instructions", ".instructions.md")
    check_md_dir("prompts", ".prompt.md")
    check_secrets()
    check_prose()

    for line in oks:
        print("  " + line)
    for line in fails:
        print("  ✗ " + line)
    print("\n%s：%d 项通过，%d 项失败" % ("FAIL" if fails else "PASS", len(oks), len(fails)))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
