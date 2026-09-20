# VBA Copilot Kit

[![ci](https://github.com/weishijun1998/vba-copilot-kit/actions/workflows/ci.yml/badge.svg)](https://github.com/weishijun1998/vba-copilot-kit/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

[English](README.md) ｜ **中文**

把一套经过实机验证的 Excel VBA 宏开发流水线搬进 **VS Code + GitHub Copilot**（也兼容任何支持 Agent Skills / MCP 的宿主），让 AI 写出来的宏**真的能跑，而且能自证**。

**循环**：Copilot 写宏 → 套装跑测试 → 给出「判定 + 真实错因 + 断言逐条」→ 按报告修 → 再跑。

**为什么需要它**：自动化跑宏的失败模式很凶 —— **模态报错框会一直阻塞到 120 秒超时清场、未保存内容全丢**，或者**根本没有弹窗的假死**，只能干等。这套工具把两类都变成**亚秒级发现 + 分级处置 + 结构化报告**，AI 才能自己迭代，不用你在旁边点按钮。

---

## 三步安装

1. **解压到你的仓库根目录**（会得到 `tools/`、`.github/`、`.vscode/` 三处）
2. **装依赖**：`pip install -r tools/vba/requirements.txt`（只需 `pywin32` + `psutil`；Windows + 本机已装桌面版 Excel）
3. **在 VS Code 里确认装载**：Chat 输入 `/skills` 应能看到 `excel-vba-automation`；输入 `/` 会看到 `/vba-dev`、`/vba-dev-zh`

> 技能目录名必须与 `SKILL.md` 里 `name` 一致，且不能含 `/` 或 `:` —— 否则会**静默加载失败**。

---

## 三种用法

### A. 让 Copilot 自主开发（推荐）

在 Copilot Chat 的 **agent 模式**里说：

> 用 excel-vba-automation 技能帮我写一个宏：把"月度数据"表按公司汇总，加透视表和柱状图，并跑测试验证。

Copilot 会按 SKILL.md 的流程走：写代码 → `python tools/vba/run_vba.py …` → 读报告 → 修 → 重跑。

### B. 你自己在终端跑（不依赖 Copilot）

```bash
python tools/vba/run_vba.py \
  --workbook "C:/work/report.xlsm" \
  --code ModReport=macro/report.bas \
  --run BuildSummaryReport \
  --expect "cell:校验!B4=0" --expect "pivot:汇总=1" --expect "named:收入数据=存在" \
  --save --keep-open
```

退出码：`0` 全过 ｜ `1` 有失败 ｜ `2` 被拦（危险代码）

打开工作簿时**默认禁用事件**（`EnableEvents=False`）—— 装载器工作簿的 `Workbook_Open` **不会**被顺带跑起来；确实想触发时加 `--allow-events`。

### C. 单独处理"卡住的弹窗 / Attribute 坑"

```bash
python tools/vba/dismiss_vba_dialog.py        # 一键清掉卡住的 VBA 报错弹窗（四种关法逐级降级）
python tools/vba/vba_attr_probe.py            # 复现/验证 Attribute 行导致的"宏不可用"
python tools/vba/vba_guard.py <excel_pid> 60  # 单独挂守卫看一个长任务
```

---

## 包含什么

| 路径 | 作用 |
|---|---|
| `tools/vba/run_vba.py` | **开发测试跑道**：剥 Attribute → 危险语句拦截 → 注入 → 守卫 → 运行 → 强制重算 → 断言 → 报告。窗口隐藏的工作簿会**自动改用限定名**；**源码编码自动识别**（UTF-8/GBK…）；**不加 `--save` 绝不写盘**；`--cleanup` 会移除注入的测试模块 |
| `tools/vba/vba_diagnose.py` | **只读诊断**（宏"跑不起来"时先跑它）：窗口隐藏 / 坏引用 / `Attribute` 行 / **未限定引用**。传文件或**目录**都行——目录模式是纯静态检查，不用开 Excel |
| `tools/vba/vba_lint_refs.py` | **静态体检**：逐行列出"未限定引用"（`Sheets(`、`Range(`、`Cells(`、`ActiveSheet`、`ActiveWorkbook`…）——它们会让**所有自动化调用**失败。能读 GBK/UTF-8 源码，不开 Excel |
| `tools/vba/vba_guard.py` | **守卫**：弹窗点掉（点 id **4800**"结束"按钮）/ 假死判定 / 跑飞判定 / 心跳保护 |
| `tools/vba/dismiss_vba_dialog.py` | 卡住弹窗一键清理（BM_CLICK → WM_COMMAND → 真实鼠标 → 杀进程） |
| `tools/vba/vba_attr_probe.py` | Attribute 行探针（复现 `0x800A03EC` 假故障） |
| `tools/validate.py` | 自检：脚本语法 + 技能/指令 frontmatter + 无个人路径与密钥前缀 |
| `tools/vba/mcp_server.py` | **零依赖** MCP server：`excel_status`（只读）/ `run_vba` / `dismiss_dialog` |
| `.github/skills/excel-vba-automation/` | **技能本体**：铁律、失败矩阵、错误归因、开发循环（中文完整版在 `references/zh-CN.md`） |
| `.github/skills/.../templates/` | 可用模板：多维汇总（建表/格式/透视表/图表/对账校验） |
| `.github/instructions/vba.instructions.md` | 按文件类型生效的常驻规则（`applyTo` 指向 `.bas`/`.vba`/`tools/vba/**`） |
| `.github/prompts/` | `/vba-dev`（English）与 `/vba-dev-zh`（中文）：写宏→测试→修 循环 |
| `.vscode/mcp.json` | （可选）把 `run_vba` 注册成 Copilot 的一等工具 |

---

## 已验证 / 未验证（诚实说明）

**本机实测**（Windows 11 + Microsoft 365 桌面版 Excel + Python 3.11）：

- 脚本行为：四轮开发流程（报错 → 修 → 全绿）、无弹窗死循环（12.7 s 判跑飞并处置）、假死判定（6.6 s 判 `IDLE_HUNG`）、心跳保护（16 s 长任务**零误杀**）
- 弹窗反应速度：出现 → 消失 **12–46 ms**；点掉本身 ≈ 0.2 s（含守卫冷启动）
- 断言与报告、Attribute 假故障复现、`MsgBox` 注入前拦截
- **隐藏窗口装载器工作簿**（窗口保存为隐藏 → 无活动工作簿 → 裸宏名必然失败）：实测裸名失败后**自动重试限定名 0.027 秒跑通**；"另存为"会继承隐藏状态（已用对照实验证实）
- 自带模板：**8/8 断言通过，0.47 s**

**未验证**（需要你的环境）：

- Skills / prompt file / MCP 在 VS Code 里的装载与调用
- 除 Microsoft 365 桌面版以外的 Excel 版本

---

## 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| 报"无法运行宏…可能宏被禁用" | **同一句话两种不相干病因**：(a) 代码里有 `Attribute` 行 —— `run_vba.py` 会自动剥，手写时别带，**不是**信任中心的问题；(b) 工作簿窗口**保存时就是隐藏的**（装载器/宿主工作簿的常见设计，**不要去改造那个文件**）→ Excel 没有活动工作簿，裸宏名解析不到。`run_vba.py` 会自动限定；`vba_diagnose.py` 可一眼分清 |
| 注入被拒 `⛔ 拒绝注入` | 代码含 `MsgBox`/`InputBox`/`Stop`/`Debug.Assert`/`.Show`（都会挂死自动化实例），改掉或加 `--allow-unsafe` |
| Excel 被守卫杀了怎么继续 | 运行前已保存工作簿，直接再跑一遍即可（脚本会重新打开） |
| 读 `.bas` 报 `UnicodeDecodeError` | 你的源码不是 UTF-8（老式中文/日文/西欧导出）。现在会自动识别编码；仍失败就加 `--encoding gbk`（或 `cp932`/`cp1252`） |
| 运行时中途报 **1004** / **91** | 宏里有"未限定引用"（`Sheets("X")`、`ActiveSheet`…），没有活动工作簿时解析不了。跑 `python tools/vba/vba_lint_refs.py <目录>` 列出来，改成 `ThisWorkbook.Worksheets(...)` |
| `> out.txt` 之后输出乱/崩 | 非 ASCII 符号撞上非 UTF-8 控制台。工具现已固定 UTF-8 + `errors="replace"`；纯文字终端可加 `--ascii` |
| 不想每次点"允许运行命令" | VS Code 设置 `chat.tools.terminal.autoApprove` 放行 `python tools/vba/*` |
| 技能没生效 | 目录名 = `SKILL.md` 的 `name`；确认技能设置已启用；用 Chat 的 **Diagnostics** 排查 |
| CI 里的 `tools/validate.py` | 会拦下不合规的技能 frontmatter、个人路径与密钥前缀；私有关键词用环境变量 `VBA_KIT_DENYLIST`（逗号分隔）补充，**不要写进仓库** |

---

## 许可证

MIT —— 见 [LICENSE](LICENSE)。
