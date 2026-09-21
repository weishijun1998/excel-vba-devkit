# VBA Copilot Kit

[![ci](https://github.com/weishijun1998/vba-copilot-kit/actions/workflows/ci.yml/badge.svg)](https://github.com/weishijun1998/vba-copilot-kit/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

[English](README.md) ｜ **中文**

把一组文件放进你的仓库，AI 就能写 Excel VBA 宏、把宏跑起来、并读懂运行结果。
环境要求：Windows + 桌面版 Excel + Python；只装两个依赖（`pywin32`、`psutil`）。

为 VS Code + GitHub Copilot 而做；任何能读 Agent Skills（`SKILL.md`）或支持 MCP 的宿主都能用。

整个东西就是一个循环：

> 写宏 → `run_vba.py` 跑它 → 给出「判定 + 真实错因 + 断言逐条结果」→ 按报告修 → 再跑

有了这个循环，AI 才能自己迭代，而不是让你守在旁边一个个点掉报错弹窗。下面是一个六模块演示工程（5000 行数据、透视表、2 张图表）跑完的真实输出：

```text
Excel: 新建实例 | pid=42552 | 工作簿=demo.xlsm
注入模块: ModDemo_Util
…
=== 判定 ===
结论: ✅ 正常运行完成
耗时: 2.09s | 宏返回: ok

=== 断言 ===
  ✅ cell:明细数据!A1=月份
  ✅ cell:校验!B4=0
  ✅ pivot:汇总=1
  ✅ chart:汇总=2
…
=== 总结: 全部通过 ✅ ===
```

失败时，报告会写清是哪一类失败，并附上弹窗原文 —— 弹窗被点掉之后就只剩报告里的这一份证据：

```text
结论: ❌ VBA 报错（弹窗已自动点掉，Excel 存活）
耗时: 0.197s
宏名: 已自动限定为 loader.xlsm!Fault_BadRef（窗口隐藏态，属设计，未改文件）
错因: 运行时错误 '1004': 方法 'Sheets' 作用于对象 '_Global' 时失败
守卫: 弹窗已点掉（Excel 存活=True）

结论: ❌ 卡死（无弹窗、零 CPU）→ 已杀进程
耗时: 6.507s
守卫: ['VERDICT IDLE_HUNG elapsed=6.5s cpu_delta=1.0s -> KILLED pid=4112']
```

控制台输出为中文；终端渲染不了非 ASCII 字符时加 `--ascii` 降级为纯 ASCII。

---

## 安装

1. 把 `tools/`、`.github/`、`.vscode/` 复制到你的仓库根目录
2. `pip install -r tools/vba/requirements.txt`
3. 在 VS Code 的 Chat 里输入 `/skills`，应能看到 `excel-vba-automation`；输入 `/` 能看到 `/vba-dev`（英文）和 `/vba-dev-zh`（中文）

技能目录名必须与 `SKILL.md` 里的 `name:` 一致，且不能含 `/` 或 `:` —— 否则加载失败而且不报错。

---

## 用法

### A. 让 Copilot 自己开发

在 Copilot Chat 的 **agent 模式**里说：

> 用 excel-vba-automation 技能帮我写一个宏：把"月度数据"表按公司汇总，加透视表和柱状图，并跑测试验证。

Copilot 会按 `SKILL.md` 里的循环走：写代码 → `python tools/vba/run_vba.py …` → 读报告 → 修 → 重跑。

### B. 自己在终端跑

```bash
python tools/vba/run_vba.py \
  --workbook "C:/work/report.xlsm" \
  --code ModReport=macro/report.bas \
  --run BuildSummaryReport \
  --expect "cell:校验!B4=0" --expect "pivot:汇总=1" --expect "named:收入数据=存在" \
  --save --keep-open
```

退出码：`0` 全过 ｜ `1` 有失败 ｜ `2` 拒绝运行（危险代码）。

- 默认禁用事件（`EnableEvents=False`）：打开装载器工作簿**不会**顺带触发它的 `Workbook_Open`；确实需要时加 `--allow-events`。
- 不加 `--save` 不写盘。
- 注入的测试模块默认留在工作簿里，报告会列出来；加 `--cleanup` 会移除。
- 自带模板用的是中文表名（`明细数据` / `汇总` / `校验`），可以随意改，只要断言跟着改。

### C. 单独用某个工具

```bash
python tools/vba/dismiss_vba_dialog.py        # 清掉卡住的 VBA 报错弹窗（四种关法逐级降级）
python tools/vba/vba_attr_probe.py            # 复现 Attribute 行导致的"宏被禁用"假故障
python tools/vba/bare_name_scope_probe.py     # 五种调用方式实测：哪些在隐藏窗口装载器上会失败
python tools/vba/hidden_app_props_probe.py    # 窗口隐藏时只有 Application.Calculation 报 1004，其余 9 项正常
python tools/vba/vba_guard.py <excel_pid> 60  # 单独挂守卫看一个长任务
```

### D. 用一个演示工程把整条工具链跑一遍

`examples/complex-demo/` 是一个六模块工程 —— 5000 行数据、预算表、SUMIF 对账、透视表、
2 张图表、命名区域、慢放入口、数组写 vs 逐格写基准 —— 外加一组故意做坏的模块：真 `MsgBox`、
无弹窗假死、跑飞死循环、必须存活 12 秒的长任务、以及同一个操作的未限定写法（1004）与限定写法（全绿）。

六个阶段，每个阶段都写了预期结果。第 5 阶段用一个办法证明断言读到的是新值：先往检查单元格
写入 `STALE`，再跑宏，断言它变回 `PASS`。入口：[`examples/complex-demo/README.md`](examples/complex-demo/README.md)。

---

## 包含什么

| 路径 | 作用 |
|---|---|
| `tools/vba/run_vba.py` | 主跑道：剥 `Attribute` 行 → 拦截危险语句 → 注入 → 守卫 → 运行 → 强制重算 → 断言 → 报告。窗口隐藏的工作簿会**自动改用限定宏名**；源码编码自动识别（UTF-8 / GBK …）；不加 `--save` 不写盘 |
| `tools/vba/vba_diagnose.py` | 只读诊断（宏"跑不起来"时先跑它）：窗口隐藏 / 坏引用 / `Attribute` 行 / 未限定引用。传文件或目录都行，目录模式纯静态、不用开 Excel |
| `tools/vba/vba_lint_refs.py` | 静态体检：逐行列出未限定引用（`Sheets(`、`Range(`、`Cells(`、`ActiveSheet`、`ActiveWorkbook`…）。能读 GBK/UTF-8 源码，不开 Excel |
| `tools/vba/vba_guard.py` | 守卫：弹窗点掉（点 id **4800**"结束"按钮）/ 假死判定 / 跑飞判定 / 心跳保护 |
| `tools/vba/dismiss_vba_dialog.py` | 卡住弹窗一次性清理（BM_CLICK → WM_COMMAND → 真实鼠标 → 杀进程） |
| `tools/vba/vba_attr_probe.py` | Attribute 行探针（复现 `0x800A03EC` 假故障） |
| `tools/vba/bare_name_scope_probe.py` | 探针：五种宏名调用方式里哪些在隐藏窗口装载器上失败（自带临时文件，不碰用户文件） |
| `tools/vba/hidden_app_props_probe.py` | 探针：窗口隐藏时 `Application.Calculation` 报 1004，其余 9 项常用设置正常 |
| `tools/vba/scan_hidden_windows.py` | 静态巡检：不打开 Excel 就能看出哪些工作簿的窗口在磁盘上就是隐藏的（读 `xl/workbook.xml`） |
| `examples/complex-demo/` | 完整六模块演示工程 + 故障模块 —— 用来回归整条工具链（见其 README） |
| `tools/validate.py` | 自检：脚本语法、技能/指令 frontmatter、个人路径与密钥前缀，以及模型可读文件中不得出现比喻/口语 |
| `tools/vba/mcp_server.py` | 零依赖 MCP server：`excel_status`（只读）/ `run_vba` / `dismiss_dialog` |
| `.github/skills/excel-vba-automation/` | 技能本体：铁律、失败矩阵、错误归因、开发循环（中文完整版在 `references/`） |
| `.github/skills/…/templates/` | 可用模板：多维汇总（建表 / 格式 / 透视表 / 图表 / 对账校验） |
| `.github/instructions/vba.instructions.md` | 按文件类型生效的常驻规则（`applyTo` 指向 `.bas` / `.vba` / `tools/vba/**`） |
| `.github/prompts/` | `/vba-dev`（英文）与 `/vba-dev-zh`（中文）：写宏 → 测试 → 修 循环 |
| `.vscode/mcp.json` | （可选）把 `run_vba` 注册成 Copilot 的一等工具 |

---

## 出问题时

有**两个互不相关**的原因会吐出同一句话 —— "无法运行宏…可能宏被禁用" —— 而且从来不是信任中心的问题：

1. 代码里有 `Attribute` 行。它是 VBE 的元数据，只在导出的 `.bas` 里合法；带它注入会让模块编译失败，并拖垮整个工程。`run_vba.py` 会自动剥。
2. 工作簿窗口**保存时就是隐藏的** —— 装载器/宿主类工作簿的常见（有意）设计。此时 Excel 没有活动工作簿，裸宏名解析不到。`run_vba.py` 会自动限定；`vba_diagnose.py` 能把两种病因分开。

| 现象 | 原因 / 处理 |
|---|---|
| 上述"无法运行宏" | 见上面两种病因，与宏安全设置无关 |
| `⛔ 拒绝注入` | 代码含 `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / `.Show`（都会阻塞自动化实例）。改掉，或加 `--allow-unsafe` |
| Excel 被守卫杀了 | 运行前已保存，直接再跑一遍即可（脚本会重新打开工作簿） |
| 读 `.bas` 报 `UnicodeDecodeError` | 源码不是 UTF-8（老式中文/日文/西欧导出）。编码已自动识别；仍失败就加 `--encoding gbk`（或 `cp932` / `cp1252`） |
| 运行中途报 **1004** / **91** | 宏里有未限定引用（`Sheets("X")`、`ActiveSheet`…），没有活动工作簿时解析不了。跑 `python tools/vba/vba_lint_refs.py <目录>` 列出来，改成 `ThisWorkbook.Worksheets(...)` |
| `> out.txt` 后输出乱掉 | 非 ASCII 字符撞上非 UTF-8 控制台。工具现固定 UTF-8 + `errors="replace"`；纯文字终端加 `--ascii` |
| 不想每次点"允许运行命令" | VS Code 设置 `chat.tools.terminal.autoApprove` 放行 `python tools/vba/*` |
| 技能没生效 | 目录名 = `SKILL.md` 的 `name`；确认技能已启用；用 Chat 的 **Diagnostics** 排查 |
| CI 跑什么 | `tools/validate.py` —— 技能 frontmatter、个人路径、密钥前缀。私有关键词用环境变量 `VBA_KIT_DENYLIST`（逗号分隔）补充，不写进仓库 |

---

## 边界

- 只支持 Windows 桌面版 Excel。Excel Online、WPS、LibreOffice、macOS 版 Excel 都不支持 —— 整条工具链基于 COM 和窗口消息。
- 宏卡死时，守卫会杀掉 Excel 进程。这是有意设计，也正是 `run_vba.py` 在运行前先保存工作簿的原因；最后一次保存之后的工作会丢失。
- 窗口隐藏的装载器工作簿保持原样。工具适配文件，不"修复"文件。
- 下面的数字都出自一台机器，你的耗时会不同。

## 本机实测

Windows 11 + Microsoft 365 桌面版 Excel + Python 3.11：

- 开发循环：四轮（报错 → 修 → 全绿）
- 弹窗反应：出现 → 消失 **12–46 ms**；点掉本身 ≈ 0.2 秒（含守卫冷启动）
- 假死判定：6.5–6.6 秒给出 `IDLE_HUNG` 并处置；跑飞在 12.7 秒被杀；16 秒长任务**不被误杀**（心跳保护）
- 隐藏窗口装载器：裸宏名失败后**自动重试限定名，0.027 秒跑通**；隐藏状态会被"另存为"继承（已用对照实验证实）
- 演示工程全链路（6 模块、11 条断言）：**2.09 秒**；自带模板：**8/8 断言通过，0.47 秒**
- 未在此验证：Skills / prompt 文件 / MCP 在 VS Code 里的装载与调用；除 Microsoft 365 桌面版以外的 Excel 版本

---

## 由来

我日常面对的 Excel 文件大、公式密、年头久，出错代价高，而且不是"重跑一遍就行"的那种。
我想从 AI 那里拿到的不只是宏代码，而是一个跑得起来、并且能证明自己跑对了的宏。
这个仓库就是为此搭出来的一条流水线：一个守卫、一份报告，以及一份"自动化会以哪些方式失败"的清单 ——
这些失败靠反复叮嘱模型是解决不掉的。

## 许可证

MIT —— 见 [LICENSE](LICENSE)。
