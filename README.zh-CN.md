# Excel VBA Devkit

[![ci](https://github.com/weishijun1998/excel-vba-devkit/actions/workflows/ci.yml/badge.svg)](https://github.com/weishijun1998/excel-vba-devkit/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

[English](README.md) ｜ **中文**

Excel VBA Devkit 是一套 Python 工具加一个 agent 技能，用来在 Windows 上开发 Excel VBA 宏。工具会把宏放进真实 Excel 里运行，点掉报错弹窗，卡死时终止进程，检查运行结果，并输出一份报告：判定、真实错因、每条断言的结果。

有了这份报告，AI 就能自己迭代：写宏 → 跑 → 读报告 → 修 → 再跑。技能是标准的 `SKILL.md` 格式，可以装进 VS Code + GitHub Copilot、Claude Code、Codex、Hermes，或任何能读技能、支持 MCP 的宿主。工具本身是纯 Python，也能在终端单独使用，不依赖 AI。

环境要求：Windows + 桌面版 Excel + Python（两个依赖：`pywin32`、`psutil`）。

一个六模块演示工程（5000 行数据、透视表、2 张图表）跑完的样子：

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

失败时，报告会写明是哪一类失败，并附上弹窗原文；弹窗被点掉之后，这段文字就是唯一的记录：

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

控制台输出默认是中文；显示不了中文的终端可以加 `--ascii`，改用纯 ASCII 输出。

## 安装

1. 把 `tools/` 和 `.github/` 复制到你的仓库根目录；用 VS Code 的话再复制 `.vscode/`。
2. 安装两个依赖：`pip install -r tools/vba/requirements.txt`
3. 在宿主里加载技能：
   - VS Code + GitHub Copilot：在 Chat 里输入 `/skills`，应看到 `excel-vba-automation`；输入 `/` 还会列出两个提示文件：`/vba-dev`（英文）和 `/vba-dev-zh`（中文）。
   - 其它宿主：把 `.github/skills/excel-vba-automation/` 复制到该宿主的技能目录。目录名各宿主不同；里面的文件是标准 `SKILL.md`。

一个容易踩的点：技能目录名必须与 `SKILL.md` 里的 `name:` 一致，且不能带 `/` 或 `:`；否则宿主会加载失败，而且不提示原因。

## 用法

### 让 AI 自己开发

在 AI 聊天窗口里说（以 VS Code + GitHub Copilot 的 agent 模式为例）：

> 用 excel-vba-automation 技能写一个宏：把"月度数据"表按公司汇总，加透视表和柱状图，并跑测试验证。

之后它会按 `SKILL.md` 里的循环走：写宏 → 跑 `python tools/vba/run_vba.py ...` → 读报告 → 修 → 再跑。

### 自己在终端跑

```bash
python tools/vba/run_vba.py \
  --workbook "C:/work/report.xlsm" \
  --code ModReport=macro/report.bas \
  --run BuildSummaryReport \
  --expect "cell:校验!B4=0" --expect "pivot:汇总=1" --expect "named:收入数据=存在" \
  --save --keep-open
```

- 退出码：`0` 全过，`1` 有失败，`2` 拒绝运行（危险代码）。
- 默认禁用事件（`EnableEvents=False`）：打开装载器工作簿不会触发它的 `Workbook_Open`；需要触发时加 `--allow-events`。
- 不加 `--save` 不写盘。
- 注入的测试模块默认留在工作簿里；加 `--cleanup` 会在结束时移除，报告里会列出残留的模块。
- 自带模板用的是中文表名（`明细数据` / `汇总` / `校验`）。可以随意改名，注意断言跟着改就行。

### 把演示工程完整跑一遍

`examples/complex-demo/` 是一个完整的六模块工程（5000 行数据、预算表、SUMIF 对账、透视表、2 张图表、命名区域、慢放入口、"逐格写 vs 数组写"对比基准），另外还故意配了几个坏模块：真 `MsgBox`、没有弹窗的假死、跑飞死循环、必须存活 12 秒的长任务，以及同一段操作的两种写法——未限定（报 1004）和限定（全绿）。

共六个阶段，每个阶段的预期结果都写在 [`examples/complex-demo/README.md`](examples/complex-demo/README.md)。第 5 阶段会先往检查单元格写入 `STALE`，再跑宏，断言它变回 `PASS`，用来证明断言读到的是新值。

## 仓库里有什么

| 路径 | 作用 |
|---|---|
| `tools/vba/run_vba.py` | 一条命令跑完整个循环：剥 `Attribute` 行、拦截危险语句、注入、守卫、运行、强制重算、断言、报告。窗口隐藏的工作簿会自动改用限定宏名；源码编码自动识别（UTF-8 / GBK …）；不加 `--save` 不写盘 |
| `tools/vba/vba_project.py` | 工程级读写：`list` / `read` 模块、按行 `patch`（只重写你指定的行，过程属性如快捷键绑定不会丢）、真编译检查、`backup`、`import`、`refs`。所有 COM 调用都在子进程里 |
| `tools/vba/vba_diagnose.py` | 宏"跑不起来"时的只读诊断：窗口隐藏 / 坏引用 / `Attribute` 行 / 未限定引用。可传文件或目录；目录模式是纯静态检查，不用打开 Excel |
| `tools/vba/vba_lint_refs.py` | 静态体检：逐行列出未限定引用（`Sheets(`、`Range(`、`Cells(`、`ActiveSheet`、`ActiveWorkbook`…），带文件和行号。能读 GBK/UTF-8 源码，不开 Excel |
| `tools/vba/vba_guard.py` | 守卫：点掉弹窗（点 id **4800** 的"结束"按钮）、假死判定、跑飞判定、心跳保护 |
| `tools/vba/dismiss_vba_dialog.py` | 一次性清掉卡住的弹窗（BM_CLICK → WM_COMMAND → 真实鼠标 → 杀进程，逐级降级） |
| `tools/vba/vba_attr_probe.py` | 探针：复现 `Attribute` 行导致的 `0x800A03EC` 假故障 |
| `tools/vba/vbe_compile_probe.py` | 探针：复现两种"看起来像编译错误"的情况——中文 VBE 上常见的程序化编译命令其实什么都没做；模块名与过程名相同时宏根本没法运行。两种情况 Excel 都报同一句"宏可能已被禁用" |
| `tools/vba/bare_name_scope_probe.py` | 探针：五种宏名调用方式里，哪些在隐藏窗口装载器上会失败 |
| `tools/vba/hidden_app_props_probe.py` | 探针：窗口隐藏时 `Application.Calculation` 报 1004，其余 9 项常用设置不受影响 |
| `tools/vba/scan_hidden_windows.py` | 静态巡检：不打开 Excel，就能看出哪些工作簿的窗口在磁盘上就是隐藏的（读 `xl/workbook.xml`） |
| `examples/complex-demo/` | 演示工程加故障模块：整条工具链的回归测试（见其 README） |
| `tools/validate.py` | 仓库自检：脚本语法、技能/指令 frontmatter、个人路径、密钥前缀，以及模型可读文件里的禁用措辞 |
| `tools/vba/mcp_server.py` | 零依赖 MCP server，9 个工具：`excel_status`（只读）/ `run_vba` / `dismiss_dialog` / `vba_modules` / `vba_patch_module` / `vba_check_compile` / `vba_backup` / `vba_import_module` / `vba_list_references`。stdio JSON-RPC；输出强制 UTF-8，中文过管道不乱码 |
| `scripts/` | MCP 冒烟测试：`mcp-smoke.ps1`（initialize + `tools/list`）和 `mcp-encoding-probe.py`（直接读 server 的 stdout 原始字节，看它输出的到底是什么编码） |
| `.github/skills/excel-vba-automation/` | 技能本体：铁律、失败矩阵、错误归因、开发循环（中文完整版在 `references/`） |
| `.github/skills/…/templates/` | 可用模板：多维汇总（建表 / 格式 / 透视表 / 图表 / 对账校验） |
| `.github/instructions/vba.instructions.md` | 按文件类型生效的常驻规则（`.bas` / `.vba` / `tools/vba/**`） |
| `.github/prompts/` | `/vba-dev`（英文）和 `/vba-dev-zh`（中文）：写 → 测 → 修 循环，VS Code 斜杠命令 |
| `.vscode/mcp.json` | （可选，VS Code）把 `run_vba` 注册成 MCP 工具 |

## 出问题怎么办

两种互不相关的原因会报同一句话："无法运行宏……可能宏被禁用"。它从来不是信任中心的问题：

1. 代码里有 `Attribute` 行。这是 VBE 的元数据，只在导出的 `.bas` 文件里合法；当成源码注入，会破坏模块，连累整个工程。`run_vba.py` 会自动剥掉。
2. 工作簿窗口在保存时是隐藏的（装载器、宿主类工作簿的常见做法）。这时 Excel 没有活动工作簿，裸宏名解析不出来。`run_vba.py` 会自动改用限定名；`vba_diagnose.py` 能把两种情况分开。

| 现象 | 原因 / 处理 |
|---|---|
| 上述"无法运行宏" | 见上面两种原因；与宏安全设置无关 |
| `⛔ 拒绝注入` | 代码里有 `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / `.Show`，都会阻塞自动化 Excel。删掉，或加 `--allow-unsafe` |
| 守卫在运行中杀掉了 Excel | 工作簿在运行前已保存；重跑一遍即可，脚本会重新打开它 |
| 读 `.bas` 报 `UnicodeDecodeError` | 源码不是 UTF-8（老式中文 / 日文 / 西欧导出）。编码会自动识别；仍失败时加 `--encoding gbk`（或 `cp932` / `cp1252`） |
| 运行中途报 **1004** 或 **91** | 宏里有未限定引用（`Sheets("X")`、`ActiveSheet`…），没有活动工作簿时解析不了。跑 `python tools/vba/vba_lint_refs.py <目录>`，把它们改成 `ThisWorkbook.Worksheets(...)` |
| `> out.txt` 后输出乱码 | 非 UTF-8 控制台上的非 ASCII 字符。工具固定用 UTF-8 + `errors="replace"`；纯文字终端加 `--ascii` |
| 每次都弹"允许运行命令" | VS Code 设置 `chat.tools.terminal.autoApprove` 放行 `python tools/vba/*` |
| 技能没有生效 | 目录名要等于 `SKILL.md` 的 `name`；确认技能已启用；用 Chat 的 Diagnostics 排查 |
| CI 跑什么 | `tools/validate.py`：技能 frontmatter、个人路径、密钥前缀。私有关键词通过环境变量 `VBA_KIT_DENYLIST`（逗号分隔）补充，不要写进仓库 |

## 边界

- 只支持 Windows 桌面版 Excel。整条工具链是 COM + 窗口消息；Excel Online、WPS、LibreOffice、macOS 版 Excel 都不支持。
- 宏卡死时，守卫会直接结束 Excel 进程（有意设计，也是 `run_vba.py` 在运行前先保存工作簿的原因）。最后一次保存之后做的工作会丢失。
- 窗口隐藏的装载器工作簿不会被改动：工具去适配文件，不"修好"文件。

## 本机实测

Windows 11 + Microsoft 365 桌面版 Excel + Python 3.11（数字都来自一台机器，你的耗时会不同）：

- 开发循环：四轮（报错 → 修 → 全绿）
- 弹窗处置：出现后 12–46 ms 被点掉（含守卫启动约 0.2 秒）
- 假死判定：6.5–6.6 秒给出 `IDLE_HUNG` 并处置；跑飞 12.7 秒被终止；16 秒的长任务不会被误杀（心跳保护）
- 隐藏窗口装载器：裸宏名失败，自动改用限定名重试，0.027 秒跑通；隐藏状态会被"另存为"继承（对照实验验证）
- 演示工程（6 模块、11 条断言）：2.09 秒；自带模板：8/8 条断言通过，0.47 秒
- 未在此验证：VS Code 里加载技能 / 提示文件 / MCP；Microsoft 365 桌面版以外的 Excel 版本

## 背景

我平时处理的 Excel 文件又大、公式又密、还会用很多年。我想从 AI 那里拿到的不是宏代码，而是一段跑得起来、能证明自己跑对了的宏。这个仓库就是这份工作的产物。

## 许可证

MIT，见 [LICENSE](LICENSE)。
