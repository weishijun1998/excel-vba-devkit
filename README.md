# Excel VBA Devkit

[![ci](https://github.com/weishijun1998/excel-vba-devkit/actions/workflows/ci.yml/badge.svg)](https://github.com/weishijun1998/excel-vba-devkit/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**English** ｜ [中文说明](README.zh-CN.md)

Excel VBA Devkit is a set of Python tools plus an agent skill for developing Excel VBA
macros on Windows. The tools run a macro in a real Excel instance, dismiss error dialogs,
stop hangs, and print a report: verdict, real error text, and the result of every assertion.

With that report, an AI can iterate by itself: write, run, read, fix, run again. The skill
uses the standard `SKILL.md` format and loads into VS Code + GitHub Copilot, Claude Code,
Codex, Hermes, or any other host that reads skills or connects over MCP. The tools are plain
Python and run from a terminal; no AI required.

Requirements: Windows, desktop Excel, and Python with `pywin32` and `psutil` installed.

A six-module demo project (5,000 rows, PivotTable, 2 charts) ends like this:

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

When something fails, the report says which kind of failure it was and includes the dialog
text; once a dialog is dismissed, that text is the only record left:

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

Console output is Chinese by default. For terminals that cannot display it, `--ascii`
forces plain-ASCII output.

## Install

1. Copy `tools/` and `.github/` into your repo root. If you use VS Code, copy `.vscode/` too.
2. Install the two dependencies: `pip install -r tools/vba/requirements.txt`
3. Load the skill in your host:
   - VS Code + GitHub Copilot: type `/skills` in Chat; `excel-vba-automation` should be listed. Typing `/` also shows the two prompts, `/vba-dev` (English) and `/vba-dev-zh` (Chinese).
   - Other hosts: copy `.github/skills/excel-vba-automation/` into that host's skills directory. The directory name varies by host; the file inside is standard `SKILL.md`.

One gotcha: the skill directory name must equal the `name:` field inside `SKILL.md` and must
not contain `/` or `:`. Otherwise the host fails to load it, without a message.

## Usage

### From an AI chat

In agent mode (the example below is VS Code + GitHub Copilot):

> Use the excel-vba-automation skill to write a macro that aggregates the "Monthly" sheet by company, adds a PivotTable and a column chart, and verifies itself with tests.

The agent then follows the loop from `SKILL.md`: write the macro, run
`python tools/vba/run_vba.py ...`, read the report, fix, run again.

### From a terminal

```bash
python tools/vba/run_vba.py \
  --workbook "C:/work/report.xlsm" \
  --code ModReport=macro/report.bas \
  --run BuildSummaryReport \
  --expect "cell:校验!B4=0" --expect "pivot:汇总=1" --expect "named:收入数据=存在" \
  --save --keep-open
```

- Exit codes: `0` all passed, `1` failures, `2` refused to run (dangerous code).
- Events are off by default (`EnableEvents=False`): opening a loader workbook does not fire its `Workbook_Open`. Pass `--allow-events` to allow that.
- Nothing is written to the workbook unless you pass `--save`.
- Injected test modules stay in the workbook unless you pass `--cleanup`; the report lists any that are left.
- The bundled template uses Chinese sheet names (`明细数据` / `汇总` / `校验`). Rename them freely; keep the assertions in sync.

### Run the demo project

`examples/complex-demo/` is a complete six-module project (5,000 rows, budget sheet, SUMIF
reconciliation, PivotTable, 2 charts, named ranges, a slow-motion entry point, a
cell-vs-array write benchmark) plus deliberately broken modules: a real `MsgBox`, a hang
with no dialog, a runaway loop, a 12 s macro that must survive, and the same operation
written with and without qualification (1004 vs green).

It runs in six stages, each with an expected result. Stage 5 proves that assertions read
fresh values: it writes `STALE` into a checkpoint cell, re-runs the macro, and expects
`PASS` back. Start at [`examples/complex-demo/README.md`](examples/complex-demo/README.md).

## What's inside

| Path | What it does |
|---|---|
| `tools/vba/run_vba.py` | The full loop in one command: strip `Attribute` lines, block dangerous statements, inject, guard, run, force recalculation, assert, report. Auto-qualifies the macro name when the workbook window is saved hidden; auto-detects source encoding (UTF-8 / GBK / …); writes only with `--save` |
| `tools/vba/vba_project.py` | Project-level read/write: `list` and `read` modules, line-level `patch` (only the lines you name are rewritten, so procedure attributes like shortcut-key bindings survive), a real compile check, `backup`, `import`, `refs`. All COM work runs in a child process |
| `tools/vba/vba_diagnose.py` | Read-only diagnosis for "cannot run the macro": hidden window, broken references, `Attribute` lines, unqualified references. Accepts a file or a directory; the directory mode is static and needs no Excel |
| `tools/vba/vba_lint_refs.py` | Static lint: lists every unqualified reference (`Sheets(`, `Range(`, `Cells(`, `ActiveSheet`, `ActiveWorkbook`, …) with file and line number. Reads GBK/UTF-8 sources; never opens Excel |
| `tools/vba/vba_guard.py` | The guard: dismisses dialogs (clicks the id **4800** "End" button), detects idle hangs and runaway CPU, watches a heartbeat file |
| `tools/vba/dismiss_vba_dialog.py` | One-shot cleanup of stuck dialogs (BM_CLICK → WM_COMMAND → real mouse → kill, escalating) |
| `tools/vba/vba_attr_probe.py` | Probe that reproduces the `0x800A03EC` false failure caused by an `Attribute` line |
| `tools/vba/vbe_compile_probe.py` | Probe that reproduces two cases that look like compile errors: on a Chinese VBE, the common programmatic compile command does nothing; and a module whose name equals its procedure cannot be run at all. Excel shows the same "macros may be disabled" message for both |
| `tools/vba/bare_name_scope_probe.py` | Probe: which of five macro-name call styles fail on a hidden-window loader |
| `tools/vba/hidden_app_props_probe.py` | Probe: on a hidden-window loader, `Application.Calculation` raises 1004 while 9 other common settings are unaffected |
| `tools/vba/scan_hidden_windows.py` | Static scan for workbooks whose window is hidden on disk (reads `xl/workbook.xml`; no Excel needed) |
| `examples/complex-demo/` | The demo project and fault modules; regression test for every layer (see its own README) |
| `tools/validate.py` | Repo self-check: script syntax, skill and instruction frontmatter, personal paths, key prefixes, and banned wording in model-facing files |
| `tools/vba/mcp_server.py` | Zero-dependency MCP server, 9 tools: `excel_status` (read-only) / `run_vba` / `dismiss_dialog` / `vba_modules` / `vba_patch_module` / `vba_check_compile` / `vba_backup` / `vba_import_module` / `vba_list_references`. stdio JSON-RPC; output is forced to UTF-8 so Chinese text survives the pipe |
| `scripts/` | MCP smoke tests: `mcp-smoke.ps1` (initialize + `tools/list`) and `mcp-encoding-probe.py` (reads the server's raw stdout to show which encoding it emits) |
| `.github/skills/excel-vba-automation/` | The skill itself: rules, failure matrix, error attribution, dev loop (Chinese version in `references/`) |
| `.github/skills/…/templates/` | A working multi-dimensional summary template (sheets, formats, PivotTable, chart, reconciliation checks) |
| `.github/instructions/vba.instructions.md` | Always-on rules for `.bas` / `.vba` / `tools/vba/**` |
| `.github/prompts/` | `/vba-dev` (English) and `/vba-dev-zh` (中文): the write → test → fix loop as VS Code slash commands |
| `.vscode/mcp.json` | (optional, VS Code) registers `run_vba` as an MCP tool |

## Troubleshooting

Two unrelated causes produce the same message, *"Cannot run the macro… macros may be
disabled"*. It is never a Trust Center problem:

1. The code contains an `Attribute` line. That line is VBE metadata; it is only valid in an
   exported `.bas` file. Injected as source, it breaks the module and the whole project.
   `run_vba.py` strips it.
2. The workbook window was **saved hidden**, which is common for loader and host workbooks.
   Excel then has no active workbook, and a bare macro name cannot resolve. `run_vba.py`
   auto-qualifies the name; `vba_diagnose.py` tells the two cases apart.

| Symptom | Cause / fix |
|---|---|
| The message above | See the two causes. It is not your macro security settings |
| `⛔ 拒绝注入` (injection refused) | `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / `.Show` in the code; all of them block an automated Excel. Remove them, or pass `--allow-unsafe` |
| Guard killed Excel during a run | The workbook was saved before the run started. Run the command again; the script reopens it |
| `UnicodeDecodeError` reading a `.bas` | The source is not UTF-8 (older Chinese / Japanese / European exports). Encoding is auto-detected; if it still fails, pass `--encoding gbk` (or `cp932` / `cp1252`) |
| Runtime error **1004** or **91** during a run | Unqualified references (`Sheets("X")`, `ActiveSheet`, …) cannot resolve with no active workbook. Run `python tools/vba/vba_lint_refs.py <dir>` and qualify them with `ThisWorkbook.Worksheets(...)` |
| Garbled output after `> out.txt` | Non-ASCII characters on a non-UTF-8 console. The tools force UTF-8 with `errors="replace"`; add `--ascii` for plain terminals |
| Approving the command every time | VS Code setting `chat.tools.terminal.autoApprove` for `python tools/vba/*` |
| Skill has no effect | The directory name must equal `SKILL.md`'s `name`; check that the skill is enabled; use Chat Diagnostics |
| What CI runs | `tools/validate.py`: skill frontmatter, personal paths, key prefixes. Add private terms through the `VBA_KIT_DENYLIST` environment variable (comma-separated); never commit them |

## Limits

- Windows desktop Excel only. The toolchain is COM and window messages; Excel Online, WPS,
  LibreOffice and Excel for macOS are not supported.
- When a macro hangs, the guard kills the Excel process. That is intended, and it is why
  `run_vba.py` saves the workbook before starting. Work done after the last save is lost.
- Loader workbooks with a hidden window are left as they are. The tools adapt to the file
  rather than "fixing" it.

## Measured on one machine

Windows 11, Microsoft 365 desktop Excel, Python 3.11 (numbers from one machine; yours will
differ):

- Development loop: 4 rounds (error → fix → all green)
- Dialog reaction: dismissed 12–46 ms after appearing (about 0.2 s including guard startup)
- Hang detection: `IDLE_HUNG` verdict and kill at 6.5–6.6 s; runaway CPU killed at 12.7 s; a 16 s macro is not killed (heartbeat protection, no false kills)
- Hidden-window loader: a bare macro name fails; the auto-retry with a qualified name passes in 0.027 s. The hidden state is inherited by "Save As" (verified with a control experiment)
- Demo project (6 modules, 11 assertions): 2.09 s; bundled template: 8/8 assertions in 0.47 s
- Not tested here: loading skills, prompt files or MCP inside VS Code, and Excel versions other than Microsoft 365 desktop

## Background

I work with Excel workbooks that are large, formula-heavy, and old. What I wanted from an AI
was not macro code but a macro that runs and proves it ran. This devkit is what came out of
that work.

## License

MIT; see [LICENSE](LICENSE).
