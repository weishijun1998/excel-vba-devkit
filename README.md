# Excel VBA Devkit

[![ci](https://github.com/weishijun1998/excel-vba-devkit/actions/workflows/ci.yml/badge.svg)](https://github.com/weishijun1998/excel-vba-devkit/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**English** ｜ [中文说明](README.zh-CN.md)

Files you drop into a repo so an AI can write Excel VBA macros, run them, and read back what
happened. Windows + desktop Excel + Python; two dependencies (`pywin32`, `psutil`).

No AI required: the tools in `tools/vba/` are plain Python and run from a terminal. The skill is
plain `SKILL.md` markdown, so any agent host can load it — VS Code + GitHub Copilot, Claude Code,
Codex, Hermes, or anything else that reads skills or speaks MCP.

The loop is the whole idea:

> write a macro → `run_vba.py` runs it → verdict + real error text + per-assertion results → fix → run again

That's what lets an agent iterate without you sitting there clicking error dialogs. A six-module
demo chain (5 000 rows, PivotTable, 2 charts) ends like this:

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

When it fails, the report names the failure — dialog text included, because a dismissed dialog
is otherwise invisible:

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

Console output is Chinese, with ASCII fallback (`--ascii`) for terminals that can't render it.

---

## Install

1. Copy `tools/` and `.github/` into your repo root (add `.vscode/` too if you use VS Code)
2. `pip install -r tools/vba/requirements.txt`
3. Load the skill in your host. With VS Code + Copilot, type `/skills` in Chat — `excel-vba-automation` should be listed, and `/` lists `/vba-dev` (EN) and `/vba-dev-zh` (中文). With another agent, copy `.github/skills/excel-vba-automation/` into that host's skills directory (the folder name differs per host; the file itself is standard `SKILL.md`).

The skill directory name must equal `name:` inside `SKILL.md`, and contain no `/` or `:` —
otherwise the host fails to load it without saying anything.

---

## Use it

### A. Let the agent drive

In your agent's chat — the example below is VS Code + GitHub Copilot in **agent** mode:

> Use the excel-vba-automation skill to write a macro that aggregates the "Monthly" sheet by company, adds a PivotTable and a column chart, and verifies itself with tests.

The agent follows the loop in `SKILL.md`: write → `python tools/vba/run_vba.py …` → read the report → fix → re-run.

### B. Drive it yourself

```bash
python tools/vba/run_vba.py \
  --workbook "C:/work/report.xlsm" \
  --code ModReport=macro/report.bas \
  --run BuildSummaryReport \
  --expect "cell:校验!B4=0" --expect "pivot:汇总=1" --expect "named:收入数据=存在" \
  --save --keep-open
```

Exit codes: `0` all passed ｜ `1` failures ｜ `2` refused to run (dangerous code).

- Events are off by default (`EnableEvents=False`): opening a loader workbook will not fire its `Workbook_Open`. Add `--allow-events` when you *want* that to happen.
- Nothing is written to the workbook unless you pass `--save`.
- Injected test modules stay behind unless you pass `--cleanup`; the report lists what's left.
- The bundled template uses Chinese sheet names (`明细数据` / `汇总` / `校验`). Rename freely — the assertions just have to match.

### C. Single-purpose tools

```bash
python tools/vba/dismiss_vba_dialog.py        # clear a stuck VBA error dialog (4 escalating strategies)
python tools/vba/vba_attr_probe.py            # reproduce the Attribute-line "macros are disabled" false failure
python tools/vba/bare_name_scope_probe.py     # which macro-name call styles fail on a hidden-window loader
python tools/vba/hidden_app_props_probe.py    # hidden window: Application.Calculation raises 1004, 9 other settings do not
python tools/vba/vba_guard.py <excel_pid> 60  # attach the guard to one long-running macro
```

### D. Run the whole toolchain once

`examples/complex-demo/` is a six-module project — 5 000 rows, budget sheet, SUMIF
reconciliation, PivotTable, 2 charts, named ranges, a slow-motion entry point, an array-vs-cell
benchmark — plus deliberately broken modules: a real `MsgBox`, a hang with no dialog, a runaway
loop, a 12 s macro that must survive, and the same operation written unqualified (1004) vs
qualified (green).

Six stages, each with its expected result. Stage 5 proves the assertions read fresh values:
it writes `STALE` into a checkpoint cell, re-runs the macro, and expects `PASS` back. Start at
[`examples/complex-demo/README.md`](examples/complex-demo/README.md).

---

## What's inside

| Path | What it does |
|---|---|
| `tools/vba/run_vba.py` | The runway: strip `Attribute` lines → block dangerous statements → inject → guard → run → force recalculation → assert → report. Auto-qualifies the macro name when the workbook window is saved hidden; auto-detects source encoding (UTF-8 / GBK / …); writes only with `--save` |
| `tools/vba/vba_diagnose.py` | Read-only diagnosis when a macro "cannot run": hidden window / broken references / `Attribute` lines / unqualified references. Takes a file or a directory — the directory mode is static and needs no Excel |
| `tools/vba/vba_lint_refs.py` | Static lint: every unqualified reference (`Sheets(`, `Range(`, `Cells(`, `ActiveSheet`, `ActiveWorkbook`, …) by file and line. Reads GBK/UTF-8 sources, never opens Excel |
| `tools/vba/vba_guard.py` | Guard: dismiss dialogs (clicks the id **4800** "End" button) / idle-hang verdict / runaway-CPU verdict / heartbeat protection |
| `tools/vba/dismiss_vba_dialog.py` | One-shot cleanup of stuck dialogs (BM_CLICK → WM_COMMAND → real mouse → kill) |
| `tools/vba/vba_attr_probe.py` | Attribute-line probe (reproduces the `0x800A03EC` false failure) |
| `tools/vba/bare_name_scope_probe.py` | Probe: which macro-name call styles fail on a hidden-window loader (5 cases, self-contained) |
| `tools/vba/hidden_app_props_probe.py` | Probe: on a hidden-window loader `Application.Calculation` raises 1004 — 9 other common settings are unaffected |
| `tools/vba/scan_hidden_windows.py` | Static scan for workbooks whose window is hidden *on disk* (reads `xl/workbook.xml`, no Excel needed) |
| `examples/complex-demo/` | The demo project + fault modules — regression test for every layer (see its README) |
| `tools/validate.py` | Self-check: script syntax, skill/instruction frontmatter, no personal paths or key prefixes, and no metaphor-heavy wording in model-facing files |
| `tools/vba/mcp_server.py` | Zero-dependency MCP server: `excel_status` (read-only) / `run_vba` / `dismiss_dialog` |
| `.github/skills/excel-vba-automation/` | The skill itself: rules, failure matrix, error attribution, dev loop (Chinese version in `references/`) |
| `.github/skills/…/templates/` | A working multi-dimensional summary template (sheets, formats, PivotTable, chart, reconciliation checks) |
| `.github/instructions/vba.instructions.md` | Always-on rules for `.bas` / `.vba` / `tools/vba/**` |
| `.github/prompts/` | `/vba-dev` (EN) and `/vba-dev-zh` (中文): the write → test → fix loop (VS Code + Copilot slash commands) |
| `.vscode/mcp.json` | (optional, VS Code) registers `run_vba` as a first-class tool over MCP |

---

## When it goes wrong

Two unrelated problems print the exact same sentence — *"Cannot run the macro… macros may be
disabled"* — and it is never a Trust Center problem:

1. The code contains an `Attribute` line. That line is VBE metadata, legal only in an exported
   `.bas`; injected, it breaks the module and the whole project. `run_vba.py` strips it.
2. The workbook window was **saved hidden** — normal for loader / host workbooks. Excel then has
   no active workbook, and a bare macro name cannot resolve. `run_vba.py` auto-qualifies;
   `vba_diagnose.py` tells the two apart.

| Symptom | Cause / fix |
|---|---|
| Same message as above | See the two causes; it is not your macro security settings |
| `⛔ 拒绝注入` (injection refused) | `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / `.Show` in the code — all of them block an automated instance. Remove, or pass `--allow-unsafe` |
| Guard killed Excel mid-run | The workbook was saved before the run — run again, the script reopens it |
| `UnicodeDecodeError` reading a `.bas` | The source isn't UTF-8 (older Chinese/Japanese/European exports). Encoding is auto-detected; if it still fails, pass `--encoding gbk` (or `cp932` / `cp1252`) |
| Runtime error **1004** / **91** *during* a run | Unqualified references (`Sheets("X")`, `ActiveSheet`, …) cannot resolve with no active workbook. Run `python tools/vba/vba_lint_refs.py <dir>` and qualify them with `ThisWorkbook.Worksheets(...)` |
| Broken output after `> out.txt` | Non-ASCII symbols on a non-UTF-8 console. The tools force UTF-8 with `errors="replace"`; add `--ascii` for plain terminals |
| Approving the command every time | VS Code setting `chat.tools.terminal.autoApprove` for `python tools/vba/*` |
| Skill has no effect | Directory name must equal `SKILL.md`'s `name`; check the skill is enabled; use Chat **Diagnostics** |
| What CI runs | `tools/validate.py` — skill frontmatter, personal paths, key prefixes. Add private terms via the `VBA_KIT_DENYLIST` env var (comma-separated), never in the repo |

---

## Limits

- Desktop Excel on Windows only. Excel Online, WPS, LibreOffice and Excel for macOS are not supported — the whole toolchain is COM + window messages.
- When a macro hangs, the guard kills the Excel process. That is the intended behaviour, and why `run_vba.py` saves the workbook before starting. Work done after the last save is lost.
- Loader workbooks with a hidden window are left as they are. The tools adapt to the file; they don't "fix" it.
- Everything below was measured on one machine. Your timings will differ.

## Measured here

Windows 11, Microsoft 365 desktop Excel, Python 3.11:

- Development loop: 4 rounds (error → fix → all green)
- Dialog reaction: appears → dismissed in **12–46 ms**; dismissal ≈ 0.2 s including guard cold start
- Hang detection: `IDLE_HUNG` verdict and kill at 6.5–6.6 s; runaway-CPU killed at 12.7 s; a 16 s long-running macro is **not** killed (heartbeat protection, zero false kills)
- Hidden-window loader: bare macro name fails → auto-retried qualified, green in **0.027 s**; hidden state is inherited by "Save As" (verified with a control experiment)
- Demo chain (6 modules, 11 assertions): **2.09 s**; bundled template: **8/8 assertions in 0.47 s**
- Not tested here: loading Skills / prompt files / MCP inside VS Code, and Excel versions other than Microsoft 365 desktop

---

## Background

I work with Excel files that are large, formula-heavy and old — the kind where a mistake is
expensive and "just re-run it" isn't obvious. What I wanted from an AI was not macro code, but a
macro that runs and proves it ran. This devkit is the pipeline that came out of that: a guard, a
report, and a list of the ways automation fails that no amount of careful prompting fixes.

## License

MIT — see [LICENSE](LICENSE).
