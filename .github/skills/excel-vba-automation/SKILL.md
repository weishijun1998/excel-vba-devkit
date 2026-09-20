---
name: excel-vba-automation
description: Use when developing, debugging or batch-running Excel VBA macros (开发/调试/批量运行 Excel VBA 宏：写宏、修编译错、跑测试、处理报错弹窗或卡死、批量刷新工作簿). Runs macros through a guarded dev-test runway instead of hand-rolled win32com.
allowed-tools: shell
---

# Excel VBA development & execution (guarded pipeline)

Always run macros through **`tools/vba/run_vba.py`**. Never hand-roll win32com injection — that just re-creates every trap listed below.

```bash
python tools/vba/run_vba.py --workbook <x.xlsm> --code <ModuleName=file> --run <MacroName> \
  --expect "cell:Sheet!A1=value" [--expect ...] [--save] [--keep-open] [--visible] [--allow-events]
```

Workbook events are **disabled by default** (`EnableEvents=False`), so opening a loader / host workbook does **not** run its `Workbook_Open`. Pass `--allow-events` when you actually *want* the workbook's own startup logic to fire. `vba_diagnose.py` always suppresses events — diagnosis must never run the user's program.

## Hard rules (each one is a real incident)

1. **Strip every `Attribute` line before injecting.** `Attribute VB_Name = "X"` is VBE-owned metadata, valid only in exported `.bas` files. Injected as source text it makes that module fail to compile — **and takes the whole project down with it** (a clean module added next to it won't run either). Excel then reports `0x800A03EC` + "the macro may be disabled", which is **not** a Trust Center problem. `run_vba.py` strips it; when injecting by hand, strip it yourself — and **overwrite or delete the broken module in place**; adding a new one beside it does not help.
2. **A hidden workbook window is a design, not a defect — and it produces the very same error text.** Loader / host workbooks (open-and-run, keep the user out of the way) are commonly saved with `wb.Windows(1).Visible = False`. Excel then has **no active workbook** (`xl.ActiveWorkbook is None`), and an *unqualified* macro name cannot be resolved, because `Run "Macro"` looks the name up in the **active** workbook. You get the exact `0x800A03EC` + "the macro may not be available in this workbook or all macros may be disabled" message from rule 1, but for an unrelated reason — so **diagnosing it along rule 1 leads nowhere**. Do not "fix" the workbook: automate with the qualified `"Workbook.xlsm!Macro"` form. `run_vba.py` detects the hidden state and qualifies automatically, and if a bare name fails it retries qualified. Check any file with `tools/vba/vba_diagnose.py` (static, works without Excel). Note this state is **contagious through "Save As"** — one hidden file makes every descendant hidden.
3. **Never emit `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / UserForm `.Show`.** Each one hangs an automated instance (`Stop` / `Debug.Assert` are the sneakiest: they drop the VBE into break mode → **no dialog, pure hang**). `run_vba.py` blocks them before injection.
4. **Give every macro an error trap**: `On Error GoTo EH`, and in `EH` write `Err.Number & ": " & Err.Description` into a cell **and** append a line to a log file (`Open ... For Append`). That record is the only error source available when no dialog ever appears. Read those files with `encoding="gbk"` (VBA writes ANSI).
5. **Save before running** — `run_vba.py` does this by default. When the guard decides a macro is hung and kills Excel, you then lose one session instead of unsaved work.
6. **Performance**: compute inside arrays and **write the range in one shot**; disable `ScreenUpdating` and set `Calculation = xlCalculationManual` (restore afterwards); prefer formulas / SUMIF / PivotTables over VBA loops.
   - Measured: one cell write = **64 µs**, one file append = **359 µs**. **Cell-by-cell loops are the #1 cause of slow macros** (100 k writes ≈ 6.4 s, 40 k writes ≈ 2058 ms in a real workbook), while pure VBA arithmetic does 3 M iterations in 49 ms. Writing 40 k cells as one array = **42.7 ms → 48× faster**, with the checksum matching the closed-form result (so it is a real speed-up, not less work).
   - Put progress / step markers only at **phase boundaries** (≤10), **never inside a loop**.
7. **A macro that finished is not a macro that is correct.** Assert on key cells / sheets / PivotTables / named ranges. Numbers compare numerically; formula results must be read **after a forced recalculation** (`run_vba.py` handles that).

## Failure matrix (`vba_guard.py`)

| Symptom | Detection | Handling | Measured |
|---|---|---|---|
| Error dialog | class `#32770`, title contains `Visual Basic` | click the id **4800** "End" button (**not `WM_CLOSE` — that does nothing**) | appears → gone in **12–46 ms**, Excel survives |
| `"Cannot run the macro…"` — **two unrelated causes** | (a) a module contains an `Attribute` line / the project does not compile | (a) strip it, overwrite the bad module in place | reproduce with `tools/vba/vba_attr_probe.py` |
| ↑ same message | (b) the workbook window is saved **hidden** → `ActiveWorkbook is None` | (b) **do not modify the file** — use the qualified `"Book.xlsm!Macro"`; `run_vba.py` does it automatically | bare name fails, qualified name passes in **0.016 s**; static check needs no Excel |
| Idle hang | message pump unresponsive + ≈0 CPU over the last 2 s | kill after `hang_after` (default 6 s) | `IDLE_HUNG` at **6.6 s** |
| Runaway loop | unresponsive + CPU > 0.3 s per 2 s | kill after the hard budget | `RUNAWAY_CPU` at **20.8 s** |
| Long job making progress | heartbeat file is fresh | **do not kill**, keep waiting | 16 s macro, **zero false kills** |
| Slow but pumping messages | liveness probe responds | do not kill | — |

Reading the cause out of a dialog: the `Static` control **id 4803** holds the message (e.g. `Runtime error '9': Subscript out of range`). Button ids: **4800 = End, 4801 = Debug, 4802 = Continue (usually greyed out), 4902 = Help**.

## Error attribution order

1. Dialog text captured by the guard (`ERRTEXT`) — most precise
2. The in-macro error record (cell / log file) — when no dialog appears
3. `com_error` hresult — **never trust the code alone, the same fault surfaces as different codes**: `0x800A03EC` (VBA "cannot run the macro") / `0x80020009` wrapped around it (pywin32, `DISP_E_EXCEPTION`) = either an `Attribute` line **or** a hidden-window workbook → tell them apart with the two checks in the failure matrix; `0x800A9C68` = compile error; `0x800706BE` (`RPC_S_CALL_FAILED`) = **the guard just killed the process**, not a macro bug
4. Static checks (blocked before injection) + `vba_diagnose.py` (window / references / Attribute lines)
5. Assertion failures, reported with the actual value + context (e.g. `sheet:差异分析=存在 ❌ (existing sheets: ['月度数据','Sheet1'])` — an instant typo diagnosis)

**One decisive trick when the cause is still unclear**: run the very same macro in a brand-new empty workbook. Green there ⇒ the toolchain and the guard are fine and the problem is in *that* file (hidden window / project / policy); failing there too ⇒ the problem is in your code path or the macro name.

## Development loop (write → test → fix)

1. Write the module to `*.bas` / `*.txt` (no `Attribute` lines)
2. Run `run_vba.py`: you get **verdict + error + per-assertion results**
3. Fix from the error (`--keep-open` reuses the open Excel instance → **1–3 s per round**)
4. Repeat until `exit 0`; **max 3 rounds**, then hand the raw error plus the list of attempts to a human

## Reference

- `templates/vba_summary_report.bas.txt` — multi-dimensional summary template (macro `BuildSummaryReport`): 3 sheets / generated data / formatting + 3-colour scale / autofilter / freeze panes / SUMIF / PivotTable / column chart / reconciliation checks / named range (**0.47 s**, 8/8 assertions green)
- Scripts live in this kit's `tools/vba/`: `run_vba.py` (dev-test runway, auto-qualifies macro names for hidden-window loader workbooks), `vba_guard.py` (dialog / hang guard), `vba_diagnose.py` (read-only diagnosis: hidden window / broken references / `Attribute` lines; also scans a whole directory — static, no Excel needed), `dismiss_vba_dialog.py`, `vba_attr_probe.py`
- 中文完整版：[references/zh-CN.md](references/zh-CN.md)
- Use `/` path separators only, for cross-platform portability
