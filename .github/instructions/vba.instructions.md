---
name: 'VBA macro rules'
description: 'Non-negotiable rules when writing or debugging Excel VBA macros (中文规则见技能 references/zh-CN.md)'
applyTo: ['**/*.bas', '**/*.vba', '**/vba/**/*.py', 'tools/vba/**']
---

# VBA macro rules

- **Strip every `Attribute` line before injecting.** It is VBE-owned metadata, valid only in exported `.bas` files. Injected as source it breaks compilation and takes the whole VBA project down (`0x800A03EC` + "macro may be disabled" — **not** a Trust Center issue). Overwrite or delete broken modules in place.
- **A hidden workbook window is a design (loader / host workbook), not a defect** — but it leaves Excel with no active workbook, so **unqualified macro names cannot run**, and the error text is *identical* to the Attribute rule above. Never "fix" the file: call `"Workbook.xlsm!Macro"` instead. `run_vba.py` qualifies automatically; inspect any file with `python tools/vba/vba_diagnose.py <file>`.
- **Qualify every reference *inside* macros**: `Sheets("X")` / `Range("A1")` / `Cells(...)` / `ActiveSheet` / `ActiveWorkbook` resolve through `_Global`, which needs an active workbook/sheet — with none (hidden-window loader) `Sheets(...)` fails with **1004**, and `Set ws = ActiveSheet` yields `Nothing` silently, failing later with **91**. Always write `ThisWorkbook.Worksheets("X")`; audit a whole tree with `python tools/vba/vba_lint_refs.py <dir>`.
- **And guard `Application.Calculation`**: with a hidden window it raises `1004` on read *and* write (measured; every other common setting — `ScreenUpdating` / `DisplayAlerts` / `StatusBar` / `EnableEvents` — is unaffected). It is a speed switch, so wrap it and fall back to automatic; otherwise a perf toggle aborts the whole macro. Reproduce with `python tools/vba/hidden_app_props_probe.py`.
- ⚠️ **Rewriting an *existing* business module? Keep its procedure attributes.** Export → strip `Attribute` lines → write the whole module back silently drops procedure-level attributes — e.g. `Attribute ProcName.VB_ProcData.VB_Invoke_Func = "P\n14"` **is a keyboard-shortcut binding**. Observed live: after such an "optimisation" of a 1,874-line production module, that macro's shortcut stopped firing (nobody noticed, because only the run result was verified). Rule: strip attributes only from code you *inject*; before rewriting an existing module, export and diff it, and carry `VB_ProcData.*` lines over verbatim. To restore: insert the line back into the exported `.bas` → `VBComponents.Remove` → `Import` (**`Import` accepts `Attribute` lines; only `AddFromString` breaks**), and write the file as system ANSI (cp936 here) or every Chinese comment is destroyed.
- **Never emit `MsgBox`, `InputBox`, `Stop`, `Debug.Assert`, or UserForm `.Show`** — each hangs an automated instance (`Stop` / `Debug.Assert` produce a hang with **no dialog**).
- **Give every macro an error trap**: `On Error GoTo EH`, writing `Err.Number & ": " & Err.Description` to a cell **and** a log file (read those files as `gbk`).
- **Run macros only through `python tools/vba/run_vba.py`** — it strips Attribute lines, blocks dangerous statements, attaches the guard, forces recalculation, evaluates assertions and prints a report. Do not hand-roll win32com injection.
- **Always verify**: use `--expect` on key cells / sheets / PivotTables / named ranges. "No error" is not "correct".
- **Performance**: compute in arrays and write ranges in one shot; disable `ScreenUpdating` / `Calculation`; never loop cell-by-cell (one write = 64 µs). Step markers only at phase boundaries, never inside loops.
- **Read the report before changing code**: it already carries the verdict (ok / error dialog / idle hang / runaway) + the real error + per-assertion results. Don't re-run on a hunch.

中文版：`.github/skills/excel-vba-automation/references/zh-CN.md`
