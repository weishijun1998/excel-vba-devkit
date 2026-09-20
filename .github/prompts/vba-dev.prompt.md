---
description: 'Develop or modify an Excel VBA macro and test it automatically (write → test → fix loop)'
argument-hint: 'What should the macro do? Workbook path + expected results (e.g. aggregate revenue by company with a PivotTable)'
agent: 'agent'
---

# Develop a VBA macro and test it automatically

Follow these steps using the terminal tool. **Do not** hand-roll win32com injection.

## 1. Establish the inputs

Ask with `#tool:vscode/askQuestions` (ask for what's missing — don't guess):

- Target workbook path (`.xlsm`; create it if missing)
- Output requirements: which sheets / fields / definitions / PivotTable or chart?
- Verifiable expectations: at least one or two (a cell that must equal a value, a sheet that must exist)

## 2. Write the code

- Write to `macro/<ModuleName>.bas` (**no `Attribute` lines**)
- Every macro gets an `On Error GoTo EH` trap: write `Err.Number & ": " & Err.Description` to a cell + append a log file
- Performance: arrays written in one shot, `ScreenUpdating` / `Calculation` off, no cell-by-cell loops
- Start from `.github/skills/excel-vba-automation/templates/vba_summary_report.bas.txt`

## 3. Run the tests

```bash
python tools/vba/run_vba.py --workbook "<workbook path>" --code <ModuleName>=macro/<ModuleName>.bas \
  --run <MacroName> --expect "cell:<Sheet>!<cell>=<expected>" [more --expect] --keep-open --save
```

Read three things in the report: the **verdict**, the **error**, and the **per-assertion results**.

## 4. Fix from the error — max 3 rounds

- Verdict `vba_error_dialog*`: read the error (raw VBA text) and fix that statement
- Verdict `idle_hung` / `runaway_cpu`: an infinite loop or a wait on an external resource — change the logic (add an exit condition / timeout)
- `⛔ injection refused`: remove `MsgBox` / `Stop` / `Debug.Assert` and friends
- Assertion failed: the report gives the actual value + context (e.g. existing sheet names) — decide whether it's a code bug or a definition (口径) question
- Re-run after each change; **reuse the open Excel instance** (`--keep-open`) to keep rounds at 1–3 s

## 5. Wrap up

- On all green (`exit 0`) report to the user: what the macro does, the test report conclusion, key numbers, workbook path
- Still failing after 3 rounds: hand the **raw error + the changes you tried** to a human — don't keep guessing
- Definition questions (are these numbers right?) **must be confirmed by the user**

中文版：`/vba-dev-zh`
