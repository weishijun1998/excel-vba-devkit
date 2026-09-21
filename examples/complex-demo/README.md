# examples/complex-demo — a full-project smoke test for the whole toolchain

A small but **complete** engineering project (6 modules, 5 000 data rows, budget sheet, SUMIF
reconciliation, PivotTable, 2 charts, named ranges, a slow-motion entry point, and a
cell-write vs array-write benchmark) plus a set of **deliberately broken modules**.

Use it to verify the toolchain itself — after changing `run_vba.py`, the guard, the linters or
the interceptors, run these stages and compare with the expected results below.
Everything is generic: Chinese sheet names only, no customer data, no personal paths.

## Modules

| File | Role |
|---|---|
| `ModDemo_Util.bas.txt` | helpers: logging, sheet (re)creation, array write-back, header style, guarded `Application.Calculation` |
| `ModDemo_Data.bas.txt` | generates the 5 000-row fact table (array write) + the budget sheet |
| `ModDemo_Report.bas.txt` | SUMIF summary, PivotTable, variance sheet, 3 named ranges |
| `ModDemo_Format.bas.txt` | number formats, conditional formatting, freeze panes, 2 charts |
| `ModDemo_Main.bas.txt` | `BuildDemoReport` (full chain, writes the reconciliation sheet), `Demo_SlowRun` (screen updates on, pauses between acts) |
| `ModDemo_Bench.bas.txt` | `Bench_Run` — cell-by-cell write vs one-shot array write, same data both sides |
| `ModFault_Hang.bas.txt` | `Fault_IdleHang` (blocks the pump, ~0 CPU) and `Fault_Runaway` (burns a core forever) |
| `ModFault_MsgBox.bas.txt` | a real `MsgBox` — must be **blocked before injection** |
| `ModFault_Safe.bas.txt` | dangerous words only inside comments/strings (must pass) + a 12 s macro that keeps reporting progress (must not be killed) |
| `ModFault_Unqualified.bas.txt` / `ModFault_Fixed.bas` | the same operation written unqualified (`Sheets("明细数据")`, `ActiveSheet`) vs qualified — 1004 on a hidden-window loader, green once qualified |
| `ModDemo_Scratch.bas.txt` | writes a marker value, used to prove an assertion read a *fresh* result |

## Stage 1 — build the project (visible workbook)

```bash
# ① data layer
python tools/vba/run_vba.py --workbook build/demo.xlsm --save --visible \
  --code ModDemo_Util=examples/complex-demo/ModDemo_Util.bas.txt \
  --code ModDemo_Data=examples/complex-demo/ModDemo_Data.bas.txt \
  --run Step_Generate --expect "cell:明细数据!A1=月份" --expect "sheet:预算=存在"

# ② analysis layer (SUMIF + PivotTable + variance + named ranges)
python tools/vba/run_vba.py --workbook build/demo.xlsm --save --visible \
  --code ModDemo_Util=examples/complex-demo/ModDemo_Util.bas.txt \
  --code ModDemo_Data=examples/complex-demo/ModDemo_Data.bas.txt \
  --code ModDemo_Report=examples/complex-demo/ModDemo_Report.bas.txt \
  --run Step_Analyze --expect "sheet:汇总=存在" --expect "named:收入数据=存在" --expect "cell:步骤!B4=0"

# ③ presentation layer (formats + 2 charts)
python tools/vba/run_vba.py --workbook build/demo.xlsm --save --visible \
  --code ModDemo_Format=examples/complex-demo/ModDemo_Format.bas.txt \
  --run Step_Format --expect "chart:汇总=2" --expect "pivot:汇总=1"
```

## Stage 2 — full chain, one call, 11 assertions

```bash
python tools/vba/run_vba.py --workbook build/demo.xlsm --save --visible \
  --code ModDemo_Util=examples/complex-demo/ModDemo_Util.bas.txt \
  --code ModDemo_Data=examples/complex-demo/ModDemo_Data.bas.txt \
  --code ModDemo_Report=examples/complex-demo/ModDemo_Report.bas.txt \
  --code ModDemo_Format=examples/complex-demo/ModDemo_Format.bas.txt \
  --code ModDemo_Main=examples/complex-demo/ModDemo_Main.bas.txt \
  --code ModDemo_Bench=examples/complex-demo/ModDemo_Bench.bas.txt \
  --run BuildDemoReport \
  --expect "cell:明细数据!A1=月份" --expect "cell:校验!B2=5000" --expect "cell:校验!B4=0" \
  --expect "cell:校验!B5=0" --expect "cell:校验!B6=通过" --expect "cell:校验!B7=通过" \
  --expect "cell:校验!B18=PASS" --expect "pivot:汇总=1" --expect "chart:汇总=2" \
  --expect "named:达成率=存在" --expect "sheet:校验=存在"
```

Expected: `exit 0`, all 11 assertions green, ≈1.6–1.9 s for the whole chain.
The reconciliation sheet compares an independently accumulated total against the SUMIF /
formula results — both deltas must be `0`.

`Demo_SlowRun` does the same with `ScreenUpdating = True` and a pause between acts — for
watching it happen (e.g. in a demo).

## Stage 3 — the guard (hang / runaway / long-but-alive)

```bash
# fake hang: no dialog, pump blocked, ~0 CPU  → expect verdict idle_hung
python tools/vba/run_vba.py --workbook build/demo.xlsm --save \
  --code ModFault_Hang=examples/complex-demo/ModFault_Hang.bas.txt \
  --run Fault_IdleHang --hang-after 6 --budget 30

# runaway: pump blocked, CPU pinned        → expect verdict runaway_cpu
python tools/vba/run_vba.py --workbook build/demo.xlsm --save \
  --code ModFault_Hang=examples/complex-demo/ModFault_Hang.bas.txt \
  --run Fault_Runaway --hang-after 6 --hard-budget 22 --budget 45

# 12 s macro that keeps pumping + reporting → must NOT be killed (zero false kills)
python tools/vba/run_vba.py --workbook build/demo.xlsm --save \
  --code ModFault_Safe=examples/complex-demo/ModFault_Safe.bas.txt \
  --run Fault_LongButAlive --budget 60
```

## Stage 4 — interceptors (block real dialogs, never false-positive)

```bash
# a real MsgBox → blocked before injection, exit 2, nothing runs
python tools/vba/run_vba.py --workbook build/demo.xlsm --save \
  --code ModFault_MsgBox=examples/complex-demo/ModFault_MsgBox.bas.txt \
  --run Fault_MsgBox

# the same words inside comments/strings → allowed, exit 0
python tools/vba/run_vba.py --workbook build/demo.xlsm --save \
  --code ModFault_Safe=examples/complex-demo/ModFault_Safe.bas.txt \
  --run Fault_CommentOnly
```

## Stage 5 — hidden-window loader (the three faces)

Turn a copy of the workbook into a loader: open it, set `wb.Windows(1).Visible = False`, save.
Then:

```bash
python tools/vba/vba_diagnose.py build/loader.xlsm          # static: hidden window + unqualified refs

# unqualified reference → the real landmine: 1004, dialog auto-dismissed, error text read from it
python tools/vba/run_vba.py --workbook build/loader.xlsm --save --visible \
  --code ModFault_Unqualified=examples/complex-demo/ModFault_Unqualified.bas.txt \
  --run Fault_BadRef

# qualified version of the same operation → green, 0.01 s
python tools/vba/run_vba.py --workbook build/loader.xlsm --save --visible \
  --code ModFault_Fixed=examples/complex-demo/ModFault_Fixed.bas.txt \
  --run Fault_GoodRef --expect "cell:明细数据!Z10=qualified-ok"

# the whole project on the loader: bare macro name is qualified automatically, 11/11 green
python tools/vba/run_vba.py --workbook build/loader.xlsm --save \
  --code ModDemo_Util=examples/complex-demo/ModDemo_Util.bas.txt \
  --run BuildDemoReport --expect "cell:校验!B18=PASS" --expect "cell:校验!B4=0"
```

Proof that the assertions read a *fresh* result, not a stale one:

```bash
python tools/vba/run_vba.py --workbook build/loader.xlsm --save --cleanup \
  --code ModDemo_Scratch=examples/complex-demo/ModDemo_Scratch.bas.txt \
  --run Scratch_Touch --expect "cell:校验!B18=STALE"       # seed a fake value
# now re-run BuildDemoReport — B18 must be back to PASS
```

## Stage 6 — reproduce the hidden-window findings

```bash
python tools/vba/bare_name_scope_probe.py         # which call styles need a qualified name
python tools/vba/hidden_app_props_probe.py        # Application.Calculation fails; 9 other settings are fine
python tools/vba/vba_attr_probe.py                # the Attribute-line false "macros are disabled"
python tools/vba/scan_hidden_windows.py <dir>     # which workbooks are hidden on disk (no Excel needed)
python tools/vba/vba_lint_refs.py <dir>           # every unqualified reference, file + line
```

All probes are self-contained: they build their own temporary workbooks and never touch
your files.

## A note on the deliberate bug in `Demo_SlowRun`

`ModDemo_*` recreates sheets by deleting and re-adding them (`EnsureSheet`). If a formula on
another sheet points at a deleted sheet, it becomes `#REF!` — which is exactly what an early
version of the reconciliation sheet did (assertion `cell:校验!B4=0` caught it, actual value
`-2146826265` = error 2023 `#REF!`). The fix is `RefreshCheckSheet`: rewrite the
reconciliation formulas *after* the sheets are rebuilt. Kept in the example on purpose — it
is a realistic failure the assertion layer is supposed to catch.
