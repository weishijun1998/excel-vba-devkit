# VBA Copilot Kit

[![ci](https://github.com/weishijun1998/vba-copilot-kit/actions/workflows/ci.yml/badge.svg)](https://github.com/weishijun1998/vba-copilot-kit/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**English** ｜ [中文说明](README.zh-CN.md)

Bring a battle-tested Excel VBA development pipeline into **VS Code + GitHub Copilot** (or any Agent Skills / MCP host), so an AI can write macros that actually run — and prove it.

**The loop:** Copilot writes the macro → the kit runs it → you get a verdict + the real error + per-assertion results → fix → re-run.

**Why this exists:** when macros are driven by automation, the failure modes are brutal — a **modal error dialog that blocks until a 120 s timeout wipes the unsaved session**, or a **hang with no dialog at all** that you can only wait out. This kit turns both into **sub-second detection + graded handling + a structured report**, which is what lets an AI iterate on its own instead of leaving you to click buttons.

---

## Install (3 steps)

1. **Drop the kit into your repo root** — you get `tools/`, `.github/`, `.vscode/`
2. **Install deps**: `pip install -r tools/vba/requirements.txt` (only `pywin32` + `psutil`; needs Windows + desktop Excel)
3. **Confirm it loaded in VS Code**: type `/skills` in Chat — you should see `excel-vba-automation`; `/` lists `/vba-dev`

> The skill directory name must equal `name` in `SKILL.md`, and must not contain `/` or `:` — otherwise it **silently fails to load**.

---

## Three ways to use it

### A. Let Copilot develop the macro (recommended)

In Copilot Chat (**agent** mode):

> Use the excel-vba-automation skill to write a macro that aggregates the "Monthly" sheet by company, adds a PivotTable and a column chart, and verify it with tests.

Copilot follows the SKILL.md loop: write code → `python tools/vba/run_vba.py …` → read the report → fix → re-run.

### B. Run it yourself in a terminal (no Copilot needed)

```bash
python tools/vba/run_vba.py \
  --workbook "C:/work/report.xlsm" \
  --code ModReport=macro/report.bas \
  --run BuildSummaryReport \
  --expect "cell:校验!B4=0" --expect "pivot:汇总=1" --expect "named:收入数据=存在" \
  --save --keep-open
```

Exit codes: `0` all passed ｜ `1` failures ｜ `2` blocked (dangerous code).

The bundled template uses Chinese sheet names — `明细数据` (data) / `汇总` (summary) / `校验` (checks). Rename freely; the assertions just have to match.

### C. Handle a stuck dialog or the Attribute trap

```bash
python tools/vba/dismiss_vba_dialog.py        # clear a stuck VBA error dialog (4 escalating strategies)
python tools/vba/vba_attr_probe.py            # reproduce / verify the Attribute-line false failure
python tools/vba/vba_guard.py <excel_pid> 60  # attach the guard to one long-running macro
```

---

## What's inside

| Path | What it does |
|---|---|
| `tools/vba/run_vba.py` | **The dev-test runway**: strip `Attribute` lines → block dangerous statements → inject → guard → run → force recalculation → assert → report. Auto-qualifies the macro name when the workbook window is saved hidden |
| `tools/vba/vba_diagnose.py` | **Read-only diagnosis** when a macro "cannot run": hidden window / broken references / `Attribute` lines. Takes a file *or a directory* — the directory mode is static and needs no Excel |
| `tools/vba/vba_guard.py` | **Guard**: dismiss dialogs (clicks the id **4800** "End" button) / idle-hang verdict / runaway-CPU verdict / heartbeat protection |
| `tools/vba/dismiss_vba_dialog.py` | One-shot cleanup of stuck dialogs (BM_CLICK → WM_COMMAND → real mouse → kill) |
| `tools/vba/vba_attr_probe.py` | Attribute-line probe (reproduces the `0x800A03EC` false failure) |
| `tools/validate.py` | Self-check: script syntax + skill/instruction frontmatter + no personal paths or key prefixes |
| `tools/vba/mcp_server.py` | Zero-dependency MCP server: `excel_status` (read-only) / `run_vba` / `dismiss_dialog` |
| `.github/skills/excel-vba-automation/` | **The Agent Skill**: rules, failure matrix, error attribution, dev loop (+ Chinese version in `references/`) |
| `.github/skills/.../templates/` | A working multi-dimensional summary template (sheets, formats, PivotTable, chart, reconciliation checks) |
| `.github/instructions/vba.instructions.md` | Always-on rules applied to `.bas` / `.vba` / `tools/vba/**` |
| `.github/prompts/` | `/vba-dev` (EN) and `/vba-dev-zh` (中文): the write → test → fix loop |
| `.vscode/mcp.json` | (optional) registers `run_vba` as a first-class Copilot tool |

---

## Verified / not verified (no hand-waving)

**Measured on this machine** (Windows 11, Microsoft 365 desktop Excel, Python 3.11):

- Script behaviour: 4-round development loop (error → fix → all green), no-dialog infinite loop (runaway detected and killed at 12.7 s), idle hang (`IDLE_HUNG` at 6.6 s), heartbeat protection (16 s macro, **zero false kills**)
- Dialog reaction: dialog appears → dismissed in **12–46 ms**; dismissal itself ≈ 0.2 s including guard cold start
- Assertions + report, Attribute false-failure reproduction, `MsgBox` blocked before injection
- The bundled template: **8/8 assertions pass in 0.47 s**

**Not verified** (needs your environment):

- Loading Skills / prompt files / MCP inside VS Code
- Excel versions other than Microsoft 365 desktop

---

## FAQ

| Symptom | Cause / fix |
|---|---|
| "Cannot run the macro… macros may be disabled" | Two **unrelated** causes share this exact text: (a) the code contains an `Attribute` line — `run_vba.py` strips it, don't hand-write it, and it is **not** a Trust Center problem; (b) the workbook window is **saved hidden** (normal for loader / host workbooks — by design, don't "fix" the file) → Excel has no active workbook, so bare macro names cannot resolve. `run_vba.py` auto-qualifies; `vba_diagnose.py` tells the two apart |
| `⛔ 拒绝注入` (injection refused) | Code contains `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / `.Show` — all of which hang an automated instance. Remove them or pass `--allow-unsafe` |
| Guard killed Excel mid-run | The workbook was saved before the run — just run again; the script reopens it |
| Don't want to approve the command every time | VS Code setting `chat.tools.terminal.autoApprove` for `python tools/vba/*` |
| Skill not taking effect | Directory name must equal `SKILL.md`'s `name`; check the skill settings are enabled; use Chat **Diagnostics** |
| CI runs `tools/validate.py` | It fails on bad skill frontmatter and on personal paths / key prefixes. Add private terms via the `VBA_KIT_DENYLIST` env var (comma-separated) — never in the repo |

---

## License

MIT — see [LICENSE](LICENSE).
