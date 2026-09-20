# VBA Copilot Kit

把「Hermes 里实测验证过的 Excel VBA 宏开发流水线」搬到 **VS Code + GitHub Copilot**，让 AI 写宏这件事变得**快、可自证、不会挂死**。

一句话：**Copilot 写宏 → 自动跑测试 → 出报告（判定 + 错因 + 断言逐条）→ 按报告修 → 再跑**。

---

## 三步安装

1. **解压到你的仓库根目录**（会得到 `tools/`、`.github/`、`.vscode/` 三处）
2. **装依赖**：`pip install -r tools/vba/requirements.txt`（只需 `pywin32` + `psutil`；Windows + 本机已装 Excel）
3. **在 VS Code 里确认装载**：Chat 输入 `/skills` 应能看到 `excel-vba-automation`；
   若装了 prompt file，输入 `/` 会看到 `/vba-dev`

> 技能目录名必须与 `SKILL.md` 里 `name` 一致；不要在名字里加 `org/` 前缀（会静默加载失败）。

---

## 三种用法

### A. 让 Copilot 自主开发（推荐）

在 Copilot Chat 的 **agent 模式**里说：

> 用 excel-vba-automation 技能帮我写一个宏：把"月度数据"表按公司汇总，加透视表和柱状图，并跑测试验证。

Copilot 会按 SKILL.md 的流程走：写代码 → `python tools/vba/run_vba.py ...` → 读报告 → 修 → 重跑。

### B. 你自己在终端跑（不依赖 Copilot）

```bash
python tools/vba/run_vba.py \
  --workbook "C:/work/report.xlsm" \
  --code ModReport=macro/report.bas \
  --run BuildReport \
  --expect "cell:汇总!B2=623685.6" --expect "pivot:汇总=1" --expect "named:收入数据=存在" \
  --save --keep-open
```

退出码：`0` 全过 ｜ `1` 有失败 ｜ `2` 被拦（危险代码）

### C. 单独处理"卡住的弹窗 / attribute 坑"

```bash
python tools/vba/dismiss_vba_dialog.py        # 一键清掉卡住的 VBA 报错弹窗（四种关法逐级降级）
python tools/vba/vba_attr_probe.py            # 复现/验证 Attribute 行导致的"宏不可用"
python tools/vba/vba_guard.py <excel_pid> 60  # 单独挂守卫看一个长任务
```

---

## 包含什么

| 路径 | 作用 |
|---|---|
| `tools/vba/run_vba.py` | **开发测试跑道**：剥 Attribute → 危险语句拦截 → 注入 → 守卫 → 运行 → 强制重算 → 断言 → 报告 |
| `tools/vba/vba_guard.py` | **守卫**：弹窗点掉（点 id 4800"结束"）/ 假死判定 / 跑飞判定 / 心跳保护 |
| `tools/vba/dismiss_vba_dialog.py` | 卡住弹窗一键清理（BM_CLICK → WM_COMMAND → 真实鼠标 → 杀进程） |
| `tools/vba/vba_attr_probe.py` | Attribute 行探针（复现 0x800A03EC 假故障） |
| `.github/skills/excel-vba-automation/SKILL.md` | **技能**：整套流程、铁律、失败矩阵（Agent Skills 开放标准） |
| `.github/skills/excel-vba-automation/templates/` | 复杂报表宏模板（建表/造数/格式/透视表/图表/校验） |
| `.github/instructions/vba.instructions.md` | 按文件类型生效的铁律（`applyTo` 指向 `.bas`/`.vba`/`tools/vba/**`） |
| `.github/prompts/vba-dev.prompt.md` | `/vba-dev` 一键进入"写宏→测试→修"循环 |
| `.vscode/mcp.json` | （可选）把 run_vba 变成 Copilot 的一等工具 |

---

## 已验证 / 未验证（诚实说明）

**已实测验证**（Win11 + Excel + Python 3.11，本机）：

- 脚本行为：四轮开发流程（报错→修→全绿）、无弹窗死循环判定（12.7s 判跑飞并处置）、假死判定（6.6s 判 IDLE_HUNG）、心跳保护（16s 长任务零误杀）
- 弹窗反应速度：出现→消失 12–46 ms；点掉本身 ~0.2s（含冷启动）
- 断言与报告、Attribute 假故障复现、`MsgBox` 拦截

**未验证**（需要你的 VS Code + Copilot 环境）：

- Skills / prompt file / MCP 在 VS Code 里的装载与调用
- 不同 Excel 版本（本机为 Microsoft 365 桌面版）

---

## 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| 报"无法运行宏…可能宏被禁用" | 代码里有 `Attribute` 行 —— 脚本会自动剥；手写时别带 |
| 注入被拒 `⛔ 拒绝注入` | 代码含 `MsgBox`/`InputBox`/`Stop`/`Debug.Assert`/`.Show`（会挂死），改掉或 `--allow-unsafe` |
| Excel 被守卫杀了怎么继续 | 运行前已保存工作簿，直接再跑一遍即可（脚本会重新打开） |
| 不想每次点"允许运行命令" | VS Code 设置 `chat.tools.terminal.autoApprove` 放行 `python tools/vba/*` |
| 技能没生效 | 目录名 = `SKILL.md` 的 `name`；确认 `chat.useAgentsMdFile`/技能设置已启用；用 Chat 的 **Diagnostics** 排查 |

---

## 为什么不用 VBA 调试器 / 手工 Alt+F8

因为自动化跑宏的失败模式是**弹窗阻塞**（默认拖到 120 秒超时清场、未保存内容全丢）和**无弹窗假死**（只能干等）。
这套工具把两类都变成**亚秒级发现 + 分级处置 + 结构化报告**，所以 AI 才能自己迭代而不需要你在旁边点按钮。
