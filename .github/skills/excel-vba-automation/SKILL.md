---
name: excel-vba-automation
description: Use when 开发/调试/批量运行 Excel VBA 宏（写宏、修编译错、跑测试、处理报错弹窗或卡死、批量刷新工作簿）。
allowed-tools: shell
---

# Excel VBA 宏开发与运行（自动化流水线）

用 **`tools/vba/run_vba.py`** 作为唯一入口跑宏，**不要**手写 win32com/pywin32 的注入步骤 —— 那会重复踩下面这些坑。

```bash
python tools/vba/run_vba.py --workbook <x.xlsm> --code <模块名=文件> --run <宏名> \
  --expect "cell:表!A1=值" [--expect ...] [--save] [--keep-open] [--visible]
```

## 铁律（每条都对应一次实机事故）

1. **注入前剥掉所有 `Attribute` 行**。`Attribute VB_Name = "X"` 是 VBE 自己的元数据，只在导出的 `.bas` 里合法；当源码注入会让该模块编译失败，**并且拖垮整个工程**（旁边新加的干净模块也跑不了），Excel 报 `0x800A03EC` + "宏可能被禁用" —— 极易误判成宏安全设置。`run_vba.py` 会自动剥；手写注入时要自己剥，且**坏模块必须就地覆写或删除**，旁边加新模块无效。
2. **代码里禁止 `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / UserForm `.Show`**。前两者和后三者都会让自动化实例**挂死**（`Stop`/`Debug.Assert` 最隐蔽：把 VBE 拉进断点模式，**无弹窗、纯假死**）。`run_vba.py` 注入前静态拦截。
3. **每个宏都加兜底**：`On Error GoTo EH`，EH 里把 `Err.Number & ": " & Err.Description` 写进单元格，并追加一行到日志文件（`Open ... For Append`）。这是"没有弹窗时"唯一的错因来源。读日志文件用 `encoding="gbk"`（ANSI 写入）。
4. **运行前保存**（`run_vba.py` 默认做）。这样守卫判定卡死而杀进程时，只损失一次重开会话，不是未保存内容全丢。
5. **性能**：数据先在数组里算好、**一次性写回区域**；关 `ScreenUpdating`、`Calculation = xlCalculationManual`（结束前恢复）；能靠公式/SUMIF/透视表的不用 VBA 循环。
   - 实测：单次单元格写入 **64 µs**、文件追加 **359 µs**。**逐格循环是宏变慢的头号原因**（10 万次 = 6.4 秒）；VBA 纯计算 300 万次只要 49 ms。
   - 进度/步骤标记只放在**阶段边界**（≤10 处），**绝不放进循环体**。
6. **宏跑完不等于跑对**：必须写断言回读关键单元格/表/透视表/命名区域。断言数字按数值比较；公式值要在**强制重算后**读（`run_vba.py` 已处理）。

## 失败判定矩阵（`vba_guard.py`）

| 现象 | 判据 | 处置 | 实测 |
|---|---|---|---|
| 报错弹窗 | `#32770` 且标题含 `Visual Basic` | 点 id **4800**"结束"按钮（**不要 WM_CLOSE，无效**） | 出现→消失 12–46 ms，Excel 存活 |
| 假死等 IO | 消息泵无响应 + 近 2 秒 CPU≈0 | 超 `hang_after`（默认 6s）杀 | 6.6 s 判 `IDLE_HUNG` |
| 死循环跑飞 | 消息泵无响应 + CPU 持续 >0.3s/2s | 超硬预算杀 | 20.8 s 判 `RUNAWAY_CPU` |
| 长任务在推进 | 心跳文件新鲜 | **不杀**，继续等 | 16 s 宏零误杀 |
| 慢但泵消息 | 活性探测有响应 | 不杀 | — |

弹窗里读错因：`Static` 控件 **id 4803** 就是错误原文（如 `运行时错误 '9': 下标越界`）。按钮控件 ID：**4800=结束、4801=调试、4802=继续（通常灰）、4902=帮助**。

## 错误归因顺序

1. 守卫抓到的弹窗文本（`ERRTEXT`）→ 最准
2. 宏内兜底写的单元格 / 日志文件 → 无弹窗时
3. `com_error` 的 hresult：`0x800A03EC`=Attribute 行；`0x800A9C68`=编译错；`0x800706BE`(RPC_S_CALL_FAILED)=**守卫刚杀的进程**，不是宏的错
4. 静态检查（注入前就拦）
5. 断言失败时的实际值 + 上下文（例：`sheet:差异分析=存在 ❌ (现有工作表: ['月度数据','Sheet1'])` —— 一眼看出表名写错）

## 开发循环（写 → 测 → 修）

1. 写模块到 `*.bas`/`*.txt`（不带 `Attribute` 行）
2. `run_vba.py` 跑：拿到 **判定 + 错因 + 断言逐条**
3. 按错因改代码（`--keep-open` 复用已开的 Excel 实例，每轮 1–3 秒）
4. 重复至 `exit 0`；**最多 3 轮**，第 4 轮还不过就把原文错误 + 改动清单交给人

## 参考

- `templates/vba_summary_report.bas.txt` — 通用多维汇总模板（宏名 `BuildSummaryReport`）：3 张表 / 造数 / 格式 + 3 色阶条件格式 / 自动筛选 / 冻结窗格 / SUMIF / 透视表 / 柱状图 / 对账校验 / 命名区域（实测 0.73 s）
- `scripts/` 在本 kit 的 `tools/vba/`（`run_vba.py`、`vba_guard.py`、`dismiss_vba_dialog.py`、`vba_attr_probe.py`）
- 路径一律用 `/`，保持跨平台可移植
