# Excel VBA 宏开发与运行（自动化流水线）· 中文完整版

> 英文主文件在 `../SKILL.md`；这里是同一套内容的中文完整版。

用 **`tools/vba/run_vba.py`** 作为唯一入口运行宏，**不要**手写 win32com/pywin32 注入步骤 —— 那会重复引入下列全部失败模式。

```bash
python tools/vba/run_vba.py --workbook <x.xlsm> --code <模块名=文件> --run <宏名> \
  --expect "cell:表!A1=值" [--expect ...] [--save] [--keep-open] [--visible] [--allow-events] \
  [--cleanup] [--encoding gbk] [--ascii]
```

打开工作簿时**默认禁用事件**（`EnableEvents=False`）：装载器/宿主工作簿的 `Workbook_Open` 不会被顺带跑起来；确实想让它触发时加 `--allow-events`。`vba_diagnose.py` 恒禁用事件 —— 诊断绝不跑用户的程序。

**源码编码自动识别**（BOM → UTF-8 → 系统 ANSI 代码页 → cp936/cp932/cp949/cp1252）：VBA 导出文件在新环境是 UTF-8，**老式中文环境是 GBK/CP936**；`--encoding` 可手动指定。**不加 `--save` 绝不写盘**；`--cleanup` 会在结束前移除注入的测试模块（处理正式工作簿建议加）。终端显示异常时加 `--ascii`。

## 铁律（每条都对应一次实机事故）

1. **注入前剥掉所有 `Attribute` 行**。`Attribute VB_Name = "X"` 是 VBE 自己的元数据，只在导出的 `.bas` 里合法；当源码注入会使该模块编译失败，**并使整个工程编译失败**（旁边新增的干净模块同样无法运行），Excel 报 `0x800A03EC` + "宏可能被禁用" —— 容易被误判为宏安全设置。`run_vba.py` 自动剥离；手写注入时须自行剥离，且**坏模块必须就地覆写或删除**，旁边新增模块无效。
   - ⚠️ **对称风险：修改「已有业务模块」**：导出 → 剥 Attribute → 整模块回写会**连带丢失过程级属性** —— 例如 `Attribute 过程名.VB_ProcData.VB_Invoke_Func = "P\n14"` 即**快捷键绑定**。实测一个 1874 行的生产模块经此改造后，该宏的快捷键静默失效（只验证了运行结果，未发现该问题）。规则：**剥 Attribute 只针对要新注入的代码**；修改已有模块前先导出 diff，把 `VB_ProcData.*` 这类过程属性行原样保留。恢复方法：把该行插回导出的 `.bas` → `VBComponents.Remove` → `Import`（**Import 接受 Attribute 行，只有 `AddFromString` 会失败**），文件必须按系统 ANSI(cp936) 写入，否则中文注释全部损坏。
2. **工作簿窗口隐藏是装载器/宿主工作簿的有意设计，不是故障，但它会输出完全相同的报错文本**。装载器/宿主类工作簿（打开即运行程序，避免用户被工作簿干扰）通常保存为 `wb.Windows(1).Visible = False`。这时 Excel **没有活动工作簿**（`xl.ActiveWorkbook is None`），而 `Run "裸宏名"` 是按"活动工作簿"解析的 → 必然失败，报文与铁律 1 **一字不差**（`0x800A03EC` +「无法运行…可能所有宏都被禁用」）。**不要为迁就自动化而改造该文件**：自动化侧一律用限定名 `"工作簿名!宏名"`。`run_vba.py` 会检测隐藏态并自动限定；裸名失败时还会自动重试限定名。任意文件可用 `tools/vba/vba_diagnose.py` 检查（静态检查，不需要开 Excel）。该状态**会通过「另存为」传播** —— 隐藏状态下另存为会使所有副本同样隐藏。
3. **宏**内部**的每一次对象引用都要限定 —— 同一成因的第二种表现（出现在运行中途）。** `Sheets("X")`、`Range("A1")`、`Cells(...)`、`ActiveSheet`、`ActiveWorkbook` 都要经 VBA 的 `_Global` 解析，而 `_Global` 需要**活动的**工作簿/工作表。没有活动工作簿（隐藏窗口的装载器）时：`Sheets("X")` 报运行时错误 **1004「方法 'Sheets' 作用于对象 '_Global' 时失败」**；`Set ws = ActiveSheet` **不报错**但得到 `Nothing`，错误延迟到**下一行真正使用 ws 时**才出现（**91**）。手工运行时永远正常 —— 这正是这类缺陷能潜伏多年的原因，而**任何自动化调用都会失败**。一律写 `ThisWorkbook.Worksheets("X")` 或显式的工作簿/工作表变量；整仓扫描用 `python tools/vba/vba_lint_refs.py <目录或文件>`（静态只读，支持 GBK/UTF-8 源码）。**同一成因的第三种表现：`Application.Calculation`。** 同一文件 A/B 实测（`python tools/vba/hidden_app_props_probe.py`）：窗口隐藏时 `ScreenUpdating`/`DisplayAlerts`/`StatusBar`/`EnableEvents`/限定写法写单元格**全部正常**，**只有 `Application.Calculation = …`（读或写）报 1004「应用程序定义或对象定义错误」**。它只是性能开关、不影响计算结果：读写一律用 `On Error Resume Next` 包裹并按「自动」处理，**避免一个性能开关中断整个宏**（`run_vba.py` 断言前仍会强制 `CalculateFull`，不影响判定）。
4. **代码里禁止 `MsgBox` / `InputBox` / `Stop` / `Debug.Assert` / UserForm `.Show`**。都会使自动化实例**阻塞**（`Stop`/`Debug.Assert` 最难检测：将 VBE 置入断点模式，**无弹窗、纯阻塞**）。`run_vba.py` 注入前静态拦截。
5. **每个宏都加错误处理**：`On Error GoTo EH`，EH 里把 `Err.Number & ": " & Err.Description` 写进单元格，并追加一行到日志文件（`Open ... For Append`）。这是"没有弹窗时"唯一的错因来源。读日志文件用 `encoding="gbk"`（ANSI 写入）。
6. **运行前保存**（`run_vba.py` 默认做）。这样守卫判定卡死而杀进程时，只损失一次重开会话，而不是未保存内容全部丢失。
7. **性能**：数据先在数组里算好、**一次性写回区域**；关 `ScreenUpdating`、`Calculation = xlCalculationManual`（结束前恢复）；能靠公式/SUMIF/透视表的不用 VBA 循环。
   - 实测：单次单元格写入 **64 µs**、文件追加 **359 µs**。**逐格循环是宏变慢的首要原因**（10 万次 ≈ 6.4 秒；真实工作簿里 4 万格逐格写 = **2058 ms**），而 VBA 纯计算 300 万次只要 49 ms。同样 4 万格改成数组一次性写回 = **42.7 ms → 48× 提速**，且校验和与闭式解完全一致（属于真实提速，而非减少工作量）。
   - 进度/步骤标记只放在**阶段边界**（≤10 处），**不放入循环体**。
8. **宏跑完不等于跑对**：必须写断言回读关键单元格/表/透视表/命名区域。断言数字按数值比较；公式值要在**强制重算后**读取（`run_vba.py` 已处理）。

## 失败判定矩阵（`vba_guard.py`）

| 现象 | 判据 | 处置 | 实测 |
|---|---|---|---|
| 报错弹窗 | `#32770` 且标题含 `Visual Basic` | 点 id **4800**"结束"按钮（**不要 WM_CLOSE，无效**） | 出现→消失 12–46 ms，Excel 存活 |
| 「无法运行"X"宏…」**两种不相干的病因** | (a) 某模块含 `Attribute` 行 / 工程编译不过 | (a) 剥掉；坏模块**就地覆写** | 用 `vba_attr_probe.py` 复现 |
| ↑ 同一句话 | (b) 工作簿窗口保存时**就是隐藏的** → `ActiveWorkbook is None` | (b) **不要动这个文件**，用限定名 `"簿名!宏名"`；`run_vba.py` 已自动处理 | 裸名失败 / 限定名 0.016 s 通过；静态检查不用开 Excel |
| 运行时错误 **1004**（方法 'Sheets' 作用于对象 '_Global' 时失败）或 **91**，**出现在宏运行中途** | 宏里用了未限定的引用，且当时没有活动工作簿（窗口隐藏） | 一律限定：`ThisWorkbook.Worksheets(...)`；用 `vba_lint_refs.py` 列出所有问题点 | 实测：`Sheets("S1")` → 1004；`Set ws = ActiveSheet` → 拿到 `Nothing`（不报错），**下一行**才失败 |
| 假死等 IO | 消息泵无响应 + 近 2 秒 CPU≈0 | 超 `hang_after`（默认 6s）杀 | 6.6 s 判 `IDLE_HUNG` |
| 死循环跑飞 | 消息泵无响应 + CPU 持续 >0.3s/2s | 超硬预算杀 | 20.8 s 判 `RUNAWAY_CPU` |
| 长任务在推进 | 心跳文件新鲜 | **不杀**，继续等 | 16 s 宏零误杀 |
| 慢但泵消息 | 活性探测有响应 | 不杀 | — |

弹窗里读错因：`Static` 控件 **id 4803** 就是错误原文（如 `运行时错误 '9': 下标越界`）。按钮控件 ID：**4800=结束、4801=调试、4802=继续（通常灰）、4902=帮助**。

## 错误归因顺序

1. 守卫抓到的弹窗文本（`ERRTEXT`）→ 最精确
2. 宏内错误处理写入的单元格 / 日志文件 → 无弹窗时
3. `com_error` 的 hresult —— **不可只看错误码：同一故障会以不同错误码出现**：`0x800A03EC`（VBA「无法运行宏」）/ 外层包裹的 `0x80020009`（pywin32 的 `DISP_E_EXCEPTION`）= 可能是 `Attribute` 行**也可能是**窗口隐藏的工作簿 → 用失败矩阵中两条检查区分；`0x800A9C68`=编译错；`0x800706BE`(RPC_S_CALL_FAILED)=**守卫刚杀的进程**，不是宏的错
4. 静态检查（注入前就拦）+ `vba_diagnose.py`（窗口 / 引用 / Attribute 行）
5. 断言失败时的实际值 + 上下文（例：`sheet:差异分析=存在 ❌ (现有工作表: ['月度数据','Sheet1'])` —— 可直接判定表名写错）

**病因不明时的决定性检查**：把这个宏拿到**全新空工作簿**里运行一遍。那边全绿 ⇒ 工具链和守卫都没问题，问题在**那个文件**（窗口隐藏 / 工程 / 策略）；那边也失败 ⇒ 问题在你的代码路径或宏名本身。

## 开发循环（写 → 测 → 修）

1. 写模块到 `*.bas`/`*.txt`（不带 `Attribute` 行）
2. `run_vba.py` 跑：拿到 **判定 + 错因 + 断言逐条**
3. 按错因改代码（`--keep-open` 复用已开的 Excel 实例，每轮 1–3 秒）
4. 重复至 `exit 0`；**最多 3 轮**，第 4 轮还不过就把原文错误 + 改动清单交给人

## 参考

- `templates/vba_summary_report.bas.txt` — 通用多维汇总模板（宏名 `BuildSummaryReport`）：3 张表 / 造数 / 格式 + 3 色阶条件格式 / 自动筛选 / 冻结窗格 / SUMIF / 透视表 / 柱状图 / 对账校验 / 命名区域（实测 0.47 s，8/8 断言通过）
- 脚本在本 kit 的 `tools/vba/`：`run_vba.py`（开发测试跑道：隐藏窗口装载器自动补限定名、源码编码自动识别、仅 `--save` 时写盘、`--cleanup` 移除注入模块）、`vba_guard.py`（弹窗 / 卡死守卫）、`vba_diagnose.py`（只读诊断）、`vba_lint_refs.py`（未限定引用静态扫描）、`scan_hidden_windows.py`（隐藏窗口静态巡检）、`bare_name_scope_probe.py`（宏名解析边界探针）、`hidden_app_props_probe.py`（隐藏窗口 Application 设置探针）、`dismiss_vba_dialog.py`、`vba_attr_probe.py`
- 路径一律用 `/`，保持跨平台可移植
