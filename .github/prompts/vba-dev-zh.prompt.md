---
description: '开发/修改一个 Excel VBA 宏并自动跑测试（写→测→修循环，中文版）'
argument-hint: '要做什么宏？目标工作簿路径 + 期望结果（如：汇总各公司收入并加透视表）'
agent: 'agent'
---

# 开发一个 VBA 宏并自动测试

按下面流程做，全程用终端工具，**不要**手写 win32com 注入步骤。

## 1. 确认输入

用 `#tool:vscode/askQuestions` 问清（缺什么问什么，别猜）：

- 目标工作簿路径（`.xlsm`，不存在则新建）
- 输出要求：建哪些表 / 哪些字段 / 什么口径 / 要不要透视表·图表
- 可验证的期望值（哪怕一两个：某个单元格应等于多少、应出现哪张表）

## 2. 写代码

- 写到 `macro/<模块名>.bas`（**不带 `Attribute` 行**）
- 每个宏带 `On Error GoTo EH` 兜底：把 `Err.Number & ": " & Err.Description` 写进单元格 + 追加日志文件
- 性能：数组一次性写回、关 `ScreenUpdating`/`Calculation`、不逐格循环
- 参考 `.github/skills/excel-vba-automation/templates/vba_summary_report.bas.txt`（通用多维汇总模板）

## 3. 跑测试

```bash
python tools/vba/run_vba.py --workbook "<工作簿路径>" --code <模块名>=macro/<模块名>.bas \
  --run <宏名> --expect "cell:<表>!<单元格>=<期望值>" [更多 --expect] --keep-open --save
```

读报告里的三样东西：**判定**、**错因**、**断言逐条**。

## 4. 按错因修，最多 3 轮

- 判定是 `vba_error_dialog*`：看 `错因`（就是 VBA 原文），定位到具体语句改
- 判定是 `idle_hung` / `runaway_cpu`：宏里有死循环或等待外部资源，改逻辑（必要时加超时/退出条件）
- `⛔ 拒绝注入`：删掉 `MsgBox`/`Stop`/`Debug.Assert` 之类
- 断言失败：报告给了实际值 + 上下文（如现有工作表列表）→ 对照期望值判断是代码问题还是口径问题
- 每轮改完重跑；**复用一个 Excel 实例**（`--keep-open`）保持 1–3 秒一轮

## 5. 收尾

- 全绿（`exit 0`）后向用户汇报：做了什么宏、测试报告结论、关键数字、工作簿路径
- 3 轮还不过：把**原文错误 + 已尝试的改动**交给人，不要盲目继续
- 口径类问题（数字对不对）**必须让用户确认**，不要自行认定

English version: `/vba-dev`
