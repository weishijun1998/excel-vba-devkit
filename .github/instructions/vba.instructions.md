---
name: 'VBA 宏开发铁律'
description: '开发/调试 Excel VBA 宏时必须遵守的规则（踩过事故的每一条）'
applyTo: ['**/*.bas', '**/*.vba', '**/vba/**/*.py', 'tools/vba/**']
---

# VBA 宏开发铁律

- **注入前剥掉所有 `Attribute` 行**：它是 VBE 的元数据，只在导出的 `.bas` 里合法；带它注入会让模块编译失败并拖垮整个工程，报 `0x800A03EC` +"宏可能被禁用"（**不是**宏安全设置问题）。坏模块必须就地覆写或删除。
- **代码里禁止** `MsgBox`、`InputBox`、`Stop`、`Debug.Assert`、UserForm `.Show` —— 都会让自动化实例挂死（`Stop`/`Debug.Assert` 是无弹窗假死）。
- **每个宏加兜底**：`On Error GoTo EH`，EH 把 `Err.Number & ": " & Err.Description` 写单元格 + 追加日志文件（读取用 `encoding="gbk"`）。
- **跑宏只走 `python tools/vba/run_vba.py`**：它会剥 Attribute、拦危险语句、起守卫、强制重算、跑断言、出报告；不要手写 win32com 注入步骤。
- **跑完必须验证**：用 `--expect` 断言关键单元格 / 表 / 透视表 / 命名区域；"没报错"不等于"算对了"。
- **性能**：数据在数组里算好一次性写回；关 `ScreenUpdating`/`Calculation`；不逐格循环（单次写入 64 µs）；步骤标记只放阶段边界，不放循环体。
- **失败先看报告再改**：报告里已有判定（正常/报错弹窗/假死/跑飞）+ 错因 + 断言逐条，不要凭猜重跑。
