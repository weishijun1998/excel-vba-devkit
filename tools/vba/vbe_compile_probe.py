# -*- coding: utf-8 -*-
"""vbe_compile_probe.py —— 两个"看起来像编译错误"的假象，实机复现（不需要外部参数）

复现两件事（都用真实 Excel，临时工作簿，跑完即删）：

  ① 网上通行的程序化编译写法 `Application.VBE.CommandBars.FindControl(Id:=578).Execute()`
     **在本机中文 VBE 上彻底无效**：Id 578 拿回来的是"文件(&F)"菜单，扫遍 VBE 菜单也
     找不到任何编译命令，`Execute()` 秒返回、不报错、什么都不编译 —— 于是"检查"会给
     出一个"编译通过"的假结论。**比不检查更坏**，因为它让人以为工程是干净的。

  ② **过程名与模块名相同时，宏无法被 Run 解析**。模块 `SameName` 里的 `Sub SameName()`
     用 `SameName` / `book.xlsm!SameName` / `'book.xlsm'!SameName` **三种写法全部失败**，
     报的还是「无法运行"X"宏…可能所有宏都被禁用」—— 与"工程编译不过"**一字不差**。
     所以 vba_project.py 的编译探针刻意让模块名 ≠ 过程名。

这两条合起来解释了为什么"工程编译不过"必须靠**真跑一次探针**来判定，而不能靠菜单命令。

用法：python tools/vba/vbe_compile_probe.py
"""
import glob
import os
import sys
import tempfile

try:
    import win32com.client as w32
except ImportError:
    print("需要 pywin32")
    sys.exit(2)

MOD, PROC = "CompileProbeMod", "CompileProbeRun"


def _setup_stdio():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def brief(e):
    return str(e).replace("\n", " ")[:110]


def main():
    tmp = tempfile.mkdtemp(prefix="vbe_compile_probe_")
    wb_path = os.path.join(tmp, "probe.xlsm")
    xl = w32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    wb = None
    try:
        wb = xl.Workbooks.Add()
        wb.SaveAs(wb_path, FileFormat=52)
        vbe = xl.VBE

        print("=" * 68)
        print("[1] FindControl(Id:=578) 到底拿到了什么")
        c578 = vbe.CommandBars.FindControl(Id=578)
        if c578 is None:
            print("    578 -> None")
        else:
            try:
                print("    578 -> Id=%s Caption=%r Enabled=%s Type=%s"
                      % (c578.Id, c578.Caption, c578.Enabled, c578.Type))
            except Exception as e:
                print("    578 -> 属性读取失败 %s" % brief(e))
        hits = []
        cbar = vbe.CommandBars
        for i in range(1, cbar.Count + 1):
            try:
                cb, n = cbar.Item(i), 0
                n = cb.Controls.Count
            except Exception:
                continue
            for j in range(1, n + 1):
                try:
                    c = cb.Controls.Item(j)
                    if "ompile" in c.Caption or "编译" in c.Caption:
                        hits.append((cb.Name, c.Id, c.Caption))
                except Exception:
                    pass
        print("    VBE 菜单里含 compile/编译 的控件: %s" % (hits[:8] or "（一个都没有）"))
        print("    => 结论：%s"
              % ("578 不是编译命令，用它做编译检查等于没检查"
                 if not hits else "本机 VBE 里有编译控件，可用"))

        print("=" * 68)
        print("[2] 模块名 ≠ 过程名：Run 应当成功")
        vbp = wb.VBProject
        comp = vbp.VBComponents.Add(1)
        comp.Name = MOD
        comp.CodeModule.AddFromString("Sub %s()\r\nEnd Sub\r\n" % PROC)
        for spec in ["%s.%s" % (MOD, PROC), "%s!%s" % (wb.Name, PROC),
                     "'%s'!%s" % (wb.Name, PROC)]:
            try:
                xl.Run(spec)
                print("    OK   %-34s" % spec)
            except Exception as e:
                print("    FAIL %-34s %s" % (spec, brief(e)))

        print("=" * 68)
        print("[3] 过程名 == 模块名：Run 应当全部失败（假象来源）")
        same = vbp.VBComponents.Add(1)
        same.Name = "SameName"
        same.CodeModule.AddFromString("Sub SameName()\r\nEnd Sub\r\n")
        for spec in ["SameName", "%s!SameName" % wb.Name, "'%s'!SameName" % wb.Name]:
            try:
                xl.Run(spec)
                print("    OK   %-34s  <- 与预期不符！" % spec)
            except Exception as e:
                print("    FAIL %-34s %s" % (spec, brief(e)))
        print("    => 结论：名字歧义的失败文案与『工程编译不过』一字不差，")
        print("       所以探针的模块名与过程名必须不同。")

        print("=" * 68)
        print("[4] 真编译检查的正确做法：注入空探针 → Run")
        try:
            xl.Run("%s!%s" % (wb.Name, PROC))
            print("    OK   探针跑通 => 整个 VBAProject 可编译（VBA 编译是工程级的）")
        except Exception as e:
            print("    FAIL 探针跑不通 => 工程编译不过 / 宏不可用：%s" % brief(e))
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass
        for f in glob.glob(os.path.join(tmp, "*")):
            try:
                os.remove(f)
            except OSError:
                pass
        try:
            os.rmdir(tmp)
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    _setup_stdio()
    sys.exit(main())
