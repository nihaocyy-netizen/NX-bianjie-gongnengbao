# -*- coding: utf-8 -*-
# NX 2506 插件：打开计算器（优先插件 exefile 目录，找不到则用系统计算器）
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

CALC_NAMES = ("计算器.exe", "calculator.exe", "calc.exe")


def find_calculator():
    """优先使用插件根 exefile 目录里的计算器，否则回落到系统 calc.exe。"""
    plugin_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    exe_dir = os.path.join(plugin_root, "exefile")
    for name in CALC_NAMES:
        path = os.path.join(exe_dir, name)
        if os.path.isfile(path):
            return path
    return "calc.exe"


def main():
    session, ui, part, lw = c.ctx()
    try:
        os.startfile(find_calculator())
    except Exception as ex:
        message = _L.T("msg.calculator_failed", str(ex))
        lw.WriteLine(message)
        c.show(ui, _L.T("ui.error"), message, is_error=True)


if __name__ == '__main__':
    main()
