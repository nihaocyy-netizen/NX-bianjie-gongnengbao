# -*- coding: utf-8 -*-
# NX 2506 插件：启动公差查询手册 GongCha.exe
# 只认插件根 exefile\GongCha.exe；缺失时明确报错，不猜测其他路径。
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

EXE_NAME = "GongCha.exe"


def gongcha_path():
    plugin_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(plugin_root, "exefile", EXE_NAME)


def main():
    session, ui, part, lw = c.ctx()
    exe_path = gongcha_path()
    if not os.path.isfile(exe_path):
        message = _L.T("msg.tolerance_missing", exe_path)
        lw.WriteLine(message)
        c.show(ui, _L.T("ui.error"), message, is_error=True)
        return
    try:
        os.startfile(exe_path)
    except Exception as ex:
        message = _L.T("msg.tolerance_failed", str(ex))
        lw.WriteLine(message)
        c.show(ui, _L.T("ui.error"), message, is_error=True)


if __name__ == '__main__':
    main()
