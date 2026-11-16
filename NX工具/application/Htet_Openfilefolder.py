# -*- coding: utf-8 -*-
# NX 2506 插件：在资源管理器中打开当前工作零件所在文件夹
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return

    file_path = part.FullPath
    if not file_path:
        lw.WriteLine(_L.T("msg.no_saved_part"))
        c.show(ui, _L.T("ui.info"), _L.T("msg.no_saved_part"))
        return
    # Teamcenter 管理模式中的零件标识不一定对应本地文件路径。
    if not os.path.isabs(file_path):
        lw.WriteLine(_L.T("msg.no_local_path", file_path))
        c.show(ui, _L.T("ui.info"), _L.T("msg.no_local_path", file_path))
        return
    folder_path = os.path.dirname(file_path)
    if not os.path.isdir(folder_path):
        lw.WriteLine(_L.T("msg.folder_missing", folder_path))
        c.show(ui, _L.T("ui.error"), _L.T("msg.folder_missing", folder_path),
               is_error=True)
        return

    try:
        os.startfile(folder_path)
    except OSError as ex:
        message = _L.T("msg.open_folder_failed", folder_path) + "\n" + str(ex)
        lw.WriteLine(message)
        c.show(ui, _L.T("ui.error"), message, is_error=True)


if __name__ == '__main__':
    main()
