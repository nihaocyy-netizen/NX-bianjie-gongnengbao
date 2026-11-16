# -*- coding: utf-8 -*-
# NX 2506 插件：工作图层切换到上一个有物体图层（层号变小，自动跳过空层，不循环）
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return

    current = part.Layers.WorkLayer
    target = None
    for number in range(current - 1, 0, -1):
        if c.layer_has_objects(part, number):
            target = number
            break
    if target is None:
        c.notice(lw, _L.T("msg.no_layer_target", current))
        return

    mark = session.SetUndoMark(c.NXOpen_MarkVisibility_Visible(),
                               _L.T("undo.prev_layer"))
    try:
        part.Layers.WorkLayer = target
    except Exception as ex:
        session.UndoToMark(mark, "")
        lw.WriteLine(_L.T("msg.layer_op_failed", str(ex)))
        c.show(ui, _L.T("ui.error"), _L.T("msg.layer_op_failed", str(ex)),
               is_error=True)


if __name__ == '__main__':
    main()
