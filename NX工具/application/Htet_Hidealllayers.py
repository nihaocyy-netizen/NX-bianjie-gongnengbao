# -*- coding: utf-8 -*-
# NX 2506 插件：关闭除工作层以外的所有图层（工作层不能隐藏，保留不动）
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import NXOpen.Layer
import Htet_common as c
import Htet_lang as _L


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return

    mark = session.SetUndoMark(c.NXOpen_MarkVisibility_Visible(),
                               _L.T("undo.hide_layers"))
    try:
        work_layer = part.Layers.WorkLayer
        for number in range(1, 257):
            if number != work_layer:
                part.Layers.SetState(number, NXOpen.Layer.State.Hidden, False)
    except Exception as ex:
        session.UndoToMark(mark, "")
        lw.WriteLine(_L.T("msg.layer_op_failed", str(ex)))
        c.show(ui, _L.T("ui.error"), _L.T("msg.layer_op_failed", str(ex)),
               is_error=True)


if __name__ == '__main__':
    main()
