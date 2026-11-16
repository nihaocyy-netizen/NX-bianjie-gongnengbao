# -*- coding: utf-8 -*-
# NX 2506 插件：把选定对象移到第 255 层，并把该层修正为可见可选
# 未选中任何对象时，先让用户原生选择；仍为空则确认后对当前零件的全部实体执行。
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

TARGET_LAYER = 255


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return
    targets = c.resolve_targets(ui, part, lw, solid_only=False)
    if targets is None:
        return
    if not targets:
        c.notice(lw, _L.T("msg.no_objects_move"))
        return

    mark = session.SetUndoMark(c.NXOpen_MarkVisibility_Visible(),
                               _L.T("undo.move255"))
    try:
        c.apply_layer_to(session, targets, TARGET_LAYER)
    except Exception as ex:
        session.UndoToMark(mark, "")
        lw.WriteLine(_L.T("msg.layer_op_failed", str(ex)))
        c.show(ui, _L.T("ui.error"), _L.T("msg.layer_op_failed", str(ex)),
               is_error=True)


if __name__ == '__main__':
    main()
