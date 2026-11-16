# -*- coding: utf-8 -*-
# NX 2506 插件：把实体逐个分配到不同图层（从第 150 层起，每个 +1，上限 256）
# 未选中任何实体时，先让用户原生选择；仍为空则确认后对当前零件的全部实体执行。
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

START_LAYER = 150
LAYER_STEP = 1
MAX_LAYER = 256


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return
    bodies = c.resolve_targets(ui, part, lw, solid_only=True)
    if bodies is None:
        return
    if not bodies:
        c.notice(lw, _L.T("msg.no_bodies_layer"))
        return

    last_layer = START_LAYER + (len(bodies) - 1) * LAYER_STEP
    if last_layer > MAX_LAYER:
        c.notice(lw, _L.T("msg.layer_too_many", MAX_LAYER))
        return

    mark = session.SetUndoMark(c.NXOpen_MarkVisibility_Visible(),
                               _L.T("undo.layer"))
    layers = []
    try:
        for idx, body in enumerate(bodies):
            target_layer = START_LAYER + idx * LAYER_STEP
            c.apply_layer(session, body, target_layer)
            layers.append(target_layer)
    except Exception as ex:
        # 中途失败要整批回滚，避免留下“部分实体已分层”的半成品。
        session.UndoToMark(mark, "")
        lw.WriteLine(_L.T("msg.layer_op_failed", str(ex)))
        c.show(ui, _L.T("ui.error"), _L.T("msg.layer_op_failed", str(ex)),
               is_error=True)
        return




if __name__ == '__main__':
    main()
