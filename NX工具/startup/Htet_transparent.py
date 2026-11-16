# -*- coding: utf-8 -*-
# NX 2506 插件：把对象设为半透明（50%）
# 未选中任何对象时，先让用户原生选择；仍为空则确认后对当前零件的全部实体执行。
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

TRANSLUCENCY = 50


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return
    targets = c.resolve_targets(ui, part, lw, solid_only=False)
    if targets is None:
        return
    if not targets:
        c.notice(lw, _L.T("msg.no_bodies_translucent"))
        return

    if not c.apply_with_undo(
            session, ui, lw, _L.T("undo.translucent"),
            lambda: c.apply_translucency(session, targets, TRANSLUCENCY),
            "msg.translucent_failed"):
        return

    # 吸收内测版效果：打开会话级半透明开关，清掉选中高亮并刷新视图，
    # 让透明度立即显现；这些是显示收尾，失败不影响已完成的透明设置。
    try:
        session.Preferences.VisualizationVisualPreferences.Translucency = True
    except Exception:
        pass
    try:
        c.clear_selection(ui)
        for obj in targets:
            try:
                obj.Unhighlight()
            except Exception:
                pass
        display_part = session.Parts.Display
        if display_part is not None:
            display_part.ModelingViews.WorkView.Regenerate()
    except Exception:
        pass


if __name__ == '__main__':
    main()
