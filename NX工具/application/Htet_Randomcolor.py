# -*- coding: utf-8 -*-
# NX 2506 插件：给实体随机改色（颜色号 1~100，NX 色板共 217 个）
# 未选中任何实体时，先让用户原生选择；仍为空则确认后对当前零件的全部实体执行。
import os
import sys
import random
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

COLOR_MIN = 1
COLOR_MAX = 100


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return
    bodies = c.resolve_targets(ui, part, lw, solid_only=True)
    if bodies is None:
        return
    if not bodies:
        c.notice(lw, _L.T("msg.no_bodies_color"))
        return

    def _paint():
        for body in bodies:
            c.apply_color(session, [body], random.randint(COLOR_MIN, COLOR_MAX))

    if not c.apply_with_undo(session, ui, lw, _L.T("undo.random"), _paint,
                             "msg.color_failed"):
        return


if __name__ == '__main__':
    main()
