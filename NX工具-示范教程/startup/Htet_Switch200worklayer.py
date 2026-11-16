# -*- coding: utf-8 -*-
# NX 2506 插件：把当前部件的工作图层切换到第 200 层
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

TARGET_LAYER = 200


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return
    session.SetUndoMark(c.NXOpen_MarkVisibility_Visible(),
                        _L.T("undo.worklayer200"))
    part.Layers.WorkLayer = TARGET_LAYER


if __name__ == '__main__':
    main()
