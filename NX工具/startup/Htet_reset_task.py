# -*- coding: utf-8 -*-
# NX 2506 插件：重置出图任务（清除进度文件）
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Htet_common as c
import Htet_lang as _L

PROGRESS_FILES = (
    "Htet_autodraw_progress.txt",     # 批量出图（自动）
    "Htet_export_body_progress.txt",  # 逐个实体出图（人工）
)


def restore_all_visible(session, part, lw):
    """清除隐藏状态：把当前部件的全部实体恢复为可见。"""
    if part is None:
        return 0
    try:
        bodies = list(part.Bodies)
    except Exception:
        bodies = []
    if not bodies:
        return 0
    restored = 0
    try:
        session.DisplayManager.UnblankObjects(bodies)
        restored = len(bodies)
    except Exception:
        uf = c.ufs()
        for body in bodies:
            try:
                uf.Obj.SetBlankStatus(body.Tag, 0)
                restored += 1
            except Exception:
                pass
    return restored


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    removed = []
    for name in PROGRESS_FILES:
        path = c.progress_path(name)
        if os.path.exists(path):
            c.clear_progress(path)
            if not os.path.exists(path):
                removed.append(name)
    restore_all_visible(session, part, lw)


if __name__ == '__main__':
    main()
