# -*- coding: utf-8 -*-
# ==============================================================================
# Htet_export_body01.py —— 逐个实体出图（人工）
#
# 每次运行只显示当前实体、隐藏其余实体，由用户自己出图/操作；再次运行切换下一个。
# 与 Htet_autodwg.py（自动出图）是两套独立功能，进度文件互不干扰。
# ==============================================================================
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import NXOpen
import NXOpen.BlockStyler
import Htet_common as c
import Htet_lang as _L

PROGRESS_FILE = c.progress_path("Htet_export_body_progress.txt")
DLX_ABS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Htet_export_body01.dlx")


class export_body01(object):
    def __init__(self):
        session, ui, part, lw = c.ctx()
        self.theSession, self.theUI, self.workPart, self.lw = session, ui, part, lw
        self.selectObject0 = None
        self._selected_body_tags = None
        self.theDialog = None
        if self.workPart is not None:
            self.theDialog = ui.CreateDialog(DLX_ABS_PATH)
            self.theDialog.AddOkHandler(self.ok_cb)
            self.theDialog.AddUpdateHandler(self.update_cb)
            self.theDialog.AddInitializeHandler(self.initialize_cb)
            self.theDialog.AddDialogShownHandler(self.dialogShown_cb)

    def Launch(self):
        import NXOpen
        response = NXOpen.BlockStyler.BlockDialog.DialogResponse.Invalid
        try:
            response = self.theDialog.Launch()
        except Exception as ex:
            c.show(self.theUI, _L.T("eb.launch_failed"), ex, True)
        return response

    def Dispose(self):
        if self.theDialog is not None:
            try:
                self.theDialog.Dispose()
            except Exception:
                pass
            self.theDialog = None
        if self.lw is not None:
            try:
                self.lw.Close()
            except Exception:
                pass

    def initialize_cb(self):
        try:
            self.selectObject0 = self.theDialog.TopBlock.FindBlock("selectObject0")
        except Exception as ex:
            self.lw.WriteLine(_L.T("eb.init_failed", str(ex)))

    def dialogShown_cb(self):
        pass

    def update_cb(self, block):
        return 0

    def ok_cb(self):
        try:
            self._selected_body_tags = c.body_tags(
                c.ufs(), self.selectObject0.GetSelectedObjects())
            if not self._selected_body_tags:
                self.lw.WriteLine(_L.T("eb.no_selection"))
                self._selected_body_tags = None
                return 1
            return 0
        except Exception as ex:
            self.lw.WriteLine(_L.T("eb.read_failed", str(ex)))
            self._selected_body_tags = None
            return 1

    # ---------------------------- 主流程 ----------------------------
    def _part_path(self):
        try:
            return self.workPart.FullPath or ""
        except Exception:
            return ""

    def _part_key(self):
        path = self._part_path()
        if not path:
            return ""
        try:
            return os.path.normcase(os.path.abspath(path))
        except Exception:
            return path

    def run_next_export(self):
        all_bodies = list(self.workPart.Bodies)
        state = c.read_task_state(PROGRESS_FILE)
        if state is None:
            self.lw.WriteLine(_L.T("eb.no_progress"))
            return
        start_idx, tag_list = state["index"], state["tags"]
        not_recorded = _L.T("eb.not_recorded")
        self.lw.WriteLine(_L.T("eb.state_info", state["part"] or not_recorded,
                               state["stamp"] or not_recorded))
        if state["legacy"]:
            c.write_task_state(PROGRESS_FILE, self._part_path(), start_idx, tag_list)

        if start_idx >= len(tag_list):
            c.clear_progress(PROGRESS_FILE)
            self.lw.WriteLine(_L.T("eb.all_done"))
            return

        current = None
        for body in all_bodies:
            if body.Tag == tag_list[start_idx]:
                current = body
                break

        if current is None:
            c.write_task_state(PROGRESS_FILE, self._part_path(), start_idx + 1, tag_list)
            self.lw.WriteLine(_L.T("eb.skip_missing", start_idx + 1))
            self.run_next_export()
            return

        self.theSession.SetUndoMark(c.NXOpen_MarkVisibility_Visible(),
                                    _L.T("eb.undo_mark", start_idx + 1))
        # 切换实体前抑制隐式图纸更新：否则切到下一个实体后再建“基本视图”，NX 会把
        # 已出好的旧视图标记过期并按“当前可见对象”重算，旧视图内容被覆盖成新实体。
        suppressed = c.set_suppress_view_update(True)
        self.lw.WriteLine("已抑制图纸视图隐式更新（原值=%s）：切换实体不会重算已有视图。"
                          % ("不可用" if suppressed is None else int(suppressed)))
        c.set_blank(c.ufs(), all_bodies, current.Tag)

        next_idx = start_idx + 1
        self.lw.WriteLine("")
        self.lw.WriteLine(_L.T("eb.current", start_idx + 1, len(tag_list), current.Tag))
        if next_idx >= len(tag_list):
            c.clear_progress(PROGRESS_FILE)
            self.lw.WriteLine(_L.T("eb.last_body"))
        else:
            c.write_task_state(PROGRESS_FILE, self._part_path(), next_idx, tag_list)
            self.lw.WriteLine(_L.T("eb.next_hint"))


def main():
    import NXOpen
    session, ui, part, lw = c.ctx()
    if not c.require_part(lw, part):
        c.show(ui, _L.T("ui.info"), _L.T("eb.no_part"), True)
        return

    app = None
    try:
        app = export_body01()
        state = c.read_task_state(PROGRESS_FILE)
        if state is not None:
            other_part = False
            if state["part"]:
                saved = state["part"]
                try:
                    saved = os.path.normcase(os.path.abspath(saved))
                except Exception:
                    pass
                other_part = (saved != app._part_key())
            if other_part:
                c.clear_progress(PROGRESS_FILE)
                app.lw.WriteLine(_L.T("eb.other_part", state["part"]))
            elif state["index"] < len(state["tags"]):
                app.lw.WriteLine(_L.T("eb.resume", state["index"] + 1,
                                      len(state["tags"]),
                                      state["stamp"] or _L.T("eb.not_recorded")))
                app.run_next_export()
                return
            else:
                c.clear_progress(PROGRESS_FILE)
                app.lw.WriteLine(_L.T("eb.prev_done"))

        ret = app.Launch()
        if (ret == NXOpen.BlockStyler.BlockDialog.DialogResponse.Ok
                and app._selected_body_tags):
            c.write_task_state(PROGRESS_FILE, app._part_path(), 0,
                               app._selected_body_tags)
            app.run_next_export()
        else:
            # 选择窗取消/关闭即结束：只落日志，不输出界面提示。
            c.log_event("Htet_drawing.log", _L.T("eb.no_start"))
    except Exception as ex:
        c.show(ui, _L.T("eb.exception_title"), ex, True)
    finally:
        if app is not None:
            app.Dispose()


if __name__ == '__main__':
    main()
