# -*- coding: utf-8 -*-
# NX 2506 插件：仅显示选中的实体，隐藏当前部件里的其他实体
# 点菜单后若已有选中实体则直接执行；否则弹出选择窗口，选好点“确定”，取消则不改动。
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import NXOpen
import NXOpen.BlockStyler
import Htet_common as c
import Htet_lang as _L

DLX_ABS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Htet_select_bodies.dlx")


def apply_display_only(ui, session, lw, part, keep_bodies):
    mark = session.SetUndoMark(c.NXOpen_MarkVisibility_Visible(),
                               _L.T("undo.display"))
    try:
        # 吸收内测版做法：按“全部类型”隐藏，曲线、点、草图、基准、注释一并收起，
        # 而不只是遍历实体；随后恢复保留对象及其父组件。
        session.DisplayManager.HideByType(
            NXOpen.DisplayManager.ShowHideType.All,
            NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
        session.DisplayManager.UnblankObjects(keep_bodies)
        parents = []
        parent_tags = set()
        for obj in keep_bodies:
            parent = getattr(obj, "OwningComponent", None)
            while parent is not None and parent.Tag not in parent_tags:
                parent_tags.add(parent.Tag)
                parents.append(parent)
                parent = parent.Parent
        if parents:
            session.DisplayManager.UnblankObjects(parents)
        try:
            c.clear_selection(ui)
            for body in keep_bodies:
                body.Unhighlight()
        except Exception:
            pass
    except Exception as ex:
        # 失败必须回滚：否则会留下“藏了一半”的中间状态。
        try:
            session.UndoToMark(mark, "")
        except Exception:
            pass
        lw.WriteLine(_L.T("msg.display_rolled_back", str(ex)))
        c.show(ui, _L.T("ui.error"),
               _L.T("msg.display_rolled_back", str(ex)), is_error=True)


class DisplayOnlyDialog(object):
    def __init__(self, ui, session):
        import NXOpen
        self.lw = session.ListingWindow
        self.selectObject0 = None
        self.selected_bodies = []
        self.theDialog = ui.CreateDialog(DLX_ABS_PATH)
        self.theDialog.AddOkHandler(self.ok_cb)
        self.theDialog.AddUpdateHandler(self.update_cb)
        self.theDialog.AddInitializeHandler(self.initialize_cb)
        self.theDialog.AddDialogShownHandler(self.dialog_shown_cb)

    def initialize_cb(self):
        self.selectObject0 = self.theDialog.TopBlock.FindBlock("selectObject0")

    def dialog_shown_cb(self):
        pass

    def update_cb(self, block):
        return 0

    def ok_cb(self):
        # 直接保留实体对象本身，装配组件内的实体也能正确恢复显示。
        bodies = [obj for obj in self.selectObject0.GetSelectedObjects()
                  if isinstance(obj, NXOpen.Body)]
        if not bodies:
            self.lw.WriteLine(_L.T("eb.no_selection"))
            return 1
        self.selected_bodies = bodies
        return 0


def main():
    session, ui, part, lw = c.ctx(open_listing=False)
    if not c.require_part(lw, part):
        return

    import NXOpen
    # 只接受实体：若把面/曲线也当成保留对象，part.Bodies 里没有对应 Tag，
    # 会导致「只选一个面 → 全部实体被隐藏」。
    # solid_only=True 时公共入口会把「选中的面/边」换回所属实体，并按 Tag
    # 还原选择管理器返回的基类对象，图形区预选因此能直接生效（v3.3.8）。
    bodies = c.selected_objects(ui, solid_only=True, part=part)
    if bodies:
        c.clear_selection(ui)
        apply_display_only(ui, session, lw, part, bodies)
        return

    dlg = DisplayOnlyDialog(ui, session)
    try:
        ret = dlg.theDialog.Launch()
        if (ret == NXOpen.BlockStyler.BlockDialog.DialogResponse.Ok
                and dlg.selected_bodies):
            apply_display_only(ui, session, lw, part, dlg.selected_bodies)
        else:
            # 对话框取消即结束：只落日志，不输出界面提示。
            c.log_event("Htet_drawing.log", _L.T("msg.display_cancelled"))
    finally:
        try:
            dlg.theDialog.Dispose()
        except Exception:
            pass


if __name__ == '__main__':
    main()
