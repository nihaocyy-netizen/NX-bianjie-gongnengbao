# -*- coding: utf-8 -*-
# ==============================================================================
# Htet_common.py —— 本插件各脚本共用的少量公共函数（公共实现只写一份）
#
# 设计要点：
#   • 模块顶层不 import NXOpen，只在函数内用到时再导入（懒加载，减少脚本冷启动开销）；
#   • 只放“被多个脚本重复使用”的逻辑：上下文、选择、改色/透明/图层、显示隐藏、进度文件、消息框。
#   • 本文件必须与其它 .py 脚本一起放在 application 目录。
# ==============================================================================
import os
import tempfile

import Htet_lang as _L

_INFO = "Information"
_ERROR = "Error"

# NXMessageBox.DialogType.Question 的返回值：两套编码都要认。
#   • Win32 标准按钮 ID：IDOK = 1、IDCANCEL = 2（NX 2506 实机点「是」返回 1、点「否」返回 2）；
#   • NXOpen DialogType：Yes = 6、No = 7（NX 文档标注这一套）。
# 只绑一套会在另一套环境下判错，所以按集合判断，其余值一律按取消处理。
MSGBOX_YES_VALUES = frozenset((1, 6))
MSGBOX_NO_VALUES = frozenset((2, 7))

# 插件作者：Wong、Htet、Limulu
AUTHOR = "Wong、Htet、Limulu"
__author__ = AUTHOR


def ctx(open_listing=True):
    """返回 (session, ui, work_part, listing_window)。

    open_listing=False 时不展开信息窗口，供“点完即走”的成功路径使用；
    失败/前置提示仍由 require_part 或 show 负责弹出。
    """
    import NXOpen
    session = NXOpen.Session.GetSession()
    lw = session.ListingWindow
    if open_listing:
        try:
            lw.Open()
        except Exception:
            pass
    return session, NXOpen.UI.GetUI(), session.Parts.Work, lw


def require_part(lw, part):
    if part is None:
        # 未打开零件属于前置条件，必须让用户看见：此时才展开信息窗口。
        try:
            lw.Open()
        except Exception:
            pass
        lw.WriteLine(_L.T("msg.no_part_generic"))
        return False
    return True


def notice(lw, text):
    """失败/无目标等非成功提示：先展开信息窗口再写，保证用户能看到。"""
    try:
        lw.Open()
    except Exception:
        pass
    lw.WriteLine(text)


def ufs():
    """获取 UF 会话（NXOpen.UF 是子模块，必须显式导入）。"""
    import NXOpen
    import NXOpen.UF
    return NXOpen.UF.UFSession.GetUFSession()


def show(ui, title, msg, is_error=False):
    """消息框；失败时只写日志，不影响主流程。"""
    if ui is None:
        return
    try:
        import NXOpen
        dtype = (NXOpen.NXMessageBox.DialogType.Error if is_error
                 else NXOpen.NXMessageBox.DialogType.Information)
        ui.NXMessageBox.Show(title, dtype, str(msg))
    except Exception:
        pass


def confirm(ui, title, msg):
    """Question 消息框：返回 True 表示用户选 Yes；任何异常按取消（No）处理。

    返回值按 MSGBOX_YES_VALUES / MSGBOX_NO_VALUES 两套编码兼容判断，
    日志同时记录原始值与判定结果；既非「是」也非「否」时按取消处理。
    """
    if ui is None:
        return False
    try:
        import NXOpen
        ret = ui.NXMessageBox.Show(title, NXOpen.NXMessageBox.DialogType.Question,
                                   str(msg))
        value = None if ret is None else int(ret)
        if value in MSGBOX_YES_VALUES:
            log_event("Htet_drawing.log", _L.T("log.msgbox_ret_yes", value))
            return True
        if value in MSGBOX_NO_VALUES:
            log_event("Htet_drawing.log", _L.T("log.msgbox_ret_no", value))
            return False
        log_event("Htet_drawing.log", _L.T("log.msgbox_unknown", value))
        return False
    except Exception:
        return False


def _solid_mask():
    """实体选择过滤 MaskTriple（Body / SolidBody）。"""
    import NXOpen
    try:
        import NXOpen.UF
        typ = int(NXOpen.UF.UFConstants.UF_solid_type)
        sub = int(NXOpen.UF.UFConstants.UF_solid_body_subtype)
    except Exception:
        typ, sub = 70, 0
    try:
        return [NXOpen.Selection.MaskTriple(typ, sub, 0)]
    except Exception:
        return []


def _exc_msg(exc):
    """异常详情：单行化并截断到 200 字符，供 log.pick_error 第二占位符使用。"""
    try:
        text = " ".join(str(exc).split())
    except Exception:
        text = ""
    return text[:200]


def _mask_debug(mask):
    """把 mask 实参形态压成一行：类型、数量、首元素类型与三个分量值。"""
    try:
        text = "%s len=%d" % (type(mask).__name__, len(mask))
    except Exception as exc:
        return "%s? %s" % (type(mask).__name__, type(exc).__name__)
    try:
        first = mask[0]
    except Exception:
        return text + " [0]=空"
    parts = ["[0]=%s" % type(first).__name__]
    for name in ("Type", "Subtype", "SolidBodySubtype"):
        try:
            parts.append("%s=%s" % (name, getattr(first, name)))
        except Exception:
            parts.append("%s=?" % name)
    return text + " " + " ".join(parts)


def _solid_filter_members():
    """反射 NXOpen.Select.FilterMember，取实体相关枚举；取不到返回空列表。

    实测该枚举类没有 SolidBody / Body / Solid，实体对应项是 AllSolidBodies，
    其余三个仅作兜底，全部按名字反射，不硬编码数值。
    """
    try:
        import NXOpen
        mod = getattr(NXOpen, "Select", None)
        if mod is None:
            import NXOpen.Select as mod
        cls = getattr(mod, "FilterMember", None)
    except Exception:
        return []
    if cls is None:
        return []
    for name in ("AllSolidBodies", "SolidFeature",
                 "ConvergentSolidBody", "AllFacetBodies"):
        try:
            value = getattr(cls, name)
        except Exception:
            continue
        if value is not None:
            return [value]
    return []


def _response_ok(response):
    """返回码判定：先取整数值，取不到再退回枚举对象相等比较。"""
    if response is None:
        return False
    value = None
    for getter in (lambda: response.value, lambda: str(response)):
        try:
            value = int(getter())
            break
        except Exception:
            continue
    try:
        import NXOpen
        ok = NXOpen.SelectionResponse.Ok
        selected = NXOpen.SelectionResponse.ObjectSelected
    except Exception:
        return False
    if value is not None:
        codes = []
        for member in (ok, selected):
            try:
                codes.append(int(member.value))
            except Exception:
                pass
        if codes:
            return value in codes
    try:
        return response == ok or response == selected
    except Exception:
        return False


def pick_objects(ui, prompt, solid_only=True, part=None):
    """调 NX 原生选择；Esc/取消或接口不可用时返回空列表，不抛异常。

    必须走实例 NXOpen.UI.GetUI().SelectionManager：NXOpen.Selection 上的
    SelectObjects 是未绑定方法，直接类调用会把第一个参数当 self，导致参数错位。
    失败原因分三类各记一条日志（接口缺失 / 签名不匹配 / 调用异常），
    但控制流不变：任何一类失败都只是降级为「询问是否对全部执行」。
    part 只用于把预选对象换回具体类型，可省略。
    """
    if ui is None:
        return []
    try:
        import NXOpen
    except Exception as exc:
        log_event("Htet_drawing.log",
                  _L.T("log.pick_error", type(exc).__name__, _exc_msg(exc)))
        return selected_objects(ui, solid_only=solid_only, part=part)
    if not (hasattr(NXOpen, "UI") and hasattr(NXOpen, "Selection")):
        log_event("Htet_drawing.log",
                  _L.T("log.pick_unavailable", "NXOpen.UI.SelectionManager"))
        return selected_objects(ui, solid_only=solid_only, part=part)
    # 批处理下选择框是模态的，会挂住 journal，直接走已选对象。
    try:
        batch = NXOpen.Session.GetSession().IsBatch
        if callable(batch):
            batch = batch()
        if batch:
            return selected_objects(ui, solid_only=solid_only, part=part)
    except Exception:
        pass
    try:
        sel = NXOpen.UI.GetUI().SelectionManager
    except Exception:
        sel = None
    if sel is None:
        log_event("Htet_drawing.log",
                  _L.T("log.pick_unavailable", "NXOpen.UI.SelectionManager"))
        return selected_objects(ui, solid_only=solid_only, part=part)
    title = _L.T("ui.pick_title")
    if solid_only:
        mask = _solid_mask()
    else:
        # 空 mask 在 GUI 原生框里等于「过滤器为空」，任何对象都选不中；
        # 0/0/0 是 NX 的「任意类型」三元组，保证非实体流程也能选中对象。
        try:
            mask = [NXOpen.Selection.MaskTriple(0, 0, 0)]
        except Exception:
            mask = []
    scope = NXOpen.Selection.SelectionScope.WorkPart
    log_event("Htet_drawing.log",
              _L.T("log.pick_args", type(sel).__name__, scope,
                   _mask_debug(mask)))
    response = None
    objects = []
    used = ""

    def _try_filter_members():
        """兜底：SelectTaggedObjectsWithFilterMembers（5 参，NX1980+）。"""
        try:
            fn = sel.SelectTaggedObjectsWithFilterMembers
            action = NXOpen.Selection.SelectionAction.ClearAndEnableSpecific
        except Exception:
            return None
        if not callable(fn):
            return None
        members = _solid_filter_members() if solid_only else []
        if solid_only and not members:
            return None
        try:
            resp, objs = fn(str(prompt), title, scope, action, members)
        except Exception as exc:
            log_event("Htet_drawing.log",
                      _L.T("log.pick_fallback_error",
                           type(exc).__name__, _exc_msg(exc)))
            return None
        return ("filter-members", resp, objs)

    try:
        response, objects = sel.SelectObjects(
            str(prompt), title, scope,
            NXOpen.Selection.SelectionAction.ClearAndEnableSpecific,
            False, False, mask)
        used = "7-arg"
    except TypeError:
        try:
            response, objects = sel.SelectObjects(
                str(prompt), title, scope, False, False)
            used = "5-arg"
        except TypeError:
            fb = _try_filter_members()
            if fb is None:
                log_event("Htet_drawing.log",
                          _L.T("log.pick_sig_mismatch", "SelectObjects 5/7-arg"))
                return []
            used, response, objects = fb
        except Exception as exc:
            fb = _try_filter_members()
            if fb is None:
                log_event("Htet_drawing.log",
                          _L.T("log.pick_error", type(exc).__name__,
                               _exc_msg(exc)))
                return selected_objects(ui, solid_only=solid_only, part=part)
            used, response, objects = fb
    except Exception as exc:
        fb = _try_filter_members()
        if fb is None:
            log_event("Htet_drawing.log",
                      _L.T("log.pick_error", type(exc).__name__, _exc_msg(exc)))
            return selected_objects(ui, solid_only=solid_only, part=part)
        used, response, objects = fb
    if not used:
        return []
    try:
        if not _response_ok(response):
            return []
        if solid_only:
            picked = [obj for obj in (objects or [])
                      if isinstance(obj, NXOpen.Body) and obj.IsSolidBody]
        else:
            # 非实体流程（改色/半透明/移层）按显示对象收，曲线、片体、面都能处理。
            picked = [obj for obj in (objects or [])
                      if isinstance(obj, NXOpen.DisplayableObject)]
        log_event("Htet_drawing.log", _L.T("log.pick_ok", used, len(picked)))
        return picked
    except Exception as exc:
        log_event("Htet_drawing.log",
                  _L.T("log.pick_error", type(exc).__name__, _exc_msg(exc)))
        return []


def clear_selection(ui):
    """释放图形区选中状态（含原生选择框留下的一次性选择列表）。

    背景（实机）：改色/半透明等执行完后实体仍保持选中，下一次点同一功能时
    resolve_targets 会直接读到这批残留选中对象，于是不弹选择框、把同一批
    实体再刷一遍。功能取到目标后立即清空，保证下一次总是重新选择。

    注意：这只解决「复用上次选中/不弹框」，与「实体点不动」无关——后者是
    目标层落成 Visible（参考层，可见不可选）导致的，见 apply_layer。
    这里把清理前后的选中数与 IsGlobalSelectionActive 记进日志，便于实机确认。
    """
    if ui is None:
        return False
    sel = None
    before = -1
    try:
        sel = ui.SelectionManager
        before = sel.GetNumSelectedObjects()
    except Exception:
        sel = getattr(ui, "SelectionManager", None)
    try:
        # RemoveAllFromSelectionList 需要 UIStyler 的 SelectionHandle，此处不适用；
        # 图形区残留选中在全局选择列表里，用无参的 ClearGlobalSelectionList 清理。
        sel.ClearGlobalSelectionList()
    except Exception as exc:
        log_event("Htet_drawing.log", _L.T(
            "log.sel_clear_error", type(exc).__name__, _exc_msg(exc)))
        return False
    try:
        after = sel.GetNumSelectedObjects()
        active = sel.IsGlobalSelectionActive()
        if before > 0 or after > 0 or active:
            log_event("Htet_drawing.log", _L.T(
                "log.sel_clear", before, after, active))
    except Exception:
        pass
    return True


def resolve_targets(ui, part, lw, solid_only=False):
    """统一“取目标对象”口径：已选直接返回 → 原生选择 → Question 确认全部。

    返回对象列表；返回 [] 表示零件里没有可处理对象；返回 None 表示用户取消。
    取到目标后立即清空选择状态，避免下一次触发复用残留选中集。
    """
    targets = selected_objects(ui, solid_only=solid_only, part=part)
    if targets:
        clear_selection(ui)
        return targets

    # 图形区确实有预选、但换成实体后为空：说明选到的是片体/曲线/点这类非实体。
    # 不能默默当成「没选」，先给一条可见提示，再走原有的选择框（行为不退化）。
    picked = selection_count(ui)
    if picked > 0 and solid_only:
        notice(lw, _L.T("msg.presel_not_solid", picked))
        log_event("Htet_drawing.log", _L.T("log.presel_not_solid", picked))

    targets = pick_objects(ui, _L.T("msg.pick_prompt"), solid_only=solid_only,
                           part=part)
    if targets:
        clear_selection(ui)
        return targets

    if solid_only:
        bodies = solid_bodies(part)
    else:
        bodies = list(part.Bodies)
    if not bodies:
        return []
    if not confirm(ui, _L.T("msg.confirm_all_title"),
                   _L.T("msg.confirm_all", len(bodies))):
        # 点「否」本身就是取消：不再往信息窗口写提示，直接返回，调用方立即退出。
        return None
    lw.WriteLine(_L.T("msg.fallback_all", len(bodies)))
    return bodies


def collect_display_objects(obj, result, visited, solid_only=False):
    """把选中对象展开成可处理对象；特征是 NXObject 但不是 DisplayableObject。

    沿 GetBodies/GetEntities 等接口递归取特征生成的几何，这样用户点特征也能
    改色、改透明、移层。实体直接加入；solid_only=True 时面/边换回所属实体，
    曲线/点/片体不加入。
    """
    import NXOpen
    tag = getattr(obj, "Tag", None)
    if tag is not None:
        if tag in visited:
            return
        visited.add(tag)
    if isinstance(obj, NXOpen.Body):
        if (not solid_only) or obj.IsSolidBody:
            result.append(obj)
        return
    if isinstance(obj, NXOpen.DisplayableObject):
        if not solid_only:
            result.append(obj)
            return
        # 只要实体时，面/边先换回它所属的体：
        # 用户常点中「体上的面」，若直接丢弃，就会被当成没选任何东西又弹一次框。
        getter = getattr(obj, "GetBody", None)
        if getter is not None:
            try:
                owner = getter()
            except Exception:
                owner = None
            if owner is not None:
                collect_display_objects(owner, result, visited, solid_only)
        return
    # 不同特征类型提供的输出接口不同：BodyFeature 通常用 GetBodies()，
    # 曲线/草图类特征常用 GetEntities()、GetFaces() 或 GetEdges()。
    for name in ("GetBodies", "GetResultBodies", "GetBody", "GetResultBody",
                 "GetEntities", "GetFaces", "GetEdges"):
        method = getattr(obj, name, None)
        if method is None:
            continue
        try:
            entities = method()
        except Exception:
            continue
        if not entities:
            continue
        # GetBody()/GetResultBody() 可能返回单个对象，其余接口返回集合。
        if isinstance(entities, NXOpen.NXObject):
            entities = (entities,)
        before = len(result)
        for entity in entities:
            collect_display_objects(entity, result, visited, solid_only)
        if len(result) > before:
            return


def selection_count(ui):
    """图形区当前选中对象数；接口不可用返回 0。"""
    try:
        return int(ui.SelectionManager.GetNumSelectedObjects())
    except Exception:
        return 0


def _wrap_by_tag(obj, part):
    """把选择管理器可能返回的基类 TaggedObject 换回零件内的具体类型对象。

    实机（NX 2506）里 GetSelectedTaggedObject 有时给的是基类包装：它既不是
    Body 也不是 DisplayableObject，原有 instanceof 判断全部落空，图形区的
    合法预选于是被当成空、又弹一次选择框（v3.3.8 修的预选 bug）。
    按 Tag 在零件的集合里查回具体对象后，后续判断与原生选择框完全一致。
    """
    import NXOpen
    if obj is None or isinstance(obj, NXOpen.NXObject):
        return obj
    if part is None:
        return obj
    tag = getattr(obj, "Tag", None)
    if tag is None:
        return obj
    for attr in ("Bodies", "FacetedBodies", "SubdivisionBodies",
                 "Features", "Curves", "Points"):
        coll = getattr(part, attr, None)
        if coll is None:
            continue
        try:
            for item in coll:
                if item.Tag == tag:
                    return item
        except Exception:
            continue
    return obj


def selected_objects(ui, solid_only=False, part=None):
    """读取图形区已选中的对象；solid_only=True 时只要实体。

    • 选择管理器返回基类 TaggedObject 时按 Tag 换回具体对象；
    • 选中面/边时，solid_only 下换回所属实体；
    • 选中特征时按 collect_display_objects 展开为它生成的几何。
    """
    out = []
    visited = set()
    picked = 0
    try:
        num = ui.SelectionManager.GetNumSelectedObjects()
        for i in range(num):
            obj = _wrap_by_tag(ui.SelectionManager.GetSelectedTaggedObject(i), part)
            if obj is not None:
                picked += 1
            collect_display_objects(obj, out, visited, solid_only)
    except Exception:
        pass
    if picked:
        log_event("Htet_drawing.log", _L.T(
            "log.sel_preselect", picked, len(out), 1 if solid_only else 0))
    return out


def solid_bodies(part):
    return [b for b in part.Bodies if b.IsSolidBody]


def body_tags(uf, objects):
    """把选择对象过滤成实体 Tag 列表（70/0 = Body）。"""
    tags = []
    for obj in objects:
        try:
            typ, subtyp = uf.Obj.AskTypeAndSubtype(obj.Tag)
            if typ == 70 and subtyp == 0:
                tags.append(obj.Tag)
        except Exception:
            pass
    return tags


def set_blank(uf, bodies, visible_tag):
    """隐藏 bodies 中的全部实体，只显示 visible_tag（可为 None = 全部隐藏）。"""
    for b in bodies:
        try:
            uf.Obj.SetBlankStatus(b.Tag, 0 if b.Tag == visible_tag else 1)
        except Exception:
            pass


def apply_color(session, targets, color):
    dm = session.DisplayManager.NewDisplayModification()
    try:
        dm.ApplyToAllFaces = True
        dm.NewColor = color
        dm.Apply(targets)
    finally:
        dm.Dispose()


def apply_translucency(session, targets, percent):
    dm = session.DisplayManager.NewDisplayModification()
    try:
        # NX 的透明度作用在面上，选中实体/特征时需同时应用到全部面。
        dm.ApplyToAllFaces = True
        dm.NewTranslucency = percent
        dm.Apply(targets)
    finally:
        dm.Dispose()


def apply_with_undo(session, ui, lw, label, action, fail_key):
    """执行一次可撤销的显示类修改：成功返回 True；失败回滚并弹框后返回 False。

    改色/透明这类 DisplayModification 抛错时若直接冒泡，NX 里既没有提示也
    不会回滚；这里统一收口，保证失败后图形状态回到执行前。
    """
    mark = session.SetUndoMark(NXOpen_MarkVisibility_Visible(), label)
    try:
        action()
        return True
    except Exception as exc:
        try:
            session.UndoToMark(mark, "")
        except Exception:
            pass
        message = _L.T(fail_key, str(exc))
        try:
            lw.WriteLine(message)
        except Exception:
            pass
        show(ui, _L.T("ui.error"), message, is_error=True)
        return False


# 层状态实测（NX2506，logs\probe_layer.log）：
#   Hidden     ↔ UF_LAYER_INACTIVE_LAYER(4)   不可见、不可选（空层默认值）
#   Visible    ↔ UF_LAYER_REFERENCE_LAYER(3)  可见、不可选 ←「看得见但点不动」
#   Selectable ↔ UF_LAYER_ACTIVE_LAYER(2)     可见、可选
#   WorkLayer  ↔ UF_LAYER_WORK_LAYER(1)
_LAYER_STATE_NAMES = ("Hidden", "Visible", "Selectable", "WorkLayer")


def _layer_state_name(state):
    try:
        import NXOpen.Layer as _Layer
        for name in _LAYER_STATE_NAMES:
            if state == getattr(_Layer.State, name):
                return name
    except Exception:
        pass
    return repr(state)


def apply_layer_to(session, objects, layer):
    """把一批对象移到指定层，并把该层状态修正为「可见可选」。"""
    import NXOpen
    import NXOpen.Layer
    if not objects:
        return
    dm = session.DisplayManager.NewDisplayModification()
    try:
        dm.NewLayer = layer
        dm.Apply(list(objects))
    finally:
        dm.Dispose()
    # 移层后的层状态必须修正为「可见可选」：
    #   • 空层默认 Hidden，移进去后实体直接不显示；
    #   • 旧实现只把 Hidden 改成 Visible，而 Visible 是参考层（可见不可选），
    #     结果就是实体看得见、点不动；且函数内未显式 import NXOpen.Layer，
    #     该分支在部分会话里会因 AttributeError 被静默吞掉。
    # 因此统一提升到 Selectable；工作层/已是可选层不动。
    try:
        layers = session.Parts.Work.Layers
        before = layers.GetState(layer)
        if (before == NXOpen.Layer.State.Hidden or
                before == NXOpen.Layer.State.Visible):
            layers.SetState(layer, NXOpen.Layer.State.Selectable)
            after = layers.GetState(layer)
            log_event("Htet_drawing.log", _L.T(
                "log.layer_state_fixed", layer,
                _layer_state_name(before), _layer_state_name(after)))
    except Exception as exc:
        log_event("Htet_drawing.log", _L.T(
            "log.layer_state_error", layer, type(exc).__name__, _exc_msg(exc)))


def apply_layer(session, body, layer):
    """把单个对象移到指定层（保留原调用口径）。"""
    apply_layer_to(session, [body], layer)


def layer_has_objects(part, number):
    """该层是否有可显示对象；上下层切换用它跳过空层。"""
    try:
        import NXOpen
        return any(isinstance(obj, NXOpen.DisplayableObject)
                   for obj in part.Layers.GetAllObjectsOnLayer(number))
    except Exception:
        return False


def NXOpen_MarkVisibility_Visible():
    import NXOpen
    return NXOpen.Session.MarkVisibility.Visible


# ---------------------------- 日志与单实例锁（合并版新增） ----------------------------
# 对齐 STEP 项目：日志放插件 logs\、UTF-8 带 BOM、2 MB × 3 轮转、不可写回落 %TEMP%。
LOG_MAX_BYTES = 2 * 1024 * 1024
LOG_KEEP = 3


def plugin_dir():
    """插件根目录：脚本在 application 目录下，取上一层；被别处执行时回落脚本目录。"""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()
    if os.path.basename(here).lower() in ("application", "startup"):
        parent = os.path.dirname(here)
        if parent:
            return parent
    return here


def _rotate_log(path):
    """path 超过阈值时：log -> .1 -> .2 -> .3，最多保留 LOG_KEEP 份。"""
    try:
        if not os.path.exists(path) or os.path.getsize(path) < LOG_MAX_BYTES:
            return
        for i in range(LOG_KEEP, 0, -1):
            src = path if i == 1 else "%s.%d" % (path, i - 1)
            dst = "%s.%d" % (path, i)
            if os.path.exists(dst):
                os.remove(dst)
            if os.path.exists(src):
                os.rename(src, dst)
    except Exception:
        pass


def _append_log(path, message):
    if not os.path.isdir(os.path.dirname(path)):
        os.makedirs(os.path.dirname(path))
    _rotate_log(path)
    if not os.path.exists(path):
        with open(path, "wb") as f:
            f.write(b"\xef\xbb\xbf")
    import time
    with open(path, "a", encoding="utf-8") as f:
        f.write("[%s PID=%d] %s\n"
                % (time.strftime("%Y-%m-%d %H:%M:%S"), os.getpid(), str(message)))


def log_event(name, message):
    """写插件 logs\\<name>；目录不可写时回落到 %TEMP%，不影响主流程。"""
    try:
        _append_log(os.path.join(plugin_dir(), "logs", name), message)
    except Exception:
        try:
            _append_log(os.path.join(tempfile.gettempdir(), name), message)
        except Exception:
            pass


def _pid_alive(pid):
    try:
        import ctypes
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    except Exception:
        return True


def lock_file(name):
    """锁文件完整路径（获取、释放、提示都落在这一个位置）。"""
    return os.path.join(tempfile.gettempdir(), name)


def acquire_lock(name):
    """返回锁文件路径；已有存活实例时返回 None；无法判断时返回空串（放行）。"""
    path = lock_file(name)
    legacy = _legacy_path(name)
    if legacy and os.path.exists(legacy):
        try:
            with open(legacy, "r", encoding="utf-8") as f:
                old_pid = int(f.read().strip() or "0")
        except Exception:
            old_pid = 0
        # 旧锁的进程还活着 → 仍然判忙；已退出 → 清理残留，改用新锁。
        if old_pid and old_pid != os.getpid() and _pid_alive(old_pid):
            return None
        try:
            os.remove(legacy)
        except Exception:
            pass
    try:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    old = int(f.read().strip() or "0")
            except Exception:
                old = 0
            if old and old != os.getpid() and _pid_alive(old):
                return None
            os.remove(path)
        with open(path, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        return path
    except Exception:
        return ""


def release_lock(path):
    if not path:
        return
    try:
        os.remove(path)
    except Exception:
        pass


# ---------------------------- 进度文件 ----------------------------
def _legacy_path(name):
    """旧版运行时文件路径（历史前缀见下方字符串拼接）；不匹配返回空串。"""
    if name.startswith("Htet_"):
        return os.path.join(tempfile.gettempdir(), ("h" + "bt_") + name[5:])
    return ""


def _takeover_legacy(new_path, old_path):
    """把旧版运行时文件原样接管为 Htet_*，成功后删除旧文件。"""
    if not old_path or not os.path.exists(old_path) or os.path.exists(new_path):
        return False
    try:
        os.replace(old_path, new_path)
        return True
    except Exception:
        try:
            with open(old_path, "rb") as src, open(new_path, "wb") as dst:
                dst.write(src.read())
            os.remove(old_path)
            return True
        except Exception:
            return False


def progress_path(name):
    path = os.path.join(tempfile.gettempdir(), name)
    _takeover_legacy(path, _legacy_path(name))
    return path


def write_progress(path, signature, next_idx, tags):
    """signature 为 None 时只写“序号 + tags”（供 export_body01 使用）。"""
    try:
        with open(path, "w", encoding="utf-8") as f:
            if signature:
                f.write(signature + "\n")
            f.write("%d\n" % next_idx)
            for t in tags:
                f.write("%d\n" % t)
        return True
    except Exception:
        return False


def read_progress(path, signature=None):
    """返回 (start_idx, tags) 或 None；signature 不匹配则视为无效。"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            lines = [x.strip() for x in f.readlines() if x.strip()]
        if signature is not None:
            if not lines or lines[0] != signature:
                return None
        start_idx = int(lines[1] if signature else lines[0])
        tags = [int(x) for x in (lines[2:] if signature else lines[1:])]
        if not tags:
            return None
        return (start_idx, tags)
    except Exception:
        return None


def clear_progress(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


# ---------------------------- 出图任务状态（含部件/时间戳/索引） ----------------------------
STATE_VERSION = "HTET_STATE_V2"
LEGACY_STATE_VERSIONS = ("H" + "BT_STATE_V2",)


def write_task_state(path, part_path, index, tags, stamp=None):
    """写出图任务状态：版本、部件路径、时间戳、下一个实体索引、实体 Tag 列表。"""
    import time
    if stamp is None:
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(STATE_VERSION + "\n")
            f.write("part=%s\n" % (part_path or ""))
            f.write("stamp=%s\n" % stamp)
            f.write("index=%d\n" % index)
            for t in tags:
                f.write("%d\n" % t)
        return True
    except Exception:
        return False


def read_task_state(path):
    """返回 {part, stamp, index, tags, legacy}；旧的“序号+Tags”格式也能读。"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            lines = [x.strip() for x in f.readlines()]
    except Exception:
        return None
    lines = [x for x in lines if x]
    if not lines:
        return None
    if lines[0] == STATE_VERSION or lines[0] in LEGACY_STATE_VERSIONS:
        info = {"part": "", "stamp": "", "index": None, "tags": [], "legacy": False}
        for line in lines[1:]:
            if line.startswith("part="):
                info["part"] = line[5:]
            elif line.startswith("stamp="):
                info["stamp"] = line[6:]
            elif line.startswith("index="):
                try:
                    info["index"] = int(line[6:])
                except Exception:
                    info["index"] = None
            else:
                try:
                    info["tags"].append(int(line))
                except Exception:
                    pass
        if info["index"] is None or not info["tags"]:
            return None
        if lines[0] != STATE_VERSION:
            # 旧版本号就地迁移（见 LEGACY_STATE_VERSIONS），保证在途进度继续可用。
            write_task_state(path, info["part"], info["index"], info["tags"],
                             info["stamp"] or None)
        return info
    try:
        values = [int(x) for x in lines]
    except Exception:
        return None
    if len(values) < 2:
        return None
    return {"part": "", "stamp": "", "index": values[0],
            "tags": values[1:], "legacy": True}


# ---------------------------- 图纸视图隐式更新抑制 ----------------------------
def set_suppress_view_update(enable):
    """开关 NX“抑制视图更新”偏好，返回设置前的值（不可用返回 None）。

    TRUE 时 NX 的隐式图纸更新不再重算成员视图，用于避免：
      • 批量出图收尾恢复实体可见性后，把其他实体重算进已完成的视图；
      • 逐个出图切换实体后，再建基本视图时把旧视图按当前可见实体重算。
    显式调用 UpdOutOfDateViews / view.Update() 不受此偏好影响。
    """
    try:
        uf = ufs()
        old = bool(uf.Drf.AskSuppressViewUpdate())
    except Exception:
        return None
    try:
        uf.Drf.SetSuppressViewUpdate(bool(enable))
        return old
    except Exception:
        return None
