"""NX 原生界面的单孔/批量腰孔工具；与 Quickhole.dlx 放在同一目录。"""
import itertools
import math
import os
import sys
import traceback

_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
if _DIRECTORY not in sys.path:
    sys.path.insert(0, _DIRECTORY)

import NXOpen
import NXOpen.BlockStyler
import NXOpen.Features
import NXOpen.UF

# 毫米制参考：螺杆直径、头部公称直径、头高、预设通孔宽度、台阶宽度。
# 内六角圆柱头常用公称尺寸；孔宽及头下余量为可修改的工具预设。
SCREW_PRESETS = {
    "M3": (3, 5.5, 3, 3.4, 6.5),
    "M4": (4, 7, 4, 4.5, 8),
    "M5": (5, 8.5, 5, 5.5, 10),
    "M6": (6, 10, 6, 6.6, 11),
    "M8": (8, 13, 8, 9, 15),
    "M10": (10, 16, 10, 11, 18),
    "M12": (12, 18, 12, 13.5, 20),
}


class UFGeometryQueries:
    """解析当前 NX Python 版本的 UF 几何查询接口并缓存方法。"""

    def __init__(self, uf_session):
        self.session = uf_session
        self.methods = {}

    def call(self, method_name, *args):
        method = self.methods.get(method_name)
        if method is None:
            # Python 与 .NET 的 UF 子接口划分可能不同，不能只照搬
            # .NET 文档中的 UFSession.Modl；按运行时公开接口查找。
            names = ["Modl", "Modeling", "ModlGeneral"]
            names.extend(name for name in dir(self.session)
                         if not name.startswith("_") and name not in names)
            for name in names:
                try:
                    interface = getattr(self.session, name, None)
                    candidate = getattr(interface, method_name, None)
                except (AttributeError, RuntimeError):
                    continue
                if callable(candidate):
                    method = candidate
                    self.methods[method_name] = method
                    break
            if method is None:
                raise RuntimeError("当前 NX Python 未提供 UF 几何查询接口：" + method_name)
        # 不吞掉方法执行时的建模错误，也不在失败后重复调用其他接口。
        return method(*args)


def positive(value, label):
    value = float(value)
    if not math.isfinite(value) or value <= 0:
        raise ValueError(label + "必须为大于零的有限数值。")
    return value


def validate_parameters(diameter, depth, counterbore, cb_diameter, cb_depth, through):
    diameter = positive(diameter, "腰孔宽度")
    if not through:
        depth = positive(depth, "孔深")
    if counterbore:
        cb_diameter = positive(cb_diameter, "上层腰孔宽度")
        cb_depth = positive(cb_depth, "台阶深度")
        if cb_diameter <= diameter:
            raise ValueError("上层腰孔宽度必须大于下层腰孔宽度。")
        if not through and cb_depth >= depth:
            raise ValueError("台阶深度必须小于盲孔总深度。")
    return diameter, depth, cb_diameter, cb_depth


def projected_center(point, plane_point, normal):
    coords = (point.X, point.Y, point.Z)
    distance = sum((coords[i] - plane_point[i]) * normal[i] for i in range(3))
    return tuple(coords[i] - distance * normal[i] for i in range(3))


def through_length(center, direction, box, tolerance):
    corners = itertools.product(*[(box[i], box[i + 3]) for i in range(3)])
    extent = max(sum((p[i] - center[i]) * direction[i] for i in range(3))
                 for p in corners)
    if extent <= tolerance:
        raise ValueError("打孔方向未穿入目标实体，请检查入口面或反转方向。")
    return extent + tolerance * 10


def slot_axis(normal, angle):
    """入口面内以绝对 X 投影为零度；垂直 X 的面改用绝对 Y。"""
    reference = (1.0, 0.0, 0.0) if abs(normal[0]) < 0.95 else (0.0, 1.0, 0.0)
    dot = sum(reference[i] * normal[i] for i in range(3))
    axis = tuple(reference[i] - dot * normal[i] for i in range(3))
    length = math.sqrt(sum(v*v for v in axis))
    axis = tuple(v / length for v in axis)
    cross = (normal[1]*axis[2]-normal[2]*axis[1],
             normal[2]*axis[0]-normal[0]*axis[2],
             normal[0]*axis[1]-normal[1]*axis[0])
    radians = math.radians(angle)
    return tuple(axis[i]*math.cos(radians)+cross[i]*math.sin(radians) for i in range(3))


class QuickHoleDialog:
    def __init__(self):
        self.session = NXOpen.Session.GetSession()
        self.ui = NXOpen.UI.GetUI()
        self.uf = NXOpen.UF.UFSession.GetUFSession()
        self.geometry = UFGeometryQueries(self.uf)
        self.part = self.session.Parts.Work
        self.edit_feature = self.session.CustomFeatureClassManager.GetEditedCustomFeature()
        self.edit_data = None
        self.replace_centers = False
        self.preview_state = None
        self.preview_placement = None
        self.owned_position_tags = set()
        self.derived_points = {}
        self.last_preview_error = None
        if self.edit_feature is not None:
            from QuickholeFeature import CLASS_NAME, read
            if self.edit_feature.FeatureData.ClassName != CLASS_NAME:
                raise ValueError('当前特征不是快捷腰孔参数特征。')
            if self.edit_feature.OwningPart != self.part:
                raise ValueError('请先将腰孔所在零件设为工作零件。')
            self.edit_data = read(self.edit_feature)[0]
        self.count = 0
        self.blocks = {}
        self.busy = True
        self.point_selection_active = False
        self.dialog = self.ui.CreateDialog(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                       "Quickhole.dlx"))
        self.dialog.AddInitializeHandler(self.initialize)
        self.dialog.AddUpdateHandler(self.update)
        self.dialog.AddApplyHandler(self.apply)
        self.dialog.AddOkHandler(self.apply)
        self.dialog.AddCancelHandler(self.cancel)
        self.dialog.AddFocusNotifyHandler(self.focus_notify)
        self.dialog.AddKeyboardFocusNotifyHandler(self.keyboard_focus_notify)
        self.dialog.AddDialogShownHandler(self.dialog_shown)

    def message(self, text):
        try:
            with open(os.path.join(_DIRECTORY, 'Quickhole_ui_error.log'), 'a', encoding='utf-8') as stream:
                stream.write(str(text) + '\n' + traceback.format_exc() + '\n')
        except OSError:
            pass
        self.ui.NXMessageBox.Show("腰型孔向导", NXOpen.NXMessageBox.DialogType.Error, str(text))

    def initialize(self):
        names = ("face", "center", "hole_type", "end_type", "diameter", "depth",
                 "cb_diameter", "cb_depth", "slot_length", "slot_angle", "reverse", "hint",
                 "screw_size", "screw_info", "clear_centers", "point_count",
                 "preview", "show_result")
        self.dialog.TopBlock.Label = "腰型孔向导"
        self.blocks = {name: self.dialog.TopBlock.FindBlock(name) for name in names}
        # StepStatus is initialization-only in NX 2506. Cached preview placement
        # can have empty UI selections; create_hole validates the actual inputs.
        self.blocks['face'].StepStatusAsString = 'Optional'
        self.blocks['center'].StepStatusAsString = 'Optional'
        # 模板原来的 Width/NumberOfColumns 为 1，导致下拉框只显示一个字。
        for name in ("hole_type", "end_type", "screw_size"):
            self.blocks[name].NumberOfColumns = 18
        self.blocks["screw_size"].SetEnumMembers(["自定义"] + list(SCREW_PRESETS))
        self.blocks["screw_size"].ValueAsString = "自定义"
        self.blocks["hole_type"].SetEnumMembers(["普通腰孔", "台阶腰孔"])
        self.blocks["hole_type"].ValueAsString = "普通腰孔"
        self.blocks["end_type"].SetEnumMembers(["通孔", "盲孔（平底）"])
        self.blocks["end_type"].ValueAsString = "通孔"
        self.blocks['face'].AutomaticProgression = True
        self.blocks["face"].SetSelectionFilter(
            NXOpen.Selection.SelectionAction.ClearAndEnableSpecific,
            [NXOpen.Selection.MaskTriple(70, 0, 20)])
        self.blocks["diameter"].Value = 6.0
        self.blocks["depth"].Value = 20.0
        self.blocks["cb_diameter"].Value = 10.0
        self.blocks["cb_depth"].Value = 5.0
        self.blocks["slot_length"].Value = 5.0
        self.blocks["slot_angle"].Value = 0.0
        self.blocks["reverse"].Value = False
        self.blocks['preview'].Value = True
        self.blocks['preview'].Enable = True
        # 与 Siemens ListPointProperties 示例一致，使用原生多点/草图定位。
        # 官方示例注明 SuperPoint 只支持 0x8(仅点)；现场反馈此时圆边不出现推断点，
        # 故放开为 0xC(点 + 边/曲线)，让光标能吸附到圆边并由 NX 推断出圆心。
        # 若当前版本拒绝 0xC，则回退到官方取值 0x8，避免整个对话框初始化失败。
        try:
            self.blocks["center"].EntityType = 0xC
        except Exception:
            self.blocks["center"].EntityType = 0x8
        self.blocks["center"].CurveRules = 0x481
        self.blocks["center"].DefaultCurveRulesAsString = "Infer Curves"
        self.blocks["center"].AutomaticProgression = False
        self.blocks["center"].PointOverlay = True
        self.blocks["center"].SketchOnPath = False
        self.blocks["center"].CreateInterpartLink = False
        self.blocks["center"].InterpartSelectionAsString = "Simple"
        self.blocks["center"].LabelString = "指定点"
        self.blocks["center"].Cue = "连续指定点，或使用草图入口在入口平面画点"
        self.blocks["center"].SnapPointTypesEnabled = 0x1ffff
        self.blocks["center"].SnapPointTypesOnByDefault = 0x298
        if self.edit_data:
            self.dialog.TopBlock.Label = '编辑快捷腰孔参数'
            for name in ('diameter','depth','cb_diameter','cb_depth','slot_angle'):
                self.blocks[name].Value = self.edit_data[name]
            self.blocks['slot_length'].Value = self.edit_data['span']
            self.blocks['reverse'].Value = self.edit_data['reverse']
            self.blocks['screw_size'].ValueAsString = self.edit_data['screw_size']
            self.blocks['hole_type'].ValueAsString = ('台阶腰孔' if self.edit_data['counterbore'] else '普通腰孔')
            self.blocks['hole_type'].Enable = False
            self.blocks['end_type'].ValueAsString = '通孔' if self.edit_data['through'] else '盲孔（平底）'
            face = self.edit_feature.FeatureData.CustomTagAttributeByName('entry_face').Value
            self.blocks['face'].SetSelectedObjects([face])
            self.set_hint('修改尺寸后更新整个孔组；沿用原中心，可重选相同数量的点。孔类型保持不变。')
        self.sync_fields()
        self.update_point_count()
        self.busy = False

    def dialog_shown(self):
        if not self.edit_feature:
            self.blocks['face'].Focus()
        self.set_hint('预览保持勾选；选好位置和参数后点击“显示结果”，检查后点击“应用”。')

    def focus_notify(self, block, focus):
        if block == self.blocks.get('center'):
            self.point_selection_active = bool(focus)

    def keyboard_focus_notify(self, block, focus):
        # Focusing a parameter never creates a preview; the result button does.
        return

    def set_label(self, name, text):
        props = self.blocks[name].GetProperties()
        try:
            props.SetString("Label", text)
        finally:
            props.Dispose()

    def set_hint(self, text):
        self.set_label("hint", text)

    def element_center(self, element):
        """整条圆弧/圆形边被整体选中时取圆心；NX 推断出的点不会走到这里。"""
        if isinstance(element, NXOpen.Arc):
            return element.CenterPoint
        if not isinstance(element, (NXOpen.Edge, NXOpen.Curve)):
            raise ValueError("所选对象不是点、圆弧或曲线，无法确定孔中心；请重新选择。")
        eval_id = self.uf.Eval.Initialize(element.Tag)
        try:
            if not self.uf.Eval.IsArc(eval_id):
                raise ValueError("所选边/曲线不是圆弧，取不到圆心；请把鼠标停在圆弧上，"
                                 "等圆心标记出现后再点选。")
            arc = self.uf.Eval.AskArc(eval_id)
            return NXOpen.Point3d(*[float(value) for value in arc.Center])
        finally:
            # NX 2506 的 Python UF 未导出 Eval.Free（实测 dir(uf.Eval) 无此方法，
            # PointerWrapper 也没有 __del__）；有则成对释放，没有则交由会话回收。
            free = getattr(self.uf.Eval, "Free", None)
            if free is not None:
                free(eval_id)

    def selected_point_objects(self):
        """SuperPoint 返回 Section，需遍历 Section 才能取得实际定位点。"""
        points = []
        seen = set()

        def append_point(point):
            if not isinstance(point, NXOpen.Point):
                return
            if point.IsOccurrence or point.OwningPart != self.part:
                raise ValueError("请选择当前工作零件的定位点；组件请先设为工作零件。")
            if point.Tag not in seen:
                seen.add(point.Tag)
                points.append(point)

        def point_for(element):
            # 推断点已由 NX 落成 Point，直接用，不再做几何计算覆盖吸附结果。
            if isinstance(element, NXOpen.Point):
                return element
            center = self.element_center(element)
            cached = self.derived_points.get(element.Tag)
            if cached is not None:
                try:
                    coords = cached.Coordinates
                    if all(abs(float(value) - float(target)) < 1e-9
                           for value, target in zip((coords.X, coords.Y, coords.Z),
                                                     (center.X, center.Y, center.Z))):
                        return cached
                except (NXOpen.NXException, SystemError):
                    pass
            point = self.part.Points.CreatePoint(center)
            self.owned_position_tags.add(point.Tag)
            self.derived_points[element.Tag] = point
            return point

        for obj in self.blocks["center"].GetSelectedObjects():
            if isinstance(obj, NXOpen.Section):
                for data in obj.GetSectionData():
                    for element in data.GetSectionElementsData():
                        append_point(point_for(element.GetSectionElementData1()[0]))
            elif isinstance(obj, NXOpen.Point):
                append_point(obj)
            else:
                raise ValueError("无法读取定位对象，请在指定点控件中选择点或草图点。")
        return points

    def update_point_count(self):
        placement = getattr(self, 'preview_placement', None)
        count = len(placement['points']) if placement else len(self.selected_point_objects())
        if not count and self.edit_data and not self.replace_centers:
            count = len(self.edit_data['points'])
        self.set_label("point_count", "已指定 {} 个孔中心；可连续选点或进入草图画点。".format(count))

    def selected_centers(self):
        placement = getattr(self, 'preview_placement', None)
        if placement:
            return [tuple(coords) for coords in placement['points']]
        centers = []
        for point in self.selected_point_objects():
            coord = point.Coordinates
            values = (float(coord.X), float(coord.Y), float(coord.Z))
            if not all(math.isfinite(v) for v in values):
                raise ValueError("孔中心坐标必须为有限数值。")
            centers.append(values)
        if not centers:
            if self.edit_data and not self.replace_centers:
                return [tuple(coords) for coords in self.edit_data['points']]
            raise ValueError("请在“指定点”中连续选点，或进入草图画点后完成草图。")
        return centers

    def sync_fields(self):
        custom = self.blocks["screw_size"].ValueAsString == "自定义"
        slot = self.blocks["hole_type"].ValueAsString in ("普通腰孔", "台阶腰孔")
        cb = self.blocks["hole_type"].ValueAsString == "台阶腰孔"
        self.blocks["slot_length"].Enable = slot
        self.blocks["slot_angle"].Enable = slot
        self.blocks["diameter"].Enable = custom
        self.blocks["cb_diameter"].Enable = cb and custom
        self.blocks["cb_depth"].Enable = cb and custom
        self.blocks["depth"].Enable = self.blocks["end_type"].ValueAsString != "通孔"

    def apply_screw_preset(self):
        size = self.blocks["screw_size"].ValueAsString
        if size == "自定义":
            text = "自定义：可编辑尺寸。螺丝类型：ISO 4762。"
        else:
            shaft, head, height, hole, recess = SCREW_PRESETS[size]
            scale = (1.0 / 25.4 if self.part.PartUnits == NXOpen.BasePart.Units.Inches else 1.0)
            self.blocks["diameter"].Value = hole * scale
            self.blocks["cb_diameter"].Value = recess * scale
            self.blocks["cb_depth"].Value = (height + 0.5) * scale
            # 圆心距独立于孔宽，切换规格时保留输入。
            text = "{}：杆径 {} / 头径 {} / 头高 {} mm".format(
                size, shaft, head, height)
        props = self.blocks["screw_info"].GetProperties()
        try:
            props.SetString("Label", text)
        finally:
            props.Dispose()

    def update(self, block):
        if self.busy:
            return 0
        self.busy = True
        try:
            was_previewing = self.preview_state is not None
            self.blocks['preview'].Value = True
            if block == self.blocks['preview']:
                return 0
            self.clear_preview()
            if block in (self.blocks['face'], self.blocks['center']):
                self.preview_placement = None
            if block == self.blocks['screw_size']:
                self.apply_screw_preset()
            if block == self.blocks['clear_centers']:
                self.blocks['center'].SetSelectedObjects([])
                self.drop_position_markers(update_ui=False)
                self.preview_placement = None
                self.replace_centers = True
            self.update_point_count()
            self.sync_fields()
            if block == self.blocks['face'] and len(self.blocks['face'].GetSelectedObjects()) == 1:
                self.point_selection_active = True
                self.blocks['center'].Focus()
            if block == self.blocks['show_result'] and not was_previewing:
                status = self.create_hole(preview=True)
                if status:
                    self.message(self.last_preview_error or '请先完成入口面、孔中心和尺寸设置。')
                    return status
                self.lock_placement(True)
                self.set_label('show_result', '撤销')
                self.set_hint('正在显示临时结果；可点击“撤销”保留孔中心，检查完成后点击“应用”。')
            elif block == self.blocks['show_result']:
                self.set_hint('临时结果已撤销，孔中心已保留；修改参数后可再次点击“显示结果”。')
            else:
                self.set_hint('位置和参数已保留；点击“显示结果”查看，检查完成后点击“应用”。')
        except Exception as error:
            self.message(error)
            return 1
        finally:
            self.busy = False
        return 0

    def apply(self):
        if self.busy:
            return 1
        self.busy = True
        try:
            self.clear_preview(update_ui=False)
            result = self.create_hole()
            if result == 0:
                self.drop_position_markers()
                self.blocks['preview'].Value = True
                self.set_label('show_result', '显示结果')
                self.lock_placement(False)
            return result
        except Exception as error:
            self.message(error)
            return 1
        finally:
            self.busy = False

    def lock_placement(self, locked):
        # Keep placement stable while the temporary cut changes the entry face.
        for name in ('face', 'center', 'clear_centers'):
            self.blocks[name].Enable = not locked

    def restore_position_selection(self, placement):
        tags = placement.get('point_tags', [])
        points = []
        for index, coords in enumerate(placement['points']):
            point = None
            if index < len(tags):
                try:
                    candidate = NXOpen.TaggedObjectManager.GetTaggedObject(tags[index])
                    if isinstance(candidate, NXOpen.Point):
                        p = candidate.Coordinates
                        if sum((v-float(c))**2 for v,c in zip((p.X,p.Y,p.Z),coords)) < 1e-12:
                            point = candidate
                except (NXOpen.NXException, SystemError):
                    # NX's tag manager returns NULL without a Python exception
                    # for tags removed by native selection/cleanup operations.
                    pass
            if point is None:
                point = self.part.Points.CreatePoint(NXOpen.Point3d(*map(float,coords)))
                point.Unblank()
                if not hasattr(self, 'owned_position_tags'):
                    self.owned_position_tags = set()
                self.owned_position_tags.add(point.Tag)
            points.append(point)
        placement['point_tags'] = [point.Tag for point in points]
        # SuperPoint accepts one Section containing points, not a Point array.
        # See Siemens ListPointProperties.GetSelectedPoints.
        self.drop_position_sections()
        tolerance = 0.001 / 25.4 if self.part.PartUnits == NXOpen.BasePart.Units.Inches else 0.001
        section = self.part.Sections.CreateSection(tolerance, tolerance, 0.5)
        if not hasattr(self, 'owned_position_section_tags'):
            self.owned_position_section_tags = set()
        self.owned_position_section_tags.add(section.Tag)
        section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyPoints)
        rule = self.part.ScRuleFactory.CreateRuleCurveDumbFromPoints(points)
        section.AddToSection([rule], points[0], None, None, points[0].Coordinates,
                             NXOpen.Section.Mode.Create, False)
        self.blocks['center'].SetSelectedObjects([section])
        self.update_point_count()

    def drop_position_sections(self, update_ui=True):
        tags = getattr(self, 'owned_position_section_tags', set())
        if not tags:
            return
        if update_ui:
            self.blocks['center'].SetSelectedObjects([])
        for tag in tags:
            try:
                section = NXOpen.TaggedObjectManager.GetTaggedObject(tag)
                if isinstance(section, NXOpen.Section):
                    section.Destroy()
            except (NXOpen.NXException, SystemError):
                # Native SuperPoint may already have destroyed its section.
                pass
        self.owned_position_section_tags = set()

    def drop_position_markers(self, update_ui=True):
        self.drop_position_sections(update_ui=update_ui)
        tags = getattr(self, 'owned_position_tags', set())
        if not tags:
            return
        if update_ui:
            self.blocks['center'].SetSelectedObjects([])
        points = []
        for tag in tags:
            try:
                point = NXOpen.TaggedObjectManager.GetTaggedObject(tag)
                if isinstance(point, NXOpen.Point):
                    point.Coordinates  # Check native validity before deleting owned markers.
                    points.append(point)
            except (NXOpen.NXException, SystemError):
                pass
        if points:
            mark = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, '清理临时定位标记')
            try:
                self.session.UpdateManager.AddToDeleteList(points)
                if self.session.UpdateManager.DoUpdate(mark):
                    raise RuntimeError('临时定位标记清理失败。')
            except Exception:
                if self.session.DoesUndoMarkExist(mark, None):
                    self.session.UndoToMark(mark, None)
                raise
            finally:
                if self.session.DoesUndoMarkExist(mark, None):
                    self.session.DeleteUndoMark(mark, None)
        self.owned_position_tags = set()

    def preview_face(self, placement):
        bodies = [b for b in self.part.Bodies if b.Tag == placement['body_tag']]
        if not bodies:
            raise ValueError('预览目标实体已不存在，请重新选面。')
        candidates = []
        for face in bodies[0].GetFaces():
            data = self.geometry.call('AskFaceData', face.Tag)
            if data[0] != 22:
                continue
            n = [float(v)*data[6] for v in data[2]]
            if sum(n[i]*placement['normal'][i] for i in range(3)) < 0.99999:
                continue
            distance = abs(sum((data[1][i]-placement['plane'][i])*n[i] for i in range(3)))
            if distance < 1e-5:
                candidates.append(face)
        if len(candidates) != 1:
            raise ValueError('入口面发生变化，请关闭预览并重新选择入口面。')
        return candidates[0]

    def refresh_edit_result(self, feature, mark):
        # Native edit-with-rollback can leave the custom feature's internal cuts
        # outside the active result. DoUpdate alone then only stores parameters.
        current = self.part.CurrentFeature
        restore_tag = current.Tag if current is not None else None
        feature.MakeCurrentFeature()
        self.session.UpdateManager.MakeUpToDate([feature], mark)
        if restore_tag is not None and restore_tag != feature.Tag:
            previous = next((f for f in self.part.Features if f.Tag == restore_tag), None)
            if previous is not None:
                previous.MakeCurrentFeature()

    def clear_preview(self, update_ui=True):
        if self.preview_state is not None:
            from QuickholeFeature import commit
            state = self.preview_state
            mark = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, '清除腰孔预览')
            try:
                if state['original'] is None:
                    features = [f for f in self.part.Features if f.Tag == state['feature_tag']]
                    if features:
                        self.session.UpdateManager.AddToDeleteList(features)
                else:
                    feature = next(f for f in self.part.Features if f.Tag == state['feature_tag'])
                    feature = commit(self.part, state['original'], self.preview_face(state['original_placement']), feature)
                    self.refresh_edit_result(feature, mark)
                if self.session.UpdateManager.DoUpdate(mark):
                    raise RuntimeError('恢复预览前模型失败，请保留当前窗口并检查模型。')
                curves = [c for c in self.part.Curves if c.Tag in state['curve_tags']]
                if curves:
                    self.session.UpdateManager.AddToDeleteList(curves)
                    if self.session.UpdateManager.DoUpdate(mark):
                        raise RuntimeError('清理预览截面失败。')
                if state['original'] is not None:
                    from QuickholeFeature import show_result
                    show_result(feature)
                else:
                    for body in self.part.Bodies:
                        if body.Tag == state['body_tag']:
                            body.Unblank()
                            body.RedisplayObject()
            except Exception:
                if self.session.DoesUndoMarkExist(mark, None):
                    self.session.UndoToMark(mark, None)
                raise
            finally:
                if self.session.DoesUndoMarkExist(mark, None):
                    self.session.DeleteUndoMark(mark, None)
            self.preview_state = None
            placement = getattr(self, 'preview_placement', None)
            if update_ui:
                # Restore editability before attempting to repopulate selection.
                self.lock_placement(False)
                self.sync_fields()
                self.set_label('show_result', '显示结果')
            if placement and update_ui:
                self.blocks['face'].SetSelectedObjects([self.preview_face(placement)])
                self.restore_position_selection(placement)
        if update_ui and self.blocks:
            self.lock_placement(False)
            self.set_label('show_result', '显示结果')

    def cancel(self):
        self.busy = True
        try:
            self.clear_preview(update_ui=False)
            self.drop_position_markers(update_ui=False)
            return 0
        except Exception as error:
            self.message('取消预览失败：' + str(error))
            return 1
        finally:
            self.busy = False

    def create_hole(self, preview=False):
        mark = None
        try:
            from QuickholeFeature import CLASS_NAME, commit
            try:
                self.session.CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
            except NXOpen.NXException as error:
                try:
                    import QuickholeFeature
                    QuickholeFeature.register()
                    self.session.CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
                except Exception as register_error:
                    raise ValueError(
                        '快捷腰孔特征类尚未加载。请确认 NX/UGII/menus/custom_dirs.dat '
                        '包含 D:\\mytool-1021，然后完全退出并重启 NX。'
                    ) from register_error
            if self.session.Parts.Work != self.part:
                raise ValueError("工作零件已改变，请关闭并重新运行腰型孔向导。")
            placement = getattr(self, 'preview_placement', None)
            faces = [self.preview_face(placement)] if placement else self.blocks["face"].GetSelectedObjects()
            if len(faces) != 1 or not isinstance(faces[0], NXOpen.Face):
                raise ValueError("请选择一个实体上的平面作为孔入口面。")
            face = faces[0]
            body = face.GetBody()
            if face.IsOccurrence or body.OwningPart != self.part or not body.IsSolidBody:
                raise ValueError("请选择当前工作零件的实体面；组件请先设为工作零件。")
            points = self.selected_centers()
            if self.edit_data and len(points) != len(self.edit_data['points']):
                raise ValueError('编辑孔组时请保持原孔中心数量不变。')
            data = self.geometry.call("AskFaceData", face.Tag)
            if data[0] != 22:
                raise ValueError("第一版仅支持平面入口，请选择平面。")
            normal = tuple(float(v) * data[6] for v in data[2])
            magnitude = math.sqrt(sum(v * v for v in normal))
            if magnitude < 1e-12:
                raise ValueError("无法确定入口面法向。")
            normal = tuple(v / magnitude for v in normal)
            direction = tuple(v * (1 if self.blocks["reverse"].Value else -1) for v in normal)
            box = self.geometry.call("AskBoundingBox", body.Tag)
            diagonal = math.sqrt(sum((box[i+3] - box[i]) ** 2 for i in range(3)))
            tolerance = max(diagonal * 1e-7, 1e-5)
            hole_type = self.blocks["hole_type"].ValueAsString
            if hole_type not in ("普通腰孔", "台阶腰孔"):
                raise ValueError("请选择普通腰孔或台阶腰孔。")
            through = self.blocks["end_type"].ValueAsString == "通孔"
            cb = hole_type == "台阶腰孔"
            diameter, depth, cb_diameter, cb_depth = validate_parameters(
                self.blocks["diameter"].Value, self.blocks["depth"].Value, cb,
                self.blocks["cb_diameter"].Value, self.blocks["cb_depth"].Value, through)
            span = positive(self.blocks["slot_length"].Value, "腰孔两端圆心距")
            if span <= tolerance:
                raise ValueError("腰孔两端圆心距过小，请增大圆心距。")
            angle = float(self.blocks["slot_angle"].Value)
            if not math.isfinite(angle):
                raise ValueError("方向角必须为有限数值。")
            axis = slot_axis(normal, angle)
            # 在任何减料之前读取平面、投影全部中心并校验，避免面重建后失效。
            jobs = []
            for index, coords in enumerate(points, 1):
                center = projected_center(NXOpen.Point3d(*coords), data[1], normal)
                if any(sum((center[i] - old[i]) ** 2 for i in range(3)) <= tolerance**2
                       for old, _ in jobs):
                    raise ValueError("第 {} 个孔中心与前面的中心重复（投影后），请取消重复点或修改定位草图。".format(index))
                probe = [center[i] + direction[i] * tolerance * 5 for i in range(3)]
                if not self.edit_data and self.geometry.call("AskPointContainment", probe, body.Tag) != 1:
                    raise ValueError("第 {} 个孔中心或方向未进入实体，请检查位置和方向。".format(index))
                cut_depth = through_length(center, direction, box, tolerance) if through else depth
                if cb and cb_depth >= cut_depth:
                    raise ValueError("第 {} 个孔的台阶深度必须小于总切削深度。".format(index))
                for sign in (-1, 1):
                    end_probe = [probe[i] + sign*span/2*axis[i] for i in range(3)]
                    if not self.edit_data and self.geometry.call("AskPointContainment", end_probe, body.Tag) != 1:
                        raise ValueError("第 {} 个腰孔端部超出实体，请调整位置、圆心距或方向角。".format(index))
                jobs.append((center, cut_depth))
            parameters = dict(version=1,diameter=diameter,depth=depth,cb_diameter=cb_diameter,
                              cb_depth=cb_depth,counterbore=cb,through=through,span=span,
                              slot_angle=angle,reverse=bool(self.blocks['reverse'].Value),
                              screw_size=self.blocks['screw_size'].ValueAsString,
                              normal=list(normal),direction=list(direction),axis=list(axis),
                              points=[list(p) for p in points],
                              jobs=[[list(center),cut_depth] for center,cut_depth in jobs])
            mark = self.session.SetUndoMark(
                NXOpen.Session.MarkVisibility.Invisible if preview else NXOpen.Session.MarkVisibility.Visible,
                '腰孔临时预览' if preview else '编辑快捷腰孔' if self.edit_feature else '批量参数腰孔')
            original = None
            original_placement = None
            if preview and self.edit_feature:
                from QuickholeFeature import read
                original = read(self.edit_feature)[0]
                original_face = self.edit_feature.FeatureData.CustomTagAttributeByName('entry_face').Value
                original_data = self.geometry.call('AskFaceData', original_face.Tag)
                original_placement = dict(body_tag=original_face.GetBody().Tag, plane=list(original_data[1]),
                                          normal=[float(v)*original_data[6] for v in original_data[2]])
            if preview and placement is None:
                placement = dict(body_tag=body.Tag, plane=list(data[1]), normal=list(normal),
                                 points=[list(p) for p in points],
                                 point_tags=[p.Tag for p in self.selected_point_objects()])
                self.preview_placement = placement
            if preview:
                # Detach native selection sections before boolean topology changes.
                self.blocks['center'].SetSelectedObjects([])
                self.blocks['face'].SetSelectedObjects([])
                self.drop_position_sections(update_ui=False)
                self.sync_fields()
            old_curve_tags = {curve.Tag for curve in self.part.Curves} if preview else set()
            feature = commit(self.part,parameters,face,self.edit_feature)
            if preview:
                feature.SetName('预览_快捷腰孔_数量%d' % len(jobs))
            if self.session.UpdateManager.DoUpdate(mark):
                raise RuntimeError('腰孔更新失败，本次修改已撤销。请检查 Quickhole_error.log。')
            if self.edit_feature:
                self.refresh_edit_result(feature, mark)
            from QuickholeFeature import show_result
            show_result(feature)
            if preview:
                self.preview_state = dict(feature_tag=feature.Tag, body_tag=body.Tag, original=original,
                                         original_placement=original_placement,
                                         curve_tags={curve.Tag for curve in self.part.Curves if curve.Tag not in old_curve_tags})
                if self.session.DoesUndoMarkExist(mark, None):
                    self.session.DeleteUndoMark(mark, None)
                self.last_preview_error = None
                return 0
            if self.edit_feature:
                self.edit_feature = feature
                self.edit_data = parameters
                self.replace_centers = False
        except Exception as error:
            if mark is not None:
                try:
                    if self.session.DoesUndoMarkExist(mark, None):
                        self.session.UndoToMark(mark, None)
                        if preview and self.session.DoesUndoMarkExist(mark, None):
                            self.session.DeleteUndoMark(mark, None)
                except Exception as undo_error:
                    self.message("{}\n本批次撤销失败，请检查模型：{}".format(error, undo_error))
                    return 1
            if preview:
                self.last_preview_error = str(error)
                self.set_hint('无法预览：' + str(error))
                return 1
            self.message(error)
            return 1
        if self.edit_feature:
            self.set_hint('孔组参数已更新，可继续修改尺寸。')
            return 0
        self.preview_placement = None
        self.count += len(jobs)
        try:
            self.blocks["center"].SetSelectedObjects([])
            self.blocks["face"].SetSelectedObjects([])
            self.update_point_count()
            self.set_hint("本次完成 {} 个孔，累计 {} 个；重新选面和中心可继续。".format(len(jobs), self.count))
        except Exception:
            pass
        return 0

    def run(self):
        try:
            mode = (NXOpen.BlockStyler.BlockDialog.DialogMode.Edit if self.edit_feature
                    else NXOpen.BlockStyler.BlockDialog.DialogMode.Create)
            self.dialog.LaunchInDialogMode(mode)
        finally:
            self.busy = True
            try:
                # Launch may have already destroyed the native block objects.
                # Model cleanup must never read or write those controls here.
                self.clear_preview(update_ui=False)
                self.drop_position_markers(update_ui=False)
            finally:
                self.dialog.Dispose()


def main():
    session = NXOpen.Session.GetSession()
    ui = NXOpen.UI.GetUI()
    if session.Parts.Work is None:
        ui.NXMessageBox.Show("腰型孔向导", NXOpen.NXMessageBox.DialogType.Information, "请先打开零件。")
        return
    try:
        session.ApplicationSwitchImmediate("UG_APP_MODELING")
        QuickHoleDialog().run()
    except Exception as error:
        try:
            with open(os.path.join(_DIRECTORY, 'Quickhole_ui_error.log'), 'a', encoding='utf-8') as stream:
                stream.write('Dialog lifecycle error\n' + traceback.format_exc() + '\n')
        except OSError:
            pass
        ui.NXMessageBox.Show("腰型孔向导", NXOpen.NXMessageBox.DialogType.Error, str(error))


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.AtTermination


if __name__ == "__main__":
    main()
