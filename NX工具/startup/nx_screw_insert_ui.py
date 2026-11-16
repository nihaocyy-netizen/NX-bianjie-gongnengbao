# -*- coding: utf-8 -*-
import os
import traceback
import math
import json
import NXOpen
import NXOpen.BlockStyler
import NXOpen.Features


class ScrewDialog:

    def initialize_cb(self):
        try:
            self.dialog.TopBlock.Label = "螺丝插入"
            for name in ("tree_control0", "buttonBack"):
                self.dialog.TopBlock.FindBlock(name).Show = False
            for name in ("enumSpec", "enumDir", "doubleJourney", "point0", "insertDirection", "preview", "show_result"):
                try:
                    self.blocks[name] = self.dialog.TopBlock.FindBlock(name)
                except Exception as exc:
                    self._log("控件 %s 初始化失败: %s" % (name, exc))
            points = self.blocks["point0"]
            points.EntityType = 0x8
            points.AutomaticProgression = False
            points.PointOverlay = True
            points.LabelString = "插入点（头部下方，可多选）"
            points.SetSelectedObjects([])
            self.blocks["insertDirection"].SetSelectedObjects([])
            self.blocks["insertDirection"].StepStatusAsString = "Required"
            self.configure_controls()
            self.blocks['preview'].Value = True
            self.blocks['preview'].Enable = True
            self.set_result_label('显示结果')
            if self.edit_data:
                self.dialog.TopBlock.Label = "编辑螺丝参数"
                self.blocks["enumSpec"].ValueAsString = self.edit_data["model"]
                self.blocks["enumDir"].ValueAsString = self.edit_data["kind"]
                self.blocks["doubleJourney"].Value = self.edit_data["length"]
                self.blocks["insertDirection"].Vector = NXOpen.Vector3d(*self.edit_data["direction"])
                self.blocks["insertDirection"].StepStatusAsString = "Satisfied"
                # 编辑沿用已保存定位点，不创建额外 NX 定位对象。
                self.blocks["point0"].StepStatusAsString = "Optional"
                self.sync_point_label()
        except Exception:
            self._log(traceback.format_exc())

    def shown_cb(self):
        try:
            from nx_screw_native import repair_internal_layers, schedule_display_repair
            part = self.session.Parts.Work
            if part is not None:
                repair_internal_layers(part)
                schedule_display_repair(part)
        except Exception:
            self._log(traceback.format_exc())

    def configure_controls(self):
        try:
            self.blocks["insertDirection"].EnableReverseDirection = True
            self.blocks["insertDirection"].LabelString = "插入方向（头部指向螺杆尖端）"
            if "enumSpec" in self.blocks:
                self.blocks["enumSpec"].SetEnumMembers(["M3", "M4", "M5", "M6", "M8", "M10", "M12"])
                self.blocks["enumSpec"].ValueAsString = "M6"
                self.blocks["enumSpec"].Label = "螺丝型号"
            if "enumDir" in self.blocks:
                self.blocks["enumDir"].SetEnumMembers(["圆柱头螺丝", "沉头螺丝", "圆头螺丝"])
                self.blocks["enumDir"].ValueAsString = "圆柱头螺丝"
                self.blocks["enumDir"].Label = "螺丝类型"
                types = self.blocks["enumDir"]
                types.PresentationStyleAsString = "Radio Box"
                types.LayoutAsString = "Horizontal"
                types.IconsOnly = True
                types.SetBitmaps([os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                              "screw_icons", name)
                                  for name in ("cylinder.bmp", "countersunk.bmp", "round.bmp")])
                types.SetBalloonTooltipTexts(["圆柱头螺丝", "沉头螺丝", "圆头螺丝"])
            if "doubleJourney" in self.blocks:
                self.blocks["doubleJourney"].Value = 20.0
                self.blocks["doubleJourney"].Label = "螺丝长度"
        except Exception:
            self._log(traceback.format_exc())

    def update_cb(self, block):
        if self.busy:
            return 0
        self.busy = True
        try:
            self.blocks['preview'].Value = True
            if block == self.blocks['preview']:
                return 0
            if self.preview_state is not None:
                self.record_display_state('before_preview_change')
            had_preview = self.preview_state is not None
            self.clear_preview()
            if block == self.blocks.get('enumDir'):
                self.sync_point_label()
            if block == self.blocks['show_result'] and not had_preview:
                return self.build_screws(preview=True)
            return 0
        except Exception as error:
            self._log(traceback.format_exc())
            self.ui.NXMessageBox.Show('螺丝预览失败',NXOpen.NXMessageBox.DialogType.Error,str(error))
            return 1
        finally:
            self.busy = False

    def set_result_label(self, text):
        props = self.blocks['show_result'].GetProperties()
        try:
            props.SetString('Label', text)
        finally:
            props.Dispose()

    def record_display_state(self, event, features=None):
        """Keep native-session evidence without changing the dialog or model."""
        try:
            import nx_screw_native as core
            part = self.session.Parts.Work
            if features is None:
                tags = {tag for tag,_ in self.preview_state['features']} if self.preview_state else set()
                features = [f for f in part.Features if f.Tag in tags]
            current = part.CurrentFeature
            records = []
            for feature in features:
                records.append(dict(tag=feature.Tag,
                    parameters=json.loads(feature.FeatureData.CustomStringAttributeByName('parameters').Value),
                    outputs=[dict(show=output.ShowInGraphicView,
                        bodies=[dict(tag=b.Tag,blanked=b.IsBlanked,layer=b.Layer)
                                for b in output.GetFeature().GetBodies()]) for output in feature.GetOutputFeatures()]))
            self._log(json.dumps(dict(event=event,build='native-edit-display-v3-20261113',
                ui=__file__,core=core.__file__,part=part.FullPath,
                current=current.Tag if current is not None else None,
                total_bodies=len(list(part.Bodies)),features=records),ensure_ascii=False))
        except Exception:
            # Diagnostic collection must never change a successful operation.
            pass

    def sync_point_label(self):
        top = self.blocks["enumDir"].ValueAsString == "沉头螺丝"
        self.blocks["point0"].LabelString = (
            "插入点（头部上方，可多选）" if top else "插入点（头部下方，可多选）")

    def ok_cb(self):
        return self.apply_cb()

    def selected_points(self):
        points = []
        seen = set()
        def append(obj):
            if isinstance(obj, NXOpen.Point) and obj.Tag not in seen:
                seen.add(obj.Tag)
                coord = obj.Coordinates
                if not all(math.isfinite(v) for v in (coord.X, coord.Y, coord.Z)):
                    raise ValueError("插入点坐标无效。")
                if not any(sum((a-b)**2 for a,b in zip(
                        (coord.X, coord.Y, coord.Z), (p.X, p.Y, p.Z))) < 1e-16 for p in points):
                    points.append(coord)
        for obj in self.blocks["point0"].GetSelectedObjects():
            if isinstance(obj, NXOpen.Section):
                for data in obj.GetSectionData():
                    for element in data.GetSectionElementsData():
                        append(element.GetSectionElementData1()[0])
            else:
                append(obj)
        if not points:
            # SuperPoint 在编辑模式中可能返回非空的 Section 容器，
            # 但容器内部没有实际点；应在解析完实体后恢复保存的位置。
            records = getattr(self, "edit_records", None)
            if records:
                return [NXOpen.Point3d(*coords) for record in records
                        for coords in record["points"]]
            if self.edit_data:
                if "points" in self.edit_data:
                    return [NXOpen.Point3d(*coords) for coords in self.edit_data["points"]]
                return [NXOpen.Point3d(*self.edit_data["point"])]
            raise ValueError("请选择一个或多个插入点。")
        return points


    def cancel_cb(self):
        try:
            self.clear_preview()
        except Exception as error:
            self._log(traceback.format_exc())
            self.ui.NXMessageBox.Show('取消预览失败',NXOpen.NXMessageBox.DialogType.Error,str(error))
            return 1
        return 0

    def _log(self, text):
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "nx_screw_error.log"), "a", encoding="utf-8") as stream:
            stream.write(str(text) + "\n")

    def show(self):
        try:
            mode = (NXOpen.BlockStyler.BlockDialog.DialogMode.Edit
                    if self.session.CustomFeatureClassManager.GetEditedCustomFeature() is not None
                    else NXOpen.BlockStyler.BlockDialog.DialogMode.Create)
            self.dialog.LaunchInDialogMode(mode)
        finally:
            try:
                self.clear_preview(update_ui=False)
            finally:
                self.dialog.Dispose()
                from nx_screw_native import schedule_display_repair
                part = self.session.Parts.Work
                if part is not None:
                    for feature in self.edit_features:
                        schedule_display_repair(part,feature)
                    schedule_display_repair(part)


class NativeScrewDialog(ScrewDialog):
    """一个原生特征拥有全部内部几何；支持多个节点统一改参数。"""
    def __init__(self):
        self.session = NXOpen.Session.GetSession()
        self.ui = NXOpen.UI.GetUI()
        self.blocks = {}
        self.edit_body = None
        self.edit_data = None
        self.edit_features = []
        self.edit_records = []
        self.preview_state = None
        self.busy = False
        from nx_screw_native import CLASS_NAME
        edited = self.session.CustomFeatureClassManager.GetEditedCustomFeature()
        if edited is not None:
            if edited.FeatureData.ClassName != CLASS_NAME:
                raise ValueError("旧版独立参数节点无法合并为原生螺丝特征，请使用新版重新生成。")
            self.edit_features = [edited]
        else:
            selection = self.ui.SelectionManager
            selected = {}
            for index in range(selection.GetNumSelectedObjects()):
                obj = selection.GetSelectedTaggedObject(index)
                if isinstance(obj, NXOpen.Features.CustomFeature):
                    features = [obj]
                else:
                    body = obj.GetBody() if isinstance(obj, NXOpen.Face) else obj
                    features = body.GetFeatures() if isinstance(body, NXOpen.Body) else []
                for feature in features:
                    if isinstance(feature, NXOpen.Features.CustomFeature) and feature.FeatureData.ClassName == CLASS_NAME:
                        selected[feature.Tag] = feature
            self.edit_features = list(selected.values())
        for feature in self.edit_features:
            self.edit_records.append(json.loads(feature.FeatureData.CustomStringAttributeByName("parameters").Value))
        self.parameter_feature = edited
        if self.edit_records:
            self.edit_data = self.edit_records[0]
        self.dialog = self.ui.CreateDialog(os.path.join(os.path.dirname(os.path.abspath(__file__)), "nx_screw_insert.dlx"))
        self.dialog.AddInitializeHandler(self.initialize_cb)
        self.dialog.AddDialogShownHandler(self.shown_cb)
        self.dialog.AddUpdateHandler(self.update_cb)
        self.dialog.AddOkHandler(self.ok_cb)
        self.dialog.AddApplyHandler(self.apply_cb)
        self.dialog.AddCancelHandler(self.cancel_cb)

    def selected_points(self):
        # 统一解析实际定位对象，不能用容器是否为空推断用户是否重新选点。
        return super().selected_points()

    def apply_cb(self):
        self.busy = True
        try:
            self.clear_preview()
            return self.build_screws()
        except Exception as error:
            self._log(traceback.format_exc())
            self.ui.NXMessageBox.Show('螺丝生成失败',NXOpen.NXMessageBox.DialogType.Error,str(error))
            return 1
        finally:
            self.busy = False

    def refresh_edit_result(self, feature, mark):
        part = feature.OwningPart
        current = part.CurrentFeature
        previous_tag = current.Tag if current is not None else None
        # During native redefine NX can be rolled back to the stock before the
        # screw. Restoring that earlier position removes the updated screws
        # from the active model. Only preserve a later history position.
        order = {item.Tag:index for index,item in enumerate(part.Features)}
        restore_later = (previous_tag is not None and
                         order.get(previous_tag,-1) > order.get(feature.Tag,-1))
        feature.MakeCurrentFeature()
        self.session.UpdateManager.MakeUpToDate([feature],mark)
        if restore_later:
            previous = next((f for f in part.Features if f.Tag == previous_tag),None)
            if previous is not None:
                previous.MakeCurrentFeature()

    def clear_preview(self, update_ui=True):
        state = self.preview_state
        if state is None:
            return
        from nx_screw_native import commit, show_result
        part = self.session.Parts.Work
        mark = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible,'撤销螺丝预览')
        try:
            restored = []
            for tag, original in state['features']:
                feature = next(f for f in part.Features if f.Tag == tag)
                if original is None:
                    self.session.UpdateManager.AddToDeleteList([feature])
                else:
                    feature = commit(part,json.loads(json.dumps(original)),feature)
                    self.refresh_edit_result(feature,mark)
                    restored.append(feature)
            if self.session.UpdateManager.DoUpdate(mark):
                raise RuntimeError('撤销螺丝预览失败。')
            if state['new']:
                # Some native deletion paths can retain an internal feature or
                # its body. Remove only tags owned by this preview generation.
                remaining = [f for f in part.Features if f.Tag in state['internal_tags']]
                remaining += [b for b in part.Bodies if b.Tag in state['body_tags']]
                if remaining:
                    self.session.UpdateManager.AddToDeleteList(remaining)
                    if self.session.UpdateManager.DoUpdate(mark):
                        raise RuntimeError('清理残留螺丝预览实体失败。')
                for tag in state['sections']:
                    try:
                        section = NXOpen.TaggedObjectManager.GetTaggedObject(tag)
                        section.Destroy()
                    except (NXOpen.NXException,SystemError):
                        pass
                objects = [obj for collection in (part.Curves,part.Points)
                           for obj in collection if obj.Tag in state['geometry']]
                for tag in state['smart_objects']:
                    try:
                        obj = NXOpen.TaggedObjectManager.GetTaggedObject(tag)
                        if obj is not None:
                            objects.append(obj)
                    except (NXOpen.NXException,SystemError):
                        pass
                if objects:
                    self.session.UpdateManager.AddToDeleteList(objects)
                    if self.session.UpdateManager.DoUpdate(mark):
                        raise RuntimeError('清理螺丝预览辅助对象失败。')
                if any(b.Tag in state['body_tags'] for b in part.Bodies):
                    raise RuntimeError('螺丝预览实体仍有残留，已恢复撤销前状态。')
            for feature in restored:
                show_result(feature)
            NXOpen.UF.UFSession.GetUFSession().Disp.RegenerateDisplay()
            self.session.DisplayManager.MakeUpToDate()
        except Exception:
            self.session.UndoToMark(mark,None)
            raise
        finally:
            self.session.DeleteUndoMark(mark,None)
        self.preview_state = None
        if update_ui:
            self.blocks['preview'].Value = True
            self.set_result_label('显示结果')

    def build_screws(self, preview=False):
        mark = None
        try:
            from nx_screw_native import CLASS_NAME, commit, basis, show_result
            part = self.session.Parts.Work
            if part is None:
                raise ValueError("请先打开工作零件。")
            try:
                feature_class = self.session.CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
            except NXOpen.NXException as error:
                # The menu can invoke this UI library directly, before the
                # FeatureLibrary action has run. Try the resident Python
                # registration once more before reporting a load failure.
                try:
                    import nx_screw_native
                    nx_screw_native.register()
                    feature_class = self.session.CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
                except Exception as register_error:
                    raise ValueError(
                        "新版整体螺丝特征尚未加载。请确认 NX/UGII/menus/custom_dirs.dat "
                        "包含 D:\\mytool-1021，然后完全退出并重启 NX。"
                    ) from register_error
            if feature_class is None:
                raise ValueError("新版整体螺丝特征尚未加载，请重启 NX。")
            points = self.selected_points()
            direction = self.blocks["insertDirection"].Vector
            direction = [float(direction.X), float(direction.Y), float(direction.Z)]
            basis(direction)
            kind = self.blocks["enumDir"].ValueAsString
            model = self.blocks["enumSpec"].ValueAsString
            length = float(self.blocks["doubleJourney"].Value)
            if not math.isfinite(length) or length not in (5, 10, 15, 20, 25, 30):
                raise ValueError("螺丝长度请选择 5、10、15、20、25、30 mm。")
            if self.edit_features:
                if any(feature.OwningPart != part for feature in self.edit_features):
                    raise ValueError("请将螺丝所在零件设为工作零件。")
                if len(points) != sum(len(record["points"]) for record in self.edit_records):
                    raise ValueError("编辑时请保持插入点总数不变；可重新选择相同数量的点。")
            mark = self.session.SetUndoMark(
                NXOpen.Session.MarkVisibility.Invisible if preview else NXOpen.Session.MarkVisibility.Visible,
                '螺丝临时预览' if preview else '整体编辑螺丝' if self.edit_features else '插入整体螺丝')
            baseline = {obj.Tag for collection in (part.Curves,part.Points) for obj in collection} if preview else set()
            baseline_bodies = {body.Tag for body in part.Bodies}
            jobs = list(zip(self.edit_features,self.edit_records)) if self.edit_features else [(None,{"points":points})]
            records = []
            features = []
            cursor = 0
            for existing, old in jobs:
                count = len(old["points"])
                positions = points[cursor:cursor+count]
                cursor += count
                data = dict(version=2, model=model, kind=kind, length=length,
                            direction=direction,
                            points=[[float(p.X),float(p.Y),float(p.Z)] for p in positions])
                feature = commit(part,data,existing)
                features.append(feature)
                records.append(data)
                self.refresh_edit_result(feature,mark)
            if self.session.UpdateManager.DoUpdate(mark):
                raise RuntimeError("整体螺丝更新失败，已撤销本次修改。")
            for feature in features:
                show_result(feature)
            body_tags = {body.Tag for feature in features for output in feature.GetOutputFeatures()
                         for body in output.GetFeature().GetBodies()}
            if not self.edit_features:
                actual_new = {body.Tag for body in part.Bodies} - baseline_bodies
                if actual_new != body_tags or len(body_tags) != len(points):
                    raise RuntimeError('螺丝实体与插入点数量不一致，本次操作已撤销。')
            if preview:
                sections = set()
                smart = set()
                if not self.edit_features:
                    for feature in features:
                        for name in ('outer_sections','socket_sections','axes','socket_directions'):
                            values = feature.FeatureData.CustomTagArrayAttributeByName(name).GetValues()
                            (sections if 'sections' in name else smart).update(obj.Tag for obj in values)
                self.preview_state = dict(
                    features=[(f.Tag,json.loads(json.dumps(old)) if existing is not None else None)
                              for f,(existing,old) in zip(features,jobs)],
                    new=not bool(self.edit_features),sections=sections,smart_objects=smart,
                    body_tags=body_tags,
                    internal_tags={item.GetFeature().Tag for feature in features
                                   for item in list(feature.GetConstructionFeatures())+list(feature.GetOutputFeatures())},
                    geometry={obj.Tag for collection in (part.Curves,part.Points)
                              for obj in collection if obj.Tag not in baseline})
                self.set_result_label('撤销')
                self.record_display_state('preview_ready',features)
                self.session.DeleteUndoMark(mark,None)
                return 0
            if self.edit_features:
                self.edit_features = features
                self.edit_records = records
                self.edit_data = records[0]
            self.record_display_state('apply_ready',features)
            return 0
        except Exception as error:
            if mark is not None and self.session.DoesUndoMarkExist(mark,None):
                self.session.UndoToMark(mark,None)
                if preview:
                    self.session.DeleteUndoMark(mark,None)
            self._log(traceback.format_exc())
            self.ui.NXMessageBox.Show("螺丝生成失败",NXOpen.NXMessageBox.DialogType.Error,str(error))
            return 1


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.AtTermination


if __name__ == "__main__":
    import sys
    directory = os.path.dirname(os.path.abspath(__file__))
    if directory not in sys.path:
        sys.path.insert(0,directory)
    # This file is launched directly by myDemo.men. Register the native
    # custom feature before constructing the dialog.
    try:
        import nx_screw_native
        nx_screw_native.register()
    except Exception:
        with open(os.path.join(directory, "nx_screw_error.log"), "a", encoding="utf-8") as stream:
            stream.write("Native screw feature registration failed\n" + traceback.format_exc() + "\n")
        raise
    NativeScrewDialog().show()
