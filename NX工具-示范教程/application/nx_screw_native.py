# -*- coding: utf-8 -*-
"""NX custom feature with internal construction and output features.

Input geometry is prepared by the UI. Update callbacks create/modify only
feature builders; they never create curves, launch dialogs, or call DoUpdate.
"""
import json
import math
import traceback
from pathlib import Path
import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities
import NXOpen.UF

CLASS_NAME = "Codex.Screw.Native.V2"
ARRAYS = ("outer_sections", "socket_sections", "axes", "socket_directions")
_registered = False
_display_timer_id = 0
_display_timer_proc = None
_display_pending_parts = set()
_display_pending_features = {}
_display_remaining = {}
_display_repair_running = False


def section_from_curves(part, curves, hint, tol):
    section = part.Sections.CreateSection(tol*0.95, tol, 0.5)
    section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)
    rule = part.ScRuleFactory.CreateRuleCurveDumb(curves)
    section.AddToSection([rule], curves[0], None, None, hint,
                         NXOpen.Section.Mode.Create, False)
    for curve in curves:
        curve.SetName('CODEXSCREW_PROFILE')
        curve.Blank()
    return section


def basis(direction):
    norm = math.sqrt(sum(v*v for v in direction))
    if not math.isfinite(norm) or norm < 1e-12:
        raise ValueError("请选择插入方向。")
    axis = tuple(-v/norm for v in direction)
    ref = (1.0, 0.0, 0.0) if abs(axis[0]) < 0.9 else (0.0, 1.0, 0.0)
    dot = sum(a*b for a,b in zip(ref, axis))
    radial = tuple(ref[i]-dot*axis[i] for i in range(3))
    norm = math.sqrt(sum(v*v for v in radial))
    radial = tuple(v/norm for v in radial)
    side = (axis[1]*radial[2]-axis[2]*radial[1],
            axis[2]*radial[0]-axis[0]*radial[2],
            axis[0]*radial[1]-axis[1]*radial[0])
    return radial, side, axis


def head_profile(diameter, length, kind):
    radius = diameter/2.0
    head_radius = diameter*0.85
    chamfer = min(diameter*0.08, length*0.08)
    tip = [(0.0,-length),(radius-chamfer,-length),(radius,-length+chamfer)]
    arc = None
    if kind == "沉头螺丝":
        height = head_radius-radius
        profile = tip+[(radius,0.0),(head_radius,height-chamfer*0.3),
                       (head_radius-chamfer*0.3,height),(0.0,height)]
    elif kind == "圆头螺丝":
        height = diameter*0.55
        profile = tip+[(radius,0.0),(head_radius,0.0),(0.0,height)]
        sphere = (head_radius**2+height**2)/(2.0*height)
        mid = head_radius*0.6
        arc = (len(profile)-2,(mid,height-sphere+math.sqrt(sphere**2-mid**2)))
    else:
        height = diameter
        profile = tip+[(radius,0.0),(head_radius-chamfer*0.4,0.0),
                       (head_radius,chamfer*0.4),(head_radius,height-chamfer),
                       (head_radius-chamfer,height),(0.0,height)]
    return profile, height, arc


def prepare_inputs(part, data):
    scale = 1.0/25.4 if part.PartUnits == NXOpen.BasePart.Units.Inches else 1.0
    diameter, length = float(data["model"][1:])*scale, data["length"]*scale
    radial, side, axis = basis(data["direction"])
    profile, height, arc = head_profile(diameter,length,data["kind"])
    offset = height if data["kind"] == "沉头螺丝" else 0.0
    tol = modeling_tolerance(part)
    arrays = {name: [] for name in ARRAYS}
    for coords in data["points"]:
        origin = tuple(coords[i]-offset*axis[i] for i in range(3))
        def world(r,z,s=0.0):
            return NXOpen.Point3d(*(float(origin[i]+r*radial[i]+s*side[i]+z*axis[i]) for i in range(3)))
        vertices = [world(r,z) for r,z in profile]
        curves=[]
        for i in range(len(vertices)):
            if arc and i==arc[0]:
                result=part.Curves.CreateArc(vertices[i],world(*arc[1]),vertices[i+1],False)
                curves.append(result[0] if isinstance(result,tuple) else result)
            else:
                curves.append(part.Curves.CreateLine(vertices[i],vertices[(i+1)%len(vertices)]))
        arrays["outer_sections"].append(section_from_curves(part,curves,vertices[0],tol))
        across = diameter*(0.5 if height<diameter*0.8 else 0.82)
        hex_radius = across/math.sqrt(3.0)
        hex_points=[world(hex_radius*math.cos(i*math.pi/3.0),height+tol,
                          hex_radius*math.sin(i*math.pi/3.0)) for i in range(6)]
        hex_curves=[part.Curves.CreateLine(hex_points[i],hex_points[(i+1)%6]) for i in range(6)]
        arrays["socket_sections"].append(section_from_curves(part,hex_curves,hex_points[0],tol))
        center=world(0.0,0.0)
        point=part.Points.CreatePoint(center)
        point.Blank()
        direction=part.Directions.CreateDirection(center,NXOpen.Vector3d(*axis),NXOpen.SmartObject.UpdateOption.WithinModeling)
        arrays["axes"].append(part.Axes.CreateAxis(point,direction,NXOpen.SmartObject.UpdateOption.WithinModeling))
        arrays["socket_directions"].append(part.Directions.CreateDirection(center,
            NXOpen.Vector3d(*tuple(-v for v in axis)),NXOpen.SmartObject.UpdateOption.WithinModeling))
    data["socket_depth"] = height*0.65+tol
    return arrays


def values(event):
    cf=event.GetCustomFeature()
    data=cf.FeatureData
    parameters=json.loads(data.CustomStringAttributeByName("parameters").Value)
    arrays={name:list(data.CustomTagArrayAttributeByName(name).GetValues()) for name in ARRAYS}
    return cf,parameters,arrays


def modeling_tolerance(part):
    # Millimetres or inches in the owning part, independent of screw size.
    return 0.001/25.4 if part.PartUnits == NXOpen.BasePart.Units.Inches else 0.001


def outer_builder(part,existing,arrays,index):
    tolerance=modeling_tolerance(part)
    try:
        builder=part.Features.CreateRevolveBuilder(existing)
    except NXOpen.NXException as error:
        if existing is None or error.ErrorCode != 670403:
            raise
        # Old screws were committed without an explicit revolve tolerance.
        # Repair the stored value before reopening its builder. Never update
        # the part recursively from an internal feature update callback.
        NXOpen.UF.UFSession.GetUFSession().Modl.SetFeatTolerance(
            existing.Tag,tolerance,False)
        builder=part.Features.CreateRevolveBuilder(existing)
    builder.Tolerance=tolerance
    if existing is None:
        builder.Section=arrays["outer_sections"][index]
        builder.Axis=arrays["axes"][index]
    else:
        edit_section(part,builder.Section,arrays["outer_sections"][index])
        builder.Axis.Point=arrays["axes"][index].Point
        builder.Axis.Direction=arrays["axes"][index].Direction
    builder.Limits.StartExtend.Value.RightHandSide="0"
    builder.Limits.EndExtend.Value.RightHandSide="360"
    return builder


def socket_builder(part,existing,arrays,data,index,target):
    builder=part.Features.CreateExtrudeBuilder(existing)
    builder.DistanceTolerance=modeling_tolerance(part)
    if existing is None:
        builder.Section=arrays["socket_sections"][index]
    else:
        edit_section(part,builder.Section,arrays["socket_sections"][index])
    builder.Direction=arrays["socket_directions"][index]
    builder.Limits.StartExtend.Value.RightHandSide="0"
    builder.Limits.EndExtend.Value.RightHandSide=str(data["socket_depth"])
    builder.BooleanOperation.Type=NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Subtract
    builder.BooleanOperation.SetTargetBodies([target])
    return builder


def edit_section(part,section,source):
    """Use Siemens TelescopeMirror's in-place section mapping on edit."""
    curves=list(source.GetOutputCurves())
    if not curves:
        raise ValueError('螺丝截面为空，无法更新。')
    section.PrepareMappingData()
    try:
        section.Clear()
        rule=part.ScRuleFactory.CreateRuleCurveDumb(curves)
        section.AddToSection([rule],None,None,None,NXOpen.Point3d(0.0,0.0,0.0),
                             NXOpen.Section.Mode.Create,False)
        for old,new in zip(section.GetOldOutputCurves(),section.GetOutputCurves()):
            section.MapSectionElements(old,new)
    finally:
        section.CleanMappingData()


def callback_error():
    error=__import__('sys').exc_info()[1]
    with (Path(__file__).parent/'nx_screw_native_error.log').open('a',encoding='utf-8') as stream:
        stream.write(traceback.format_exc()+'\n')
    # Preserve NX's diagnostic instead of returning generic error 1, which
    # leaves the native redefine dialog without a useful explanation.
    return getattr(error,'ErrorCode',1) or 1


def pre_update(event):
    try:
        cf,data,arrays=values(event)
        existing=list(event.GetConstructionFeatures())
        if existing:
            if len(existing)!=len(data["points"]):
                raise ValueError("编辑时请保持原插入点数量。")
            # These are transient data records. Rebuild them for this update,
            # as in NX's TelescopeMirror sample, rather than reusing records
            # returned by an earlier update/native redefine.
            records=[]
            for construction in existing:
                record=event.CreateConstructionFeatureData(construction.GetFeature())
                record.ShowInGraphicView=True
                records.append(record)
            event.SetConstructionFeatures(records)
            return 0
        part=cf.OwningPart
        construction=[]
        for i in range(len(data["points"])):
            builder=outer_builder(part,None,arrays,i)
            try:
                feature=builder.CommitFeature()
                feature.SetName("ScrewOuter_%d"%i)
                item=event.CreateConstructionFeatureData(feature)
                # Construction and socket output own the same modified body.
                # Hiding either can hide the finished screw on a native redraw.
                item.ShowInGraphicView = True
                construction.append(item)
            finally:
                builder.Destroy()
        event.SetConstructionFeatures(construction)
        return 0
    except Exception:
        return callback_error()


def post_update(event):
    try:
        cf,data,arrays=values(event)
        existing=list(event.GetOutputFeatures())
        if existing:
            records=[]
            for output in existing:
                record=event.CreateOutputFeatureData(output.GetFeature())
                record.ShowInGraphicView=True
                records.append(record)
            event.SetOutputFeatures(records)
            restore_output_layers(cf,records)
            schedule_display_repair(cf.OwningPart,cf)
            return 0
        outputs=[]
        construction=list(event.GetConstructionFeatures())
        for i in range(len(data["points"])):
            body=construction[i].GetFeature().GetBodies()[0]
            builder=socket_builder(cf.OwningPart,None,arrays,data,i,body)
            try:
                feature=builder.CommitFeature()
                feature.SetName("ScrewSocket_%d"%i)
                output=event.CreateOutputFeatureData(feature)
                output.ShowInGraphicView = True
                outputs.append(output)
            finally:
                builder.Destroy()
        event.SetOutputFeatures(outputs)
        restore_output_layers(cf,outputs)
        schedule_display_repair(cf.OwningPart,cf)
        return 0
    except Exception:
        return callback_error()


def restore_output_layers(feature, outputs):
    """Also run inside native updates after the parameter dialog has returned."""
    for output in outputs:
        for body in output.GetFeature().GetBodies():
            if not 1 <= body.Layer <= 256:
                body.Layer=feature.OwningPart.Layers.WorkLayer
                body.Unblank()


def internal_pre_update(event):
    try:
        cf,data,arrays=values(event)
        feature=event.GetFeature()
        index=int(feature.Name.rsplit('_',1)[1])
        if feature.Name.startswith('ScrewOuter_'):
            builder=outer_builder(cf.OwningPart,feature,arrays,index)
        elif feature.Name.startswith('ScrewSocket_'):
            construction=list(event.GetConstructionFeatures())
            target=construction[index].GetFeature().GetBodies()[0]
            builder=socket_builder(cf.OwningPart,feature,arrays,data,index,target)
        else:
            raise ValueError('Unknown screw internal feature')
        # NX is already updating this feature. A nested commit can re-enter
        # the native updater; only edit its parameters, as Siemens' sample does.
        builder.Destroy()
        return 0
    except Exception:
        return callback_error()


def register():
    global _registered
    if not _registered:
        cfclass=NXOpen.Session.GetSession().CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
        cfclass.AddPreUpdateHandler(pre_update)
        cfclass.AddPostUpdateHandler(post_update)
        cfclass.AddInternalFeaturePreUpdateHandler(internal_pre_update)
        _registered=True


def commit(part,data,existing=None):
    # Registered by the resident FeatureLibrary, never by the UI interpreter.
    if existing is not None:
        for name in ('outer_sections','socket_sections'):
            for section in existing.FeatureData.CustomTagArrayAttributeByName(name).GetValues():
                for curve in profile_curves(section):
                    curve.SetName('CODEXSCREW_PROFILE')
                    curve.Blank()
    arrays=prepare_inputs(part,data)
    builder=part.Features.CreateCustomFeatureBuilder(existing)
    try:
        if existing is None:
            factory=part.Features.CustomAttributeCollection
            props=[NXOpen.Features.CustomAttribute.Property.MandatoryInput]
            parameters=factory.CreateCustomStringAttribute('parameters',props)
            parameters.Value=json.dumps(data,ensure_ascii=True)
            attributes=[parameters]
            for name in ARRAYS:
                attribute=factory.CreateCustomTagArrayAttribute(name,props)
                attribute.SetValues(arrays[name])
                attributes.append(attribute)
            cls=NXOpen.Session.GetSession().CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
            feature_data=part.Features.CustomFeatureDataCollection.CreateData(cls,attributes)
        else:
            feature_data=existing.FeatureData
            old=json.loads(feature_data.CustomStringAttributeByName('parameters').Value)
            if len(old['points'])!=len(data['points']):
                raise ValueError('编辑时请保持原插入点数量。')
            feature_data.CustomStringAttributeByName('parameters').Value=json.dumps(data,ensure_ascii=True)
            for name in ARRAYS:
                feature_data.CustomTagArrayAttributeByName(name).SetValues(arrays[name])
        builder.FeatureData=feature_data
        feature=builder.CommitFeature()
        feature.SetName('螺丝_%s_L%g_数量%d'%(data['model'],data['length'],len(data['points'])))
        return feature
    finally:
        builder.Destroy()


def show_result(feature):
    """Refresh every final screw body after the native update has finished."""
    cleanup_profiles(feature.OwningPart)
    outputs=list(feature.GetOutputFeatures())
    expected=len(json.loads(feature.FeatureData.CustomStringAttributeByName('parameters').Value)['points'])
    if len(outputs)!=expected:
        raise RuntimeError('螺丝输出数量异常，请检查 nx_screw_native_error.log。')
    for construction in feature.GetConstructionFeatures():
        construction.ShowInGraphicView = True
    bodies={}
    for output in outputs:
        output.ShowInGraphicView = True
        result=list(output.GetFeature().GetBodies())
        if len(result)!=1:
            raise RuntimeError('螺丝输出实体异常。')
        bodies[result[0].Tag]=result[0]
    if len(bodies)!=expected:
        raise RuntimeError('螺丝输出实体数量与插入点数量不一致。')
    for body in bodies.values():
        # Native custom-feature construction can leave a finished body on an
        # internal layer (for example 271). Unblank does not make that layer
        # displayable. Restore only final screw bodies to a normal work layer.
        if not 1 <= body.Layer <= 256:
            body.Layer = feature.OwningPart.Layers.WorkLayer
        body.Unblank()
        body.RedisplayObject()
    # Flush the whole result after all bodies are ready; individual redraws
    # can leave an older graphics list visible during a BlockStyler callback.
    NXOpen.UF.UFSession.GetUFSession().Disp.RegenerateDisplay()
    NXOpen.Session.GetSession().DisplayManager.MakeUpToDate()


def profile_curves(section):
    """Include both rule sources and NX's evaluated section curves."""
    curves={curve.Tag:curve for curve in section.GetOutputCurves()}
    for data in section.GetSectionData():
        for element in data.GetSectionElementsData():
            source=element.GetSectionElementData1()[0]
            if isinstance(source,NXOpen.Curve):
                curves[source.Tag]=source
    return list(curves.values())


def repair_internal_layers(part, force_features=()):
    """Recover saved screw outputs stranded on NX's internal layers."""
    repaired=0
    for feature in part.Features:
        if not isinstance(feature,NXOpen.Features.CustomFeature):
            continue
        if feature.FeatureData.ClassName!=CLASS_NAME:
            continue
        force=feature.Tag in force_features
        for output in feature.GetOutputFeatures():
            for body in output.GetFeature().GetBodies():
                internal=not 1 <= body.Layer <= 256
                if internal or (force and body.IsBlanked):
                    if internal:
                        body.Layer=part.Layers.WorkLayer
                    body.Unblank()
                    body.RedisplayObject()
                    repaired+=1
    if repaired:
        NXOpen.UF.UFSession.GetUFSession().Disp.RegenerateDisplay()
        NXOpen.Session.GetSession().DisplayManager.MakeUpToDate()
    return repaired


def schedule_display_repair(part, feature=None):
    """One coalesced WM_TIMER on NX's calling UI thread, never a worker thread.

    Native redefine can move or blank bodies after the FeatureUILibrary returns.
    Three bounded passes also cover a late native redraw; the timer then stops.
    The resident module retains the timer callback until NX terminates.
    """
    import os
    if os.name!='nt':
        return False
    import ctypes
    global _display_timer_id,_display_timer_proc
    user32=ctypes.WinDLL('user32',use_last_error=True)
    timer_type=ctypes.WINFUNCTYPE(None,ctypes.c_void_p,ctypes.c_uint,
                                 ctypes.c_size_t,ctypes.c_uint32)
    user32.SetTimer.argtypes=[ctypes.c_void_p,ctypes.c_size_t,ctypes.c_uint,timer_type]
    user32.SetTimer.restype=ctypes.c_size_t
    user32.KillTimer.argtypes=[ctypes.c_void_p,ctypes.c_size_t]
    user32.KillTimer.restype=ctypes.c_int
    if _display_timer_proc is None:
        @timer_type
        def on_timer(window,message,timer_id,ticks):
            global _display_timer_id,_display_repair_running
            user32.KillTimer(None,timer_id)
            _display_timer_id=0
            pending=list(_display_pending_parts)
            _display_pending_parts.clear()
            _display_repair_running=True
            for tag in pending:
                try:
                    target=NXOpen.TaggedObjectManager.GetTaggedObject(tag)
                    if not isinstance(target,NXOpen.Part):
                        continue
                    # Native redefine may unblank section geometry after the
                    # dialog returns. Clean it only after NX finishes updating.
                    removed=cleanup_profiles(target)
                    repaired=repair_internal_layers(target,_display_pending_features.get(tag,set()))
                    states=[dict(feature=f.Tag,outputs=[dict(show=o.ShowInGraphicView,
                        bodies=[dict(tag=b.Tag,layer=b.Layer,blanked=b.IsBlanked)
                                for b in o.GetFeature().GetBodies()]) for o in f.GetOutputFeatures()])
                        for f in target.Features if isinstance(f,NXOpen.Features.CustomFeature)
                        and f.FeatureData.ClassName==CLASS_NAME]
                    with (Path(__file__).parent/'nx_screw_native_error.log').open('a',encoding='utf-8') as stream:
                        stream.write(json.dumps(dict(event='deferred_layer_repair',
                            build='native-edit-display-v3-20261113',part=target.FullPath,
                            repaired=repaired,removed_profiles=removed,features=states,
                            remaining=_display_remaining.get(tag,1)-1),ensure_ascii=False)+'\n')
                except (NXOpen.NXException,SystemError):
                    # The user may have closed the part before its timer runs.
                    pass
                except Exception:
                    try:
                        callback_error()
                    except Exception:
                        pass
                finally:
                    remaining=_display_remaining.get(tag,1)-1
                    if remaining>0:
                        _display_remaining[tag]=remaining
                        _display_pending_parts.add(tag)
                    else:
                        _display_remaining.pop(tag,None)
                        _display_pending_features.pop(tag,None)
            _display_repair_running=False
            if _display_pending_parts:
                _display_timer_id=user32.SetTimer(None,_display_timer_id,500,_display_timer_proc)
        _display_timer_proc=on_timer
    if feature is not None:
        _display_pending_features.setdefault(part.Tag,set()).add(feature.Tag)
    if _display_repair_running and part.Tag in _display_remaining:
        return True
    _display_pending_parts.add(part.Tag)
    _display_remaining[part.Tag]=3
    # Reset this same timer so the final dialog-exit request runs after any
    # earlier update request; there is only one timer with three finite passes.
    _display_timer_id=user32.SetTimer(None,_display_timer_id,250,_display_timer_proc)
    if not _display_timer_id:
        _display_pending_parts.clear()
        raise OSError(ctypes.get_last_error(),'无法安排螺丝图层刷新')
    return True


def hide_active_profiles(part):
    """Hide section inputs without removing parametric dependencies."""
    active=set()
    def protect(section):
        for curve in profile_curves(section):
            active.add(curve.Tag)
            curve.SetName('CODEXSCREW_PROFILE')
            curve.Blank()
    for custom in part.Features:
        if not isinstance(custom,NXOpen.Features.CustomFeature):
            continue
        if custom.FeatureData.ClassName!=CLASS_NAME:
            continue
        for name in ('outer_sections','socket_sections'):
            for section in custom.FeatureData.CustomTagArrayAttributeByName(name).GetValues():
                protect(section)
        for construction in custom.GetConstructionFeatures():
            builder=part.Features.CreateRevolveBuilder(construction.GetFeature())
            try:
                protect(builder.Section)
            finally:
                builder.Destroy()
        for output in custom.GetOutputFeatures():
            builder=part.Features.CreateExtrudeBuilder(output.GetFeature())
            try:
                protect(builder.Section)
            finally:
                builder.Destroy()
    return active


def cleanup_profiles(part):
    """Delete our obsolete profiles outside update callbacks; protect live inputs."""
    active=hide_active_profiles(part)
    obsolete=[curve for curve in part.Curves
              if curve.Name.upper()=='CODEXSCREW_PROFILE' and curve.Tag not in active]
    if obsolete:
        session=NXOpen.Session.GetSession()
        mark=session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible,'清理螺丝辅助曲线')
        try:
            session.UpdateManager.AddToDeleteList(obsolete)
            if session.UpdateManager.DoUpdate(mark):
                raise RuntimeError('清理旧螺丝曲线失败。')
        except Exception:
            session.UndoToMark(mark,None)
            raise
        finally:
            session.DeleteUndoMark(mark,None)
        # The delete update can restore visibility of live section curves.
        # Re-read sections after that update and hide their current geometry.
        hide_active_profiles(part)
    return len(obsolete)
