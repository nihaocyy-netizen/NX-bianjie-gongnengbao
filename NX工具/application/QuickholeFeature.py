# -*- coding: utf-8 -*-
"""Resident NX custom feature: one batch of editable capsule-shaped cuts."""
import json
import math
import os
import sys
import traceback

import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities

CLASS_NAME = 'Codex.QuickHole.Parametric.V1'
_registered = False


def tolerance(part):
    return 0.001 / 25.4 if part.PartUnits == NXOpen.BasePart.Units.Inches else 0.001


def cut_specs(parameters):
    result = []
    for center, depth in parameters['jobs']:
        result.append(dict(center=center, width=parameters['diameter'], depth=depth,
                           through=parameters['through']))
        if parameters['counterbore']:
            result.append(dict(center=center, width=parameters['cb_diameter'],
                               depth=parameters['cb_depth'], through=False))
    return result


def prepare_inputs(part, parameters):
    """All geometric inputs are created outside NX update callbacks."""
    normal, axis = parameters['normal'], parameters['axis']
    direction = parameters['direction']
    side = (normal[1]*axis[2]-normal[2]*axis[1],
            normal[2]*axis[0]-normal[0]*axis[2],
            normal[0]*axis[1]-normal[1]*axis[0])
    half_span = parameters['span'] / 2.0
    sections = []
    for spec in cut_specs(parameters):
        center, radius = spec['center'], spec['width']/2.0
        def point(u, v):
            return NXOpen.Point3d(*(float(center[i]+u*axis[i]+v*side[i]) for i in range(3)))
        a,b,c,d = (point(-half_span,radius),point(half_span,radius),
                   point(half_span,-radius),point(-half_span,-radius))
        right = part.Curves.CreateArc(b,point(half_span+radius,0.0),c,False)
        left = part.Curves.CreateArc(d,point(-half_span-radius,0.0),a,False)
        curves = [part.Curves.CreateLine(a,b),right[0] if isinstance(right,tuple) else right,
                  part.Curves.CreateLine(c,d),left[0] if isinstance(left,tuple) else left]
        tol = tolerance(part)
        section = part.Sections.CreateSection(tol*0.95,tol,0.5)
        section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)
        rule = part.ScRuleFactory.CreateRuleCurveDumb(curves)
        section.AddToSection([rule],curves[0],None,None,a,NXOpen.Section.Mode.Create,False)
        for curve in curves:
            curve.SetName('QuickHole_Profile')
            curve.Blank()
        sections.append(section)
    vector = part.Directions.CreateDirection(
        NXOpen.Point3d(*map(float,parameters['jobs'][0][0])),
        NXOpen.Vector3d(*map(float,direction)),NXOpen.SmartObject.UpdateOption.WithinModeling)
    return sections, vector


def read(feature):
    data = feature.FeatureData
    return (json.loads(data.CustomStringAttributeByName('parameters').Value),
            list(data.CustomTagArrayAttributeByName('sections').GetValues()),
            data.CustomTagAttributeByName('direction').Value,
            data.CustomTagAttributeByName('target').Value)


def replace_section(part, section, source):
    curves = list(source.GetOutputCurves())
    if not curves:
        raise ValueError('腰孔截面为空。')
    section.PrepareMappingData()
    try:
        section.Clear()
        rule = part.ScRuleFactory.CreateRuleCurveDumb(curves)
        section.AddToSection([rule],None,None,None,NXOpen.Point3d(0.0,0.0,0.0),
                             NXOpen.Section.Mode.Create,False)
        for old,new in zip(section.GetOldOutputCurves(),section.GetOutputCurves()):
            section.MapSectionElements(old,new)
    finally:
        section.CleanMappingData()


def extrude_builder(part, existing, parameters, sections, direction, target, index):
    spec = cut_specs(parameters)[index]
    builder = part.Features.CreateExtrudeBuilder(existing)
    try:
        tol = tolerance(part)
        builder.DistanceTolerance = tol
        if existing is None:
            builder.Section = sections[index]
        else:
            replace_section(part,builder.Section,sections[index])
        if existing is None:
            builder.Direction = direction
        else:
            builder.Direction.ReplaceParameters(direction)
        extend = NXOpen.GeometricUtilities.Extend.ExtendType
        builder.Limits.StartExtend.TrimType = extend.Value
        builder.Limits.StartExtend.Value.RightHandSide = format(-tol,'.15g')
        # Keep the native limit mode fixed: internal callbacks must edit existing
        # expressions rather than commit builder-only trim-mode changes. Through
        # depth is recalculated from the target bounding box by the UI on apply.
        builder.Limits.EndExtend.TrimType = extend.Value
        builder.Limits.EndExtend.Value.RightHandSide = format(spec['depth'],'.15g')
        builder.BooleanOperation.Type = NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Subtract
        builder.BooleanOperation.SetTargetBodies([target])
        return builder
    except Exception:
        builder.Destroy()
        raise


def log_error():
    error = sys.exc_info()[1]
    with open(os.path.join(os.path.dirname(__file__),'Quickhole_error.log'),'a',encoding='utf-8') as stream:
        stream.write(traceback.format_exc()+'\n')
    return getattr(error,'ErrorCode',1) or 1


def post_update(event):
    try:
        feature = event.GetCustomFeature()
        parameters, sections, direction, target = read(feature)
        outputs = list(event.GetOutputFeatures())
        if outputs:
            if len(outputs) != len(sections):
                raise ValueError('编辑时请保持孔数量和孔类型不变。')
            # All cuts share the same modified body. Hiding any output may
            # hide that body when NX reapplies graphic visibility in the UI.
            for output in outputs:
                output.ShowInGraphicView = True
            event.SetOutputFeatures(outputs)
            return 0
        for index in range(len(sections)):
            builder = extrude_builder(feature.OwningPart,None,parameters,sections,direction,target,index)
            try:
                internal = builder.CommitFeature()
                internal.SetName('QuickHoleCut_%d' % index)
                bodies = internal.GetBodies()
                if len(bodies) != 1:
                    raise ValueError('腰孔减料结果必须为单个实体。')
                target = bodies[0]
                output = event.CreateOutputFeatureData(internal)
                output.ShowInGraphicView = True
                outputs.append(output)
            finally:
                builder.Destroy()
        event.SetOutputFeatures(outputs)
        return 0
    except Exception:
        return log_error()


def internal_pre_update(event):
    try:
        custom = event.GetCustomFeature()
        parameters, sections, direction, target = read(custom)
        internal = event.GetFeature()
        index = int(internal.Name.rsplit('_',1)[1])
        if index:
            previous = event.GetOutputFeatures()[index-1].GetFeature()
            target = previous.GetBodies()[0]
        builder = extrude_builder(custom.OwningPart,internal,parameters,sections,direction,target,index)
        # NX is already updating this internal feature. Do not commit or DoUpdate.
        builder.Destroy()
        return 0
    except Exception:
        return log_error()


def register():
    global _registered
    if not _registered:
        cls = NXOpen.Session.GetSession().CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
        cls.AddPostUpdateHandler(post_update)
        cls.AddInternalFeaturePreUpdateHandler(internal_pre_update)
        _registered = True


def commit(part, parameters, face, existing=None):
    if existing is not None:
        old = read(existing)[0]
        if old['counterbore'] != parameters['counterbore'] or len(old['jobs']) != len(parameters['jobs']):
            raise ValueError('编辑时请保持孔数量和普通/台阶类型不变；请新建其他类型的孔组。')
        if read(existing)[3].Tag != face.GetBody().Tag:
            raise ValueError('编辑时请保持原目标实体不变，可选择该实体上的其他入口平面。')
    sections, direction = prepare_inputs(part,parameters)
    builder = part.Features.CreateCustomFeatureBuilder(existing)
    try:
        if existing is None:
            factory = part.Features.CustomAttributeCollection
            mandatory = [NXOpen.Features.CustomAttribute.Property.MandatoryInput]
            attr = factory.CreateCustomStringAttribute('parameters',mandatory)
            tag_array = factory.CreateCustomTagArrayAttribute('sections',mandatory)
            tags = [factory.CreateCustomTagAttribute(name,mandatory)
                    for name in ('direction','target','entry_face')]
            attr.Value = json.dumps(parameters,ensure_ascii=True)
            tag_array.SetValues(sections)
            for tag,value in zip(tags,(direction,face.GetBody(),face)):
                tag.Value = value
            cls = NXOpen.Session.GetSession().CustomFeatureClassManager.GetClassFromName(CLASS_NAME)
            data = part.Features.CustomFeatureDataCollection.CreateData(cls,[attr,tag_array]+tags)
        else:
            data = existing.FeatureData
        data.CustomStringAttributeByName('parameters').Value = json.dumps(parameters,ensure_ascii=True)
        data.CustomTagArrayAttributeByName('sections').SetValues(sections)
        for name,value in (('direction',direction),('target',face.GetBody()),('entry_face',face)):
            data.CustomTagAttributeByName(name).Value = value
        builder.FeatureData = data
        feature = builder.CommitFeature()
        feature.SetName('快捷腰孔_%s_数量%d' % (parameters['screw_size'],len(parameters['jobs'])))
        return feature
    finally:
        builder.Destroy()


def show_result(feature):
    """Display only this cut group's resulting body, never other part bodies."""
    outputs = list(feature.GetOutputFeatures())
    for output in outputs:
        output.ShowInGraphicView = True
    if not outputs:
        raise RuntimeError('腰孔没有输出减料特征。')
    bodies = list(outputs[-1].GetFeature().GetBodies())
    if len(bodies) != 1:
        raise RuntimeError('腰孔输出实体异常，无法显示预览。')
    bodies[0].Unblank()
    bodies[0].RedisplayObject()
    cleanup_profiles(feature.OwningPart)


def cleanup_profiles(part):
    """Remove only our unused profiles, retaining live parametric inputs."""
    active = set()
    for custom in part.Features:
        if not isinstance(custom, NXOpen.Features.CustomFeature):
            continue
        if custom.FeatureData.ClassName != CLASS_NAME:
            continue
        for section in read(custom)[1]:
            for data in section.GetSectionData():
                for element in data.GetSectionElementsData():
                    source = element.GetSectionElementData1()[0]
                    if isinstance(source, NXOpen.Curve):
                        active.add(source.Tag)
                        source.SetName('QuickHole_Profile')
                        source.Blank()
            for curve in section.GetOutputCurves():
                active.add(curve.Tag)
                curve.SetName('QuickHole_Profile')
                curve.Blank()
        # Protect the actual internal extrusion inputs as well.
        for output in custom.GetOutputFeatures():
            builder = part.Features.CreateExtrudeBuilder(output.GetFeature())
            try:
                for curve in builder.Section.GetOutputCurves():
                    active.add(curve.Tag)
                    curve.Blank()
            finally:
                builder.Destroy()
    obsolete = [curve for curve in part.Curves
                if curve.Name.upper() == 'QUICKHOLE_PROFILE' and curve.Tag not in active]
    if obsolete:
        session = NXOpen.Session.GetSession()
        mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, '清理腰孔辅助曲线')
        try:
            session.UpdateManager.AddToDeleteList(obsolete)
            if session.UpdateManager.DoUpdate(mark):
                raise RuntimeError('清理旧腰孔曲线失败。')
        except Exception:
            session.UndoToMark(mark, None)
            raise
        finally:
            session.DeleteUndoMark(mark, None)


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.AtTermination


if __name__ == '__main__':
    register()
