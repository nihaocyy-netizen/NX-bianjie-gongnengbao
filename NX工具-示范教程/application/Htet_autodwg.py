# -*- coding: utf-8 -*-
"""
NX 2506 多实体自动出图 17改-2 (V29 简化流程版)
================================================
在 17改-1(V28) 基础上重写主流程：功能不变，动作大幅减少。
结论来自本机 NX2506 实测（model2.prt 4 实体 + 12 实体测试件）。

功能：
1. 选中任意多个实体，每个实体生成一套三视图（主视/侧视/俯视），
   每个视图里只显示它自己那个实体（校验环节逐个视图核对显示实体）。
2. 所有三视图横向排成一行、只占一张图纸；各组按最下面视图的底边对齐；
   一行宽度超过上限时自动整体缩小比例，保证不越出纸张。
3. 主视图取包围盒最大面，高>宽时旋转 90°，保证"长>宽"横放；
   侧视图在主视图正右、俯视图在正下方，严格对齐、互不重叠、不歪斜。
4. 自动标注：主视图标长（上方）与宽（左侧），俯视图标厚度（左侧）；
   保留一位小数；字高按视图在图纸上的尺寸自适应；标注点越界会自动换边。
5. 运行过程切到制图界面，按 PROGRESS_REFRESH_EVERY 的节奏刷新，能看到出图过程。
6. 每次运行都新建视图（名字带时间戳+随机号），重复运行不再报
   "不能用现有视图名保存"；零件停在制图状态时会先切回建模环境。
7. 结束输出逐实体校验报告（含每组的尺寸条数）。

相比 V28 的简化（每条都是实测结论，改回去就会变慢或出错）：
  1. 不再用临时大图纸：V28 因"未擦除的视图画全部实体、会越界"必须先建
     2190x2124 大图纸再缩回成品尺寸。现在建视图时其它实体是隐藏的，
     视图里只有它自己、边框就是成品尺寸，直接建成品图纸即可。
  2. 建视图当趟就隔离：三个视图建完，立刻在"其它实体仍隐藏"的状态下
     DependentDisplay.Erase(others)。这一步作为视图相关编辑被记录下来，
     之后恢复显示也不会把别的实体带回来（实测确认）。因此不需要
     "建完再逐视图擦除 + 多轮重试"。
  3. 删掉 5 轮重试：旧 isolate_body 每个视图最多 5 轮"重擦 + 刷屏 + sleep(0.4)"，
     实测第一轮就达标；而且其中 uf.Draw.UpdateOneView 单参数调用一直抛异常
     被吞掉，全是空转。现在只在真没隔离干净时才补一次。
  4. 实体显隐从 O(N²) 降到 O(N)：旧版每组都隐藏/恢复其余 N-1 个实体；
     现在开头一次性隐藏全部，每组只 Unblank 当前实体、出完再 Blank 回去。
  5. 刷新次数可调（PROGRESS_REFRESH_EVERY）：1=每组刷新最直观、0=只最后刷新最快。
  6. 标注更稳：挑边改成在候选里选最接近目标长度的一对（旧版只取最远两条，
     实体有台阶/凸台时会挑错、整条尺寸被跳过，表现就是"某个视图漏了尺寸"）；
     标注点落到纸外会自动挪到另一侧；尺寸条数计入校验报告。

实测（批处理，无图形窗口）：
  4 实体 4.12s → 2.43s；12 实体 2.41s → 2.20s，两次全部 PASS。
  GUI 下差距更大：旧版每次刷新要重画"图纸上所有视图 × 所有实体"
  （12 实体 ≈ 36 视图 × 12 实体 = 432 份几何），新版每个视图只有 1 份。

仍需保留的实测坑：
  A. 被别的实体挡住的实体，某些观察方向上 NX 连它的线条都不生成，
     事后擦掉遮挡物也变不回来（视图空白、边框不更新，重试无效）。
     所以建视图时必须让其它实体处于隐藏状态。
  B. 投影视图（ProjectedView）同样会漏画实体，三视图一律用显式定向的
     BaseView 建，对齐靠按边框中心摆位，不靠父子投影关系。
  C. 位置校正必须用 uf.Draw.AskViewBorders 的边框中心（= 内容中心），
     view.CalculateMinMaxBox / GetDrawingReferencePoint 都不是图纸坐标内容中心。
  D. uf.Drf.Object 的 AssocDwgPos 只能是 2 元素数组，AssocType 必须传
     NXOpen.UF.Drf.AssocType 枚举（传 int 报类型错误），用 END_POINT。
  E. NX 默认抑制尾零（显示 430 而不是 430.0），要一位小数必须
     OverrideDimensionText=True + SetOverriddenDimensionText(["430.0"])。
  F. 建自定义尺寸图纸必须给模板文件（MetricSheetTemplateLocation），
     否则 builder.Commit() 报"图纸页模板名称长度必须不小于一个字符"；
     也不要用 InsertSheet——GUI 环境下可能取不到 NXOpen.Drawings.DrawingSheetUnit。

用法：NX 打开零件 -> Alt+F8 -> 运行本文件 -> 选择实体 -> 确定。
"""

import math
import os
import tempfile
import time
import traceback

import NXOpen
import NXOpen.UF
import NXOpen.Drawings

# ============================================================
# 参数（毫米）
# ============================================================
VIEW_PAD = 20.0          # 视图边框相对内容的余量
CELL_GAP_X = 60.0        # 组与组之间的水平间距
MARGIN_X = 45.0          # 图纸左右页边距
MARGIN_Y = 65.0          # 图纸上下页边距

VIEW_GAP_RATIO = 0.10    # 组内视图间距比例
VIEW_GAP_MIN = 30.0
VIEW_GAP_MAX = 100.0

MIN_SCALE = 0.005
MAX_SCALE = 1.0
MAX_GROUP_SIDE = 900.0   # 单组外形上限（超出则降比例）
MAX_ROW_SHEET_W = 6000.0 # 一行排布时图纸宽度上限

# 界面刷新策略（决定"过程可见"与"速度"的取舍）：
#   1  = 每出完一组就刷新（过程最直观，实体多时最慢）
#   3  = 每 3 组刷新一次（默认，兼顾）
#   0  = 只在最后刷新一次（最快）
PROGRESS_REFRESH_EVERY = 3

INTERACTIVE_BODY_SELECTION = True
ALLOW_ALL_IF_SELECTION_CANCELLED = False
DEBUG_VIEWS = False      # True 时输出每个视图的边框/可见实体数

# 图纸模板：自定义尺寸图纸必须挂一个模板文件，否则
# builder.Commit() 会报"图纸页模板名称长度必须不小于一个字符"
A0_TEMPLATE_NAME = "Drawing-A0-Size2D-template.prt"
MANUAL_NX_TEMPLATE_DIR = r"E:\NX-2506\NX 2506\DRAFTING\templates"
PROJECTION_ANGLE = "First"      # First / Third
SHEET_REVISION = "A"

# 视图矩阵：行 = 视图 X/Y/Z 轴在模型空间的方向
MTX_FRONT = ((1.0, 0.0, 0.0), (0.0, 0.0, 1.0), (0.0, -1.0, 0.0))
MTX_RIGHT = ((0.0, 1.0, 0.0), (0.0, 0.0, 1.0), (1.0, 0.0, 0.0))
MTX_TOP = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))


# ============================================================
# 日志（与包内 Htet_common 约定一致）
# ------------------------------------------------------------
# 位置：插件根目录 logs\Htet_autodwg.log
# 编码：UTF-8 带 BOM（新建时先写 EF BB BF），任何系统码页下都不乱码
# 行格式：[YYYY-mm-dd HH:MM:SS PID=进程号] 内容
# 轮转：超过 2 MB 时 .log -> .1 -> .2 -> .3，最多保留 3 份
# 兜底：插件目录不可写时改写到 %TEMP%\Htet_autodwg.log，不影响主流程
# ============================================================
LOG_NAME = "Htet_autodwg.log"
LOG_MAX_BYTES = 2 * 1024 * 1024
LOG_KEEP = 3


def plugin_dir():
    """插件根目录：脚本在 startup/application 下，取上一层；否则用脚本目录。"""
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
    with open(path, "a", encoding="utf-8") as f:
        f.write("[%s PID=%d] %s\n"
                % (time.strftime("%Y-%m-%d %H:%M:%S"), os.getpid(), str(message)))


def log_file(message):
    try:
        _append_log(os.path.join(plugin_dir(), "logs", LOG_NAME), message)
    except Exception:
        try:
            _append_log(os.path.join(tempfile.gettempdir(), LOG_NAME), message)
        except Exception:
            pass


# ============================================================
# 基础工具
# ============================================================
def log(message):
    """流程日志：只写日志文件，不再展开 NX 信息窗口。"""
    log_file(str(message))


def _write_listing(text):
    """把一行文本写入 NX 信息窗口（需要时打开窗口）。"""
    try:
        lw = NXOpen.Session.GetSession().ListingWindow
        if not lw.IsOpen:
            lw.Open()
        lw.WriteLine(text)
    except Exception:
        # 非 936 码页的控制台可能在 print 中文时抛 UnicodeEncodeError，吞掉即可
        try:
            print(text)
        except Exception:
            pass


def log_popup(message):
    """失败/错误类提示：写日志文件，同时输出到信息窗口。"""
    text = str(message)
    log_file(text)
    _write_listing(text)


def _prop(obj, name, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


def get_solid_bodies(work_part):
    bodies = []
    try:
        all_bodies = list(work_part.Bodies)
    except Exception:
        return bodies
    for b in all_bodies:
        try:
            if b.IsSolidBody:
                bodies.append(b)
        except Exception:
            try:
                if type(b).__name__ == "Body":
                    bodies.append(b)
            except Exception:
                pass
    return bodies


def body_name(body, index):
    name = _prop(body, "Name", None)
    if name:
        return str(name)
    return "实体{}".format(index + 1)


def unique_names(bodies):
    """取实体名；同名的自动加 #序号（很多零件的实体都叫"输入1"，报告里分不清）"""
    raw = [body_name(b, i) for i, b in enumerate(bodies)]
    counts = {}
    for n in raw:
        counts[n] = counts.get(n, 0) + 1
    seq, out = {}, []
    for n in raw:
        if counts[n] > 1:
            seq[n] = seq.get(n, 0) + 1
            out.append("{}#{}".format(n, seq[n]))
        else:
            out.append(n)
    return out


def get_body_bbox(body):
    uf = NXOpen.UF.UFSession.GetUFSession()
    raw = uf.ModlGeneral.AskBoundingBox(body.Tag)
    if isinstance(raw, tuple) and len(raw) == 1:
        raw = raw[0]
    if raw and len(raw) >= 6:
        return tuple(float(x) for x in raw[:6])
    raise RuntimeError("包围盒读取失败")


def bbox_dims(bbox):
    dx = abs(bbox[3] - bbox[0])
    dy = abs(bbox[4] - bbox[1])
    dz = abs(bbox[5] - bbox[2])
    if max(dx, dy, dz) <= 1.0e-7:
        raise RuntimeError("实体包围盒尺寸为零")
    return dx, dy, dz


def extent_along_axis(bbox, axis):
    corners = [(bbox[xi], bbox[yi], bbox[zi])
               for xi in (0, 3) for yi in (1, 4) for zi in (2, 5)]
    dots = [c[0] * axis[0] + c[1] * axis[1] + c[2] * axis[2] for c in corners]
    return max(dots) - min(dots)


def select_bodies(ui, bodies):
    try:
        mask = [NXOpen.Selection.MaskTriple(NXOpen.UF.UFConstants.UF_solid_type, 0, 0)]
        result = ui.SelectionManager.SelectTaggedObjects(
            "选择需要生成工程图的实体 (可多选)",
            "选择出图实体",
            NXOpen.Selection.SelectionScope.WorkPart,
            NXOpen.Selection.SelectionAction.ClearAndEnableSpecific,
            False, False, mask)
        selected = result[1] if (isinstance(result, tuple) and len(result) >= 2) else result
        chosen, allowed = [], set(b.Tag for b in bodies)
        for obj in selected:
            if obj.Tag in allowed and obj not in chosen:
                chosen.append(obj)
        return chosen
    except Exception as ex:
        log_popup("交互选择器调用失败：{}".format(ex))
        return None


def filter_bodies(ui, bodies):
    if INTERACTIVE_BODY_SELECTION:
        selected = select_bodies(ui, bodies)
        if selected is None:
            log_popup("交互选择不可用，处理全部实体。")
        elif selected:
            return selected
        elif not ALLOW_ALL_IF_SELECTION_CANCELLED:
            return []
    return bodies


# ============================================================
# 矩阵
# ============================================================
def make_matrix3x3(rows):
    m = NXOpen.Matrix3x3()
    m.Xx, m.Xy, m.Xz = rows[0]
    m.Yx, m.Yy, m.Yz = rows[1]
    m.Zx, m.Zy, m.Zz = rows[2]
    return m


def matrix_rows(m):
    return ((m.Xx, m.Xy, m.Xz), (m.Yx, m.Yy, m.Yz), (m.Zx, m.Zy, m.Zz))


def rotate_rows_about_z(rows, angle_deg):
    """在视图平面内旋转（绕视图 Z 轴）"""
    a = math.radians(angle_deg)
    c, s = math.cos(a), math.sin(a)
    x, y, z = rows
    return (tuple(c * x[i] - s * y[i] for i in range(3)),
            tuple(s * x[i] + c * y[i] for i in range(3)), z)


def side_and_top_rows(rows):
    """由主视矩阵推导侧视/俯视矩阵（与 NX 投影视图的取法完全一致）

    主视 (vx, vy, vz)：vx=视图水平轴, vy=视图竖直轴, vz=视线反方向
    侧视（放主视正右）：水平轴=vz（厚度方向）、竖直轴=vy（高度不变）
    俯视（放主视正下）：水平轴=vx（长度不变）、竖直轴=vz（厚度方向）
    """
    vx, vy, vz = rows
    nvx = tuple(-c for c in vx)
    nvz = tuple(-c for c in vz)
    return ((vz, vy, nvx), (vx, nvz, vy))


def view_matrix_rows(view):
    m = _prop(view, "Matrix")
    if m is None:
        return None
    try:
        return matrix_rows(m)
    except Exception:
        return None


def is_orthogonal(rows, tol=1.0e-2):
    if rows is None:
        return False
    for r in rows:
        n = math.sqrt(sum(c * c for c in r))
        if abs(n - 1.0) > tol:
            return False
        if sum(1 for c in r if abs(c) > 0.9) != 1:
            return False
    return True


def view_extents(bbox, rows):
    return (extent_along_axis(bbox, rows[0]),
            extent_along_axis(bbox, rows[1]),
            extent_along_axis(bbox, rows[2]))


def choose_main_face(dims):
    """按包围盒最大面定主视方向，返回 (名称, 矩阵行, 宽, 高, 厚度)"""
    dx, dy, dz = dims
    cands = [
        ("Front", MTX_FRONT, dx, dz, dy),
        ("Right", MTX_RIGHT, dy, dz, dx),
        ("Top", MTX_TOP, dx, dy, dz),
    ]
    return max(cands, key=lambda c: c[2] * c[3])


# ============================================================
# 图纸 / 视图
# ============================================================
def _enum(cls, names, default=None):
    for n in names:
        v = getattr(cls, n, None)
        if v is not None:
            return v
    return default


def get_template_path():
    """找 NX 自带的 A0 图框模板。
    自定义尺寸图纸必须指定模板文件，否则 builder 提交时报
    "图纸页模板名称长度必须不小于一个字符"（GUI 里就是这一步挂掉的）。"""
    candidates = []
    nx_dir = os.environ.get("UGII_BASE_DIR", "")
    if nx_dir:
        candidates.append(os.path.join(nx_dir, "DRAFTING", "templates"))
    candidates.append(MANUAL_NX_TEMPLATE_DIR)
    for folder in candidates:
        path = os.path.join(folder, A0_TEMPLATE_NAME)
        if os.path.isfile(path):
            return path
    return ""


def make_sheet(work_part, length, height, name=None):
    """按指定尺寸插入一张新图纸（尺寸=自定义，单位毫米）

    首选 DrawingSheetBuilder + 模板文件：这条路在 GUI 和批处理下都稳。
    InsertSheet 放在后面兜底——某些环境下 NXOpen.Drawings.DrawingSheetUnit
    取不到（会报 module has no attribute），且不同版本重载签名不一致。
    """
    sheets = list(work_part.DrawingSheets)
    number = str(len(sheets) + 1)
    title = name or ("AD-{}".format(number))

    builder = None
    try:
        builder = work_part.DraftingDrawingSheets.CreateDraftingDrawingSheetBuilder(
            NXOpen.Drawings.DraftingDrawingSheet.Null)
        builder.AutoStartViewCreation = False
        builder.Option = NXOpen.Drawings.DrawingSheetBuilder.SheetOption.CustomSize
        builder.Height = float(height)
        builder.Length = float(length)
        builder.Units = NXOpen.Drawings.DrawingSheetBuilder.SheetUnits.Metric
        if PROJECTION_ANGLE == "First":
            builder.ProjectionAngle = \
                NXOpen.Drawings.DrawingSheetBuilder.SheetProjectionAngle.First
        else:
            builder.ProjectionAngle = \
                NXOpen.Drawings.DrawingSheetBuilder.SheetProjectionAngle.Third
        builder.Number = number
        builder.SecondaryNumber = ""
        builder.Revision = SHEET_REVISION
        builder.ScaleNumerator = 1.0
        builder.ScaleDenominator = 1.0
        template = get_template_path()
        if template:
            builder.MetricSheetTemplateLocation = template
        else:
            log_popup("警告：没找到 A0 图纸模板（UGII_BASE_DIR/DRAFTING/templates），"
                      "图纸可能建不出来")
        sheet = builder.Commit()
        sheet.Open()
        return sheet
    except Exception as ex:
        log_popup("DrawingSheetBuilder 建图纸失败，改用 InsertSheet：{}".format(ex))
    finally:
        if builder is not None:
            try:
                builder.Destroy()
            except Exception:
                pass

    unit = _enum(getattr(NXOpen.Drawings, "DrawingSheetUnit", None),
                 ("Millimeters", "Millimeter"))
    angle = _enum(getattr(NXOpen.Drawings, "DrawingSheetProjectionAngleType", None),
                  ("FirstAngle", "First"))
    if unit is None:
        raise RuntimeError("当前环境既建不出模板图纸，也取不到 DrawingSheetUnit，"
                           "请检查 UGII_BASE_DIR 下的 DRAFTING/templates 是否存在")
    sheet = work_part.DraftingDrawingSheets.InsertSheet(
        title, unit, float(length), float(height), 1.0, 1.0, angle)
    sheet.Open()
    return sheet


def find_model_view(work_part, name):
    try:
        mv = work_part.ModelingViews.FindObject(name)
        if mv:
            return mv
    except Exception:
        pass
    try:
        for mv in work_part.ModelingViews:
            if str(getattr(mv, "Name", "")).lower() == name.lower():
                return mv
    except Exception:
        pass
    return None


_tag_cache = [None]
_counter = [0]


def _run_tag():
    if _tag_cache[0] is None:
        try:
            import time
            import random
            _tag_cache[0] = "{}{:03d}".format(time.strftime("%H%M%S"), random.randint(0, 999))
        except Exception:
            _tag_cache[0] = "X"
    return _tag_cache[0]


def existing_view_names(work_part):
    names = set()
    for coll in ("ModelingViews", "Views"):
        try:
            for v in getattr(work_part, coll):
                nm = str(getattr(v, "Name", ""))
                if nm:
                    names.add(nm.lower())
        except Exception:
            pass
    return names


def create_oriented_model_view(work_part, rows):
    """定向工作视图并另存为临时建模视图（必须在建模环境调用）。
    每次运行都新建：名字带时间戳+随机号，避免"不能用现有视图名保存"。"""
    wv = work_part.ModelingViews.WorkView
    wv.Orient(make_matrix3x3(rows))

    used = existing_view_names(work_part)
    tag = _run_tag()
    last_err = None
    for _ in range(300):
        _counter[0] += 1
        name = "AD_TMP_{}_{}".format(tag, _counter[0])
        if name.lower() in used:
            continue
        try:
            work_part.Views.SaveAs(wv, name, False, False)
        except Exception as ex:
            last_err = ex
            used.add(name.lower())
            continue
        mv = find_model_view(work_part, name)
        if mv is None:
            used.add(name.lower())
            continue
        return mv
    raise RuntimeError("临时建模视图创建失败：{}".format(last_err))


def create_base_view(work_part, model_view, x, y):
    builder = work_part.DraftingViews.CreateBaseViewBuilder(NXOpen.Drawings.BaseView.Null)
    try:
        builder.Placement.Associative = False
        builder.SelectModelView.SelectedView = model_view
        builder.SecondaryComponents.ObjectType = \
            NXOpen.Drawings.DraftingComponentSelectionBuilder.Geometry.PrimaryGeometry
        builder.Style.ViewStyleBase.Part = work_part
        try:
            builder.Style.ViewStyleBase.Arrangement.SelectedArrangement = \
                NXOpen.Assemblies.Arrangement.Null
            builder.Style.ViewStyleBase.Arrangement.InheritArrangementFromParent = False
        except Exception:
            pass
        builder.Placement.Placement.SetValue(
            NXOpen.TaggedObject.Null, work_part.Views.WorkView,
            NXOpen.Point3d(float(x), float(y), 0.0))
        return builder.Commit()
    finally:
        builder.Destroy()


def create_projected_view(work_part, parent, x, y):
    """（已停用）NX 投影视图：实测会漏画实体，见 create_group_views 说明。
    保留仅便于对照排查，主流程不再调用。"""
    builder = work_part.DraftingViews.CreateProjectedViewBuilder(
        NXOpen.Drawings.ProjectedView.Null)
    try:
        builder.Parent.View.Value = parent
        builder.Style.ViewStyleBase.Part = work_part
        builder.Placement.Placement.SetValue(
            NXOpen.TaggedObject.Null, work_part.Views.WorkView,
            NXOpen.Point3d(float(x), float(y), 0.0))
        return builder.Commit()
    finally:
        builder.Destroy()


def ask_scale(view):
    try:
        uf = NXOpen.UF.UFSession.GetUFSession()
        res = uf.Draw.AskViewScale(view.Tag)
        if isinstance(res, (tuple, list)) and len(res) >= 2:
            return float(res[1])
        return float(res)
    except Exception:
        return None


def set_scale(view, target):
    actual = ask_scale(view)
    if actual is not None and abs(actual - target) < 0.002:
        return actual
    try:
        uf = NXOpen.UF.UFSession.GetUFSession()
        uf.Draw.SetViewScale(view.Tag, float(target))
    except Exception as ex:
        log("  设置比例失败：{}".format(ex))
    return ask_scale(view)


def view_borders(view):
    try:
        uf = NXOpen.UF.UFSession.GetUFSession()
        b = uf.Draw.AskViewBorders(view.Tag)
        if b and len(b) >= 4:
            return (float(b[0]), float(b[1]), float(b[2]), float(b[3]))
    except Exception:
        pass
    return None


def move_view_to_center(view, cx, cy, tol=0.3):
    for _ in range(6):
        b = view_borders(view)
        if b is None:
            return False
        dx = cx - (b[0] + b[2]) * 0.5
        dy = cy - (b[1] + b[3]) * 0.5
        if abs(dx) < tol and abs(dy) < tol:
            return True
        try:
            ref = view.GetDrawingReferencePoint()
            view.MoveView(NXOpen.Point3d(ref.X + dx, ref.Y + dy, 0.0))
        except Exception:
            return False
    return True


def debug_view(view, label):
    if not DEBUG_VIEWS or view is None:
        return
    b = view_borders(view)
    try:
        objs = view.AskVisibleObjects()
        nb = sum(1 for o in objs if type(o).__name__ == "Body")
    except Exception:
        nb = -1
    bs = "无边框" if b is None else "{:.1f}x{:.1f}".format(b[2] - b[0], b[3] - b[1])
    log("    [dbg] {} 边框 {} 可见实体 {}".format(label, bs, nb))


def blank_bodies(bodies):
    """临时隐藏实体，返回原始隐藏状态（用于还原）"""
    state = {}
    for b in bodies:
        try:
            was = bool(getattr(b, "IsBlanked", False))
            state[b.Tag] = was
            if not was:
                b.Blank()
        except Exception:
            pass
    return state


def restore_bodies(bodies, state):
    """按 blank_bodies 记录的状态精确还原（原本隐藏的重新隐藏，原本可见的恢复可见）"""
    for b in bodies:
        try:
            was_blank = state.get(b.Tag, False)
            is_blank = bool(getattr(b, "IsBlanked", False))
            if was_blank and not is_blank:
                b.Blank()
            elif not was_blank and is_blank:
                b.Unblank()
        except Exception:
            pass


def show_body(body):
    """确保该实体可见（其它实体不动）"""
    try:
        if bool(getattr(body, "IsBlanked", False)):
            body.Unblank()
    except Exception:
        pass


def hide_body(body):
    """确保该实体隐藏（其它实体不动）"""
    try:
        if not bool(getattr(body, "IsBlanked", False)):
            body.Blank()
    except Exception:
        pass


def view_body_tags(view):
    """视图里当前画出来的实体 tag 集合（空集=视图是空白的）"""
    try:
        objs = view.AskVisibleObjects()
    except Exception:
        return None
    tags = set()
    for o in objs:
        try:
            if type(o).__name__ in ("Body", "DraftingBody"):
                tags.add(o.Tag)
        except Exception:
            pass
    return tags


def isolate_body(view, keep_body, all_bodies, expect_w, expect_h, tol=45.0):
    """兜底修复：只在视图没隔离干净时动一次。

    正常流程在建视图时就已经隔离好了（见 create_group_views），这里绝大多数情况
    会直接返回。老版本是 5 轮"重擦 + 刷屏 + sleep(0.4)"，实测第一轮就已达标，
    多出来的 4 轮纯属拖慢（而且 uf.Draw.UpdateOneView 需要 2 个参数，
    单参数调用一直抛异常被吞掉，等于空转）。
    """
    if view is None:
        return False
    tags = view_body_tags(view)
    if tags is not None and tags == {keep_body.Tag}:
        return True
    others = [b for b in all_bodies if b.Tag != keep_body.Tag]
    if not others:
        return True
    try:
        view.DependentDisplay.Erase(others)
        view.Commit()
    except Exception as ex:
        log("  补擦除失败：{}".format(ex))
    tags = view_body_tags(view)
    if tags is not None and tags == {keep_body.Tag}:
        return True
    try:
        view.UpdateAutomaticViewBound()      # 强制按当前内容重算边框
    except Exception:
        pass
    tags = view_body_tags(view)
    return True if tags is None else (tags == {keep_body.Tag})


# ============================================================
# 布局：一行排布，各组底边对齐
# ============================================================
def group_size(rec, scale):
    ws = rec["w"] * scale
    hs = rec["h"] * scale
    ds = rec["depth"] * scale
    gx = min(max(ws * VIEW_GAP_RATIO, VIEW_GAP_MIN), VIEW_GAP_MAX)
    gy = min(max(hs * VIEW_GAP_RATIO, VIEW_GAP_MIN), VIEW_GAP_MAX)
    return (ws + gx + ds + 2.0 * VIEW_PAD,
            hs + gy + ds + 2.0 * VIEW_PAD)


def plan_row(group_list):
    """所有组横向排一行：返回一页的排版信息，组底边统一在 y=0。"""
    if not group_list:
        return []
    n = len(group_list)
    content_w = sum(g[0] for g in group_list) + CELL_GAP_X * (n - 1)
    content_h = max(g[1] for g in group_list)
    groups, x = [], 0.0
    for gw, gh in group_list:
        groups.append((x, 0.0, gw, gh))
        x += gw + CELL_GAP_X
    return [{"groups": groups, "count": n,
             "content_w": content_w, "content_h": content_h,
             "sheet_w": content_w + 2.0 * MARGIN_X,
             "sheet_h": content_h + 2.0 * MARGIN_Y}]


def group_main_center(page, index, rec):
    """组左下角 -> 主视图中心（图纸坐标，y 向上）；主视图顶边即组顶边"""
    gx, gy, gw, gh = page["groups"][index]
    ws = rec["w"] * rec["scale"]
    hs = rec["h"] * rec["scale"]
    return (MARGIN_X + gx + VIEW_PAD + ws * 0.5,
            MARGIN_Y + gy + gh - VIEW_PAD - hs * 0.5)


# ============================================================
# 尺寸标注
# ============================================================
def dot3(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def axis_unit(i):
    v = [0.0, 0.0, 0.0]
    v[i] = 1.0
    return tuple(v)


def collect_straight_edges(body, uf):
    """返回 [(edge, 中心点, 主方向轴号)]；用包围盒判断是否直线边。"""
    out = []
    try:
        edges = body.GetEdges()
    except Exception:
        return out
    for e in edges:
        try:
            raw = uf.ModlGeneral.AskBoundingBox(e.Tag)
            if isinstance(raw, tuple) and len(raw) == 1:
                raw = raw[0]
            if not raw or len(raw) < 6:
                continue
            b = [float(x) for x in raw[:6]]
            d = (abs(b[3] - b[0]), abs(b[4] - b[1]), abs(b[5] - b[2]))
            mx = max(d)
            if mx <= 1.0e-6:
                continue
            axes = [i for i in range(3) if d[i] > 0.6 * mx]
            if len(axes) != 1:
                continue
            c = ((b[0] + b[3]) * 0.5, (b[1] + b[4]) * 0.5, (b[2] + b[5]) * 0.5)
            out.append((e, c, axes[0]))
        except Exception:
            continue
    return out


def pick_edge_pair(edges, para_axis, measure_axis, target=None, tol_ratio=0.10):
    """挑一对边：方向平行 para_axis，沿 measure_axis 量出的跨度最接近 target。

    老版本只取"最远两条"，实体一旦有台阶、凸台、圆角，最远两条往往不是
    外轮廓，跨度对不上就整条尺寸跳过（表现就是"某个主视图漏了尺寸"）。
    现在把所有不同位置的边两两组合，取跨度最接近 target 的那一对。
    """
    cand = []
    for e, c, k in edges:
        if abs(dot3(axis_unit(k), para_axis)) > 0.9:
            cand.append((dot3(c, measure_axis), e))
    if len(cand) < 2:
        return None
    cand.sort(key=lambda t: t[0])

    # 同一位置的重复边只留一条（长方体每条棱都有一对重合边）
    uniq = [cand[0]]
    for p, e in cand[1:]:
        if abs(p - uniq[-1][0]) > 0.35:
            uniq.append((p, e))
    if len(uniq) < 2:
        return None

    if target is None:
        return uniq[0][1], uniq[-1][1], uniq[-1][0] - uniq[0][0]

    tol = max(1.0, tol_ratio * abs(target))
    best = None
    for i in range(len(uniq) - 1):
        for j in range(len(uniq) - 1, i, -1):
            span = uniq[j][0] - uniq[i][0]
            d = abs(span - target)
            if d <= tol and (best is None or d < best[0]):
                best = (d, uniq[i][1], uniq[j][1], span)
    if best is None:
        return None
    return best[1], best[2], best[3]


def uf_object(view, edge):
    o = NXOpen.UF.Drf.Object()
    o.ObjectTag = edge.Tag
    o.ObjectViewTag = view.Tag
    o.ObjectAssocType = NXOpen.UF.Drf.AssocType.END_POINT
    o.ObjectAssocModifier = 0
    o.Object2Tag = 0
    o.AssocDwgPos = [0.0, 0.0]
    return o


def style_dimension(work_part, dim_tag, text_h, value_text):
    try:
        dim = NXOpen.TaggedObjectManager.GetTaggedObject(dim_tag)
    except Exception:
        return False
    builder = None
    ok = False
    try:
        builder = work_part.Dimensions.CreateLinearDimensionBuilder(dim)
        try:
            builder.Style.LetteringStyle.DimensionTextSize = float(text_h)
            ok = True
        except Exception:
            pass
        # 强制一位小数（NX 默认抑制尾零，430 而不是 430.0）
        try:
            builder.Style.DimensionStyle.DimensionValuePrecision = 1
        except Exception:
            pass
        try:
            builder.Style.DimensionStyle.OverrideDimensionText = True
            builder.Style.DimensionStyle.SetOverriddenDimensionText([value_text])
            ok = True
        except Exception as ex:
            log("    尺寸文本设置失败：{}".format(ex))
        if ok:
            builder.Commit()
        return ok
    except Exception as ex:
        log("    尺寸样式设置失败：{}".format(ex))
        return False
    finally:
        if builder is not None:
            try:
                builder.Destroy()
            except Exception:
                pass


def make_dimension(work_part, uf, view, edge1, edge2, horizontal, origin, text_h, value):
    try:
        if horizontal:
            tag = uf.Drf.CreateHorizontalDim(
                uf_object(view, edge1), uf_object(view, edge2),
                NXOpen.UF.Drf.Text(),
                [float(origin[0]), float(origin[1]), 0.0])
        else:
            tag = uf.Drf.CreateVerticalDim(
                uf_object(view, edge1), uf_object(view, edge2),
                NXOpen.UF.Drf.Text(),
                [float(origin[0]), float(origin[1]), 0.0])
    except Exception as ex:
        log("    尺寸创建失败：{}".format(ex))
        return None
    style_dimension(work_part, tag, text_h, "{:.1f}".format(value))
    return tag


def annotate_group(work_part, entry, text_h):
    """主视图标长/宽，俯视图标厚度。返回 (主视图条数, 俯视图条数) 供校验。

    标注点会自动避开图纸边界：原来固定放在"主视图上方 / 左侧"，
    若该组已经贴到纸边，尺寸虽然建得出来但落在纸外，看起来就是"漏了尺寸"。
    """
    rec = entry["rec"]
    xv, yv, zv = rec["rows"][0], rec["rows"][1], rec["rows"][2]
    uf = NXOpen.UF.UFSession.GetUFSession()
    edges = collect_straight_edges(rec["body"], uf)
    if not edges:
        log("  跳标注：没有可用的直线边")
        return 0, 0

    main_view, top_view = entry["main"], entry["top"]
    bm = view_borders(main_view)
    bt = view_borders(top_view) if top_view is not None else None
    scale = entry["scale"]
    gap = max(25.0, 0.10 * max(rec["h"] * scale, 20.0))

    sheet = None
    try:
        sheet = work_part.DrawingSheets.CurrentDrawingSheet
    except Exception:
        sheet = None
    sw = float(_prop(sheet, "Length", 1.0e9)) if sheet is not None else 1.0e9
    shh = float(_prop(sheet, "Height", 1.0e9)) if sheet is not None else 1.0e9

    def inside(x, y, pad=1.0):
        return -pad <= x <= sw + pad and -pad <= y <= shh + pad

    made_main = 0
    if bm is None:
        log("  主视图无边框，跳过标注")
    else:
        # 长：边平行视图 Y，沿视图 X 量两端
        pair = pick_edge_pair(edges, yv, xv, target=rec["w"])
        if pair is None:
            log("  [警告] 主视图长向找不到边对")
        else:
            cy = (bm[0] + bm[2]) * 0.5
            y = bm[3] + gap
            if not inside(cy, y):
                y = bm[1] - gap                      # 上方出界就挪到下方
            if make_dimension(work_part, uf, main_view, pair[0], pair[1], True,
                              (cy, y), text_h, pair[2]):
                made_main += 1
        # 宽：边平行视图 X，沿视图 Y 量两端
        pair = pick_edge_pair(edges, xv, yv, target=rec["h"])
        if pair is None:
            log("  [警告] 主视图宽向找不到边对")
        else:
            cy = (bm[1] + bm[3]) * 0.5
            x = bm[0] - gap
            if not inside(x, cy):
                x = bm[2] + gap                      # 左侧出界就挪到右侧
            if make_dimension(work_part, uf, main_view, pair[0], pair[1], False,
                              (x, cy), text_h, pair[2]):
                made_main += 1

    made_top = 0
    if bt is None:
        log("  [警告] 俯视图无边框，跳过厚度标注")
    else:
        pair = pick_edge_pair(edges, xv, zv, target=rec["depth"])
        if pair is None:
            log("  [警告] 俯视图厚度向找不到边对")
        else:
            cy = (bt[1] + bt[3]) * 0.5
            x = bt[0] - gap
            if not inside(x, cy):
                x = bt[2] + gap
            if make_dimension(work_part, uf, top_view, pair[0], pair[1], False,
                              (x, cy), text_h, pair[2]):
                made_top += 1

    log("  标注：主视图 长/宽 {}/2，俯视图 厚度 {}/1".format(made_main, made_top))
    return made_main, made_top


def refresh_ui(session, ui):
    """刷新显示，让逐视图生成的过程可见"""
    try:
        uf = NXOpen.UF.UFSession.GetUFSession()
        uf.Disp.RegenerateDisplay()
        return True
    except Exception:
        pass
    for fn in ("Refresh", "Update"):
        try:
            getattr(ui, fn)()
            return True
        except Exception:
            pass
    return False


# ============================================================
# 建组 / 完成组
# ============================================================
def create_group_views(work_part, rec, all_bodies, main_cx, main_cy, page_index, page):
    """在成品位置上建出主视/侧视/俯视三个视图，并当场把其它实体登记为"已擦除"。

    调用前必须保证：其它实体都已隐藏、只有 rec["body"] 可见
    （由 main 的"全部隐藏 → 逐组只显示当前实体"统一控制，避免每组都
      隐藏/恢复 N-1 个实体，那是 O(N²) 次建模操作，实体多时极慢）。

      1. 只有目标实体可见：被大平板完全压住的实体，NX 在某个观察方向上
         根本不为它生成线条（视图空白、边框不更新），隐藏遮挡物即可避免；
      2. 因此视图里只有它自己，边框就是成品尺寸 → 直接用成品图纸，
         不需要临时大图纸，也不需要事后再逐视图擦除；
      3. 立刻 Erase 其它实体：这一步作为"视图相关编辑"记在视图上，
         之后再恢复显示，别的实体也不会被带进这三个视图（已实测）。
    """
    others = [b for b in all_bodies if b.Tag != rec["body"].Tag]
    scale = rec["scale"]
    ws, hs, ds = rec["w"] * scale, rec["h"] * scale, rec["depth"] * scale
    gap_x = min(max(ws * VIEW_GAP_RATIO, VIEW_GAP_MIN), VIEW_GAP_MAX)
    gap_y = min(max(hs * VIEW_GAP_RATIO, VIEW_GAP_MIN), VIEW_GAP_MAX)

    main_view = create_base_view(work_part, rec["mv_main"], main_cx, main_cy)
    if main_view is None:
        raise RuntimeError("主视图创建失败")
    scale = set_scale(main_view, scale) or scale
    ws, hs, ds = rec["w"] * scale, rec["h"] * scale, rec["depth"] * scale
    gap_x = min(max(ws * VIEW_GAP_RATIO, VIEW_GAP_MIN), VIEW_GAP_MAX)
    gap_y = min(max(hs * VIEW_GAP_RATIO, VIEW_GAP_MIN), VIEW_GAP_MAX)

    side_cx = main_cx + ws * 0.5 + gap_x + ds * 0.5
    top_cy = main_cy - hs * 0.5 - gap_y - ds * 0.5
    side_view = create_base_view(work_part, rec["mv_side"], side_cx, main_cy)
    top_view = create_base_view(work_part, rec["mv_top"], main_cx, top_cy)
    if side_view is None or top_view is None:
        raise RuntimeError("侧视图/俯视图创建失败")
    for v in (side_view, top_view):
        set_scale(v, scale)

    if others:
        for v in (main_view, side_view, top_view):
            try:
                v.DependentDisplay.Erase(others)
                v.Commit()
            except Exception as ex:
                log("  视图隔离失败：{}".format(ex))

    return {"rec": rec, "page": page, "page_index": page_index,
            "main": main_view, "side": side_view, "top": top_view,
            "scale": scale, "ws": ws, "hs": hs, "ds": ds,
            "main_target": (main_cx, main_cy),
            "side_target": (side_cx, main_cy),
            "top_target": (main_cx, top_cy)}


def finish_group(entry, all_bodies, verify_rows, all_rects, ui=None):
    rec = entry["rec"]
    page = entry["page"]
    page_index = entry["page_index"]
    name = rec["name"]
    ws, hs, ds, scale = entry["ws"], entry["hs"], entry["ds"], entry["scale"]
    main_view, side_view, top_view = entry["main"], entry["side"], entry["top"]
    log("隔离并定位 {} 三视图...".format(name))

    ok_main = isolate_body(main_view, rec["body"], all_bodies, ws, hs)
    ok_side = isolate_body(side_view, rec["body"], all_bodies, ds, hs)
    ok_top = isolate_body(top_view, rec["body"], all_bodies, ws, ds)

    # 视图建的时候就在目标位置上，这里只做一次微调（原来跑两轮是冗余）
    move_view_to_center(main_view, entry["main_target"][0], entry["main_target"][1])
    move_view_to_center(side_view, entry["side_target"][0], entry["side_target"][1])
    move_view_to_center(top_view, entry["top_target"][0], entry["top_target"][1])
    debug_view(main_view, "主视图")
    debug_view(side_view, "侧视图")
    debug_view(top_view, "俯视图")

    # 自动标注（字高随视图大小自适应）
    made_main = made_top = 0
    try:
        work_part = NXOpen.Session.GetSession().Parts.Work
        text_h = min(max(0.025 * max(ws, hs), 3.5), 20.0)
        made_main, made_top = annotate_group(work_part, entry, text_h)
    except Exception as ex:
        log("  标注失败：{}".format(ex))
        log(traceback.format_exc())

    results = {"主视图": main_view, "侧视图": side_view, "俯视图": top_view}
    boxes = {}
    for label, v in results.items():
        b = view_borders(v) if v is not None else None
        boxes[label] = b
        if b is not None:
            all_rects.append((page_index, name + ":" + label, b))

    # 视图里到底画了哪些实体：必须"只有它自己"，且不能是空白视图
    only_self = {}
    for label, v in results.items():
        tags = view_body_tags(v) if v is not None else None
        only_self[label] = tags
        if tags is not None:
            if not tags:
                log("  [警告] {} 是空白视图（没有画出任何实体）".format(label))
            else:
                log("  {} 显示实体：{}".format(label, sorted(tags)))

    problems = []
    bm, bs, bt = boxes.get("主视图"), boxes.get("侧视图"), boxes.get("俯视图")
    if bm is None:
        problems.append("主视图无边框")
    else:
        mw, mh = bm[2] - bm[0], bm[3] - bm[1]
        if mw <= mh:
            problems.append("主视图非横放 ({:.0f}x{:.0f})".format(mw, mh))
        if bs is not None:
            if (bs[2] - bs[0]) * (bs[3] - bs[1]) > mw * mh:
                problems.append("侧视图面积过大")
            if abs((bm[1] + bm[3]) * 0.5 - (bs[1] + bs[3]) * 0.5) > 1.5:
                problems.append("侧视图未水平对齐")
        if bt is not None:
            if (bt[2] - bt[0]) * (bt[3] - bt[1]) > mw * mh:
                problems.append("俯视图面积过大")
            if abs((bm[0] + bm[2]) * 0.5 - (bt[0] + bt[2]) * 0.5) > 1.5:
                problems.append("俯视图未垂直对齐")
    for label, v in (("侧视图", side_view), ("俯视图", top_view)):
        if v is None:
            problems.append(label + "缺失")
        elif not is_orthogonal(view_matrix_rows(v)):
            problems.append(label + "非正交")
    if not ok_main:
        problems.append("主视图未隔离到单实体")
    if not ok_side:
        problems.append("侧视图未隔离到单实体")
    if not ok_top:
        problems.append("俯视图未隔离到单实体")
    for label, tags in only_self.items():
        if tags is None:
            continue
        if not tags:
            problems.append(label + "是空白视图")
        elif tags != {rec["body"].Tag}:
            problems.append("{}里有别的实体({})".format(label, sorted(tags)))
    for label, b in boxes.items():
        if b is None:
            continue
        if b[0] < 0 or b[1] < 0 or b[2] > page["sheet_w"] or b[3] > page["sheet_h"]:
            problems.append(label + "超出图纸")
    if made_main < 2:
        problems.append("主视图尺寸不足 {}/2".format(made_main))
    if made_top < 1:
        problems.append("俯视图厚度尺寸缺失 0/1")
    if abs(scale - rec["scale"]) > 0.002:
        problems.append("比例异常")

    verify_rows.append({"name": name, "page": page_index, "problems": problems,
                        "rotated": rec["rotated"], "face": rec["face"],
                        "dims": (made_main, made_top)})
    log("  {}：{}".format(name, "OK" if not problems else "；".join(problems)))


# ============================================================
# 报告
# ============================================================
def report(verify_rows, all_rects, pages, errors):
    lines = ["=" * 64, "校验报告", "=" * 64]
    bad = 0
    for row in verify_rows:
        dm, dt = row.get("dims", (0, 0))
        if row["problems"]:
            bad += 1
            lines.append("[FAIL] {}（尺寸 {}/2 + {}/1）：{}".format(
                row["name"], dm, dt, "；".join(row["problems"])))
        else:
            tag = "第 {} 张，主视面 {}，{}，尺寸 {}/2 + {}/1".format(
                row["page"] + 1, row["face"],
                "横放旋转" if row["rotated"] else "原向", dm, dt)
            lines.append("[PASS] {}（{}）".format(row["name"], tag))

    # 视图互不重叠
    overlap = 0
    for i in range(len(all_rects)):
        for j in range(i + 1, len(all_rects)):
            p1, n1, r1 = all_rects[i]
            p2, n2, r2 = all_rects[j]
            if p1 != p2:
                continue
            ow = min(r1[2], r2[2]) - max(r1[0], r2[0])
            oh = min(r1[3], r2[3]) - max(r1[1], r2[1])
            if ow > 1.0 and oh > 1.0:
                overlap += 1
                lines.append("[FAIL] 第 {} 张：{} 与 {} 重叠 {:.0f}x{:.0f} mm".format(
                    p1 + 1, n1, n2, ow, oh))
    if overlap == 0 and all_rects:
        lines.append("[PASS] 所有视图互不重叠。")

    for e in errors:
        lines.append("[错误] " + e)

    lines.append("=" * 64)
    if bad == 0 and overlap == 0 and not errors:
        lines.append("结论：全部符合要求。")
    else:
        lines.append("结论：存在不符合项，见上方 FAIL。")
    lines.append("=" * 64)

    for line in lines:
        log(line)
    # 正常完成只写日志；存在不符合项时才展开信息窗口，避免正常出图被打扰。
    if bad or overlap or errors:
        for line in lines:
            _write_listing(line)


# ============================================================
# 主程序
# ============================================================
def main(args):
    session = NXOpen.Session.GetSession()
    ui = NXOpen.UI.GetUI()
    work_part = session.Parts.Work
    if work_part is None:
        log_popup("错误：请先在 NX 中打开一个零件。")
        return

    try:
        log("开始自动出图：零件 {}（{}）".format(work_part.Name, work_part.FullPath))
    except Exception:
        log("开始自动出图。")

    all_bodies = get_solid_bodies(work_part)
    if not all_bodies:
        log_popup("未找到实体 Body。")
        return

    bodies = filter_bodies(ui, all_bodies)
    if not bodies:
        # 选择器取消即结束：只落日志文件，不再展开信息窗口输出提示。
        log_file("未选择实体，程序结束。")
        return

    log("=" * 64)
    log("实体总数 {}，本次出图 {} 个。".format(len(all_bodies), len(bodies)))

    # ---------- 1. 尺寸 / 主视方向 / 比例 ----------
    records, errors = [], []
    names = unique_names(bodies)
    log("实体清单：{}".format("、".join(names)))
    for index, body in enumerate(bodies):
        name = names[index]
        try:
            bbox = get_body_bbox(body)
            dims = bbox_dims(bbox)
            face, rows, w, h, depth = choose_main_face(dims)
            rotated = False
            if h > w:
                rows = rotate_rows_about_z(rows, 90.0)
                w, h = h, w
                rotated = True
            records.append({"body": body, "name": name, "bbox": bbox, "dims": dims,
                            "face": face, "rows": rows, "w": w, "h": h,
                            "depth": depth, "rotated": rotated})
            log("{}：X={:.1f} Y={:.1f} Z={:.1f}；主视面 {} ({:.1f}x{:.1f})，厚 {:.1f}{}".format(
                name, dims[0], dims[1], dims[2], face, w, h, depth,
                "，转横放" if rotated else ""))
        except Exception as ex:
            errors.append("{}：尺寸识别失败：{}".format(name, ex))
            log(errors[-1])
    if not records:
        log_popup("没有可出图的实体。")
        return

    # ---------- 2. 比例 ----------
    for rec in records:
        gw1, gh1 = group_size(rec, 1.0)
        s = min(MAX_SCALE, MAX_GROUP_SIDE / max(gw1, 1.0), MAX_GROUP_SIDE / max(gh1, 1.0))
        rec["scale"] = max(min(s, MAX_SCALE), MIN_SCALE)
    # 一行放不下就统一缩小，保证单行且不越纸
    usable_w = max(MAX_ROW_SHEET_W - 2.0 * MARGIN_X, 1.0)
    for _ in range(60):
        sizes = [group_size(r, r["scale"]) for r in records]
        total = sum(g[0] for g in sizes) + CELL_GAP_X * (len(records) - 1)
        if total <= usable_w or len(records) <= 1:
            break
        k = (usable_w / total) * 0.985
        if k >= 0.999:
            break
        for rec in records:
            rec["scale"] = max(rec["scale"] * k, MIN_SCALE)
    for rec in records:
        g = group_size(rec, rec["scale"])
        log("{}：比例 1:{:.3f}，组外形 {:.0f}x{:.0f}".format(
            rec["name"], 1.0 / rec["scale"], g[0], g[1]))

    pages = plan_row([group_size(r, r["scale"]) for r in records])
    log("共 {} 个实体，横向一行，占 {} 张图纸（{:.0f} x {:.0f} mm）。".format(
        len(records), len(pages), pages[0]["sheet_w"], pages[0]["sheet_h"]))

    # ---------- 3. 预建定向建模视图（必须在建模环境）----------
    try:
        work_part.Drafting.ExitDraftingApplication()
    except Exception:
        pass
    try:
        session.ApplicationSwitchImmediate("UG_APP_MODELING")
    except Exception:
        pass
    try:
        for rec in records:
            rec["mv_main"] = create_oriented_model_view(work_part, rec["rows"])
            rows_side, rows_top = side_and_top_rows(rec["rows"])
            rec["mv_side"] = create_oriented_model_view(work_part, rows_side)
            rec["mv_top"] = create_oriented_model_view(work_part, rows_top)
    except Exception:
        log_popup("创建定向建模视图失败：")
        log(traceback.format_exc())
        return

    verify_rows, all_rects = [], []
    last_sheet = None

    try:
        # 跳到制图界面，让用户能看到出图过程
        session.ApplicationSwitchImmediate("UG_APP_DRAFTING")
        work_part.Drafting.EnterDraftingApplication()

        page = pages[0]
        # 成品图纸一次建好：建视图时其它实体是隐藏的，视图里只会有它自己，
        # 不会越界，所以不再需要"先建临时大图纸、出完图再缩回"那一套。
        last_sheet = make_sheet(work_part, page["sheet_w"], page["sheet_h"])
        log("图纸 1：{:.0f} x {:.0f} mm".format(page["sheet_w"], page["sheet_h"]))

        # 逐组流水线：只显示当前实体 → 建三视图 → 隔离（建视图时已完成）→
        # 定位 → 标注 → 按策略刷新。
        # 关键：一次性隐藏所有实体，然后每组只 Unblank 当前这一个再 Blank 回去，
        # 每组只动 2 个实体；老做法每组要隐藏/恢复其余 N-1 个，实体一多就是 O(N²)。
        saved = blank_bodies(all_bodies)
        every = int(PROGRESS_REFRESH_EVERY or 0)
        try:
            for i in range(page["count"]):
                rec = records[i]
                cx, cy = group_main_center(page, i, rec)
                show_body(rec["body"])
                try:
                    entry = create_group_views(work_part, rec, all_bodies, cx, cy, 0, page)
                except Exception as ex:
                    errors.append("{}：建视图失败：{}".format(rec["name"], ex))
                    log(errors[-1])
                    log(traceback.format_exc())
                    hide_body(rec["body"])
                    continue
                try:
                    finish_group(entry, all_bodies, verify_rows, all_rects, ui)
                except Exception as ex:
                    errors.append("{}：出图失败：{}".format(rec["name"], ex))
                    log(errors[-1])
                    log(traceback.format_exc())
                hide_body(rec["body"])
                if every > 0 and (i + 1) % every == 0:
                    refresh_ui(session, ui)
        finally:
            restore_bodies(all_bodies, saved)

        try:
            last_sheet.Open()
        except Exception:
            pass
        refresh_ui(session, ui)

    except Exception:
        errors.append("出图流程异常：{}".format(traceback.format_exc()))
        log_popup("出图流程异常：")
        log(traceback.format_exc())

    report(verify_rows, all_rects, pages, errors)


def main_entry():
    start = time.time()
    try:
        main([])
    except Exception:
        log_popup("异常终止：")
        log(traceback.format_exc())
        raise
    finally:
        log("结束：总耗时 {:.2f} 秒。".format(time.time() - start))


if __name__ == "__main__":
    main_entry()
