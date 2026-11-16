# -*- coding: utf-8 -*-
"""格式转换 —— 单一入口（全自包含，不 import 任何同目录兄弟脚本）

原三个文件已合并进本文件，职责分区如下：
  段 1  公共后台执行器（原 nx_batch_runner.py）
  段 2  批量导入（原 batch_import.py）
  段 3  批量导出 + 交互主流程（原 batch_export_step.py）

合并只为「一个动作只指向一个 .py」：菜单、工具条、后台会话、进度窗
全部用本文件自身的 __file__ 启动，行为与原三文件版本等价。
"""

import types
# ==============================================================================
# 段：公共后台执行器（原 nx_batch_runner.py）
# ==============================================================================
# 批量导出 / 批量导入共用的后台执行器（计划清单协议 + 非模态进度窗 + 取消）。
# 
# 交互会话只做“选文件 → 算清单 → 写计划 JSON”，随后启动 N 个独立的 NX 批处理会话
# （run_journal.exe）与一个非模态进度窗，本会话立刻返回，NX 前台不再被占用。
# 
# - 子进程通过环境变量 NX_BATCH_PLAN 拿到自己的分片计划，把进度写进 progress JSON；
# - 进度窗轮询这些文件，显示“第几个/共几个、文件名、阶段”，并发默认 2、可调 1~6；
# - 取消先在文件边界优雅停；10 秒仍未退出则 taskkill 强杀。
# -*- coding: utf-8 -*-
import json
import os
import shutil
import subprocess
import tempfile
import threading
import time

ENV_PLAN = "NX_BATCH_PLAN"
MAX_WORKERS = 6
DEFAULT_WORKERS = 2
_WIN32_CACHE = {}

# 过渡方案：脚本运行期间把 NX 自己的「工作进行中」可中断进度框隐藏起来。
# 只按“本进程 + 类名 #32770 + 三种语言标题 + 子按钮含停止”精确匹配，绝不批量隐藏其它对话框。
WAIT_DLG_CLASS = "#32770"
WAIT_DLG_TITLES = ("工作进行中", "工作進行中", "Work in progress")
STOP_TEXTS = ("停止", "stop")

# ---- 文件名处理：导出 / 导入两个窗口共用同一套规则（避免各写一份、改一处漏一处）----
# 输出文件名里的全角/中文符号 -> 英文符号（只影响生成的文件名，不改源文件）
NAME_REPLACE = {
    u"\uff08": "(",   # （
    u"\uff09": ")",   # ）
    u"\u3010": "[",   # 【
    u"\u3011": "]",   # 】
}


def rn_normalize_name(name):
    """把文件名中的全角括号等字符替换为英文符号"""
    for old, new in NAME_REPLACE.items():
        name = name.replace(old, new)
    return name


def rn_safe_name_part(text):
    """前缀/后缀里 Windows 文件名不允许的字符替换成下划线，避免拼出非法路径"""
    value = str(text or "")
    for ch in '\\/:*?"<>|':
        value = value.replace(ch, "_")
    return value.strip()


# 同一套规则在窗口侧（PowerShell）的版本，模板里用 [[name_helpers]] 引入
NAME_HELPERS_PS = r"""
function Get-SafePart([string]$text) {
    $v = [string]$text
    foreach ($ch in @('\', '/', ':', '*', '?', '"', '<', '>', '|')) { $v = $v.Replace($ch, '_') }
    return $v.Trim()
}

function Get-CleanStem([string]$name) {
    $v = $name
    $v = $v.Replace([string][char]0xFF08, '(').Replace([string][char]0xFF09, ')')
    $v = $v.Replace([string][char]0x3010, '[').Replace([string][char]0x3011, ']')
    return $v
}
"""

# 两窗共用的左栏清单逻辑（重编号 / 多选删除 / 全选 / 清空 / 单步撤销 / 右键菜单），
# 模板里用 [[left_list]] 引入。约定由调用方提供：
#   $left / $right / $btnDelRows 三个控件；
#   $script:txtDelRows / txtMenuDelete / txtMenuSelectAll / txtMenuClear 四个文案；
#   function Refresh-List（重编号 + 计数 + 右栏复算 + 空表提示，各窗自己的差异都收在这里）。
LEFT_LIST_PS = r"""
$script:lastDeleted = $null
$script:scrollTop = 0
$script:scrollCare = 0

function Renumber-Left {
    for ($i = 0; $i -lt $left.Items.Count; $i++) {
        $left.Items[$i].SubItems[0].Text = [string]($i + 1)
    }
}

function Add-RightRow([int]$no, [string]$text, [string]$tip, [bool]$dim) {
    $item = New-Object System.Windows.Forms.ListViewItem([string]$no)
    $item.ImageIndex = 0
    $item.ToolTipText = $tip
    if ($dim) { $item.ForeColor = [System.Drawing.Color]::Gray }
    [void]$item.SubItems.Add($text)
    [void]$right.Items.Add($item)
}

# 造一行左栏数据（新增、撤销恢复共用）；kind 供导出窗区分 FILE / DIR，导入窗留空
function New-LeftItem([string]$path, [string]$kind = "") {
    $name = [System.IO.Path]::GetFileName($path)
    if (-not $name) { $name = $path }
    $item = New-Object System.Windows.Forms.ListViewItem("")
    $item.Name = $kind
    $item.ImageIndex = 0
    $item.Tag = $path
    $item.ToolTipText = $path
    [void]$item.SubItems.Add($name)
    [void]$item.SubItems.Add("✕")
    return $item
}

# 加一行左栏：按 Tag（导出窗再叠加 kind）去重；-NoRefresh 供批量加入后统一刷新
function Add-LeftRow([string]$path, [string]$kind = "", [switch]$NoRefresh) {
    foreach ($row in $left.Items) {
        if ([string]$row.Tag -eq $path -and (-not $kind -or [string]$row.Name -eq $kind)) { return }
    }
    [void]$left.Items.Add((New-LeftItem $path $kind))
    if ($NoRefresh) { return }
    Refresh-List
}

function Update-EmptyHints {
    $lblLeftEmpty.Visible = ($left.Items.Count -eq 0)
    $lblRightEmpty.Visible = ($right.Items.Count -eq 0)
    $lblLeftEmpty.BackColor = $uiField
    $lblRightEmpty.BackColor = $uiField
    $lblLeftEmpty.ForeColor = [System.Drawing.Color]::Gray
    $lblRightEmpty.ForeColor = [System.Drawing.Color]::Gray
    if ($lblLeftEmpty.Visible) { $lblLeftEmpty.BringToFront() }
    if ($lblRightEmpty.Visible) { $lblRightEmpty.BringToFront() }
}

function Update-DeleteState {
    $n = $left.SelectedItems.Count
    $btnDelRows.Enabled = ($n -gt 0)
    if ($n -gt 0) { $btnDelRows.Text = $script:txtDelRows + " ($n)" }
    else { $btnDelRows.Text = $script:txtDelRows }
}

function Select-AllRows {
    if ($left.Items.Count -le 0) { return }
    $left.BeginUpdate()
    try {
        foreach ($row in $left.Items) { $row.Selected = $true }
    } finally {
        $left.EndUpdate()
    }
    Update-DeleteState
}

# 删完保留大致滚动位置，并把选区落到被删位置附近
function Restore-Scroll {
    if ($left.Items.Count -le 0) { return }
    $t = $script:scrollTop
    if ($t -ge $left.Items.Count) { $t = $left.Items.Count - 1 }
    if ($t -lt 0) { $t = 0 }
    $left.TopItem = $left.Items[$t]
    if ($left.TopItem) { $t = $left.TopItem.Index }
    $vis = 1
    try { $vis = [Math]::Max(1, [int][Math]::Floor($left.ClientRectangle.Height / $script:rowHeight)) } catch { }
    $i = $script:scrollCare
    # 选区夹在可见范围内：既不跳回顶部，也不会把列表滚走
    if ($i -lt $t) { $i = $t }
    if ($i -ge ($t + $vis)) { $i = $t + $vis - 1 }
    if ($i -ge $left.Items.Count) { $i = $left.Items.Count - 1 }
    if ($i -lt 0) { $i = 0 }
    $left.Items[$i].Selected = $true
}

# 删掉左栏选中的行（支持 Ctrl/Shift 多选），序号与右栏结果名一起重算
function Remove-SelectedRows {
    if ($left.SelectedItems.Count -le 0) { return }
    $rows = @($left.SelectedItems)
    $script:lastDeleted = @()
    $script:scrollTop = 0
    if ($left.TopItem) { $script:scrollTop = $left.TopItem.Index }
    $script:scrollCare = $left.Items.Count
    foreach ($row in $rows) {
        if ($row.Index -lt $script:scrollCare) { $script:scrollCare = $row.Index }
        $script:lastDeleted += (New-Object psobject -Property @{ idx = $row.Index; kind = [string]$row.Name; path = [string]$row.Tag })
    }
    # 上面被删掉的行会让列表内容上移，滚动目标同步减去这部分偏移
    $above = 0
    foreach ($row in $rows) { if ($row.Index -lt $script:scrollTop) { $above++ } }
    $script:scrollTop = [Math]::Max(0, $script:scrollTop - $above)
    $left.BeginUpdate()
    try {
        foreach ($row in $rows) { [void]$left.Items.Remove($row) }
    } finally {
        $left.EndUpdate()
    }
    Refresh-List
    Restore-Scroll
    Update-DeleteState
}

function Clear-LeftRows {
    if ($left.Items.Count -le 0) { return }
    $script:lastDeleted = @()
    foreach ($row in $left.Items) {
        $script:lastDeleted += (New-Object psobject -Property @{ idx = $row.Index; kind = [string]$row.Name; path = [string]$row.Tag })
    }
    $left.Items.Clear()
    Refresh-List
    Update-DeleteState
}

# Ctrl+Z：恢复最近一次删掉的那批，序号按原位置插回后重排
function Restore-Deleted {
    if (-not $script:lastDeleted -or $script:lastDeleted.Count -le 0) { return }
    $left.BeginUpdate()
    try {
        foreach ($rec in @($script:lastDeleted | Sort-Object { $_.idx })) {
            $i = [int]$rec.idx
            if ($i -gt $left.Items.Count) { $i = $left.Items.Count }
            $left.Items.Insert($i, (New-LeftItem $rec.path $rec.kind))
        }
    } finally {
        $left.EndUpdate()
    }
    $script:lastDeleted = $null
    Refresh-List
    Update-DeleteState
}

# 接线：按钮 / 选中态 / Ctrl+A / Ctrl+Z / 右键菜单 / 右键选中行
# 注意：不绑定 Delete 键删除（避免误删），删除入口只有按钮与右键菜单
function Initialize-LeftList {
    $btnDelRows.add_Click({ Remove-SelectedRows })
    $left.add_SelectedIndexChanged({ Update-DeleteState })
    $left.add_KeyDown({
        param($sender, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $e.Handled = $true
            $e.SuppressKeyPress = $true
            Select-AllRows
            return
        }
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Z) {
            $e.Handled = $true
            $e.SuppressKeyPress = $true
            Restore-Deleted
        }
    })
    # 右键点在未选中的行上时先选中它（点在多选范围内则保留多选），让菜单作用于所指行
    $left.add_MouseDown({
        param($sender, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
        $hit = $left.GetItemAt($e.X, $e.Y)
        if (-not $hit) { return }
        if (-not $hit.Selected) {
            $left.SelectedItems.Clear()
            $hit.Selected = $true
        }
    })

    $script:menuLeft = New-Object System.Windows.Forms.ContextMenuStrip
    $script:miDelete = $script:menuLeft.Items.Add($script:txtMenuDelete)
    $script:miSelectAll = $script:menuLeft.Items.Add($script:txtMenuSelectAll)
    $script:miClear = $script:menuLeft.Items.Add($script:txtMenuClear)
    $script:miDelete.add_Click({ Remove-SelectedRows })
    $script:miSelectAll.add_Click({ Select-AllRows })
    $script:miClear.add_Click({ Clear-LeftRows })
    $script:menuLeft.add_Opening({
        param($sender, $e)
        $script:miDelete.Enabled = ($left.SelectedItems.Count -gt 0)
        $script:miSelectAll.Enabled = ($left.Items.Count -gt 0)
        $script:miClear.Enabled = ($left.Items.Count -gt 0)
    })
    $left.ContextMenuStrip = $script:menuLeft
    Update-DeleteState
}
"""

# 选择窗抢前台（导入/导出/预览共用）：模板里用 [[front_helpers]] 引入。
# 只在弹出那一刻置前一次，不常驻 TopMost，点回 NX 不收回、不关闭。
PICKER_FRONT_PS = r"""
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class NBActPick {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern int GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
}
"@

function Show-PickerFront([System.Windows.Forms.Form]$win, [string]$MarkFile) {
    # 与进度窗 Set-FrontOnce 同一路数：AttachThreadInput + BringWindowToTop，
    # 第 2 轮起补 ALT 键解除前台锁并瞬时 TopMost（随即取消），最多 3 轮。
    $ok = $false
    $h = [IntPtr]::Zero
    try { $h = $win.Handle } catch { $ok = $false }
    for ($i = 1; $h -ne [IntPtr]::Zero -and $i -le 3; $i++) {
        $fgTid = 0; $myTid = 0; $attached = $false
        try {
            [void][NBActPick]::GetWindowThreadProcessId([NBActPick]::GetForegroundWindow(), [ref]$fgTid)
            $myTid = [NBActPick]::GetCurrentThreadId()
            if ($fgTid -ne 0 -and $myTid -ne 0 -and $fgTid -ne $myTid) {
                $attached = [NBActPick]::AttachThreadInput($myTid, $fgTid, $true)
            }
        } catch { }
        try {
            if ($i -ge 2) {
                [NBActPick]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
                [NBActPick]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
                Start-Sleep -Milliseconds 60
            }
            [void][NBActPick]::ShowWindow($h, 9)
            [void][NBActPick]::BringWindowToTop($h)
            [void][NBActPick]::SetForegroundWindow($h)
            $win.Activate()
        } catch { }
        if ($attached) { try { [void][NBActPick]::AttachThreadInput($myTid, $fgTid, $false) } catch { } }
        if ($i -ge 2) {
            try { $win.TopMost = $true; $win.TopMost = $false } catch { }
        }
        Start-Sleep -Milliseconds 150
        try {
            $fg = [NBActPick]::GetForegroundWindow()
            $owner = 0
            [void][NBActPick]::GetWindowThreadProcessId($fg, [ref]$owner)
            $ok = ($owner -eq [int]$PID)
        } catch { $ok = $false }
        if ($ok) { break }
    }
    if ($MarkFile) {
        try { Set-Content -LiteralPath $MarkFile -Value $(if ($ok) { "ok" } else { "failed" }) -Encoding Ascii } catch { }
    }
    return $ok
}
"""

rn_MSG = {
    "en": {
        "title_fmt": "Format Convert",
        "count": "{0} / {1}",
        "file": "File: {0}",
        "stage": "Stage: {0}",
        "idle": "Preparing...",
        "kept": "You can keep working in NX.",
        "workers": "Concurrent:",
        "cancel": "Cancel",
        "close": "Close",
        "stopping": "Stopping...",
        "finished": "Finished: success {0} / skipped {1} / failed {2}",
        "cancelled": "Cancelled: success {0} / skipped {1} / failed {2}",
        "stop_title": "Stop and roll back",
        "stop_confirm": "Stop will cancel the remaining tasks and delete the files already produced by this batch.\nFiles replaced in overwrite mode are kept and reported.\n\nContinue?",
        "rollback_done": "Rolled back: deleted {0} file(s), kept {1} (overwrite mode), failed to delete {2}.",
        "s.queue": "Queued", "s.open": "Reading layers", "s.export": "Exporting",
        "s.verify": "Verifying", "s.import": "Importing", "s.save": "Saving",
        "s.done": "Done", "s.failed": "Failed", "s.skip": "Skipped",
    },
    "zh_cn": {
        "title_fmt": "格式转换",
        "count": "{0} / {1}",
        "file": "文件：{0}",
        "stage": "阶段：{0}",
        "idle": "准备中…",
        "kept": "运行期间可在 NX 中继续操作。",
        "workers": "同时执行数：",
        "cancel": "取消",
        "close": "关闭",
        "stopping": "正在停止…",
        "finished": "已完成：成功 {0} / 跳过 {1} / 失败 {2}",
        "cancelled": "已取消：成功 {0} / 跳过 {1} / 失败 {2}",
        "stop_title": "停止并回滚",
        "stop_confirm": "停止将取消剩余任务，并删除本次已经生成的文件。\n覆盖模式替换掉的原文件无法恢复，这类产物会保留并记录。\n\n是否继续？",
        "rollback_done": "已回滚：删除 {0} 个文件，保留 {1} 个（覆盖模式），删除失败 {2} 个。",
        "s.queue": "排队", "s.open": "读取图层", "s.export": "导出中",
        "s.verify": "校验", "s.import": "导入中", "s.save": "保存",
        "s.done": "完成", "s.failed": "失败", "s.skip": "跳过",
    },
    "zh_tw": {
        "title_fmt": "格式轉換",
        "count": "{0} / {1}",
        "file": "檔案：{0}",
        "stage": "階段：{0}",
        "idle": "準備中…",
        "kept": "執行期間可在 NX 中繼續操作。",
        "workers": "同時執行數：",
        "cancel": "取消",
        "close": "關閉",
        "stopping": "正在停止…",
        "finished": "已完成：成功 {0} / 略過 {1} / 失敗 {2}",
        "cancelled": "已取消：成功 {0} / 略過 {1} / 失敗 {2}",
        "stop_title": "停止並回滾",
        "stop_confirm": "停止將取消剩餘任務，並刪除本次已經產生的檔案。\n覆寫模式取代掉的原檔案無法復原，這類產物會保留並記錄。\n\n是否繼續？",
        "rollback_done": "已回滾：刪除 {0} 個檔案，保留 {1} 個（覆寫模式），刪除失敗 {2} 個。",
        "s.queue": "排隊", "s.open": "讀取圖層", "s.export": "匯出中",
        "s.verify": "驗證", "s.import": "匯入中", "s.save": "儲存",
        "s.done": "完成", "s.failed": "失敗", "s.skip": "略過",
    },
}


def rn_T(lang, key, *args):
    text = (rn_MSG.get(lang) or {}).get(key)
    if text is None:
        text = rn_MSG["en"].get(key, key)
    return text.format(*args) if args else text


def write_json(path, payload):
    """原子写：先写 .tmp 再替换，避免进度窗读到半个文件"""
    if not path:
        return
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception:
        # 目标被占用时退化为直接写，进度显示优先于原子性
        try:
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False)
        except Exception:
            pass


def load_plan(path):
    with open(path, "r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def execute(plan, handler, log=print):
    """子进程侧公共执行循环。

    handler(item, report) 返回 (status, detail)，status = ok / fail / skip；
    report(stage) 更新进度窗里的阶段文字。
    取消只在文件之间检查（文件边界优雅停），绝不跨线程打断 NX。
    """
    items = list(plan.get("items") or [])
    cancel = plan.get("cancel") or ""
    progress = plan.get("progress") or ""
    state = {"done": 0, "total": len(items), "ok": 0, "skip": 0, "fail": 0,
             "index": 0, "name": "", "stage": "queue", "status": "running",
             "updated": 0.0}
    results = []

    def flush():
        state["updated"] = time.time()
        write_json(progress, state)

    flush()
    stopped = False
    for index, item in enumerate(items, 1):
        if cancel and os.path.exists(cancel):
            stopped = True
            break
        state["index"] = index
        state["name"] = os.path.basename(str(item.get("src") or ""))
        state["stage"] = "queue"
        flush()

        def report(stage):
            state["stage"] = stage
            flush()

        try:
            response = handler(item, report)
            status = response[0]
        except Exception as exc:
            log("    item failed: %s" % exc)
            status = "fail"
        if status == "ok":
            state["ok"] += 1
        elif status == "fail":
            state["fail"] += 1
        else:
            state["skip"] += 1
        state["done"] += 1
        results.append({"src": item.get("src"), "status": status})
        flush()

    state["status"] = "cancelled" if stopped else "done"
    state["stage"] = "done"
    flush()
    return results, state


def _powershell():
    candidates = [
        os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                     "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
        "powershell",
    ]
    for candidate in candidates:
        if candidate == "powershell" or os.path.exists(candidate):
            return candidate
    return None


def _win32_api():
    """抢前台用到的最小 user32/kernel32 封装（只初始化一次）"""
    cached = _WIN32_CACHE.get("api")
    if cached is not None:
        return cached
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    wndenumproc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows.argtypes = (wndenumproc, wintypes.LPARAM)
    user32.EnumWindows.restype = wintypes.BOOL
    user32.IsWindowVisible.argtypes = (wintypes.HWND,)
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.GetWindowTextLengthW.argtypes = (wintypes.HWND,)
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
    user32.GetWindowTextW.restype = ctypes.c_int
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.AttachThreadInput.argtypes = (wintypes.DWORD, wintypes.DWORD, wintypes.BOOL)
    user32.AttachThreadInput.restype = wintypes.BOOL
    user32.BringWindowToTop.argtypes = (wintypes.HWND,)
    user32.BringWindowToTop.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = (wintypes.HWND,)
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.ShowWindow.argtypes = (wintypes.HWND, ctypes.c_int)
    user32.ShowWindow.restype = wintypes.BOOL
    user32.SetWindowPos.argtypes = (wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                    ctypes.c_int, ctypes.c_int, ctypes.c_uint)
    user32.SetWindowPos.restype = wintypes.BOOL
    user32.keybd_event.argtypes = (wintypes.BYTE, wintypes.BYTE, wintypes.DWORD, ctypes.c_void_p)
    # 隐藏「工作进行中」框用到：类名/子窗枚举/异步显示（异步避免 UI 线程忙时卡住本线程）
    user32.GetClassNameW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
    user32.GetClassNameW.restype = ctypes.c_int
    user32.EnumChildWindows.argtypes = (wintypes.HWND, wndenumproc, wintypes.LPARAM)
    user32.EnumChildWindows.restype = wintypes.BOOL
    user32.ShowWindowAsync.argtypes = (wintypes.HWND, ctypes.c_int)
    user32.ShowWindowAsync.restype = wintypes.BOOL
    user32.IsWindow.argtypes = (wintypes.HWND,)
    user32.IsWindow.restype = wintypes.BOOL
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD
    api = {"ctypes": ctypes, "wintypes": wintypes, "user32": user32,
           "kernel32": kernel32, "enum": wndenumproc}
    _WIN32_CACHE["api"] = api
    return api


def find_window_by_title(title):
    """按标题精确匹配可见顶层窗口；找到返回 hwnd，否则 0"""
    want = (title or "").strip()
    if not want:
        return 0
    api = _win32_api()
    ctypes = api["ctypes"]
    user32 = api["user32"]
    hits = []

    def _visit(hwnd, _lparam):
        try:
            if not user32.IsWindowVisible(hwnd):
                return True
            size = user32.GetWindowTextLengthW(hwnd)
            if size <= 0:
                return True
            buf = ctypes.create_unicode_buffer(size + 1)
            user32.GetWindowTextW(hwnd, buf, size + 1)
            if buf.value.strip() == want:
                hits.append(int(hwnd or 0))
                return False
        except Exception:
            pass
        return True

    try:
        user32.EnumWindows(api["enum"](_visit), 0)
    except Exception:
        return 0
    return hits[0] if hits else 0


def bring_window_front(hwnd):
    """把已存在的窗口提到前台：还原 + AttachThreadInput + BringWindowToTop，必要时瞬时置顶"""
    if not hwnd:
        return False
    api = _win32_api()
    ctypes = api["ctypes"]
    wintypes = api["wintypes"]
    user32 = api["user32"]
    kernel32 = api["kernel32"]
    try:
        user32.ShowWindow(wintypes.HWND(hwnd), 9)      # SW_RESTORE
    except Exception:
        pass
    ok = False
    for attempt in range(1, 4):
        fg_tid = wintypes.DWORD(0)
        attached = False
        my_tid = kernel32.GetCurrentThreadId()
        try:
            user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), ctypes.byref(fg_tid))
            if fg_tid.value and fg_tid.value != my_tid:
                attached = bool(user32.AttachThreadInput(my_tid, fg_tid.value, True))
        except Exception:
            attached = False
        try:
            if attempt >= 2:
                # 系统只允许前台进程改前台：先按一下 ALT 让自己“收到最近输入”
                user32.keybd_event(0x12, 0, 0, None)
                user32.keybd_event(0x12, 0, 2, None)
                time.sleep(0.06)
            user32.BringWindowToTop(wintypes.HWND(hwnd))
            user32.SetForegroundWindow(wintypes.HWND(hwnd))
        except Exception:
            pass
        if attached:
            try:
                user32.AttachThreadInput(my_tid, fg_tid.value, False)
            except Exception:
                pass
        if attempt >= 2:
            # 仍被拒绝时瞬时置顶再取消（不常驻），至少把窗抬到 NX 之上
            try:
                user32.SetWindowPos(wintypes.HWND(hwnd), wintypes.HWND(-1), 0, 0, 0, 0, 0x0003)
                user32.SetWindowPos(wintypes.HWND(hwnd), wintypes.HWND(-2), 0, 0, 0, 0, 0x0003)
            except Exception:
                pass
        time.sleep(0.15)
        try:
            ok = (int(user32.GetForegroundWindow() or 0) == int(hwnd))
        except Exception:
            ok = False
        if ok:
            break
    return ok


def activate_picker(title):
    """已有同名选择窗时把它提到前台：ok / failed / none（none = 没找到这个窗）"""
    try:
        hwnd = find_window_by_title(title)
    except Exception:
        return "none"
    if not hwnd:
        return "none"
    try:
        return "ok" if bring_window_front(hwnd) else "failed"
    except Exception:
        return "failed"


def _window_text(hwnd):
    """读窗口标题；失败返回空串"""
    api = _win32_api()
    ctypes = api["ctypes"]
    user32 = api["user32"]
    try:
        size = user32.GetWindowTextLengthW(hwnd)
        if size <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(size + 1)
        user32.GetWindowTextW(hwnd, buf, size + 1)
        return buf.value or ""
    except Exception:
        return ""


def _window_class(hwnd):
    """读窗口类名；失败返回空串"""
    api = _win32_api()
    ctypes = api["ctypes"]
    try:
        buf = ctypes.create_unicode_buffer(256)
        api["user32"].GetClassNameW(hwnd, buf, 256)
        return buf.value or ""
    except Exception:
        return ""


def _has_stop_button(hwnd):
    """子窗里是否有类名 Button 且文本含“停止 / Stop”的按钮（精确到这一个对话框）"""
    api = _win32_api()
    user32 = api["user32"]
    wintypes = api["wintypes"]
    found = []

    def _visit(child, _lparam):
        try:
            if _window_class(child) != "Button":
                return True
            text = (_window_text(child) or "").casefold()
            for token in STOP_TEXTS:
                if token in text:
                    found.append(int(child or 0))
                    return False
        except Exception:
            pass
        return True

    try:
        user32.EnumChildWindows(wintypes.HWND(hwnd), api["enum"](_visit), 0)
    except Exception:
        return False
    return bool(found)


def find_wait_dialog(pid=None):
    """定位“本进程”的「工作进行中」框：类名 #32770 + 标题三种语言 + 带停止按钮

    找到返回 hwnd，否则返回 0。只认第一个命中的窗口，绝不批量处理其它对话框。
    """
    api = _win32_api()
    ctypes = api["ctypes"]
    user32 = api["user32"]
    wintypes = api["wintypes"]
    want_pid = int(pid or os.getpid())
    titles = tuple(item.casefold() for item in WAIT_DLG_TITLES)
    hits = []

    def _visit(hwnd, _lparam):
        try:
            if not user32.IsWindowVisible(hwnd):
                return True
            owner = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if int(owner.value) != want_pid:
                return True
            if _window_class(hwnd) != WAIT_DLG_CLASS:
                return True
            title = (_window_text(hwnd) or "").strip().casefold()
            if not title or title not in titles:
                return True
            if not _has_stop_button(hwnd):
                return True
            hits.append(int(hwnd or 0))
            return False
        except Exception:
            pass
        return True

    try:
        user32.EnumWindows(api["enum"](_visit), 0)
    except Exception:
        return 0
    return hits[0] if hits else 0


def _waitdlg_enabled():
    """开关：NX_BATCH_HIDE_WAITDLG=0/off/false/no 时关闭（默认开启）"""
    value = (os.environ.get("NX_BATCH_HIDE_WAITDLG") or "").strip().lower()
    return value not in ("0", "off", "false", "no")


class WaitDialogHider(threading.Thread):
    """过渡：脚本等待窗口返回期间，把 NX 的「工作进行中」框隐藏；结束/异常时恢复

    - 只处理本进程、且只处理“类名 #32770 + 标题匹配 + 带停止按钮”的那一个窗；
    - 用 ShowWindowAsync 避免 UI 线程忙时把本线程卡死；隐藏后复查可见性，失败立即放弃；
    - 隐藏期间转换/等待逻辑照常，不参与、也不打断任何流程。
    """

    def __init__(self, logger=None):
        threading.Thread.__init__(self)
        self.daemon = True
        self._log = logger
        self._stop_event = threading.Event()
        self._hwnd = 0
        self._give_up = False

    def _write(self, message):
        try:
            if self._log is not None:
                self._log(message)
        except Exception:
            pass

    def run(self):
        try:
            api = _win32_api()
        except Exception as exc:
            self._write("[nxdlg] hide=failed reason=win32:%s" % exc)
            return
        user32 = api["user32"]
        wintypes = api["wintypes"]
        while not self._stop_event.is_set():
            if self._give_up:
                self._stop_event.wait(0.3)
                continue
            try:
                hwnd = find_wait_dialog()
                if not hwnd:
                    self._stop_event.wait(0.3)
                    continue
                title = _window_text(hwnd) or ""
                self._write("[nxdlg] hide=try hwnd=0x%X title=%s class=%s"
                            % (hwnd, title, WAIT_DLG_CLASS))
                user32.ShowWindowAsync(wintypes.HWND(hwnd), 0)   # SW_HIDE
                self._stop_event.wait(0.6)
                if self._stop_event.is_set():
                    # 刚发起隐藏就被要求收尾：记下句柄，让 _restore 负责恢复
                    self._hwnd = hwnd
                    break
                if user32.IsWindowVisible(wintypes.HWND(hwnd)):
                    self._give_up = True
                    self._write("[nxdlg] hide=failed hwnd=0x%X visible=1" % hwnd)
                    continue
                self._hwnd = hwnd
                self._write("[nxdlg] hide=ok hwnd=0x%X title=%s" % (hwnd, title))
                self._write("[nxdlg] note=隐藏期间「停止」按钮不可见；脚本结束时恢复")
                # 已隐藏成功：不再枚举其它窗，只等结束信号
                self._stop_event.wait()
                break
            except Exception as exc:
                self._give_up = True
                self._write("[nxdlg] hide=failed reason=%s" % exc)
        self._restore()

    def _restore(self):
        hwnd = self._hwnd
        if not hwnd:
            return
        try:
            api = _win32_api()
            user32 = api["user32"]
            wintypes = api["wintypes"]
            if user32.IsWindow(wintypes.HWND(hwnd)):
                user32.ShowWindowAsync(wintypes.HWND(hwnd), 5)   # SW_SHOW
                self._write("[nxdlg] restore=ok hwnd=0x%X" % hwnd)
            else:
                self._write("[nxdlg] restore=gone hwnd=0x%X" % hwnd)
        except Exception as exc:
            self._write("[nxdlg] restore=failed reason=%s" % exc)

    def stop(self):
        """结束隐藏线程并恢复窗口；最多等 0.6 秒，绝不阻塞脚本退出"""
        self._stop_event.set()
        try:
            self.join(0.6)
        except Exception:
            pass
        if self.is_alive():
            # 线程还卡在 Win32 调用里：直接恢复，失败就算了
            self._restore()


def start_waitdlg_hider(logger=None):
    """启动隐藏线程；关闭开关或初始化失败时返回 None（不抛、不重试、不弹错）"""
    if not _waitdlg_enabled():
        try:
            if logger is not None:
                logger("[nxdlg] hide=off")
        except Exception:
            pass
        return None
    try:
        hider = WaitDialogHider(logger)
        hider.start()
        return hider
    except Exception as exc:
        try:
            if logger is not None:
                logger("[nxdlg] hide=failed reason=%s" % exc)
        except Exception:
            pass
        return None


def _labels(kind, lang):
    stage = {}
    for name in ("queue", "open", "export", "verify", "import", "save",
                 "done", "failed", "skip"):
        stage[name] = rn_T(lang, "s." + name)
    return {
        "title": rn_T(lang, "title_fmt"),
        "count": rn_T(lang, "count"),
        "file": rn_T(lang, "file"),
        "stage": rn_T(lang, "stage"),
        "idle": rn_T(lang, "idle"),
        "kept": rn_T(lang, "kept"),
        "workers": rn_T(lang, "workers"),
        "cancel": rn_T(lang, "cancel"),
        "close": rn_T(lang, "close"),
        "stopping": rn_T(lang, "stopping"),
        "finished": rn_T(lang, "finished"),
        "cancelled": rn_T(lang, "cancelled"),
        "stop_title": rn_T(lang, "stop_title"),
        "stop_confirm": rn_T(lang, "stop_confirm"),
        "rollback_done": rn_T(lang, "rollback_done"),
        "stage_text": stage,
    }


def rn__pid_alive(pid):
    """进程还在吗：判断不了时按“还活着”处理，避免误删正在用的目录

    注意：进程已退出但句柄未释放时 OpenProcess 仍可能成功，
    所以必须再看退出码（259 = STILL_ACTIVE）。
    """
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.GetExitCodeProcess.argtypes = (wintypes.HANDLE,
                                                ctypes.POINTER(wintypes.DWORD))
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        handle = kernel32.OpenProcess(0x1000, False, int(pid))
        if not handle:
            return ctypes.get_last_error() != 87   # 87 = PID 不存在
        try:
            code = wintypes.DWORD()
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                return True
            return code.value == 259               # STILL_ACTIVE
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return True


def rn__proc_stamp(pid):
    """进程创建时间（FILETIME 字符串）：用来识别 PID 被系统复用，取不到返回空串"""
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        handle = kernel32.OpenProcess(0x1000, False, int(pid))
        if not handle:
            return ""
        try:
            created = wintypes.FILETIME()
            exited = wintypes.FILETIME()
            kern = wintypes.FILETIME()
            user = wintypes.FILETIME()
            if not kernel32.GetProcessTimes(handle, ctypes.byref(created),
                                            ctypes.byref(exited),
                                            ctypes.byref(kern), ctypes.byref(user)):
                return ""
            return str((created.dwHighDateTime << 32) | created.dwLowDateTime)
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return ""


def rn__proc_parent_map():
    """进程快照：pid -> (父 pid, 可执行文件名)，用于往上找发起本批次的 NX 主进程"""
    try:
        import ctypes
        from ctypes import wintypes

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_char * 260),
            ]

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        kernel32.CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
        if not snapshot or snapshot == wintypes.HANDLE(-1).value:
            return {}
        table = {}
        try:
            entry = PROCESSENTRY32()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
            ok = kernel32.Process32First(snapshot, ctypes.byref(entry))
            while ok:
                name = entry.szExeFile.decode("mbcs", "ignore")
                table[int(entry.th32ProcessID)] = (int(entry.th32ParentProcessID), name)
                ok = kernel32.Process32Next(snapshot, ctypes.byref(entry))
        finally:
            kernel32.CloseHandle(snapshot)
        return table
    except Exception:
        return {}


def rn__nx_owner_pid(start_pid=None):
    """往上找 NX 主程序（ugraf.exe）PID；找不到返回 0。

    调用方一般会显式传入（准备会话从交互 NX 上下文里取好），这里只是兜底：
    本进程的祖先进程里若还有 NX，就取它，避免 nx_pid 为空让看门狗整段失效。
    """
    pid = int(start_pid or os.getpid())
    try:
        table = rn__proc_parent_map()
    except Exception:
        return 0
    seen = set()
    for _ in range(12):
        if pid <= 0 or pid in seen:
            break
        seen.add(pid)
        info = table.get(pid)
        if not info:
            break
        parent, name = info
        if str(name).strip().lower() in ("ugraf.exe", "ugraf"):
            return pid
        pid = parent
    return 0


def _proc_name(pid):
    """进程可执行文件名（不含路径）；取不到返回空串。与启动时间一起识别 PID 复用。"""
    try:
        table = rn__proc_parent_map()
        info = table.get(int(pid))
        if not info:
            return ""
        return str(info[1]).strip()
    except Exception:
        return ""


def sweep_stale_dirs(kind):
    """清掉上次异常退出留下的 nx_batch_<kind>_<pid>[_<ms>]（NX 崩溃/被杀时进度窗来不及清理）"""
    root = tempfile.gettempdir()
    prefix = "nx_batch_%s_" % kind
    try:
        names = os.listdir(root)
    except Exception:
        return
    for name in names:
        if not name.startswith(prefix):
            continue
        tail = name[len(prefix):]
        pid_text = tail.split("_", 1)[0]
        if not pid_text.isdigit() or rn__pid_alive(int(pid_text)):
            continue
        shutil.rmtree(os.path.join(root, name), ignore_errors=True)


def purge_orphan_prt(kind, age=600):
    """清掉上次被取消/中断时 translator 来不及删除的中间 .prt（不碰 .log，日志按保留策略处理）"""
    folder = os.path.join(tempfile.gettempdir(), "nx_batch_%s" % kind)
    cutoff = time.time() - age
    try:
        names = os.listdir(folder)
    except Exception:
        return
    for name in names:
        if not name.lower().endswith(".prt"):
            continue
        path = os.path.join(folder, name)
        try:
            if os.path.getmtime(path) < cutoff:
                os.remove(path)
        except Exception:
            continue


def start(kind, journal, items, total, options, log_path, lang,
          lock="", base_dir=None, extras=None):
    """写计划并启动后台会话 + 进度窗；不可用时返回 False（调用方回落前台执行）

    extras：workers（并发数）/ nx_pid（交互会话进程号）
    """
    if not items:
        return False
    extras = extras or {}
    runner = os.path.join(os.environ.get("UGII_BASE_DIR") or "",
                          "NXBIN", "run_journal.exe")
    if not os.path.isfile(runner):
        return False
    # 每次启动使用独立目录：同一 NX 会话里连开两次同功能时，旧批次收尾只会删自己的目录，
    # 不会把新批次的计划/进度一并删掉。
    run_dir = os.path.join(tempfile.gettempdir(),
                           "nx_batch_%s_%d_%d" % (kind, os.getpid(), int(time.time() * 1000)))
    sweep_stale_dirs(kind)
    purge_orphan_prt(kind)
    try:
        if not os.path.isdir(run_dir):
            os.makedirs(run_dir)
        for name in os.listdir(run_dir):
            try:
                os.remove(os.path.join(run_dir, name))
            except Exception:
                pass
    except Exception:
        return False

    cancel = os.path.join(run_dir, "cancel.flag")
    # 会话标记：同一插件再次启动批量时改写它，旧的进度窗读到不一致就自行关闭并清理，
    # 避免多个进度窗叠在一起。必须先写标记再启动新窗——否则新窗可能读到旧标记、把自己关掉。
    token = "%d-%d" % (os.getpid(), int(time.time() * 1000))
    token_dir = os.path.join(tempfile.gettempdir(), "nx_batch_%s" % kind)
    token_file = os.path.join(token_dir, "session_token.txt")
    try:
        if not os.path.isdir(token_dir):
            os.makedirs(token_dir)
        tmp_token = token_file + ".tmp"
        with open(tmp_token, "w", encoding="utf-8") as handle:
            handle.write(token)
        os.replace(tmp_token, token_file)
    except Exception:
        token_file = ""
    # 主窗可设并发（1~6），这里按其取值切片；未传则沿用默认值
    try:
        workers = int(extras.get("workers") or DEFAULT_WORKERS)
    except Exception:
        workers = DEFAULT_WORKERS
    workers = max(1, min(MAX_WORKERS, workers))
    shard_count = min(workers, len(items))
    shards = []
    for index in range(shard_count):
        plan_path = os.path.join(run_dir, "plan_%d.json" % (index + 1))
        progress = os.path.join(run_dir, "progress_%d.json" % (index + 1))
        write_json(plan_path, {
            "kind": kind,
            "items": items[index::shard_count],
            "options": options,
            "base_dir": base_dir,
            "log_path": log_path,
            "lang": lang,
            "progress": progress,
            "cancel": cancel,
        })
        shards.append({"plan": plan_path, "progress": progress})

    state_path = os.path.join(run_dir, "run.json")
    nx_pid = int(extras.get("nx_pid") or 0)
    # 兜底：未显式传入时沿父进程链找 ugraf.exe。拿不到就如实留 0——
    # 进度窗宁可少一条判据，也不能把「死得很快的 python 宿主」当成 NX 误关窗。
    if nx_pid <= 0:
        nx_pid = rn__nx_owner_pid()
    nx_name = _proc_name(nx_pid) if nx_pid > 0 else ""
    write_json(state_path, {
        "kind": kind,
        "total": int(total or len(items)),
        "concurrency": workers,
        "max_workers": MAX_WORKERS,
        "runner": runner,
        "journal": journal,
        "cancel": cancel,
        "lock": lock or "",
        "log": log_path or "",
        "nx_pid": nx_pid,
        # 与单实例锁同判据：PID 会被系统复用，还要比对进程启动时间
        "nx_stamp": rn__proc_stamp(nx_pid) if nx_pid > 0 else "",
        "nx_name": nx_name,
        "auto_open": bool(extras.get("auto_open")),
        "session_token": token,
        "token_file": token_file,
        "shards": shards,
        "labels": _labels(kind, lang),
    })

    powershell = _powershell()
    if powershell is None:
        return False
    script = os.path.join(run_dir, "progress.ps1")
    try:
        with open(script, "w", encoding="utf-8-sig") as handle:
            handle.write(PROGRESS_PS)
        # 进度窗会再 fork worker；准备会话的环境里带着 NX_BATCH_STAGE，
        # 若原样继承，worker 会把它当成「准备阶段」重放选窗/确认清单（真机死锁）。
        child_env = dict(os.environ)
        for key in ("NX_BATCH_STAGE", "NX_BATCH_PLAN"):
            child_env.pop(key, None)
        proc = subprocess.Popen(
            [powershell, "-NoProfile", "-NonInteractive", "-STA",
             "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
             "-File", script, "-StateFile", state_path],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL, creationflags=0x08000000,
            env=child_env)
        # 立刻把单实例锁改挂到进度窗进程上，避免父进程退出后出现“假空闲”
        if lock:
            try:
                with open(lock, "w", encoding="ascii") as handle:
                    # 与 _acquire_run_lock 同格式：PID|进程创建时间，防止 PID 复用把陈旧锁判成忙
                    handle.write("%d|%s" % (proc.pid, rn__proc_stamp(proc.pid)))
            except Exception:
                pass
    except Exception:
        return False
    return True


PROGRESS_PS = r'''
param([Parameter(Mandatory=$true)][string]$StateFile)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 进度窗按普通窗口显示：可最小化、任务栏有独立条目、点 NX 会沉到后面
# 仅保留把进度窗自己提到前台所需的 Win32 封装（不操作 NX；结束态短暂置顶后自动取消）
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class NBAct {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);
    [DllImport("kernel32.dll")] public static extern int GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

    [StructLayout(LayoutKind.Sequential)]
    public struct FLASHWINFO {
        public uint cbSize;
        public IntPtr hwnd;
        public uint dwFlags;
        public uint uCount;
        public uint dwTimeout;
    }
}
"@

$state = Get-Content -LiteralPath $StateFile -Raw -Encoding UTF8 | ConvertFrom-Json
$labels = $state.labels
$stageText = $labels.stage_text
$shards = @($state.shards)

function Stage-Text([string]$key) {
    $p = $stageText.PSObject.Properties[$key]
    if ($p) { return [string]$p.Value }
    return $key
}

$form = New-Object System.Windows.Forms.Form
$form.Text = $labels.title
$form.ClientSize = New-Object System.Drawing.Size(468, 268)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.TopMost = $false
$uiFont = [System.Drawing.SystemFonts]::MessageBoxFont
try { $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9) } catch { }
$form.Font = $uiFont

$padX = 16
$innerW = 436
$lineH = [System.Windows.Forms.TextRenderer]::MeasureText("国Ag", $uiFont).Height

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Location = New-Object System.Drawing.Point($padX, 14)
$lblCount.Size = New-Object System.Drawing.Size($innerW, 20)
$lblCount.Text = $labels.idle
$form.Controls.Add($lblCount)

$bar = New-Object System.Windows.Forms.ProgressBar
$bar.Location = New-Object System.Drawing.Point($padX, 38)
$bar.Size = New-Object System.Drawing.Size($innerW, 18)
$bar.Minimum = 0
$bar.Maximum = [Math]::Max(1, [int]$state.total)
$form.Controls.Add($bar)

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Location = New-Object System.Drawing.Point($padX, 64)
$lblFile.Size = New-Object System.Drawing.Size($innerW, 20)
$lblFile.AutoEllipsis = $true
$form.Controls.Add($lblFile)

$lblStage = New-Object System.Windows.Forms.Label
$lblStage.Location = New-Object System.Drawing.Point($padX, 88)
$lblStage.Size = New-Object System.Drawing.Size($innerW, 20)
$form.Controls.Add($lblStage)

# 提示区：高度按内容自适应（最多 3 行，超出截断加省略号），不再固定 20px 切字
$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Location = New-Object System.Drawing.Point($padX, 114)
$lblHint.Size = New-Object System.Drawing.Size($innerW, 20)
$lblHint.AutoSize = $false
$lblHint.AutoEllipsis = $false
$lblHint.ForeColor = [System.Drawing.SystemColors]::GrayText
$lblHint.Text = $labels.kept
$form.Controls.Add($lblHint)

$lblWorkers = New-Object System.Windows.Forms.Label
$lblWorkers.Location = New-Object System.Drawing.Point($padX, 152)
$lblWorkers.Size = New-Object System.Drawing.Size(120, 24)
$lblWorkers.Text = $labels.workers
$form.Controls.Add($lblWorkers)

$combo = New-Object System.Windows.Forms.ComboBox
$combo.Location = New-Object System.Drawing.Point(140, 149)
$combo.Size = New-Object System.Drawing.Size(64, 24)
$combo.DropDownStyle = "DropDownList"
# 候选项按后端上限生成：硬编码 1~4 时，主窗选 6 会因找不到 "6" 而 SelectedIndex=-1，
# 并发上限被卡在 4（真机表现），更糟的是取不到值导致一个 worker 都不启动。
$maxWorkers = 6
try { $maxWorkers = [int]$state.max_workers } catch { $maxWorkers = 6 }
if ($maxWorkers -lt 1) { $maxWorkers = 6 }
for ($i = 1; $i -le $maxWorkers; $i++) { [void]$combo.Items.Add([string]$i) }
$combo.SelectedItem = [string][int]$state.concurrency
if ($combo.SelectedIndex -lt 0) { $combo.SelectedItem = [string]$maxWorkers }
$form.Controls.Add($combo)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Location = New-Object System.Drawing.Point(178, 190)
$btnCancel.Size = New-Object System.Drawing.Size(112, 32)
$btnCancel.Text = $labels.cancel
$form.Controls.Add($btnCancel)

# 结束态才出现的「关闭」按钮：与取消按钮同位置，运行中不可见
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(178, 190)
$btnClose.Size = New-Object System.Drawing.Size(112, 32)
$btnClose.Text = $labels.close
$btnClose.Visible = $false
$form.Controls.Add($btnClose)

$script:hintText = [string]$labels.kept
$tip = New-Object System.Windows.Forms.ToolTip

function Get-HintLayout([string]$text) {
    # 最多 3 行：测量界限用 3 行高度，控件高度再留 2px，避免最后一行贴边被裁
    $maxLinesH = $lineH * 3
    $boxH = $maxLinesH + 2
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($text, $uiFont,
        (New-Object System.Drawing.Size($innerW, $maxLinesH)),
        [System.Windows.Forms.TextFormatFlags]::WordBreak)
    if ($size.Height -le $maxLinesH) { return @{ Text = $text; Height = [Math]::Max($lineH + 4, [Math]::Min($boxH, $size.Height + 2)) } }
    # 超过 3 行：逐步截断到放得下为止
    $trim = $text
    while ($trim.Length -gt 0) {
        $trim = $trim.Substring(0, $trim.Length - 1)
        $cand = $trim.TrimEnd() + "…"
        $m = [System.Windows.Forms.TextRenderer]::MeasureText($cand, $uiFont,
            (New-Object System.Drawing.Size($innerW, $maxLinesH)),
            [System.Windows.Forms.TextFormatFlags]::WordBreak)
        if ($m.Height -le $maxLinesH) { return @{ Text = $cand; Height = [Math]::Max($lineH + 4, [Math]::Min($boxH, $m.Height + 2)) } }
    }
    return @{ Text = "…"; Height = $lineH + 4 }
}

# 窗口高度随内容增长：提示区/路径行按实际高度参与排版，其余控件位置顺延，不挤压
function Update-Layout {
    $fit = Get-HintLayout ([string]$script:hintText)
    $lblHint.Text = [string]$fit.Text
    $lblHint.Height = [int]$fit.Height
    $tip.SetToolTip($lblHint, [string]$script:hintText)
    $y = 14
    $lblCount.Top = $y; $y += 24
    $bar.Top = $y; $y += 26
    $lblFile.Top = $y; $y += 24
    $lblStage.Top = $y; $y += 24
    $lblHint.Top = $y; $y += $lblHint.Height + 8
    $combo.Top = $y
    $lblWorkers.Top = $y + 3
    $y += 28
    $btnCancel.Top = $y
    $btnClose.Top = $y
    $y += 32 + 14
    # 不超过屏幕工作区，留出标题栏/任务栏余量
    $maxH = 620
    try {
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $maxH = [Math]::Max(220, $wa.Height - 100)
    } catch { }
    $form.ClientSize = New-Object System.Drawing.Size(468, [Math]::Min($y, $maxH))
}

$script:closing = $false
$script:closeRequested = $false

function Clear-OwnedLock {
    # 只删属于自己的锁：锁里写的是「PID|进程创建时间」，新批次可能已接手，不能盲删。
    # 收尾时可能正好撞上别的进程在读锁，读失败就先重试几次，别把锁留成残留。
    if (-not $state.lock) { return $false }
    $path = ""
    try { $path = [string]$state.lock } catch { return $false }
    for ($i = 1; $i -le 4; $i++) {
        try {
            if (-not (Test-Path -LiteralPath $path)) { return $true }
            $raw = ""
            try { $raw = (Get-Content -LiteralPath $path -Raw -Encoding Ascii) } catch { $raw = $null }
            if ($null -eq $raw -or $raw -eq "") {
                if ($i -lt 4) { Start-Sleep -Milliseconds 120; continue }
                return $false
            }
            $owner = ([string]$raw).Trim().Split("|")[0]
            if ($owner -ne [string]$PID) { return $true }
            Remove-Item -LiteralPath $path -Force -ErrorAction Stop
            return $true
        } catch {
            if ($i -lt 4) { Start-Sleep -Milliseconds 120 }
        }
    }
    return $false
}

function Close-Run([string]$reason) {
    # 唯一的退出通道：停 worker → 删临时目录 → 写收尾日志 → 关窗 → 结束本进程
    # （手动关闭 / 取消 / NX 退出 / 被同一插件的新批次取代，四条路径共用）
    if ($script:closing) { return }
    $script:closing = $true
    try { $timer.Stop() } catch { }
    try { $watch.Stop() } catch { }
    try { $frontTimer.Stop() } catch { }
    try { $form.TopMost = $false } catch { }
    foreach ($r in @($script:running)) {
        try { & taskkill.exe /T /F /PID $r.proc.Id | Out-Null } catch { }
    }
    try { $script:running.Clear() } catch { }
    $runDir = ""
    try { $runDir = Split-Path -Parent ([string]$StateFile) } catch { }
    $removed = -1
    if ($runDir -and (Test-Path -LiteralPath $runDir)) {
        $removed = 0
        for ($i = 1; $i -le 5; $i++) {
            try {
                Remove-Item -LiteralPath $runDir -Recurse -Force -ErrorAction Stop
                $removed = 1
                break
            } catch { Start-Sleep -Milliseconds 200 }
        }
    }
    # 会话标记只在仍属于自己时删除，避免把新批次写的标记删掉
    try {
        if ($state.token_file -and $state.session_token) {
            $cur = (Get-Content -LiteralPath ([string]$state.token_file) -Raw -Encoding UTF8).Trim()
            if ($cur -eq [string]$state.session_token) {
                Remove-Item -LiteralPath ([string]$state.token_file) -Force -ErrorAction SilentlyContinue
            }
        }
    } catch { }
    $lockCleared = Clear-OwnedLock
    if ($state.log) {
        try {
            Add-Content -LiteralPath ([string]$state.log) -Encoding UTF8 -Value (
                "[{0}] [close:{1}] temp_removed={2} lock_cleared={3} path={4} pid={5}" -f
                (Get-Date -Format 'HH:mm:ss'), $reason, $removed,
                $(if ($lockCleared) { 1 } else { 0 }), $runDir, $PID)
        } catch { }
    }
    try { $form.Hide() } catch { }
    [Environment]::Exit(0)
}

Update-Layout

# ---- 本批次产物台账 / 停止回滚 / 转换后打开 ----
# 各后台分片把本片产物写到 run 目录的 products_N.json；这里只认这份清单，
# 不按目录或扩展名扫描，避免误删用户自己的文件。
$script:rollback = $false

function Get-BatchRecords {
    $list = New-Object System.Collections.ArrayList
    $runDir = ""
    try { $runDir = Split-Path -Parent ([string]$StateFile) } catch { $runDir = "" }
    if (-not $runDir) { return $list }
    foreach ($f in @(Get-ChildItem -LiteralPath $runDir -Filter "products_*.json" -ErrorAction SilentlyContinue)) {
        $p = $null
        try { $p = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $p = $null }
        if (-not $p) { continue }
        foreach ($it in @($p.items)) {
            if ($it) { [void]$list.Add($it) }
        }
    }
    # 不能写 return ,$list：调用方统一用 @(Get-BatchRecords) 取值，
    # 再套一层一元逗号会让 foreach 只拿到一个 ArrayList 整体，
    # [string]$rec.status 变成 "ok ok" 而被整批跳过（打开产物 / 停止回滚都会失效）。
    return $list
}

function Write-BatchLog([string]$text) {
    if (-not $state.log) { return }
    try {
        Add-Content -LiteralPath ([string]$state.log) -Encoding UTF8 -Value (
            "[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $text)
    } catch { }
}

function Remove-BatchFile([string]$path, [string]$kind) {
    # 逐个删除清单里点名的文件；被占用时只记录、不中断
    if (-not $path) { return $true }
    try {
        if (-not (Test-Path -LiteralPath $path)) { return $true }
        Remove-Item -LiteralPath $path -Force -ErrorAction Stop
        return $true
    } catch {
        Write-BatchLog ("[rollback] delete_failed kind={0} path={1}" -f $kind, $path)
        return $false
    }
}

function Invoke-Rollback {
    $deleted = 0
    $kept = 0
    $failedFiles = New-Object System.Collections.ArrayList
    foreach ($rec in @(Get-BatchRecords)) {
        $dst = [string]$rec.dst
        # 临时残留（*nxtmp_* 与其 .log）：成功/失败/未开始都可能留下，先清
        foreach ($tmp in @($rec.tmp)) {
            if ($tmp) { [void](Remove-BatchFile ([string]$tmp) "temp") }
        }
        if ([string]$rec.status -ne "ok" -or -not $dst) { continue }
        if ([bool]$rec.preexisting) {
            # 覆盖模式：目标原本就存在，本次产物已替换它；删掉等于连用户原有文件一起丢 → 保留
            $kept++
            Write-BatchLog ("[rollback] keep(overwrite) path={0}" -f $dst)
            continue
        }
        if (Remove-BatchFile $dst "product") { $deleted++ }
        else { [void]$failedFiles.Add($dst) }
    }
    Write-BatchLog ("[rollback] deleted={0} kept={1} failed={2}" -f $deleted, $kept, $failedFiles.Count)
    return @{ deleted = $deleted; kept = $kept; failed = $failedFiles }
}

function Wait-ProductReady([string]$path, [int]$timeoutMs = 15000) {
    # 有界等待产物真正写完：文件存在、非空、且能被独立读打开（worker 已释放写句柄）。
    # 条件一满足立即返回，不是盲延时；超过 timeoutMs 判为未就绪、不打开。
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        try {
            $fi = New-Object System.IO.FileInfo $path
            if ($fi.Exists -and $fi.Length -gt 0) {
                $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
                try { [void]$fs.ReadByte() } finally { $fs.Dispose() }
                return $true
            }
        } catch { }
        if ($sw.ElapsedMilliseconds -ge $timeoutMs) { return $false }
        Start-Sleep -Milliseconds 200
    }
}

function Open-Products([object[]]$files) {
    # 打开本批实际生成的 prt：入参是台账里的 dst，逐个复核「存在 + 非空 + 可读」后才打开。
    #
    # v3.3.8 修复（转换后打开空白零件）：旧代码优先用系统关联打开
    # （Start-Process -FilePath $path），但实机 .prt 关联是
    #   "D:\NX2506\NXBIN\ugs_router.exe" -ug -use_file_dir "%1"
    # 它只把「文件所在目录」传给 NX、不把文件本身交给 NX：shell 调用照样返回成功，
    # 结果是 NX 起了一个空会话/空窗口，看起来就是「转换后打开的是空白零件」。
    #
    # 实测（NX 2506）：
    #   ugs_router.exe -ug "<prt>"  → 已运行的 NX 会话里多出一个该部件的标签页（正确）；
    #   ugraf.exe "<prt>"           → 只进起始页，部件并未加载；
    #   系统关联 -use_file_dir      → 只切换目录，不打开文件（原 bug）。
    # 所以现在优先调用 ugs_router.exe 并把完整路径作为参数；router 缺失或失败才退
    # ugraf.exe，最后才退系统关联。转换工具本身就在 NX 里运行，router 会把文件交给
    # 当前会话打开，不会另起一个空 NX。不等待、不阻塞进度窗。
    $failed = New-Object System.Collections.ArrayList
    $binDir = ""
    try {
        if ($env:UGII_BASE_DIR) { $binDir = Join-Path ([string]$env:UGII_BASE_DIR) "NXBIN" }
    } catch { }
    if (-not ($binDir -and (Test-Path -LiteralPath $binDir))) {
        # NX 环境变量缺失时，按后台执行器所在目录（NXBIN）回推
        try {
            $binDir = Split-Path -Parent ([string]$state.runner)
        } catch { }
    }
    $routerExe = ""
    $nxExe = ""
    if ($binDir) {
        $routerExe = Join-Path $binDir "ugs_router.exe"
        $nxExe = Join-Path $binDir "ugraf.exe"
        if (-not (Test-Path -LiteralPath $routerExe)) { $routerExe = "" }
        if (-not (Test-Path -LiteralPath $nxExe)) { $nxExe = "" }
    }
    $opened = 0
    $methods = New-Object System.Collections.ArrayList
    foreach ($file in @($files)) {
        $path = [string]$file
        if (-not $path) { continue }
        if (-not (Wait-ProductReady $path 15000)) {
            Write-BatchLog ("[open] not-ready path={0}" -f $path)
            [void]$failed.Add($path)
            continue
        }
        $done = $false
        if ($routerExe) {
            try {
                Start-Process -FilePath $routerExe -ArgumentList '-ug', ('"' + $path + '"') -ErrorAction Stop | Out-Null
                $done = $true
                if (-not $methods.Contains("router")) { [void]$methods.Add("router") }
            } catch { }
        }
        if (-not $done -and $nxExe) {
            try {
                Start-Process -FilePath $nxExe -ArgumentList ('"' + $path + '"') -ErrorAction Stop | Out-Null
                $done = $true
                if (-not $methods.Contains("ugraf")) { [void]$methods.Add("ugraf") }
            } catch { }
        }
        if (-not $done) {
            try {
                Start-Process -FilePath $path -ErrorAction Stop | Out-Null
                $done = $true
                if (-not $methods.Contains("shell")) { [void]$methods.Add("shell") }
            } catch { }
        }
        if ($done) { $opened++ } else { [void]$failed.Add($path) }
    }
    Write-BatchLog ("[open] method={0} count={1} fail={2}" -f
        (($methods -join "+"), $opened, $failed.Count))
    return ,$failed
}

function Get-SuccessProducts {
    # 本批成功产出且已落盘的目标文件（去重），只认 products_*.json 台账
    $files = New-Object System.Collections.ArrayList
    foreach ($rec in @(Get-BatchRecords)) {
        if ([string]$rec.status -ne "ok" -or -not $rec.dst) { continue }
        $path = [string]$rec.dst
        if (-not (Test-Path -LiteralPath $path)) { continue }
        if (-not $files.Contains($path)) { [void]$files.Add($path) }
    }
    # 与 Get-BatchRecords 同一约定：调用方用 @() 包装，这里不能再套一元逗号，
    # 否则 $opens.Count 恒为 1、真正要打开的文件一个都进不了循环。
    return $files
}

$script:pending = New-Object System.Collections.ArrayList
foreach ($s in $shards) { [void]$script:pending.Add($s) }
$script:running = New-Object System.Collections.ArrayList
$script:done = $false
$script:stopAt = $null
$script:sumDone = 0; $script:sumOk = 0; $script:sumSkip = 0; $script:sumFail = 0
$script:deadShards = 0

function Get-Limit {
    $n = 0
    try { $n = [int]$combo.SelectedItem } catch { $n = 0 }
    if ($n -lt 1) { $n = 1 }
    return $n
}

function Start-Pending {
    $limit = Get-Limit
    while ($script:running.Count -lt $limit -and $script:pending.Count -gt 0) {
        $s = $script:pending[0]
        $script:pending.RemoveAt(0)
        $env:NX_BATCH_PLAN = [string]$s.plan
        # 进度窗环境继承自准备会话，可能带着 NX_BATCH_STAGE；worker 见到它会重入
        # 准备阶段（再弹一次「确认转换文件」）而不执行清单，于是永远停在「排队」。
        Remove-Item Env:\NX_BATCH_STAGE -ErrorAction SilentlyContinue
        $p = $null
        try {
            $p = Start-Process -FilePath $state.runner -ArgumentList ('-nx "' + [string]$state.journal + '"') -PassThru -WindowStyle Hidden
        } catch { }
        Remove-Item Env:\NX_BATCH_PLAN -ErrorAction SilentlyContinue
        if ($p) {
            $tries = 1
            try { $tries = [int]$s.tries } catch { $tries = 1 }
            if ($tries -lt 1) { $tries = 1 }
            [void]$script:running.Add(@{ proc = $p; shard = $s; tries = $tries; since = (Get-Date) })
        }
        else { [void]$script:pending.Insert(0, $s); break }
    }
}

function Update-Bar {
    $script:sumDone = 0; $script:sumOk = 0; $script:sumSkip = 0; $script:sumFail = 0
    $name = ""; $stage = "queue"; $latest = -1.0
    foreach ($s in $shards) {
        $path = [string]$s.progress
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try { $p = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
        if (-not $p) { continue }
        $script:sumDone += [Math]::Max([int]$p.done, [int]$p.index)
        $script:sumOk += [int]$p.ok
        $script:sumSkip += [int]$p.skip
        $script:sumFail += [int]$p.fail
        if ([double]$p.updated -ge $latest -and [string]$p.name -ne "") {
            $latest = [double]$p.updated; $name = [string]$p.name; $stage = [string]$p.stage
        }
    }
    # 没写出任何进度就死掉的分片计入失败，避免「总数看起来没跑完、界面却没事」
    $script:sumFail += [int]$script:deadShards
    $bar.Value = [Math]::Max(0, [Math]::Min($bar.Maximum, $script:sumDone))
    $lblCount.Text = $labels.count -f $script:sumDone, $state.total
    if ($name) { $lblFile.Text = $labels.file -f $name } else { $lblFile.Text = "" }
    $lblStage.Text = $labels.stage -f (Stage-Text $stage)
}

function Finish-Run {
    if ($script:done) { return }
    $script:done = $true
    $timer.Stop()
    Clear-OwnedLock
    Update-Bar
    $summary = $labels.finished
    if ($script:cancelled) { $summary = $labels.cancelled }
    $lblStage.Text = $summary -f $script:sumOk, $script:sumSkip, $script:sumFail
    $lblFile.Text = ""
    $combo.Enabled = $false
    $btnCancel.Enabled = $false
    if ($script:rollback) {
        # 停止=整批回滚：按台账删除本批次产物，清空进度/计数。
        # 结果只留在这一行和日志里，不弹必须先点掉的窗口；短暂停留后自动关窗回 NX。
        $result = Invoke-Rollback
        $lblStage.Text = [string]$labels.rollback_done -f $result.deleted, $result.kept, $result.failed.Count
        $lblCount.Text = ""
        $lblFile.Text = ""
        $bar.Value = $bar.Minimum
        Write-BatchLog ("[rollback] done deleted={0} kept={1} failed={2}" -f
            $result.deleted, $result.kept, $result.failed.Count)
        Start-Sleep -Milliseconds 900
        Close-Run "rollback"
        return
    }
    # 选择窗勾了「打开转换后的文件」：整批结束即打开本批成功产物，不等关窗、不再询问
    if ($script:cancelled) {
        Write-BatchLog "[open] skipped reason=cancel"
    } elseif (-not $state.auto_open) {
        Write-BatchLog ("[open] skipped reason=flag auto_open={0}" -f $(if ($state.auto_open) { 1 } else { 0 }))
    } else {
        try {
            $opens = @(Get-SuccessProducts)
            Write-BatchLog ("[open] candidates={0}" -f $opens.Count)
            if ($opens.Count -gt 0) { [void](Open-Products $opens) }
            else { Write-BatchLog "[open] skipped count=0" }
        } catch {
            Write-BatchLog ("[open] error {0}" -f $_.Exception.Message)
        }
    }
    # 结束态：显示「关闭」按钮（运行中只有取消），并把窗口提到最前
    # （抢前台最长重试 8.4s；仍被 NX 抢占则保留置顶，用户点击/关窗时解除）
    $btnCancel.Visible = $false
    $btnClose.Visible = $true
    try {
        if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
            $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        }
    } catch { }
    Update-Layout
    Start-ProgressFront -HoldTopMost
    # 用户在运行中点过关闭（已转成取消）：收尾完成后直接退出，不再留窗
    if ($script:closeRequested) { Close-Run "cancel" }
    # 正常跑完：停在结果上，点「关闭」或右上角 X 都走 Close-Run，清理一致
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 600
$timer.add_Tick({
    if ($script:done) { return }
    Start-Pending
    for ($i = $script:running.Count - 1; $i -ge 0; $i--) {
        $r = $script:running[$i]
        $exited = $false
        try { $exited = $r.proc.HasExited } catch { $exited = $true }
        $noProgress = -not (Test-Path -LiteralPath ([string]$r.shard.progress))
        # 启动超时：进程还在却从未写进度（卡在启动或被环境带偏），强制结束转入重试。
        # 阈值放宽到 180s：并发 6 时 NX 批处理冷启动本身就要几十秒，不能误杀。
        if (-not $exited -and $noProgress -and ((Get-Date) - $r.since).TotalSeconds -gt 180) {
            Write-BatchLog ("[worker] startup-timeout shard={0}" -f $r.shard.plan)
            try { & taskkill.exe /T /F /PID $r.proc.Id | Out-Null } catch { }
            $exited = $true
        }
        if (-not $exited) { continue }
        $script:running.RemoveAt($i)
        if (-not $noProgress) { continue }
        # 退出却没产出进度：重试一次；仍失败就把该分片判失败，不能让整批无限停在「排队」
        $tries = [int]$r.tries + 1
        if ($tries -le 2) {
            [void]$script:pending.Insert(0, @{ plan = $r.shard.plan; progress = $r.shard.progress; tries = $tries })
            Write-BatchLog ("[worker] retry shard={0} attempt={1}" -f $r.shard.plan, $tries)
        } else {
            $script:deadShards++
            Write-BatchLog ("[worker] failed shard={0} reason=no-progress attempts={1}" -f $r.shard.plan, $tries)
        }
    }
    Start-Pending
    Update-Bar
    if ($script:stopAt -ne $null -and $script:running.Count -gt 0 -and
        ((Get-Date) - $script:stopAt).TotalSeconds -gt 10) {
        $script:cancelled = $true
        foreach ($r in $script:running) {
            try { & taskkill.exe /T /F /PID $r.proc.Id | Out-Null } catch { }
        }
        $script:running.Clear()
    }
    if ($script:pending.Count -eq 0 -and $script:running.Count -eq 0) { Finish-Run }
})

$btnCancel.add_Click({
    if ($script:done) { return }
    # 停止=整批回滚，先确认：会删掉本次已经生成的文件（覆盖模式的原文件无法恢复）
    $choice = [System.Windows.Forms.MessageBox]::Show($form, [string]$labels.stop_confirm,
        [string]$labels.stop_title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $script:rollback = $true
    $btnCancel.Enabled = $false
    $btnCancel.Text = $labels.stopping
    $script:cancelled = $true
    $script:pending.Clear()
    try { Set-Content -LiteralPath ([string]$state.cancel) -Value "1" -Encoding Ascii } catch { }
    $script:stopAt = Get-Date
})

# 结束态的「关闭」：与右上角 X 走同一条 Close-Run 收尾路径
$btnClose.add_Click({
    if (-not $script:done) { return }
    Close-Run "close"
})

function Test-NxAlive {
    # 判据与单实例锁完全一致：进程在 + 创建时间一致 + 映像名一致。
    # 只要进程确实查不到就判死（强杀/正常退出都一样）；查得到但时间/名字对不上
    # 说明 PID 已被系统复用，也判死。读不到时间时不做这一步，避免权限问题误关窗。
    $nxPid = 0
    try { $nxPid = [int]$state.nx_pid } catch { $nxPid = 0 }
    if ($nxPid -le 0) { return $true }
    $proc = $null
    try { $proc = [System.Diagnostics.Process]::GetProcessById($nxPid) } catch { $proc = $null }
    if ($null -eq $proc) { return $false }
    $alive = $true
    try {
        $want = [string]$state.nx_stamp
        if ($want) {
            $now = ""
            try { $now = [string]$proc.StartTime.ToFileTimeUtc() } catch { $now = "" }
            if ($now -and $now -ne $want) { $alive = $false }
        }
        if ($alive) {
            $name = [string]$state.nx_name
            if ($name) {
                $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
                if ($base -and $proc.ProcessName -ne $base) { $alive = $false }
            }
        }
    } catch { $alive = $true }
    try { $proc.Dispose() } catch { }
    return $alive
}

# 看门狗：不管在哪个状态（跑批中 / 跑完停住 / 空闲等待关窗），只要发起它的 NX
# 已经不在了就立刻按 Close-Run 收尾并退进程。用独立 Timer 而不是批处理 $timer，
# 因为 $timer 在 Finish-Run 里会停掉——那正是「跑完停住」漏关窗的原因。
$watch = New-Object System.Windows.Forms.Timer
$watch.Interval = 1000
$watch.add_Tick({
    if ($script:closing) { return }
    try {
        if (-not (Test-NxAlive)) {
            Write-BatchLog ("[watch] nx_exit pid={0} stamp={1}" -f $state.nx_pid, $state.nx_stamp)
            Close-Run "nx_exit"
            return
        }
        # 同一插件再次启动批量时会改写会话标记：读到的不是自己的就说明本窗已过期
        if ($state.token_file) {
            $tok = ""
            try { $tok = (Get-Content -LiteralPath ([string]$state.token_file) -Raw -Encoding UTF8).Trim() } catch { $tok = "" }
            if ($tok -and $tok -ne [string]$state.session_token) { Close-Run "superseded" }
        }
    } catch {
        # 收尾过程中（写日志/删目录）偶发异常不能让看门狗自己死掉，下一拍继续判
        try { Write-BatchLog ("[watch] error {0}" -f $_.Exception.Message) } catch { }
    }
})

$form.add_FormClosing({
    param($sender, $e)
    if ($script:closing) { return }
    if ($script:done) { $e.Cancel = $true; Close-Run "user"; return }
    # 还没跑完就想关：先转成取消，等收尾完成再退出（Finish-Run 里会调 Close-Run）
    $e.Cancel = $true
    $script:closeRequested = $true
    if ($script:stopAt -eq $null) {
        # 关窗只是取消，不删产物；整批回滚只走「停止」按钮
        $script:cancelled = $true
        $script:pending.Clear()
        try { Set-Content -LiteralPath ([string]$state.cancel) -Value "1" -Encoding Ascii } catch { }
        $script:stopAt = Get-Date
    }
})

$form.add_FormClosed({
    # 兜底：任何非 Close-Run 的关闭路径也不留临时目录
    try {
        $runDir = Split-Path -Parent ([string]$StateFile)
        if ($runDir -and (Test-Path -LiteralPath $runDir)) {
            Remove-Item -LiteralPath $runDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { }
})

function Set-FrontOnce {
    # 单次置前：恢复最小化 → 挂到前台线程 → 抬升/激活 → 精确句柄判定
    param([bool]$TopMost = $false)
    $h = [IntPtr]::Zero
    try { $h = $form.Handle } catch { return $false }
    if ($h -eq [IntPtr]::Zero) { return $false }
    $fgTid = 0; $myTid = 0; $attached = $false
    try {
        if ([NBAct]::IsIconic($h)) { [void][NBAct]::ShowWindow($h, 9) }
        $fg = [NBAct]::GetForegroundWindow()
        [void][NBAct]::GetWindowThreadProcessId($fg, [ref]$fgTid)
        $myTid = [NBAct]::GetCurrentThreadId()
        if ($fgTid -ne 0 -and $myTid -ne 0 -and $fgTid -ne $myTid) {
            $attached = [NBAct]::AttachThreadInput($myTid, $fgTid, $true)
        }
    } catch { }
    try {
        if ($TopMost) {
            # 抬到 NX 之上（不改位置/大小/激活状态用后面的 SetForegroundWindow 补）
            [void][NBAct]::SetWindowPos($h, [IntPtr](-1), 0, 0, 0, 0, 0x0043)
        }
        [void][NBAct]::BringWindowToTop($h)
        [void][NBAct]::SetForegroundWindow($h)
        [void][NBAct]::SetFocus($h)
        $form.Activate()
    } catch { }
    if ($attached) { try { [void][NBAct]::AttachThreadInput($myTid, $fgTid, $false) } catch { } }
    try { return ([NBAct]::GetForegroundWindow().ToInt64() -eq $h.ToInt64()) } catch { return $false }
}

function Stop-ProgressFront {
    param([string]$Result, [int]$Tries, [bool]$Self, [bool]$Visible, [bool]$Iconic, [switch]$KeepTopMost)
    try { $frontTimer.Stop() } catch { }
    if ($KeepTopMost) {
        # 抢不到前台时保留置顶：窗口始终压在 NX 之上，用户点一下或关窗时再解除
        $script:frontHold = $true
        try { $form.TopMost = $true } catch { }
    } elseif ($script:frontHold) {
        try { $form.TopMost = $false } catch { }
        try { [void][NBAct]::SetWindowPos($form.Handle, [IntPtr](-2), 0, 0, 0, 0, 0x0043) } catch { }
        $script:frontHold = $false
    }
    if ($state.log) {
        try {
            Add-Content -LiteralPath ([string]$state.log) -Encoding UTF8 -Value (
                "[{0}] [progress] front={1} try={2} fg_self={3} visible={4} iconic={5}" -f
                (Get-Date -Format 'HH:mm:ss'), $Result, $Tries,
                $(if ($Self) { 1 } else { 0 }), $(if ($Visible) { 1 } else { 0 }), $(if ($Iconic) { 1 } else { 0 }))
        } catch { }
    }
}

function Start-ProgressFront {
    # 用 Timer 分拍重试：每拍之间消息泵真正跑起来，窗口管理器才有机会落实激活。
    # 结束态首拍前短暂置顶，保证即使 NX 立刻抢回焦点，窗口也压在 NX 之上可见。
    param([switch]$HoldTopMost)
    $script:frontTries = 0
    $script:frontSeen = $false
    $script:frontHold = [bool]$HoldTopMost
    if ($script:frontHold) {
        try { $form.TopMost = $true } catch { }
    }
    $script:frontTimer.Start()
}

$script:frontTries = 0
$script:frontHold = $false
# 结束态重试窗口：24 拍 × 350ms ≈ 8.4s，够覆盖「打开产物后 NX 冷启动抢焦点」的整段
$script:frontMaxTries = 24
$script:frontTimer = New-Object System.Windows.Forms.Timer
$script:frontTimer.Interval = 350
$script:frontTimer.add_Tick({
    if ($script:closing) { try { $frontTimer.Stop() } catch { }; return }
    $script:frontTries++
    $ok = Set-FrontOnce ([bool]$script:frontHold)
    $vis = $false; $ico = $false
    try { $vis = [NBAct]::IsWindowVisible($form.Handle); $ico = [NBAct]::IsIconic($form.Handle) } catch { }
    # 至少跑满 3 拍（≈1s）再收工：覆盖 worker 退出、NX 抢回焦点的那几百毫秒
    if ($ok -and $script:frontTries -ge 3) {
        Stop-ProgressFront "ok" $script:frontTries $true $vis $ico
        return
    }
    if ($script:frontTries -ge $script:frontMaxTries) {
        if ($ok) {
            Stop-ProgressFront "ok" $script:frontTries $true $vis $ico
            return
        }
        # 长时间抢不到前台（NX 持续抢占）：不撤置顶，保持窗口压在 NX 之上并闪烁提示
        try {
            $fi = [NBAct+FLASHWINFO]::new()
            $fi.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf($fi)
            $fi.hwnd = $form.Handle
            $fi.dwFlags = [uint32](0x00000003 -bor 0x0000000C)
            $fi.uCount = 3
            $fi.dwTimeout = 0
            [void][NBAct]::FlashWindowEx([ref]$fi)
        } catch { }
        Stop-ProgressFront "topmost" $script:frontTries $false $vis $ico -KeepTopMost
    }
})

$form.add_Shown({
    if ($state.lock) {
        # 进度窗接手锁时同样写「PID|进程创建时间」，与 _acquire_run_lock 保持一致；
        # 只写 PID 会让陈旧锁在 PID 被复用后永远判忙。
        try {
            $stamp = ""
            try { $stamp = [string]([Diagnostics.Process]::GetCurrentProcess().StartTime.ToFileTimeUtc()) } catch { }
            Set-Content -LiteralPath ([string]$state.lock) -Value ("{0}|{1}" -f $PID, $stamp) -Encoding Ascii
        } catch { }
    }
    Start-ProgressFront
    Start-Pending
    Update-Bar
    $timer.Start()
    $watch.Start()
})

$form.add_Activated({
    # 重试结束后仍保留的置顶（front=topmost）：用户一旦点到自己窗口就解除，避免一直压着 NX
    if ($script:frontHold -and -not $script:frontTimer.Enabled) {
        try { $form.TopMost = $false } catch { }
        try { [void][NBAct]::SetWindowPos($form.Handle, [IntPtr](-2), 0, 0, 0, 0, 0x0043) } catch { }
        $script:frontHold = $false
    }
})

[void]$form.ShowDialog()
'''


def _runner_api():
    """原 nx_batch_runner.py 的对外接口（合并后按命名空间暴露，调用点无需改动）"""
    return types.SimpleNamespace(
        execute=execute,
        load_plan=load_plan,
        write_json=write_json,
        start=start,
        activate_picker=activate_picker,
        start_waitdlg_hider=start_waitdlg_hider,
        normalize_name=rn_normalize_name,
        safe_name_part=rn_safe_name_part,
        LEFT_LIST_PS=LEFT_LIST_PS,
        NAME_HELPERS_PS=NAME_HELPERS_PS,
        PICKER_FRONT_PS=PICKER_FRONT_PS)

# ==============================================================================
# 段：批量导入（原 batch_import.py）
# ==============================================================================
# 批量导入多个文件：每个文件各自成为一个独立部件。
# 
# 支持格式：
#     .prt                 直接打开为显示部件（已在会话中打开则切换为活动显示）
#     .stp / .step         STEP AP214 导入到新建部件
#     .igs / .iges         IGES 导入到新建部件
#     .x_t / .x_b          Parasolid 导入到新建部件
# 
# 行为约定：
#     1. 多选使用 Windows 系统文件对话框（ctypes 调 comdlg32.GetOpenFileNameW），
#        不依赖 tkinter —— NX 内嵌 Python 不一定带 tkinter。
#     2. 每个文件一个独立部件；不把几何导入到当前工作部件里。
#     3. 绝不覆盖磁盘上的用户文件：.prt 输入只打开、只显示、不另存；
#        转换类（STEP/IGES/Parasolid）另存为源文件同目录下的同名 .prt，
#        同名已存在则加 _2/_3。
#     4. 单个文件失败只记录、继续下一个，不中断整批。
#     5. 进度写 NX 信息窗口（ListingWindow），日志写插件目录 logs\batch_import.log。
#     6. 取消选择即静默退出。
# 
# 注意：不要在本文件里定义 startup(argc, args)。
# startup 目录下的 .py 会被 NX 当成“启动应用”导入；一旦它有 startup 函数，
# 之后点菜单用 journal 播放同一模块时 nxpythonldr 会崩（syslog 里是 signal 11，
# 界面弹空白的 Journal execution error）。没有 startup 时 NX 只会在 syslog 记一条
# “no attribute startup”的提示，功能正常。
# -*- coding: utf-8 -*-

import os
import re
import time
import tempfile
import shutil
import ctypes

import NXOpen

# ================= 多语言 =================
# UGII_LANG 的常见取值：simpl_chinese / trad_chinese / english
# 未知或未设置时回落英文；三张表缺项时逐项回落英文。


def imp__detect_lang():
    value = (os.environ.get("UGII_LANG") or "").strip().lower()
    if value in ("simpl_chinese", "simplified_chinese", "zh_cn", "zh-cn", "chs",
                 "schinese", "zh_hans"):
        return "zh_cn"
    if value in ("trad_chinese", "traditional_chinese", "zh_tw", "zh-tw", "cht",
                 "chinese", "zh_hant"):
        return "zh_tw"
    return "en"


imp_LANG = imp__detect_lang()

imp_MSG = {
    "en": {
        "imp.title": "Select files to import",
        "imp.filter": "Supported files",
        "imp.filter_all": "All files",
        "imp.ui.title": "Select files to import",
        "imp.ui.tree_label": "Folders",
        "imp.ui.list_label": "Files (Ctrl/Shift multi-select, double-click to add)",
        "imp.ui.col_name": "File name",
        "imp.ui.col_size": "Size (KB)",
        "imp.ui.this_pc": "This PC",
        "imp.ui.denied": "Cannot access this folder (permission denied or invalid path):",
        "imp.ui.denied_hint": "The folder does not exist or access is denied.",
        "imp.ui.empty_dir": "(No importable file in this folder)",
        "imp.ui.grid_label": "Selected files ({0})",
        "imp.ui.col_type": "Type",
        "imp.ui.col_path": "Path",
        "imp.ui.btn_add_file": "Add import files",
        "imp.ui.btn_add_folder": "+ Add folder",
        "imp.ui.name_prefix": "Add prefix",
        "imp.ui.name_suffix": "Add suffix",
        "imp.ui.btn_clear": "Clear",
        "imp.ui.btn_del_rows": "Delete selected",
        "imp.ui.menu_delete": "Delete selected",
        "imp.ui.menu_select_all": "Select all",
        "imp.ui.menu_clear": "Clear",
        "imp.ui.btn_cancel": "Cancel",
        "imp.ui.btn_ok": "OK",
        "imp.ui.mode_open": "Open in NX after import",
        "imp.ui.mode_files": "Only create files",
        "imp.ui.mode_group": "Import mode",
        "imp.ui.note_line": "(.prt opens as is; prefix/suffix on saved files)",
        "imp.ui.on_exists_group": "If the target exists",
        "imp.ui.on_exists_rename": "Rename (_2, _3 ...)",
        "imp.ui.on_exists_skip": "Skip",
        "imp.ui.on_exists_overwrite": "Overwrite",
        "imp.ui.out_group": "Output folder",
        "imp.ui.out_src": "Output to source folder",
        "imp.ui.out_dir": "Output to a specified folder",
        "imp.ui.out_browse": "Browse...",
        "imp.ui.out_invalid": "The specified output folder is invalid; falling back to the source folder.",
        "imp.ui.skip_exists": " (skipped)",
        "imp.ui.overwrite_exists": " (overwritten)",
        "imp.ui.right_title": "Converted name",
        "imp.ui.col_no": "No.",
        "imp.ui.col_file": "File name",
        "imp.ui.left_empty": "(Double-click a file above to add)",
        "imp.ui.right_empty": "(Shown after files are added)",
        "imp.ui.open_only": "(open only)",
        "imp.ui.save_rename": " (renamed)",
        "imp.ui.workers_label": "Concurrent processes",
        "imp.ui.msg_remove_title": "Remove file",
        "imp.ui.msg_remove_text": "Remove this file from the list?",
        "imp.ui.msg_empty_title": "Notice",
        "imp.ui.msg_empty_one": "Nothing has been added yet.",
        "imp.ui.msg_empty_two": "Select files on the left first, then start importing.",
        "imp.err.dialog": "Cannot open the system file dialog: {0}",
        "imp.err.background": "The background NX session could not be started, so the import was not run. Please try again or check the log.",
        "imp.err.picker_hidden": "File picker window is not visible (pid {0}). It was likely launched with SW_HIDE; the console must be hidden by CREATE_NO_WINDOW only.",
        "imp.err.locked": "Target file is in use or not writable: {0}",
        "imp.err.src_readonly": "The source folder is not writable",
        "imp.err.out_readonly": "The specified output folder is not writable",
        "imp.err.readonly_title": "Folder not writable",
        "imp.err.readonly_body": "These folders cannot be written; the affected files will be reported as failed:\n\n{0}",
        "imp.err.not_displayed": "The file was not displayed (bad file or unsupported format)",
        "imp.err.no_geometry": "Translation produced no solid/sheet, treated as failed",
        "imp.log.session_ok": "  Displayed {0} in {1:.2f}s",
        "imp.log.session_skip": "  Skipped (target exists): {0}",
        "imp.log.session_already": "  Already open, switched to it: {0}",
        "imp.log.cancelled": "Cancelled: {0} file(s) not processed",
        "imp.sum.interactive": "Opened in the current session: {0} file(s) in {1:.1f}s.",
        "imp.sum.cancelled_note": "Cancelled; {0} file(s) were not processed.",
        "imp.prog.title": "Importing into this session",
        "imp.prog.cancel": "Cancel",
        "imp.prog.cancelling": "Cancelling...",
        "imp.prog.stage_convert": "Translating and opening...",
        "imp.prog.stage_save": "Saving...",
        "imp.prog.stage_done": "Done",
        "imp.prog.stage_skip": "Skipped (target exists)",
        "imp.prog.stage_fail": "Failed",
        "imp.confirm.title": "Batch Import",
        "imp.confirm.body": "{0} file(s) to import:\n\n{1}",
        "imp.confirm.more": "\n... and {0} more (see the log)",
        "imp.confirm.proceed_log": "Each file becomes an independent part. .prt is opened only; STEP/IGES/Parasolid is saved as <prefix><source><suffix>.prt next to the source (_2/_3 if taken).",
        "imp.sum.title": "Batch Import Finished",
        "imp.sum.body": "Succeeded {0} / Failed {1} / Skipped {2}\n\n{3}\n\nLog: {4}",
        "imp.sum.fail_item": "- {0}: {1}",
        "imp.sum.no_fail": "(no failures)",
        "imp.sum.more_fail": "... and {0} more failure(s), see the log",
        "imp.sum.skipped": "{0} file(s) skipped (duplicate or missing), see the log.",
        "imp.sum.disk": "Saved to file {0} / memory only {1}",
        "imp.sum.temp_logs": "Failed items keep translator logs in: {0}",
        "imp.log.header": "Batch import started: {0} file(s); log {1}",
        "imp.log.file": "[{0}/{1}] {2}",
        "imp.log.skip_dup": "  Skipped (duplicate): {0}",
        "imp.log.skip_missing": "  Skipped (file not found): {0}",
        "imp.log.skip_same": "  Skipped (target exists, keep existing): {0}",
        "imp.log.prt_open": "  Open .prt: {0}",
        "imp.log.prt_already": "  Already open in this session, switched to active display",
        "imp.log.prt_done": "  Done {0:.1f}s -> part {1}",
        "imp.log.new_part": "  New display part: {0}",
        "imp.log.import_begin": "  Import {0} -> part {1}",
        "imp.log.import_done": "  Done {0:.1f}s",
        "imp.log.saveas_start": "  Save as: {0}",
        "imp.log.saveas_ok": "  Saved as: {0}",
        "imp.log.saveas_rename": "  Target exists, renamed to: {0}",
        "imp.log.backup": "  Original backed up as: {0}",
        "imp.err.backup": "  Could not back up {0} -> {1} ({2}); overwrite aborted, original left untouched",
        "imp.err.saveas": "  Save failed ({0}): {1}; geometry kept in memory (part not closed)",
        "imp.err.saveas_missing": "  Save failed: target file not written {0}",
        "imp.log.load_status": "  Part load status: {0}",
        "imp.log.fail": "  Failed: {0}",
        "imp.log.fit_fail": "  Fit view failed (ignored): {0}",
        "imp.log.destroy_fail": "  Destroy importer failed (ignored): {0}",
        "imp.log.work_part": "  Work part: {0}",
        "imp.log.close_empty": "  Empty part closed: {0}",
        "imp.log.close_fail": "  Cannot close the empty part (kept in session): {0}",
        "imp.log.import_to": "  ImportTo not set (continuing with default work part): {0}",
        "imp.log.prep": "  Importer prepared: {0}; OutputFile={1}",
        "imp.log.prep_fail": "  Cannot prepare importer (continuing with defaults): {0}",
        "imp.log.settings_missing": "  STEP214 settings file not found, using the importer defaults",
        "imp.log.sew_probe": "  {0} importer default SewSurfaces = {1}",
        "imp.log.sew_probe_fail": "  Cannot read {0} importer SewSurfaces: {1}",
        "imp.log.sew_set": "  {0} importer SewSurfaces set to True",
        "imp.log.sew_set_fail": "  Cannot set {0} importer SewSurfaces: {1}",
        "imp.log.sew_unavailable": "  {0} importer has no SewSurfaces property, skipped",
        "imp.log.temp_clean": "  Removed {0} translator log(s) of this run from: {1}",
        "imp.log.temp_clean_fail": "  Failed to remove {0} translator log(s) of this run: {1}",
        "imp.log.temp_clean_old": "  Removed {0} translator log(s) older than {1} day(s)",
        "imp.log.summary": "Summary: succeeded {0} / failed {1} / skipped {2}",
        "imp.log.background": "Started {0} job(s) in background sessions; you can keep working in NX.",
        "imp.log.background_fail": "Background session unavailable ({0}), running in the foreground.",
        "imp.log.log_fallback": "Plugin folder not writable, log redirected to: {0}",
        "imp.log.no_confirm": "Confirmation dialog unavailable, continuing",
        "imp.log.open_mode": "  Open mode: {0}",
        "imp.log.keep_open": "  Work/display switch: {0}",
        "imp.log.parts_check": "  Session check: {0} part(s) expected, {1} still loaded",
        "imp.log.parts_restored": "  Re-opened {0} part(s) that the display switch had closed",
        "imp.log.parts_missing": "  Parts no longer in the session and cannot be re-opened: {0}",
        "imp.log.restore_fail": "  Re-open failed ({0}): {1}",
        "imp.reason.open": "Cannot open part: {0}",
        "imp.reason.import": "Import failed: {0}",
        "imp.reason.no_part": "Cannot create new part: {0}",
        "imp.reason.work_part": "Cannot set the new part as work part: {0}",
        "imp.reason.ext": "Unsupported file type",
        "imp.msg.cancelled": "Cancelled, nothing imported.",
        "imp.msg.no_files": "No importable file.",
    },
    "zh_cn": {
        "imp.title": "选择要导入的文件",
        "imp.filter": "支持的文件",
        "imp.filter_all": "所有文件",
        "imp.ui.title": "选择要导入的文件",
        "imp.ui.tree_label": "文件夹",
        "imp.ui.list_label": "文件（Ctrl/Shift 多选，双击加入）",
        "imp.ui.col_name": "文件名称",
        "imp.ui.col_size": "大小(KB)",
        "imp.ui.this_pc": "此电脑",
        "imp.ui.denied": "无法访问该目录（权限不足或路径无效）：",
        "imp.ui.denied_hint": "文件夹不存在或没有访问权限。",
        "imp.ui.empty_dir": "（该目录下没有可导入的文件）",
        "imp.ui.grid_label": "已选文件清单（共 {0} 个）",
        "imp.ui.col_type": "类型",
        "imp.ui.col_path": "路径",
        "imp.ui.btn_add_file": "添加导入文件",
        "imp.ui.btn_add_folder": "＋ 添加文件夹",
        "imp.ui.name_prefix": "文件添加前缀",
        "imp.ui.name_suffix": "文件添加后缀",
        "imp.ui.btn_clear": "清空",
        "imp.ui.btn_del_rows": "删除选中",
        "imp.ui.menu_delete": "删除选中",
        "imp.ui.menu_select_all": "全选",
        "imp.ui.menu_clear": "清空",
        "imp.ui.btn_cancel": "取消",
        "imp.ui.btn_ok": "确定",
        "imp.ui.mode_open": "导入后打开",
        "imp.ui.mode_files": "仅生成文件",
        "imp.ui.mode_group": "导入方式",
        "imp.ui.note_line": "（.prt 仅打开；前后缀只对另存生效）",
        "imp.ui.on_exists_group": "同名文件已存在时",
        "imp.ui.on_exists_rename": "改名（_2、_3…）",
        "imp.ui.on_exists_skip": "跳过",
        "imp.ui.on_exists_overwrite": "覆盖",
        "imp.ui.out_group": "输出目录",
        "imp.ui.out_src": "输出到源文件目录",
        "imp.ui.out_dir": "输出到指定目录",
        "imp.ui.out_browse": "浏览…",
        "imp.ui.out_invalid": "指定输出目录无效，已回落到源文件目录。",
        "imp.ui.skip_exists": "（跳过）",
        "imp.ui.overwrite_exists": "（覆盖）",
        "imp.ui.right_title": "转换后名称",
        "imp.ui.col_no": "序号",
        "imp.ui.col_file": "文件名称",
        "imp.ui.left_empty": "（双击上方文件加入）",
        "imp.ui.right_empty": "（添加文件后显示）",
        "imp.ui.open_only": "（仅打开）",
        "imp.ui.save_rename": "（改名）",
        "imp.ui.workers_label": "同时执行数",
        "imp.ui.msg_remove_title": "移除文件",
        "imp.ui.msg_remove_text": "从已选清单中移除这个文件？",
        "imp.ui.msg_empty_title": "提示",
        "imp.ui.msg_empty_one": "尚未添加任何文件。",
        "imp.ui.msg_empty_two": "请先在左侧选择文件，再开始导入。",
        "imp.err.dialog": "无法打开系统文件选择框：{0}",
        "imp.err.background": "后台 NX 会话未能启动，本次导入没有执行。请重试或查看日志。",
        "imp.err.picker_hidden": "选择窗进程已启动但窗口不可见（pid {0}），疑似被 SW_HIDE 启动；隐藏控制台只能靠 CREATE_NO_WINDOW。",
        "imp.err.locked": "目标文件被占用或无权限：{0}",
        "imp.err.src_readonly": "源文件所在目录不可写",
        "imp.err.out_readonly": "指定的输出目录不可写",
        "imp.err.readonly_title": "目录不可写",
        "imp.err.readonly_body": "以下目录不可写，相关文件将直接判为失败：\n\n{0}",
        "imp.err.not_displayed": "文件未被显示（可能是坏文件或格式不支持）",
        "imp.err.no_geometry": "翻译完成但没有实体/片体，判为失败",
        "imp.log.session_ok": "  已显示 {0}（{1:.2f}s）",
        "imp.log.session_skip": "  跳过（目标已存在）：{0}",
        "imp.log.session_already": "  已在会话中，切换为活动显示：{0}",
        "imp.log.cancelled": "已取消：剩余 {0} 个文件未处理",
        "imp.sum.interactive": "已在当前会话打开 {0} 个文件，总耗时 {1:.1f}s。",
        "imp.sum.cancelled_note": "已取消，另有 {0} 个文件未处理。",
        "imp.prog.title": "正在导入到当前会话",
        "imp.prog.cancel": "取消",
        "imp.prog.cancelling": "正在取消…",
        "imp.prog.stage_convert": "正在翻译并打开…",
        "imp.prog.stage_save": "正在另存…",
        "imp.prog.stage_done": "完成",
        "imp.prog.stage_skip": "已跳过（目标已存在）",
        "imp.prog.stage_fail": "失败",
        "imp.confirm.title": "批量导入",
        "imp.confirm.body": "共 {0} 个文件待导入：\n\n{1}",
        "imp.confirm.more": "\n…… 其余 {0} 个见日志",
        "imp.confirm.proceed_log": "确认后每个文件各自打开为独立部件；.prt 仅打开、不另存，STEP/IGES/Parasolid 另存为源文件同目录下的 .prt（重名加 _2/_3）。",
        "imp.sum.title": "批量导入完成",
        "imp.sum.body": "成功 {0} / 失败 {1} / 跳过 {2}\n\n{3}\n\n日志：{4}",
        "imp.sum.fail_item": "· {0}：{1}",
        "imp.sum.no_fail": "（没有失败项）",
        "imp.sum.more_fail": "…… 另有 {0} 个失败项，见日志",
        "imp.sum.skipped": "另有 {0} 个文件被跳过（重复选择或文件不存在），详见日志。",
        "imp.sum.disk": "已保存文件 {0} / 仅内存 {1}",
        "imp.sum.temp_logs": "失败项的转换器日志保留在：{0}",
        "imp.log.header": "开始批量导入：共 {0} 个文件；日志 {1}",
        "imp.log.file": "[{0}/{1}] {2}",
        "imp.log.skip_dup": "  跳过（重复选择）：{0}",
        "imp.log.skip_missing": "  跳过（文件不存在）：{0}",
        "imp.log.skip_same": "  跳过（目标已存在，保留原文件）：{0}",
        "imp.log.prt_open": "  打开 .prt：{0}",
        "imp.log.prt_already": "  该部件已在会话中打开，切换为活动显示",
        "imp.log.prt_done": "  完成 {0:.1f}s -> 部件 {1}",
        "imp.log.new_part": "  新建显示部件：{0}",
        "imp.log.import_begin": "  导入 {0} -> 部件 {1}",
        "imp.log.import_done": "  完成 {0:.1f}s",
        "imp.log.saveas_start": "  开始另存：{0}",
        "imp.log.saveas_ok": "  另存成功：{0}",
        "imp.log.saveas_rename": "  同名文件已存在，改名为：{0}",
        "imp.log.backup": "  原文件已备份为：{0}",
        "imp.err.backup": "  无法备份已存在的文件 {0} → {1}（{2}）；已中止覆盖，原文件保持不动",
        "imp.err.saveas": "  保存失败（{0}）：{1}；几何保留在内存，未关闭部件",
        "imp.err.saveas_missing": "  保存失败：目标文件未生成 {0}",
        "imp.log.load_status": "  部件加载状态：{0}",
        "imp.log.fail": "  失败：{0}",
        "imp.log.fit_fail": "  拟合视图失败（忽略）：{0}",
        "imp.log.destroy_fail": "  销毁导入器失败（忽略）：{0}",
        "imp.log.work_part": "  工作部件：{0}",
        "imp.log.close_empty": "  已关闭空部件：{0}",
        "imp.log.close_fail": "  空部件关闭失败（保留在会话中）：{0}",
        "imp.log.import_to": "  ImportTo 设置失败（按默认工作部件继续）：{0}",
        "imp.log.prep": "  已补齐导入器参数：{0}；OutputFile={1}",
        "imp.log.prep_fail": "  导入器参数补齐失败（按默认继续）：{0}",
        "imp.log.settings_missing": "  未找到 STEP214 设置文件，按导入器默认值处理",
        "imp.log.sew_probe": "  {0} 导入器当前 SewSurfaces 默认值 = {1}",
        "imp.log.sew_probe_fail": "  读取 {0} 导入器 SewSurfaces 失败：{1}",
        "imp.log.sew_set": "  {0} 导入器 SewSurfaces 已设为 True",
        "imp.log.sew_set_fail": "  设置 {0} 导入器 SewSurfaces 失败：{1}",
        "imp.log.sew_unavailable": "  {0} 导入器没有 SewSurfaces 属性，已跳过",
        "imp.log.temp_clean": "  已清理本次产生的 {0} 个转换器日志：{1}",
        "imp.log.temp_clean_fail": "  有 {0} 个本次产生的转换器日志未能清理：{1}",
        "imp.log.temp_clean_old": "  已清理 {1} 天前的 {0} 个转换器日志",
        "imp.log.summary": "汇总：成功 {0} / 失败 {1} / 跳过 {2}",
        "imp.log.background": "已在后台启动 {0} 个任务，NX 可继续操作。",
        "imp.log.background_fail": "后台会话不可用（{0}），改为前台执行。",
        "imp.log.log_fallback": "插件目录不可写，日志改写到：{0}",
        "imp.log.no_confirm": "无法弹出确认框，按继续处理",
        "imp.log.open_mode": "  打开方式：{0}",
        "imp.log.keep_open": "  工作/显示切换方式：{0}",
        "imp.log.parts_check": "  会话核对：应有 {0} 个部件，当前仍在会话中 {1} 个",
        "imp.log.parts_restored": "  已重新打开被显示切换关闭的部件 {0} 个",
        "imp.log.parts_missing": "  以下部件已不在会话中且无法重新打开：{0}",
        "imp.log.restore_fail": "  重新打开失败（{0}）：{1}",
        "imp.reason.open": "无法打开部件：{0}",
        "imp.reason.import": "导入失败：{0}",
        "imp.reason.no_part": "未能创建新部件：{0}",
        "imp.reason.work_part": "无法把新部件设为工作部件：{0}",
        "imp.reason.ext": "不支持的文件类型",
        "imp.msg.cancelled": "已取消，未导入任何文件。",
        "imp.msg.no_files": "没有可导入的文件。",
    },
    "zh_tw": {
        "imp.title": "選擇要匯入的檔案",
        "imp.filter": "支援的檔案",
        "imp.filter_all": "所有檔案",
        "imp.ui.title": "選擇要匯入的檔案",
        "imp.ui.tree_label": "資料夾",
        "imp.ui.list_label": "檔案（Ctrl/Shift 多選，雙擊加入）",
        "imp.ui.col_name": "檔案名稱",
        "imp.ui.col_size": "大小(KB)",
        "imp.ui.this_pc": "本機",
        "imp.ui.denied": "無法存取此目錄（權限不足或路徑無效）：",
        "imp.ui.denied_hint": "資料夾不存在或沒有存取權限。",
        "imp.ui.empty_dir": "（此目錄下沒有可匯入的檔案）",
        "imp.ui.grid_label": "已選檔案清單（共 {0} 個）",
        "imp.ui.col_type": "類型",
        "imp.ui.col_path": "路徑",
        "imp.ui.btn_add_file": "新增匯入檔案",
        "imp.ui.btn_add_folder": "＋ 新增資料夾",
        "imp.ui.name_prefix": "檔案新增前綴",
        "imp.ui.name_suffix": "檔案新增後綴",
        "imp.ui.btn_clear": "清空",
        "imp.ui.btn_del_rows": "刪除選取",
        "imp.ui.menu_delete": "刪除選取",
        "imp.ui.menu_select_all": "全選",
        "imp.ui.menu_clear": "清空",
        "imp.ui.btn_cancel": "取消",
        "imp.ui.btn_ok": "確定",
        "imp.ui.mode_open": "匯入後開啟",
        "imp.ui.mode_files": "僅產生檔案",
        "imp.ui.mode_group": "匯入方式",
        "imp.ui.note_line": "（.prt 僅開啟；前後綴只對另存生效）",
        "imp.ui.on_exists_group": "同名檔案已存在時",
        "imp.ui.on_exists_rename": "改名（_2、_3…）",
        "imp.ui.on_exists_skip": "略過",
        "imp.ui.on_exists_overwrite": "覆蓋",
        "imp.ui.out_group": "輸出目錄",
        "imp.ui.out_src": "輸出到來源檔案目錄",
        "imp.ui.out_dir": "輸出到指定目錄",
        "imp.ui.out_browse": "瀏覽…",
        "imp.ui.out_invalid": "指定輸出目錄無效，已回落到來源檔案目錄。",
        "imp.ui.skip_exists": "（略過）",
        "imp.ui.overwrite_exists": "（覆蓋）",
        "imp.ui.right_title": "轉換後名稱",
        "imp.ui.col_no": "序號",
        "imp.ui.col_file": "檔案名稱",
        "imp.ui.left_empty": "（雙擊上方檔案加入）",
        "imp.ui.right_empty": "（新增檔案後顯示）",
        "imp.ui.open_only": "（僅開啟）",
        "imp.ui.save_rename": "（改名）",
        "imp.ui.workers_label": "同時執行數",
        "imp.ui.msg_remove_title": "移除檔案",
        "imp.ui.msg_remove_text": "從已選清單中移除這個檔案？",
        "imp.ui.msg_empty_title": "提示",
        "imp.ui.msg_empty_one": "尚未新增任何檔案。",
        "imp.ui.msg_empty_two": "請先在左側選擇檔案，再開始匯入。",
        "imp.err.dialog": "無法開啟系統檔案選擇框：{0}",
        "imp.err.background": "後台 NX 工作階段未能啟動，本次匯入沒有執行。請重試或查看日誌。",
        "imp.err.picker_hidden": "選擇窗程序已啟動但視窗不可見（pid {0}），疑似被 SW_HIDE 啟動；隱藏主控台只能靠 CREATE_NO_WINDOW。",
        "imp.err.locked": "目標檔案被占用或無權限：{0}",
        "imp.err.src_readonly": "原檔案所在目錄不可寫",
        "imp.err.out_readonly": "指定的輸出目錄不可寫",
        "imp.err.readonly_title": "目錄不可寫",
        "imp.err.readonly_body": "以下目錄不可寫，相關檔案將直接判為失敗：\n\n{0}",
        "imp.err.not_displayed": "檔案未被顯示（可能是壞檔案或格式不支援）",
        "imp.err.no_geometry": "翻譯完成但沒有實體/片體，判為失敗",
        "imp.log.session_ok": "  已顯示 {0}（{1:.2f}s）",
        "imp.log.session_skip": "  略過（目標已存在）：{0}",
        "imp.log.session_already": "  已在會話中，切換為活動顯示：{0}",
        "imp.log.cancelled": "已取消：剩餘 {0} 個檔案未處理",
        "imp.sum.interactive": "已在目前會話開啟 {0} 個檔案，總耗時 {1:.1f}s。",
        "imp.sum.cancelled_note": "已取消，另有 {0} 個檔案未處理。",
        "imp.prog.title": "正在匯入到目前會話",
        "imp.prog.cancel": "取消",
        "imp.prog.cancelling": "正在取消…",
        "imp.prog.stage_convert": "正在翻譯並開啟…",
        "imp.prog.stage_save": "正在另存…",
        "imp.prog.stage_done": "完成",
        "imp.prog.stage_skip": "已略過（目標已存在）",
        "imp.prog.stage_fail": "失敗",
        "imp.confirm.title": "批次匯入",
        "imp.confirm.body": "共 {0} 個檔案待匯入：\n\n{1}",
        "imp.confirm.more": "\n…… 其餘 {0} 個見日誌",
        "imp.confirm.proceed_log": "確認後每個檔案各自開啟為獨立部件；.prt 僅開啟、不另存，STEP/IGES/Parasolid 另存為源檔案同目錄下的 .prt（重名加 _2/_3）。",
        "imp.sum.title": "批次匯入完成",
        "imp.sum.body": "成功 {0} / 失敗 {1} / 略過 {2}\n\n{3}\n\n日誌：{4}",
        "imp.sum.fail_item": "· {0}：{1}",
        "imp.sum.no_fail": "（沒有失敗項）",
        "imp.sum.more_fail": "…… 另有 {0} 個失敗項，見日誌",
        "imp.sum.skipped": "另有 {0} 個檔案被略過（重複選擇或檔案不存在），詳見日誌。",
        "imp.sum.disk": "已儲存檔案 {0} / 僅記憶體 {1}",
        "imp.sum.temp_logs": "失敗項的轉換器日誌保留在：{0}",
        "imp.log.header": "開始批次匯入：共 {0} 個檔案；日誌 {1}",
        "imp.log.file": "[{0}/{1}] {2}",
        "imp.log.skip_dup": "  略過（重複選擇）：{0}",
        "imp.log.skip_missing": "  略過（檔案不存在）：{0}",
        "imp.log.skip_same": "  略過（目標已存在，保留原檔案）：{0}",
        "imp.log.prt_open": "  開啟 .prt：{0}",
        "imp.log.prt_already": "  該部件已在會話中開啟，切換為活動顯示",
        "imp.log.prt_done": "  完成 {0:.1f}s -> 部件 {1}",
        "imp.log.new_part": "  新建顯示部件：{0}",
        "imp.log.import_begin": "  匯入 {0} -> 部件 {1}",
        "imp.log.import_done": "  完成 {0:.1f}s",
        "imp.log.saveas_start": "  開始另存：{0}",
        "imp.log.saveas_ok": "  另存成功：{0}",
        "imp.log.saveas_rename": "  同名檔案已存在，改名為：{0}",
        "imp.log.backup": "  原檔案已備份為：{0}",
        "imp.err.backup": "  無法備份已存在的檔案 {0} → {1}（{2}）；已中止覆蓋，原檔案保持不動",
        "imp.err.saveas": "  儲存失敗（{0}）：{1}；幾何保留在記憶體，未關閉部件",
        "imp.err.saveas_missing": "  儲存失敗：目標檔案未產生 {0}",
        "imp.log.load_status": "  部件載入狀態：{0}",
        "imp.log.fail": "  失敗：{0}",
        "imp.log.fit_fail": "  擬合視圖失敗（忽略）：{0}",
        "imp.log.destroy_fail": "  銷毀匯入器失敗（忽略）：{0}",
        "imp.log.work_part": "  工作部件：{0}",
        "imp.log.close_empty": "  已關閉空部件：{0}",
        "imp.log.close_fail": "  空部件關閉失敗（保留於會話中）：{0}",
        "imp.log.import_to": "  ImportTo 設定失敗（依預設工作部件繼續）：{0}",
        "imp.log.prep": "  已補齊匯入器參數：{0}；OutputFile={1}",
        "imp.log.prep_fail": "  匯入器參數補齊失敗（依預設繼續）：{0}",
        "imp.log.settings_missing": "  未找到 STEP214 設定檔，依匯入器預設值處理",
        "imp.log.sew_probe": "  {0} 匯入器目前 SewSurfaces 預設值 = {1}",
        "imp.log.sew_probe_fail": "  讀取 {0} 匯入器 SewSurfaces 失敗：{1}",
        "imp.log.sew_set": "  {0} 匯入器 SewSurfaces 已設為 True",
        "imp.log.sew_set_fail": "  設定 {0} 匯入器 SewSurfaces 失敗：{1}",
        "imp.log.sew_unavailable": "  {0} 匯入器沒有 SewSurfaces 屬性，已略過",
        "imp.log.temp_clean": "  已清理本次產生的 {0} 個轉換器日誌：{1}",
        "imp.log.temp_clean_fail": "  有 {0} 個本次產生的轉換器日誌未能清理：{1}",
        "imp.log.temp_clean_old": "  已清理 {1} 天前的 {0} 個轉換器日誌",
        "imp.log.summary": "彙總：成功 {0} / 失敗 {1} / 略過 {2}",
        "imp.log.background": "已在背景啟動 {0} 個任務，NX 可繼續操作。",
        "imp.log.background_fail": "背景工作階段不可用（{0}），改為前景執行。",
        "imp.log.log_fallback": "外掛目錄不可寫，日誌改寫到：{0}",
        "imp.log.no_confirm": "無法彈出確認框，按繼續處理",
        "imp.log.open_mode": "  開啟方式：{0}",
        "imp.log.keep_open": "  工作/顯示切換方式：{0}",
        "imp.log.parts_check": "  會話核對：應有 {0} 個部件，目前仍在會話中 {1} 個",
        "imp.log.parts_restored": "  已重新開啟被顯示切換關閉的部件 {0} 個",
        "imp.log.parts_missing": "  以下部件已不在會話中且無法重新開啟：{0}",
        "imp.log.restore_fail": "  重新開啟失敗（{0}）：{1}",
        "imp.reason.open": "無法開啟部件：{0}",
        "imp.reason.import": "匯入失敗：{0}",
        "imp.reason.no_part": "未能建立新部件：{0}",
        "imp.reason.work_part": "無法把新部件設為工作部件：{0}",
        "imp.reason.ext": "不支援的檔案類型",
        "imp.msg.cancelled": "已取消，未匯入任何檔案。",
        "imp.msg.no_files": "沒有可匯入的檔案。",
    },
}


def imp_T(key, *args):
    """取当前语言文案；该语言缺项时回落英文。"""
    table = imp_MSG.get(imp_LANG) or {}
    text = table.get(key)
    if text is None:
        text = imp_MSG["en"].get(key, key)
    if args:
        return text.format(*args)
    return text


# 非 Unicode 控制台（英文 cp437、繁中 cp950 等）打印中文会抛 UnicodeEncodeError。
try:
    import sys
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# ================= 常量 =================
SUPPORTED_PRT = (".prt",)
SUPPORTED_STEP = (".stp", ".step")
SUPPORTED_IGES = (".igs", ".iges")
SUPPORTED_PARASOLID = (".x_t", ".x_b")
ALL_EXTENSIONS = SUPPORTED_PRT + SUPPORTED_STEP + SUPPORTED_IGES + SUPPORTED_PARASOLID
FILTER_PATTERN = ";".join("*" + ext for ext in ALL_EXTENSIONS)

MAX_NAME_LEN = 80
LAST_DIR_FILE = "import_last_dir.txt"

# 自动缝合开关：默认关，保持与改动前完全一致的导入结果。
# 只在 True 时才写 importer.SewSurfaces；False 时仅打印一次转换器里的当前默认值。
SEW_SURFACES = False
# 会话启动时清理 translator 临时日志的保留天数（只清 .log，不动 .prt 中间产物）
TEMP_LOG_DAYS = 7

imp__log_handle = None
_trace_handle = None
_sew_probed = set()          # 已探查过 SewSurfaces 默认值的转换器类型（每种只记一次日志）


# ================= 日志 =================

def imp_plugin_dir():
    r"""插件根目录（脚本在 startup\ 下，取它的上一层；被 exec 执行时回落当前目录）"""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()
    if os.path.basename(here).lower() == "startup":
        parent = os.path.dirname(here)
        if parent:
            return parent
    return here


def imp_log(msg):
    """同时打印到 NX 信息窗口/控制台和日志文件"""
    text = str(msg)
    try:
        print(text)
    except Exception:
        pass
    if imp__log_handle:
        try:
            imp__log_handle.write(text + "\n")
            imp__log_handle.flush()
        except Exception:
            pass


def trace(step, extra=""):
    """行级打点：绝对时间戳 + 立即落盘（fsync）。

    用于定位“NX 弹工作进行中但无进展”这类卡死：日志停在哪个 step，
    就说明卡在下一步调用上。无 trace 文件时静默跳过。
    """
    if not _trace_handle:
        return
    try:
        now = time.time()
        stamp = "%s.%03d" % (time.strftime("%H:%M:%S", time.localtime(now)),
                             int(now * 1000) % 1000)
        _trace_handle.write("%s | %s | %s\n" % (stamp, step, extra))
        _trace_handle.flush()
        os.fsync(_trace_handle.fileno())
    except Exception:
        pass


def _force_remove(path):
    try:
        os.remove(path)
    except PermissionError:
        try:
            os.chmod(path, 0o666)
            os.remove(path)
        except Exception:
            pass
    except Exception:
        pass


def imp__rotate_log(path, max_bytes=2 * 1024 * 1024, keep=3):
    """日志超过上限时轮转：当前 -> .1，旧的依次后移，最多保留 keep 份"""
    try:
        if (not os.path.exists(path)) or os.path.getsize(path) < max_bytes:
            return
        oldest = "%s.%d" % (path, keep)
        if os.path.exists(oldest):
            _force_remove(oldest)
        for index in range(keep - 1, 0, -1):
            src = "%s.%d" % (path, index)
            if os.path.exists(src):
                os.replace(src, "%s.%d" % (path, index + 1))
        os.replace(path, path + ".1")
    except Exception:
        pass


def open_log():
    r"""日志放到插件目录 logs\；插件目录不可写时回落 %TEMP%"""
    imp__log_handle, _trace_handle
    try:
        log_dir = os.path.join(imp_plugin_dir(), "logs")
        if not os.path.isdir(log_dir):
            os.makedirs(log_dir)
        log_path = os.path.join(log_dir, "batch_import.log")
        imp__rotate_log(log_path)
        imp__log_handle = open(log_path, "w", encoding="utf-8-sig")
        try:
            _trace_handle = open(os.path.join(log_dir, "batch_import_trace.log"),
                                 "a", encoding="utf-8")
            _trace_handle.write("\n===== run %s =====\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
            _trace_handle.flush()
        except Exception:
            _trace_handle = None
        return log_path
    except Exception:
        import tempfile
        fallback = os.path.join(tempfile.gettempdir(),
                                "nx_batch_import_%d.log" % os.getpid())
        imp__log_handle = open(fallback, "w", encoding="utf-8-sig")
        try:
            _trace_handle = open(os.path.join(tempfile.gettempdir(),
                                              "nx_batch_import_trace_%d.log" % os.getpid()),
                                 "a", encoding="utf-8")
        except Exception:
            _trace_handle = None
        return fallback


def imp__last_dir_path():
    return os.path.join(imp_plugin_dir(), "logs", LAST_DIR_FILE)


def _read_prefs():
    """读上次目录、处理方式、并发数、前后缀、同名处理与输出位置。

    第 1 行目录，其后为 key=value：open=1/0、workers=N、prefix=…、suffix=…、
    collision=rename/skip/overwrite、out_mode=source/dir、outdir=…
    旧文件只有 2~5 行也能读，缺的键按默认值兜底。
    """
    try:
        with open(imp__last_dir_path(), "r", encoding="utf-8") as handle:
            lines = [item.strip() for item in handle.read().splitlines()]
    except Exception:
        return "", True, 2, "", "", "skip", "source", ""
    folder = lines[0] if lines and os.path.isdir(lines[0]) else ""
    opened = True
    workers = 2
    prefix = ""
    suffix = ""
    collision = "skip"
    out_mode = "source"
    outdir = ""
    for line in lines[1:]:
        key, _, value = line.partition("=")
        key = key.strip().lower()
        if key == "open":
            opened = value.strip() != "0"
        elif key == "workers":
            try:
                workers = max(1, min(6, int(value.strip())))
            except Exception:
                workers = 2
        elif key == "prefix":
            prefix = value
        elif key == "suffix":
            suffix = value
        elif key == "collision":
            value = value.strip().lower()
            if value in ("rename", "skip", "overwrite"):
                collision = value
        elif key == "out_mode":
            out_mode = "dir" if value.strip().lower() == "dir" else "source"
        elif key == "outdir":
            outdir = value.strip()
    return folder, opened, workers, prefix, suffix, collision, out_mode, outdir


def _write_prefs(folder=None, opened=None, workers=None, prefix=None, suffix=None,
                 collision=None, out_mode=None, outdir=None):
    try:
        old = _read_prefs()
        if folder is None:
            folder = old[0]
        if opened is None:
            opened = old[1]
        if workers is None:
            workers = old[2]
        if prefix is None:
            prefix = old[3]
        if suffix is None:
            suffix = old[4]
        if collision is None:
            collision = old[5]
        if out_mode is None:
            out_mode = old[6]
        if outdir is None:
            outdir = old[7]
        path = imp__last_dir_path()
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        # 前后缀/目录按单行保存，含换行会把记忆文件写坏
        prefix = str(prefix).replace("\r", " ").replace("\n", " ")
        suffix = str(suffix).replace("\r", " ").replace("\n", " ")
        outdir = str(outdir).replace("\r", " ").replace("\n", " ")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("%s\nopen=%d\nworkers=%d\nprefix=%s\nsuffix=%s\n"
                         "collision=%s\nout_mode=%s\noutdir=%s"
                         % (folder, 1 if opened else 0, workers, prefix, suffix,
                            collision, "dir" if out_mode == "dir" else "source",
                            outdir))
    except Exception:
        pass


def imp_read_last_dir():
    return _read_prefs()[0]


def imp_write_last_dir(folder):
    _write_prefs(folder=folder)


def read_open_after():
    return _read_prefs()[1]


def write_open_after(value):
    _write_prefs(opened=bool(value))


def read_workers():
    return _read_prefs()[2]


def write_workers(value):
    try:
        value = max(1, min(6, int(value)))
    except Exception:
        value = 2
    _write_prefs(workers=value)


def read_affixes():
    """(前缀, 后缀)：只影响另存出来的 .prt 文件名"""
    prefs = _read_prefs()
    return prefs[3], prefs[4]


def write_affixes(prefix, suffix):
    _write_prefs(prefix=str(prefix or ""), suffix=str(suffix or ""))


def read_collision():
    """同名处理：rename / skip / overwrite；默认与导出窗一致（skip）"""
    return _read_prefs()[5]


def write_collision(value):
    value = str(value or "").strip().lower()
    if value not in ("rename", "skip", "overwrite"):
        value = "skip"
    _write_prefs(collision=value)


def read_output():
    """(输出模式, 指定目录)：source / dir"""
    prefs = _read_prefs()
    return prefs[6], prefs[7]


def write_output(out_mode, outdir):
    _write_prefs(out_mode=("dir" if out_mode == "dir" else "source"),
                 outdir=str(outdir or ""))


# ================= Windows 系统多选对话框 =================
# 不依赖 tkinter：NX 内嵌 Python 未必带。直接调 comdlg32 的 GetOpenFileNameW。

OFN_READONLY = 0x00000001
OFN_OVERWRITEPROMPT = 0x00000002
OFN_HIDEREADONLY = 0x00000004
OFN_NOCHANGEDIR = 0x00000008
OFN_PATHMUSTEXIST = 0x00000800
OFN_FILEMUSTEXIST = 0x00001000
OFN_ALLOWMULTISELECT = 0x00000200
OFN_EXPLORER = 0x00080000

_FILE_BUFFER_CHARS = 32768


class _OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", ctypes.c_uint32),
        ("hwndOwner", ctypes.c_void_p),
        ("hInstance", ctypes.c_void_p),
        ("lpstrFilter", ctypes.c_wchar_p),
        ("lpstrCustomFilter", ctypes.c_wchar_p),
        ("nMaxCustFilter", ctypes.c_uint32),
        ("nFilterIndex", ctypes.c_uint32),
        ("lpstrFile", ctypes.c_wchar_p),
        ("nMaxFile", ctypes.c_uint32),
        ("lpstrFileTitle", ctypes.c_wchar_p),
        ("nMaxFileTitle", ctypes.c_uint32),
        ("lpstrInitialDir", ctypes.c_wchar_p),
        ("lpstrTitle", ctypes.c_wchar_p),
        ("Flags", ctypes.c_uint32),
        ("nFileOffset", ctypes.c_uint16),
        ("nFileExtension", ctypes.c_uint16),
        ("lpstrDefExt", ctypes.c_wchar_p),
        ("lCustData", ctypes.c_void_p),
        ("lpfnHook", ctypes.c_void_p),
        ("lpTemplateName", ctypes.c_wchar_p),
        ("pvReserved", ctypes.c_void_p),
        ("dwReserved", ctypes.c_uint32),
        ("FlagsEx", ctypes.c_uint32),
    ]


def _file_filter():
    return "%s (*.prt;*.stp;*.step;*.igs;*.iges;*.x_t;*.x_b)\0%s\0%s (*.*)\0*.*\0\0" % (
        imp_T("imp.filter"), FILTER_PATTERN, imp_T("imp.filter_all"))


# 导入选择窗：布局、控件、交互与导出侧的 SELECT_UI_PS 保持一致
# （左侧目录树 / 右侧文件列表 / 下方已选清单 / 右下 取消+确认），
# 差别只有：只收文件、清单上方带数量、下方说明改成导入行为说明。
IMPORT_UI_PS = r'''
param(
    [string]$Root = "",
    [int]$Workers = 2,
    [string]$Prefix = "",
    [string]$Suffix = "",
    [string]$Collision = "skip",
    [string]$OutMode = "source",
    [string]$OutDir = "",
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 弹出即抢前台（AttachThreadInput + BringWindowToTop + 瞬时 TopMost），片段在公共执行器里
[[front_helpers]]

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[imp.ui.title]]"
$form.ClientSize = New-Object System.Drawing.Size(944, 726)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.TopMost = $false
$uiFont = [System.Drawing.SystemFonts]::MessageBoxFont
try { $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9) } catch { }
$form.Font = $uiFont
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)

# ---- NX 2506 观感基准（实测值与来源见交付参数表）----
# 内容背景 #FFFFFF、分隔线 #DCDCDC、正文 #1E1E1E 取自 NX 2506 界面截图采样；
# 深色主题无本机 NX 截图可量，用 Windows 深色系统色近似；
# 环境变量 NX_PICKER_THEME=light|dark 可强制切换（仅用于渲染核对，功能无关）。
$script:isDark = $false
try {
    $themeKey = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction Stop
    if ($themeKey.AppsUseLightTheme -eq 0) { $script:isDark = $true }
} catch { }
if ($env:NX_PICKER_THEME -eq 'dark') { $script:isDark = $true }
elseif ($env:NX_PICKER_THEME -eq 'light') { $script:isDark = $false }
if ($script:isDark) {
    $uiBg = [System.Drawing.Color]::FromArgb(32, 32, 32)
    $uiField = [System.Drawing.Color]::FromArgb(24, 24, 24)
    $uiFg = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $uiLine = [System.Drawing.Color]::FromArgb(63, 63, 70)
    $uiBtn = [System.Drawing.Color]::FromArgb(51, 51, 51)
} else {
    $uiBg = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $uiField = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $uiFg = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $uiLine = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $uiBtn = [System.Drawing.Color]::FromArgb(240, 240, 240)
}
$form.BackColor = $uiBg
$form.ForeColor = $uiFg

function Add-Separator([int]$x, [int]$y, [int]$w) {
    $line = New-Object System.Windows.Forms.Panel
    $line.Location = New-Object System.Drawing.Point($x, $y)
    $line.Size = New-Object System.Drawing.Size($w, 1)
    $line.BackColor = $uiLine
    $form.Controls.Add($line)
}

function Apply-Theme([System.Windows.Forms.Control]$root) {
    foreach ($c in $root.Controls) {
        if ($c -is [System.Windows.Forms.GroupBox]) {
            $c.ForeColor = $uiFg
            $c.BackColor = $uiBg
        } elseif ($c -is [System.Windows.Forms.Label] -or $c -is [System.Windows.Forms.RadioButton]) {
            $c.ForeColor = $uiFg
            $c.BackColor = [System.Drawing.Color]::Transparent
        } elseif ($c -is [System.Windows.Forms.ComboBox]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
            # 深色下标准 Combobox 会自绘白底，改 Flat 才吃 BackColor
            if ($script:isDark) { $c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
        } elseif ($c -is [System.Windows.Forms.TextBox]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
        } elseif ($c -is [System.Windows.Forms.NumericUpDown]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
        } elseif ($c -is [System.Windows.Forms.TreeView] -or $c -is [System.Windows.Forms.ListView]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
            # 列表网格线是系统色，深色下会变成刺眼白线，直接关掉
            if ($script:isDark -and $c -is [System.Windows.Forms.ListView]) { $c.GridLines = $false }
        } elseif ($c -is [System.Windows.Forms.Button]) {
            $c.BackColor = $uiBtn
            $c.ForeColor = $uiFg
            $c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $c.FlatAppearance.BorderColor = $uiLine
        }
        if ($c.HasChildren) { Apply-Theme $c }
    }
}

$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = "[[imp.ui.mode_group]]"
$grpMode.Location = New-Object System.Drawing.Point(8, 6)
$grpMode.Size = New-Object System.Drawing.Size(300, 100)
$form.Controls.Add($grpMode)

# 转换一律交给后台会话完成，结果不再自动显示；这里只保留唯一的处理方式说明
$rbFiles = New-Object System.Windows.Forms.RadioButton
$rbFiles.Text = "[[imp.ui.mode_files]]"
$rbFiles.Location = New-Object System.Drawing.Point(14, 30)
$rbFiles.Size = New-Object System.Drawing.Size(270, 22)
$rbFiles.Checked = $true
$grpMode.Controls.Add($rbFiles)

# ---- 同名处理：与导出窗同位同样式；默认跳过 ----
$grpExists = New-Object System.Windows.Forms.GroupBox
$grpExists.Text = "[[imp.ui.on_exists_group]]"
$grpExists.Location = New-Object System.Drawing.Point(314, 6)
$grpExists.Size = New-Object System.Drawing.Size(200, 100)
$form.Controls.Add($grpExists)

$rbRename = New-Object System.Windows.Forms.RadioButton
$rbRename.Text = "[[imp.ui.on_exists_rename]]"
$rbRename.Location = New-Object System.Drawing.Point(14, 18)
$rbRename.Size = New-Object System.Drawing.Size(180, 22)
$grpExists.Controls.Add($rbRename)

$rbSkip = New-Object System.Windows.Forms.RadioButton
$rbSkip.Text = "[[imp.ui.on_exists_skip]]"
$rbSkip.Location = New-Object System.Drawing.Point(14, 42)
$rbSkip.Size = New-Object System.Drawing.Size(180, 22)
$grpExists.Controls.Add($rbSkip)

$rbOverwrite = New-Object System.Windows.Forms.RadioButton
$rbOverwrite.Text = "[[imp.ui.on_exists_overwrite]]"
$rbOverwrite.Location = New-Object System.Drawing.Point(14, 66)
$rbOverwrite.Size = New-Object System.Drawing.Size(180, 22)
$grpExists.Controls.Add($rbOverwrite)

try { $col = [string]$Collision } catch { $col = "skip" }
if ($col -eq "rename") { $rbRename.Checked = $true }
elseif ($col -eq "overwrite") { $rbOverwrite.Checked = $true }
else { $rbSkip.Checked = $true }

# ---- 输出位置：源目录 / 指定目录 ----
$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = "[[imp.ui.out_group]]"
$grpOut.Location = New-Object System.Drawing.Point(520, 6)
$grpOut.Size = New-Object System.Drawing.Size(416, 100)
$form.Controls.Add($grpOut)

# 两个单选项并排一行两列（左：源文件目录；右：指定目录）；路径行仍在下方
$rbOutSrc = New-Object System.Windows.Forms.RadioButton
$rbOutSrc.Text = "[[imp.ui.out_src]]"
$rbOutSrc.Location = New-Object System.Drawing.Point(14, 22)
$rbOutSrc.Size = New-Object System.Drawing.Size(176, 22)
$rbOutSrc.TabIndex = 0
$grpOut.Controls.Add($rbOutSrc)

$rbOutDir = New-Object System.Windows.Forms.RadioButton
$rbOutDir.Text = "[[imp.ui.out_dir]]"
$rbOutDir.Location = New-Object System.Drawing.Point(198, 22)
$rbOutDir.Size = New-Object System.Drawing.Size(204, 22)
$rbOutDir.TabIndex = 1
$grpOut.Controls.Add($rbOutDir)

$txtOutDir = New-Object System.Windows.Forms.TextBox
$txtOutDir.Location = New-Object System.Drawing.Point(14, 50)
$txtOutDir.Size = New-Object System.Drawing.Size(300, 24)
$txtOutDir.Text = $OutDir
$txtOutDir.TabIndex = 2
$grpOut.Controls.Add($txtOutDir)

$btnOutBrowse = New-Object System.Windows.Forms.Button
$btnOutBrowse.Text = "[[imp.ui.out_browse]]"
$btnOutBrowse.Location = New-Object System.Drawing.Point(320, 49)
$btnOutBrowse.Size = New-Object System.Drawing.Size(86, 26)
$btnOutBrowse.TabIndex = 3
$grpOut.Controls.Add($btnOutBrowse)

if ([string]$OutMode -eq "dir") { $rbOutDir.Checked = $true } else { $rbOutSrc.Checked = $true }

# 支持格式 / .prt 直接打开 / 前后缀只作用于另存文件：压缩成一行挂在前后缀右侧
$lblNoteLine = New-Object System.Windows.Forms.Label
$lblNoteLine.Text = "[[imp.ui.note_line]]"
$lblNoteLine.Location = New-Object System.Drawing.Point(484, 108)
$lblNoteLine.Size = New-Object System.Drawing.Size(452, 20)
$lblNoteLine.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblNoteLine)

# ---- 输出文件名前后缀：只作用于会落盘的 .prt（stp/igs/x_t → prt），.prt 源不受影响 ----
$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text = "[[imp.ui.name_prefix]]"
$lblPrefix.Location = New-Object System.Drawing.Point(8, 108)
$lblPrefix.Size = New-Object System.Drawing.Size(84, 20)
$form.Controls.Add($lblPrefix)

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Location = New-Object System.Drawing.Point(96, 106)
$txtPrefix.Size = New-Object System.Drawing.Size(140, 24)
$txtPrefix.Text = $Prefix
$form.Controls.Add($txtPrefix)

$lblSuffix = New-Object System.Windows.Forms.Label
$lblSuffix.Text = "[[imp.ui.name_suffix]]"
$lblSuffix.Location = New-Object System.Drawing.Point(248, 108)
$lblSuffix.Size = New-Object System.Drawing.Size(84, 20)
$form.Controls.Add($lblSuffix)

$txtSuffix = New-Object System.Windows.Forms.TextBox
$txtSuffix.Location = New-Object System.Drawing.Point(336, 106)
$txtSuffix.Size = New-Object System.Drawing.Size(140, 24)
$txtSuffix.Text = $Suffix
$form.Controls.Add($txtSuffix)

$lblTree = New-Object System.Windows.Forms.Label
$lblTree.Text = "[[imp.ui.tree_label]]"
$lblTree.Location = New-Object System.Drawing.Point(8, 130)
$lblTree.Size = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($lblTree)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Location = New-Object System.Drawing.Point(8, 148)
$tree.Size = New-Object System.Drawing.Size(340, 178)
$tree.HideSelection = $false
$tree.ItemHeight = 22
# NX 树形控件不用点状连线，展开箭头由系统绘制
$tree.ShowLines = $false
$tree.ShowRootLines = $false
$tree.Indent = 16
$tree.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($tree)

$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = "[[imp.ui.list_label]]"
$lblList.Location = New-Object System.Drawing.Point(364, 130)
$lblList.Size = New-Object System.Drawing.Size(568, 20)
$form.Controls.Add($lblList)

$list = New-Object System.Windows.Forms.ListView
$list.Location = New-Object System.Drawing.Point(364, 148)
$list.Size = New-Object System.Drawing.Size(568, 178)
$list.View = "Details"
$list.FullRowSelect = $true
$list.MultiSelect = $true
$list.GridLines = $true
[void]$list.Columns.Add("[[imp.ui.col_name]]", 440)
[void]$list.Columns.Add("[[imp.ui.col_size]]", 100)
$list.ShowItemToolTips = $true
$form.Controls.Add($list)

# ---- 已选清单（左）与转换后文件名对照（右）----
$btnAddFile = New-Object System.Windows.Forms.Button
$btnAddFile.Text = "[[imp.ui.btn_add_file]]"
$btnAddFile.Location = New-Object System.Drawing.Point(8, 328)
$btnAddFile.Size = New-Object System.Drawing.Size(124, 26)
$form.Controls.Add($btnAddFile)

# 与导出窗同位置同样式：把所选目录下的支持格式一次性加入左栏
$btnAddFolder = New-Object System.Windows.Forms.Button
$btnAddFolder.Text = "[[imp.ui.btn_add_folder]]"
$btnAddFolder.Location = New-Object System.Drawing.Point(136, 328)
$btnAddFolder.Size = New-Object System.Drawing.Size(110, 26)
$form.Controls.Add($btnAddFolder)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "[[imp.ui.btn_clear]]"
$btnClear.Location = New-Object System.Drawing.Point(250, 328)
$btnClear.Size = New-Object System.Drawing.Size(62, 26)
$form.Controls.Add($btnClear)

$lblGrid = New-Object System.Windows.Forms.Label
$lblGrid.Location = New-Object System.Drawing.Point(316, 332)
$lblGrid.Size = New-Object System.Drawing.Size(152, 20)
$form.Controls.Add($lblGrid)

$lblRightTitle = New-Object System.Windows.Forms.Label
$lblRightTitle.Text = "[[imp.ui.right_title]]"
$lblRightTitle.Location = New-Object System.Drawing.Point(492, 332)
$lblRightTitle.Size = New-Object System.Drawing.Size(444, 20)
$form.Controls.Add($lblRightTitle)

$lblArrow = New-Object System.Windows.Forms.Label
$lblArrow.Text = "→"
$lblArrow.Location = New-Object System.Drawing.Point(470, 474)
$lblArrow.Size = New-Object System.Drawing.Size(20, 20)
$form.Controls.Add($lblArrow)

$left = New-Object System.Windows.Forms.ListView
$left.Location = New-Object System.Drawing.Point(8, 358)
$left.Size = New-Object System.Drawing.Size(460, 256)
$left.View = "Details"
$left.FullRowSelect = $true
$left.MultiSelect = $true
$left.GridLines = $true
$left.HideSelection = $false
[void]$left.Columns.Add("[[imp.ui.col_no]]", 46)
[void]$left.Columns.Add("[[imp.ui.col_file]]", 370)
[void]$left.Columns.Add("✕", 36)
$left.ShowItemToolTips = $true
$form.Controls.Add($left)

$right = New-Object System.Windows.Forms.ListView
$right.Location = New-Object System.Drawing.Point(492, 358)
$right.Size = New-Object System.Drawing.Size(444, 256)
$right.View = "Details"
$right.FullRowSelect = $true
$right.MultiSelect = $false
$right.GridLines = $true
[void]$right.Columns.Add("[[imp.ui.col_no]]", 46)
[void]$right.Columns.Add("[[imp.ui.col_file]]", 390)
$right.ShowItemToolTips = $true
$form.Controls.Add($right)

$lblLeftEmpty = New-Object System.Windows.Forms.Label
$lblLeftEmpty.Text = "[[imp.ui.left_empty]]"
$lblLeftEmpty.Location = New-Object System.Drawing.Point(16, 384)
$lblLeftEmpty.Size = New-Object System.Drawing.Size(440, 20)
$lblLeftEmpty.BackColor = $uiField
$lblLeftEmpty.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblLeftEmpty)

$lblRightEmpty = New-Object System.Windows.Forms.Label
$lblRightEmpty.Text = "[[imp.ui.right_empty]]"
$lblRightEmpty.Location = New-Object System.Drawing.Point(500, 384)
$lblRightEmpty.Size = New-Object System.Drawing.Size(430, 20)
$lblRightEmpty.BackColor = $uiField
$lblRightEmpty.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblRightEmpty)

$lblWorkers = New-Object System.Windows.Forms.Label
$lblWorkers.Text = "[[imp.ui.workers_label]]"
$lblWorkers.Location = New-Object System.Drawing.Point(8, 622)
$lblWorkers.Size = New-Object System.Drawing.Size(176, 22)
$form.Controls.Add($lblWorkers)

$numWorkers = New-Object System.Windows.Forms.NumericUpDown
$numWorkers.Location = New-Object System.Drawing.Point(188, 619)
$numWorkers.Size = New-Object System.Drawing.Size(60, 24)
$numWorkers.Minimum = 1
$numWorkers.Maximum = 6
$w = 2
try { $w = [int]$Workers } catch { }
if ($w -lt 1) { $w = 1 }
if ($w -gt 6) { $w = 6 }
$numWorkers.Value = [decimal]$w
$form.Controls.Add($numWorkers)

# 多选删除：本按钮或左栏右键菜单，删完重编号并重算右栏
$btnDelRows = New-Object System.Windows.Forms.Button
$btnDelRows.Text = "[[imp.ui.btn_del_rows]]"
$btnDelRows.Location = New-Object System.Drawing.Point(262, 620)
$btnDelRows.Size = New-Object System.Drawing.Size(148, 26)
$form.Controls.Add($btnDelRows)

# 行高按 NX 实测行距（22px@125% ≈ 17.6 逻辑 + 行间留白）统一为 22：
# ListView 只能借 ImageList 撑高行，图标列宽 1px 不占位
$script:rowHeight = 22
$script:rowImages = New-Object System.Windows.Forms.ImageList
$script:rowImages.ImageSize = New-Object System.Drawing.Size(1, $script:rowHeight)
$script:rowImages.Images.Add((New-Object System.Drawing.Bitmap(1, $script:rowHeight)))
$list.SmallImageList = $script:rowImages
$left.SmallImageList = $script:rowImages
$right.SmallImageList = $script:rowImages

# Get-SafePart / Get-CleanStem 与导出窗共用，定义在 nx_batch_runner.NAME_HELPERS_PS
[[name_helpers]]

# 左栏清单逻辑与导出窗共用，定义在 nx_batch_runner.LEFT_LIST_PS；
# 这里只提供文案与各窗自己的统一刷新入口 Refresh-List
$script:txtDelRows = "[[imp.ui.btn_del_rows]]"
$script:txtMenuDelete = "[[imp.ui.menu_delete]]"
$script:txtMenuSelectAll = "[[imp.ui.menu_select_all]]"
$script:txtMenuClear = "[[imp.ui.menu_clear]]"
[[left_list]]

function Update-Count {
    $lblGrid.Text = "[[imp.ui.grid_label]]".Replace("{0}", [string]$left.Items.Count)
}

# 支持的导入格式（左侧列表、「添加文件夹」、系统多选框三处共用）
function Test-PickFile([string]$path) {
    $l = $path.ToLower()
    return ($l.EndsWith(".prt") -or $l.EndsWith(".stp") -or $l.EndsWith(".step") -or
            $l.EndsWith(".igs") -or $l.EndsWith(".iges") -or $l.EndsWith(".x_t") -or
            $l.EndsWith(".x_b"))
}

# 结果名按 plan_target 的规则复算：目录 = 指定输出目录（有效时）或源目录；
# 名字 = 前缀 + 清洗后 stem + 后缀 + .prt；同名处理 跳过（灰）/ 覆盖 / 改名（_2/_3）；
# .prt 源只打开、不参与改名与输出目录
function Get-OutFolder([string]$src) {
    if ($rbOutDir.Checked) {
        $d = ([string]$txtOutDir.Text).Trim()
        if ($d -and [System.IO.Directory]::Exists($d)) { return $d }
    }
    return [System.IO.Path]::GetDirectoryName($src)
}

function Update-Preview {
    $right.BeginUpdate()
    try {
        $right.Items.Clear()
        $pre = Get-SafePart $txtPrefix.Text
        $suf = Get-SafePart $txtSuffix.Text
        $used = @{}
        $no = 0
        foreach ($row in $left.Items) {
            $no++
            $src = [string]$row.Tag
            $srcExt = [System.IO.Path]::GetExtension($src).ToLower()
            if ($srcExt -eq ".prt") {
                Add-RightRow $no "[[imp.ui.open_only]]" $src $true
                continue
            }
            $folder = Get-OutFolder $src
            $base = $pre + (Get-CleanStem ([System.IO.Path]::GetFileNameWithoutExtension($src))) + $suf
            $name = $base + ".prt"
            $full = [System.IO.Path]::Combine($folder, $name)
            $key = $full.ToLower()
            $taken = [System.IO.File]::Exists($full) -or $used.ContainsKey($key)
            if ($rbSkip.Checked) {
                if ($taken) {
                    Add-RightRow $no ($name + "[[imp.ui.skip_exists]]") $full $true
                } else {
                    $used[$key] = $src
                    Add-RightRow $no $name $full $false
                }
                continue
            }
            if ($rbOverwrite.Checked) {
                $state = ""
                if ($taken) { $state = "[[imp.ui.overwrite_exists]]" }
                $used[$key] = $src
                Add-RightRow $no ($name + $state) $full $false
                continue
            }
            $serial = 2
            $renamed = $false
            while ([System.IO.File]::Exists([System.IO.Path]::Combine($folder, $name)) -or $used.ContainsKey(([System.IO.Path]::Combine($folder, $name)).ToLower())) {
                $name = $base + "_" + $serial + ".prt"
                $serial++
                $renamed = $true
            }
            $full = [System.IO.Path]::Combine($folder, $name)
            $used[$full.ToLower()] = $src
            $state = ""
            if ($renamed) { $state = "[[imp.ui.save_rename]]" }
            Add-RightRow $no ($name + $state) $full $renamed
        }
    } finally {
        $right.EndUpdate()
        Update-EmptyHints
    }
}

# 各窗自己的统一刷新入口：清单变动后走这里（共享片段内部只认这个名字）
function Refresh-List {
    Renumber-Left
    Update-Count
    Update-Preview
    Update-EmptyHints
}

Update-Count
Update-EmptyHints

$script:PH = "__placeholder__"
# 同一目录只枚举一次：重复点击/折叠再展开不再走磁盘
$script:dirCache = @{}

function New-DirNode([System.Windows.Forms.TreeNode]$parent, [string]$path) {
    # 只加一层：子目录等到展开时才枚举，避免一上来递归整块硬盘
    $node = New-Object System.Windows.Forms.TreeNode([System.IO.Path]::GetFileName($path.TrimEnd('\')))
    if (-not $node.Text) { $node.Text = $path }
    $node.Tag = $path
    [void]$parent.Nodes.Add($node)
    $hasChild = $false
    try {
        foreach ($d in [System.IO.Directory]::GetDirectories($path)) { $hasChild = $true; break }
    } catch { }
    if ($hasChild) {
        $ph = New-Object System.Windows.Forms.TreeNode("...")
        $ph.Tag = $script:PH
        [void]$node.Nodes.Add($ph)
    }
    return $node
}

# 从「此电脑 → 盘符」沿路径逐级建节点并展开，返回目标目录节点（用于「上次位置」记忆）
function Expand-ToPath([string]$path) {
    if (-not $path) { return $null }
    $full = $path.TrimEnd('\')
    if (-not $full) { return $null }
    try {
        if (-not [System.IO.Directory]::Exists($full)) { return $null }
    } catch { return $null }
    $rootPart = [System.IO.Path]::GetPathRoot($full)
    if (-not $rootPart) { return $null }
    $node = $null
    # 盘符节点挂在「此电脑」下面，先查根节点本身，再查它的一级子节点
    foreach ($n in $tree.Nodes) {
        if (([string]$n.Tag).TrimEnd('\') -eq $rootPart.TrimEnd('\')) { $node = $n; break }
        foreach ($c in $n.Nodes) {
            if (([string]$c.Tag).TrimEnd('\') -eq $rootPart.TrimEnd('\')) { $node = $c; break }
        }
        if ($node) { break }
    }
    if (-not $node) { return $null }
    if ($node.Parent) { $node.Parent.Expand() }
    [void]$node.Expand()
    $rest = $full.Substring($rootPart.Length).Trim('\')
    if ($rest) {
        $walk = $rootPart
        foreach ($seg in $rest.Split('\')) {
            if (-not $seg) { continue }
            $walk = [System.IO.Path]::Combine($walk, $seg)
            $child = $null
            foreach ($c in $node.Nodes) {
                if (([string]$c.Tag).TrimEnd('\') -eq $walk.TrimEnd('\')) { $child = $c; break }
            }
            if (-not $child) { return $null }
            $node = $child
            [void]$node.Expand()
        }
    }
    return $node
}
# 顶层只有「此电脑」+ 所有盘符（含可移动盘）；上次位置靠展开定位，不再单独挂节点
$pcNode = New-Object System.Windows.Forms.TreeNode("[[imp.ui.this_pc]]")
$pcNode.Tag = ""
try {
    foreach ($drv in [System.IO.DriveInfo]::GetDrives()) {
        try {
            if (-not $drv.IsReady) { continue }
            $label = $drv.Name
            if ($drv.VolumeLabel) { $label = "$($drv.Name)  $($drv.VolumeLabel)" }
            $node = New-DirNode $pcNode $drv.Name
            $node.Text = $label
        } catch { continue }
    }
} catch { }
[void]$tree.Nodes.Add($pcNode)
$pcNode.Expand()

$tree.add_BeforeExpand({
    param($sender, $e)
    $node = $e.Node
    if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Tag -eq $script:PH) {
        $node.Nodes.Clear()
        $path = $node.Tag
        if ($path -and [System.IO.Directory]::Exists($path)) {
            try {
                $dirKey = $path + "|dirs"
                if ($script:dirCache.ContainsKey($dirKey)) {
                    $dirs = $script:dirCache[$dirKey]
                } else {
                    $dirs = [System.IO.Directory]::GetDirectories($path)
                    $script:dirCache[$dirKey] = $dirs
                }
                foreach ($d in $dirs) { [void](New-DirNode $node $d) }
            } catch {
                $lblList.Text = "[[imp.ui.denied]]"
            }
        }
    }
})

$tree.add_AfterSelect({
    param($sender, $e)
    $path = $e.Node.Tag
    $list.Items.Clear()
    $lblList.Text = "[[imp.ui.list_label]]"
    if (-not $path) { return }
    if (-not [System.IO.Directory]::Exists($path)) {
        $lblList.Text = "[[imp.ui.denied]]"
        try {
            $errLog = Join-Path $env:TEMP "nx_batch_import_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: cannot access '{1}' (Directory.Exists=False; missing or permission denied)" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path)
        } catch { }
        $hint = "[[imp.ui.denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[imp.ui.title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
        return
    }
    try {
        $dirKey = $path + "|files"
        if ($script:dirCache.ContainsKey($dirKey)) {
            $files = $script:dirCache[$dirKey]
        } else {
            $files = [System.IO.Directory]::GetFiles($path, "*.*") | Where-Object { Test-PickFile $_ }
            $script:dirCache[$dirKey] = $files
        }
        $arr = New-Object System.Collections.ArrayList
        foreach ($f in $files) {
            $fi = New-Object System.IO.FileInfo($f)
            $item = New-Object System.Windows.Forms.ListViewItem($fi.Name)
            $item.ImageIndex = 0
            [void]$item.SubItems.Add([string]([math]::Round($fi.Length / 1KB, 1)))
            $item.Tag = $f
            $item.ToolTipText = $f
            [void]$arr.Add($item)
        }
        if ($arr.Count -gt 0) {
            $list.BeginUpdate()
            $list.Items.AddRange([System.Windows.Forms.ListViewItem[]]$arr.ToArray())
            $list.EndUpdate()
        }
        if ($arr.Count -eq 0) { $lblList.Text = "[[imp.ui.empty_dir]]" }
    } catch {
        $lblList.Text = "[[imp.ui.denied]]"
        try {
            $errLog = Join-Path $env:TEMP "nx_batch_import_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: failed to list files in '{1}'`r`n{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path, $_.Exception.ToString())
        } catch { }
        $hint = "[[imp.ui.denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[imp.ui.title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
    }
})

$list.add_DoubleClick({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-LeftRow ([string]$sel.Tag)
    }
})

$left.add_MouseClick({
    param($sender, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $hit = $left.HitTest($e.X, $e.Y)
    if (-not $hit.Item) { return }
    if (-not $hit.SubItem) { return }
    if ($hit.Item.SubItems.IndexOf($hit.SubItem) -ne 2) { return }
    $path = [string]$hit.Item.Tag
    $ans = [System.Windows.Forms.MessageBox]::Show("[[imp.ui.msg_remove_text]]`n`n$path", "[[imp.ui.msg_remove_title]]", "YesNo", "Question")
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        # 逐行删除会打乱撤销记录的原始位置，这里清空单步撤销
        $script:lastDeleted = $null
        $left.Items.RemoveAt($hit.Item.Index)
        Refresh-List
        Update-DeleteState
    }
})

$btnAddFile.add_Click({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-LeftRow ([string]$sel.Tag)
    }
})

# 添加文件夹：把所选目录（含子目录）里支持的格式一次性加入左栏；
# 已在清单中的路径由 Add-LeftRow 自身去重，与单个添加同一规则
$btnAddFolder.add_Click({
    param($sender, $e)
    if (-not $tree.SelectedNode) { return }
    $dir = [string]$tree.SelectedNode.Tag
    if (-not $dir -or -not [System.IO.Directory]::Exists($dir)) { return }
    $found = @()
    try {
        $found = [System.IO.Directory]::GetFiles($dir, "*.*", [System.IO.SearchOption]::AllDirectories)
    } catch {
        # 个别子目录无权限时退回只管当前层，避免整次添加失败
        try { $found = [System.IO.Directory]::GetFiles($dir, "*.*") } catch { $found = @() }
    }
    $found = @($found | Where-Object { Test-PickFile $_ } | Sort-Object)
    foreach ($f in $found) { Add-LeftRow $f -NoRefresh }
    Refresh-List
})

$txtPrefix.add_TextChanged({ param($sender, $e) Update-Preview })
$txtSuffix.add_TextChanged({ param($sender, $e) Update-Preview })

# 同名处理 / 输出位置：任一变化都重算右栏预演
$rbRename.add_CheckedChanged({ param($sender, $e) Update-Preview })
$rbSkip.add_CheckedChanged({ param($sender, $e) Update-Preview })
$rbOverwrite.add_CheckedChanged({ param($sender, $e) Update-Preview })

function Update-OutState {
    $on = $rbOutDir.Checked
    $txtOutDir.Enabled = $on
    $btnOutBrowse.Enabled = $on
}

$rbOutSrc.add_CheckedChanged({ param($sender, $e) Update-OutState; Update-Preview })
$rbOutDir.add_CheckedChanged({ param($sender, $e) Update-OutState; Update-Preview })
$txtOutDir.add_TextChanged({ param($sender, $e) Update-Preview })
$btnOutBrowse.add_Click({
    param($sender, $e)
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    try {
        $dlg.Description = "[[imp.ui.out_group]]"
        $cur = ([string]$txtOutDir.Text).Trim()
        if ($cur -and [System.IO.Directory]::Exists($cur)) { $dlg.SelectedPath = $cur }
    } catch { }
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutDir.Text = $dlg.SelectedPath
        Update-Preview
    }
})
Update-OutState

# 多选删除的按钮 / 快捷键 / 右键菜单接线都在共享片段里：Initialize-LeftList
Initialize-LeftList

$btnClear.add_Click({
    param($sender, $e)
    Clear-LeftRows
})

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[imp.ui.btn_cancel]]"
$btnCancel.Location = New-Object System.Drawing.Point(742, 650)
$btnCancel.Size = New-Object System.Drawing.Size(80, 28)
$form.Controls.Add($btnCancel)
$btnCancel.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btnOK.Text = "[[imp.ui.btn_ok]]"
$btnOK.Location = New-Object System.Drawing.Point(832, 650)
$btnOK.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnOK)
$btnOK.add_Click({
    param($sender, $e)
    if ($left.Items.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("[[imp.ui.msg_empty_one]]`n[[imp.ui.msg_empty_two]]", "[[imp.ui.msg_empty_title]]", "OK", "Warning")
        return
    }
    # 指定目录为空/不存在：提示后回落到源目录，不阻断
    if ($rbOutDir.Checked) {
        $d = ([string]$txtOutDir.Text).Trim()
        if (-not $d -or -not [System.IO.Directory]::Exists($d)) {
            [void][System.Windows.Forms.MessageBox]::Show("[[imp.ui.out_invalid]]", "[[imp.ui.title]]", "OK", "Warning")
            $rbOutSrc.Checked = $true
        }
    }
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# 先建树句柄再定位：展开/选中会触发 BeforeExpand 建子节点、AfterSelect 填充右侧列表，
# 保证窗口显示时列表已是上次目录的内容（而不是空列表）
[void]$tree.Handle

# 上次位置：沿「此电脑 → 盘符 → 逐级子目录」自动展开并选中，右侧列表随即显示该目录
$targetNode = Expand-ToPath $Root
if ($targetNode) {
    try {
        $tree.SelectedNode = $targetNode
        $targetNode.EnsureVisible()
    } catch { }
} else {
    try { $tree.SelectedNode = $pcNode } catch { }
}

$form.add_Shown({
    # 弹出即在前台：只在显示那一刻抢一次，之后点回 NX 不收回、不关闭
    Show-PickerFront $form ($OutFile + ".front") | Out-Null
})

Apply-Theme $form
Update-EmptyHints
[void]$form.ShowDialog()

$files = @()
$ok = ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
if ($ok) {
    foreach ($row in $left.Items) { $files += [string]$row.Tag }
}
$workersValue = 2
try { $workersValue = [int]$numWorkers.Value } catch { }

$collisionValue = "skip"
if ($rbRename.Checked) { $collisionValue = "rename" }
elseif ($rbOverwrite.Checked) { $collisionValue = "overwrite" }
$outModeValue = "source"
$outDirValue = ""
if ($rbOutDir.Checked) {
    $outModeValue = "dir"
    $outDirValue = ([string]$txtOutDir.Text).Trim()
}

$payload = [ordered]@{ ok = $ok; files = $files;
    workers = $workersValue; prefix = [string]$txtPrefix.Text; suffix = [string]$txtSuffix.Text;
    collision = $collisionValue; out_mode = $outModeValue; outdir = $outDirValue }
$json = $payload | ConvertTo-Json -Depth 3 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''


def has_visible_window(pid):
    """目标进程有没有可见顶层窗口；只用于诊断“窗口被隐藏导致卡死”，判断不了按可见处理"""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        user32.EnumWindows.argtypes = (ctypes.c_void_p, wintypes.LPARAM)
        user32.GetWindowThreadProcessId.argtypes = (wintypes.HWND,
                                                    ctypes.POINTER(wintypes.DWORD))
        found = []

        def cb(hwnd, _):
            owner = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if (owner.value == int(pid) and user32.IsWindowVisible(hwnd)
                    and user32.GetWindowTextLengthW(hwnd) > 0):
                found.append(hwnd)
            return True

        user32.EnumWindows(proc(cb), 0)
        return bool(found)
    except Exception:
        return True


def imp__powershell_path():
    """返回 powershell.exe 的可执行路径；找不到时返回 None"""
    candidates = [
        os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                     "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
        "powershell",
    ]
    for candidate in candidates:
        if os.path.exists(candidate) or candidate == "powershell":
            return candidate
    return None


def imp_render_ps(template):
    """把 PowerShell 模板里的 [[key]] 占位符替换成当前语言文案（缺项回落英文）"""
    merged = dict(imp_MSG["en"])
    merged.update(imp_MSG.get(imp_LANG) or {})
    for key, value in merged.items():
        template = template.replace("[[%s]]" % key, str(value))
    # 文件名清洗片段与导出窗共用一份（定义在公共执行器里）
    try:
        shared = imp__shared()
        template = template.replace("[[name_helpers]]", shared.NAME_HELPERS_PS)
        template = template.replace("[[left_list]]", shared.LEFT_LIST_PS)
        template = template.replace("[[front_helpers]]", shared.PICKER_FRONT_PS)
    except Exception:
        pass
    return template


def _ui_error_log(text):
    """选择窗内部异常写独立小日志，避免污染导入主日志"""
    import tempfile
    try:
        with open(os.path.join(tempfile.gettempdir(),
                               "nx_batch_import_ui_errors.log"), "a",
                  encoding="utf-8") as handle:
            handle.write("%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), text))
    except Exception:
        pass


def pick_files_ui(initial_dir, workers=2, prefix="", suffix="",
                  collision="skip", out_mode="source", outdir=""):
    """弹出与导出侧同款的导入选择窗。

    返回 (文件列表, 是否走了新窗口, 并发数, 前缀, 后缀,
          同名处理, 输出模式, 指定输出目录)。
    窗口本身选完即确认，不再二次弹框。
    任何异常都抛给调用方，由调用方回落系统多选框。
    """
    import json
    import subprocess
    import tempfile

    if os.name != "nt":
        raise RuntimeError("import picker needs Windows")

    temp_dir = tempfile.gettempdir()
    ps_path = os.path.join(temp_dir, "nx_batch_import_picker.ps1")
    out_path = os.path.join(temp_dir, "nx_batch_import_picker.json")
    front_path = out_path + ".front"
    imp__rotate_log(os.path.join(temp_dir, "nx_batch_import_ui_errors.log"))

    # 上次异常退出留下的结果/前台标记先清掉，避免读到过期内容
    for stale in (out_path, front_path):
        try:
            if os.path.exists(stale):
                os.remove(stale)
        except Exception:
            pass

    # 导入窗是同步等待的，NX 在此期间点不到菜单，不存在“窗已开着再点一次”；
    # 这里只管把新窗抢到前台（结果由窗口自己写 .front 标记，下面读）。

    with open(ps_path, "w", encoding="utf-8-sig") as handle:
        handle.write(imp_render_ps(IMPORT_UI_PS))

    powershell = imp__powershell_path()
    if powershell is None:
        raise RuntimeError("powershell.exe not found")

    arguments = ["-File", ps_path, "-OutFile", out_path,
                 "-Workers", str(max(1, min(6, int(workers or 2)))),
                 "-Collision", (collision if collision in ("rename", "skip", "overwrite") else "skip"),
                 "-OutMode", ("dir" if out_mode == "dir" else "source")]
    if outdir:
        arguments += ["-OutDir", outdir]
    if prefix:
        arguments += ["-Prefix", prefix]
    if suffix:
        arguments += ["-Suffix", suffix]
    if initial_dir:
        arguments += ["-Root", initial_dir]
    command = [powershell, "-NoProfile", "-NonInteractive", "-STA",
               "-WindowStyle", "Hidden", "-ExecutionPolicy",
               "Bypass"] + arguments

    kwargs = {"stdout": subprocess.PIPE, "stderr": subprocess.PIPE,
              "creationflags": 0x08000000}
    # 这里绝不能设 STARTUPINFO 的 wShowWindow=0/SW_HIDE：
    # WinForms 窗体的第一次 ShowWindow 会被它覆盖，窗口永远不可见，
    # 而 NX 在 subprocess 里一直等 → 表现为「点批量导入后卡死」。
    # 控制台黑框已经由 CREATE_NO_WINDOW 处理，不需要再隐藏窗口。
    process = subprocess.Popen(command, **kwargs)
    started = time.time()
    trace("PICK_START", "pid=%s script=%s" % (process.pid, ps_path))
    flagged = False
    while process.poll() is None:
        if not flagged and time.time() - started > 8:
            flagged = True
            visible = has_visible_window(process.pid)
            trace("PICK_WINDOW", "visible=%s pid=%s" % (visible, process.pid))
            if not visible:
                imp_log(imp_T("imp.err.picker_hidden", process.pid))
        time.sleep(0.2)
    trace("PICK_EXIT", "rc=%s elapsed=%.2fs" % (process.returncode,
                                                time.time() - started))
    try:
        stderr = process.stderr.read() if process.stderr else b""
    except Exception:
        stderr = b""

    # 抢前台验收结果（窗口自己写的 ok/failed），日志里可判断
    try:
        if os.path.exists(front_path):
            with open(front_path, "r", encoding="ascii") as handle:
                front_state = (handle.read() or "").strip() or "unknown"
            trace("DLG_FRONT", front_state)
            imp_log("[dlg] front=%s" % front_state)
    except Exception:
        pass

    if not os.path.exists(out_path):
        detail = ""
        try:
            detail = (stderr or b"").decode("utf-8", "ignore")[:300]
        except Exception:
            pass
        raise RuntimeError("picker produced no result: %s" % detail)

    with open(out_path, "r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)
    if not payload.get("ok"):
        return [], True, workers, prefix, suffix, collision, out_mode, outdir
    raw = payload.get("files") or []
    if isinstance(raw, str):
        # 只有一个文件时 ConvertTo-Json 会退化成字符串
        raw = [raw]
    files = []
    for item in raw:
        text = str(item or "").strip()
        if text:
            files.append(text)
    prefix = str(payload.get("prefix") or "")
    suffix = str(payload.get("suffix") or "")
    collision = str(payload.get("collision") or collision).strip().lower()
    if collision not in ("rename", "skip", "overwrite"):
        collision = "skip"
    out_mode = "dir" if str(payload.get("out_mode") or out_mode).strip().lower() == "dir" else "source"
    outdir = str(payload.get("outdir") or "")
    try:
        workers = max(1, min(6, int(payload.get("workers") or workers)))
    except Exception:
        workers = max(1, min(6, int(workers or 2)))
    trace("PICK_OK", "%d file(s)" % len(files))
    return files, True, workers, prefix, suffix, collision, out_mode, outdir



def select_files(initial_dir):
    """弹系统多选框；取消返回 []，异常抛给调用方。"""
    if os.name != "nt":
        raise RuntimeError("只支持 Windows 系统文件对话框")

    buffer = ctypes.create_unicode_buffer(_FILE_BUFFER_CHARS)
    ofn = _OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(_OPENFILENAMEW)
    ofn.hwndOwner = None
    ofn.lpstrFilter = _file_filter()
    ofn.nFilterIndex = 1
    ofn.lpstrFile = ctypes.cast(buffer, ctypes.c_wchar_p)
    ofn.nMaxFile = _FILE_BUFFER_CHARS
    ofn.lpstrTitle = imp_T("imp.title")
    ofn.lpstrInitialDir = initial_dir or None
    ofn.Flags = (OFN_ALLOWMULTISELECT | OFN_EXPLORER | OFN_FILEMUSTEXIST |
                 OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR)

    try:
        comdlg32 = ctypes.WinDLL("comdlg32")
    except Exception:
        comdlg32 = ctypes.windll.comdlg32
    get_open = comdlg32.GetOpenFileNameW
    get_open.argtypes = [ctypes.POINTER(_OPENFILENAMEW)]
    get_open.restype = ctypes.c_int

    if not get_open(ctypes.byref(ofn)):
        # 0 = 用户取消；非 0 错误码说明对话框本身有问题
        try:
            code = comdlg32.CommDlgExtendedError()
        except Exception:
            code = 0
        if code:
            raise RuntimeError("CommDlgExtendedError=0x%08X" % code)
        return []

    try:
        raw = ctypes.string_at(ctypes.addressof(buffer),
                               _FILE_BUFFER_CHARS * 2).decode("utf-16-le")
    except Exception:
        raw = ctypes.wstring_at(ctypes.addressof(buffer), _FILE_BUFFER_CHARS)
    parts = [item for item in raw.split("\x00") if item]
    if not parts:
        return []
    if len(parts) == 1:
        return [parts[0]]
    folder = parts[0]
    return [os.path.join(folder, item) for item in parts[1:]]


# ================= NX 交互 =================

def message_box(kind, title, text):
    """NXMessageBox；kind = 'Information' / 'Warning' / 'Error' / 'Question'。UI 不可用时返回 None。"""
    try:
        ui = NXOpen.UI.GetUI()
        if ui is None:
            return None
        kinds = getattr(NXOpen.NXMessageBox, "DialogType", None)
        value = getattr(kinds, kind, None) if kinds is not None else None
        if value is None:
            return None
        return ui.NXMessageBox.Show(title, value, text)
    except Exception:
        return None


def confirm_yes(title, text):
    response = message_box("Question", title, text)
    if response is None:
        imp_log(imp_T("imp.log.no_confirm"))
        return True
    try:
        return response == NXOpen.NXMessageBox.DialogResponse.Yes
    except Exception:
        return True


def get_listing(session):
    try:
        listing = session.ListingWindow
        if not listing.IsOpen:
            listing.Open()
        return listing
    except Exception:
        return None


def listing_write(listing, text):
    if listing is None:
        return
    try:
        listing.WriteLine(text)
    except Exception:
        pass


def imp_find_open_part(session, path):
    """已打开的部件按 FullPath 匹配（NXOpen 没有 FindPartByPathName 这类接口）"""
    target = os.path.normcase(os.path.abspath(path))
    try:
        parts = list(session.Parts)
    except Exception:
        return None
    for part in parts:
        try:
            current = part.FullPath
        except Exception:
            continue
        if current and os.path.normcase(os.path.abspath(current)) == target:
            return part
    return None


def _split_open_result(result):
    """OpenDisplay / NewDisplay 可能返回 Part，也可能返回 (Part, PartLoadStatus)"""
    if isinstance(result, (tuple, list)):
        if len(result) >= 2:
            return result[0], result[1]
        if len(result) == 1:
            return result[0], None
    return result, None


def _describe_load_status(status):
    if status is None:
        return ""
    fields = []
    for attr in ("NumberUnloadedParts", "NumberParts", "NumberErrors"):
        try:
            value = getattr(status, attr)
        except Exception:
            continue
        if isinstance(value, int):
            fields.append("%s=%d" % (attr, value))
    detail = []
    try:
        getter = getattr(status, "GetPartLoadStatus", None)
        count = 0
        try:
            count = int(getattr(status, "NumberUnloadedParts", 0) or 0)
        except Exception:
            count = 0
        if getter is not None:
            for index in range(min(count, 5)):
                try:
                    info = getter(index)
                except Exception:
                    break
                text = []
                for name in ("PartName", "FileName", "Description", "Status"):
                    try:
                        value = getattr(info, name)
                    except Exception:
                        continue
                    if value:
                        text.append(str(value))
                if text:
                    detail.append(" / ".join(text))
    except Exception:
        pass
    if detail:
        return "; ".join(fields + detail)
    if fields:
        return ", ".join(fields)
    try:
        return str(status)[:200]
    except Exception:
        return "PartLoadStatus(unprintable)"


def display_option_keep_others():
    """取“在已有显示部件之外再显示一个”的枚举值。

    DisplayPartOption.ReplaceExisting 会把上一个显示部件替换掉：该部件
    未被修改时 NX 直接把它从会话里关掉，于是批量导入 N 个文件最后只剩
    最后一个（现场截图就是这个现象）。AllowAdditional 是追加显示。
    """
    options = getattr(NXOpen, "DisplayPartOption", None)
    return getattr(options, "AllowAdditional", None) if options is not None else None


def ensure_multi_display(session):
    """打开会话的“允许多个显示部件”选项。

    该选项被关掉时，任何切换显示部件的操作都会替换当前窗口内容，
    即使传 AllowAdditional 也只会剩一个部件。返回 (状态, 动作) 供日志。
    """
    parts = session.Parts
    try:
        status = parts.AllowMultipleDisplayedParts
    except Exception:
        return "unknown", "none"
    enum_cls = getattr(NXOpen.PartCollection, "MultipleDisplayedPartStatus", None)
    enabled = getattr(enum_cls, "Enabled", None)
    if enabled is not None:
        try:
            if status == enabled:
                return "Enabled", "none"
        except Exception:
            pass
    value = getattr(status, "value", None)
    setter = getattr(parts, "SetAllowMultipleDisplayedParts", None)
    if setter is not None:
        try:
            setter(True)
            return "value=%r enabled" % value, "SetAllowMultipleDisplayedParts(True)"
        except Exception as exc:
            return "value=%r failed %s" % (value, exc), "none"
    return "value=%r" % value, "none"


def activate_display(session, part):
    """把 part 设为显示部件，但不挤掉其它已加载部件。"""
    keep = display_option_keep_others()
    if keep is not None:
        session.Parts.SetActiveDisplay(
            part, keep, NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
        return "SetActiveDisplay(AllowAdditional)"
    session.Parts.SetActiveDisplay(
        part,
        NXOpen.DisplayPartOption.ReplaceExisting,
        NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
    return "SetActiveDisplay(ReplaceExisting)"


def set_work_part(session, part):
    """NewDisplay 只保证显示部件；这里显式把新部件设为工作部件。
    优先走 SetActiveDisplay(AllowAdditional)：新部件同时成为显示/工作部件，
    又不会关掉前面已经导入的部件。"""
    errors = []
    keep = display_option_keep_others()
    setter = getattr(session.Parts, "SetActiveDisplay", None)
    if setter is not None and keep is not None:
        try:
            setter(part, keep,
                   NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
            return "SetActiveDisplay(AllowAdditional)"
        except Exception as exc:
            errors.append("SetActiveDisplay(AllowAdditional): %s" % exc)
    setter = getattr(session.Parts, "SetDisplay", None)
    if setter is not None:
        try:
            setter(part, False, True)
            return "SetDisplay"
        except Exception as exc:
            errors.append("SetDisplay: %s" % exc)
    setter = getattr(session.Parts, "SetWork", None)
    if setter is not None:
        try:
            setter(part)
            return "SetWork"
        except Exception as exc:
            errors.append("SetWork: %s" % exc)
    try:
        return activate_display(session, part)
    except Exception as exc:
        errors.append("SetActiveDisplay: %s" % exc)
    raise RuntimeError("; ".join(errors) or "no work-part API")


def close_part(session, part):
    """尽力关闭导入失败留下的新部件；返回是否成功。

    枚举是扁平的 NXOpen.BasePartCloseWholeTree / BasePartCloseModified。
    CloseModified 必须用 CloseModified（空部件是“已修改”状态，用
    DontCloseModified 会报 “Modified part not saved” 而关不掉）。
    """
    closer = getattr(part, "Close", None)
    whole = getattr(getattr(NXOpen, "BasePartCloseWholeTree", None),
                    "TrueValue", None)
    modified = getattr(getattr(NXOpen, "BasePartCloseModified", None),
                       "CloseModified", None)
    if closer is None or whole is None or modified is None:
        return False
    responses = None
    try:
        responses = session.Parts.NewPartCloseResponses()
    except Exception:
        responses = None
    try:
        closer(whole, modified, responses)
        return True
    except Exception:
        return False


def new_display_part(session, name):
    maker = getattr(session.Parts, "NewDisplay", None)
    if maker is None:
        raise RuntimeError("PartCollection.NewDisplay 不存在")
    units = getattr(NXOpen, "PartUnits", None)
    value = getattr(units, "Millimeters", None) if units is not None else None
    result = maker(name) if value is None else maker(name, value)
    part, status = _split_open_result(result)
    if part is None:
        raise RuntimeError("NewDisplay 没有返回部件")
    return part, status


def align_new_part_layers(part):
    """把新建部件的图层状态铺成「NX 直接打开 STEP」时的那套默认值。

    实测（NX 2506）：直接打开 .step 得到的新部件，工作层 = 1，1~255 层为可选、
    256 层为可见；而 PartCollection.NewDisplay 建出的空部件除工作层外全部隐藏，
    STEP 里 PRESENTATION_LAYER_ASSIGNMENT 指定的层（本例 151）导入后仍然不可见，
    图形区因此空白。新建部件没有用户自己的图层约定可破坏，这里按同一套默认值
    铺开，让转换产物与手动打开的图层状态一致；工作层保持不动。
    """
    try:
        import NXOpen.Layer
        state = NXOpen.Layer.State
    except Exception as exc:
        trace("LAYER_ALIGN_FAIL", exc)
        return 0
    try:
        work = int(part.Layers.WorkLayer)
    except Exception as exc:
        trace("LAYER_ALIGN_FAIL", exc)
        return 0
    changed = 0
    for layer in range(1, 257):
        if layer == work:
            continue
        want = state.Visible if layer == 256 else state.Selectable
        try:
            if str(part.Layers.GetState(layer)) == str(want):
                continue
            part.Layers.SetState(layer, want)
            changed += 1
        except Exception:
            continue
    trace("LAYER_ALIGN", "work=%d changed=%d" % (work, changed))
    return changed


def open_part_file(session, path):
    """打开 .prt，返回 (部件, loadStatus, 使用的方式)。

    OpenDisplay / OpenBaseDisplay 在 NX 里等价于 ReplaceExisting：上一个
    显示部件没被修改时会被直接关掉，批量打开 N 个 .prt 最后只剩最后一个。
    优先用 OpenActiveDisplay(AllowAdditional) 追加打开，全部留在会话里。
    """
    keep = display_option_keep_others()
    opener = getattr(session.Parts, "OpenActiveDisplay", None)
    if opener is not None and keep is not None:
        try:
            result = opener(path, keep)
        except Exception as exc:
            raise RuntimeError(imp_T("imp.reason.open", exc))
        part, status = _split_open_result(result)
        if part is None:
            raise RuntimeError(imp_T("imp.reason.open", "OpenActiveDisplay 没有返回部件"))
        return part, status, "OpenActiveDisplay(AllowAdditional)"
    opener = getattr(session.Parts, "OpenDisplay", None)
    if opener is None:
        opener = getattr(session.Parts, "OpenBaseDisplay", None)
    if opener is None:
        raise RuntimeError("PartCollection.OpenDisplay 不存在")
    try:
        result = opener(path)
    except Exception as exc:
        raise RuntimeError(imp_T("imp.reason.open", exc))
    part, status = _split_open_result(result)
    if part is None:
        raise RuntimeError(imp_T("imp.reason.open", "OpenDisplay 没有返回部件"))
    return part, status, "OpenDisplay(ReplaceExisting)"


def part_file_on_disk(part):
    """部件对应的磁盘文件；内存态部件返回空串"""
    try:
        full = part.FullPath or ""
    except Exception:
        return ""
    if full and os.path.isfile(full):
        return full
    return ""


def part_still_loaded(session, part, path):
    """部件是否仍在会话中：有磁盘路径按路径比对，否则按对象身份比对"""
    if path:
        return imp_find_open_part(session, path) is not None
    try:
        for item in list(session.Parts):
            if item is part:
                return True
    except Exception:
        return True
    return False


def displayed_part_names(session):
    """当前显示集合里的部件名；老版本没有 MDP 时返回空列表"""
    try:
        shown = list(session.Parts.GetDisplayedParts())
    except Exception:
        return []
    names = []
    for item in shown:
        try:
            names.append(item.Leaf)
        except Exception:
            names.append(str(item))
    return names


def part_still_displayed(session, part, path):
    """部件是否还在显示集合里。

    STEP/IGS 分支用 Parts.NewDisplay 建新部件，NX 会把之前的显示集合
    换掉，部件仍留在会话里但不再显示。这里按对象身份或磁盘路径比对。
    """
    try:
        shown = list(session.Parts.GetDisplayedParts())
    except Exception:
        return True
    for item in shown:
        if item is part:
            return True
        try:
            if path and (item.FullPath or "") == path:
                return True
        except Exception:
            pass
    return False


def set_part_displayed(session, part):
    """把仍在会话里的部件重新加入显示集合（AllowAdditional，不挤掉别的部件）"""
    try:
        action = activate_display(session, part)
        trace("SETDISPLAY_OK", "%s %s" % (getattr(part, "Leaf", part), action))
        return True
    except Exception as exc:
        trace("SETDISPLAY_FAIL", "%s: %s" % (getattr(part, "Leaf", part), exc))
        return False


def reopen_missing_parts(session, records):
    """兜底：把被显示部件切换挤掉的部件补回会话，并补上显示。

    records = [(名称, 磁盘路径或空, 部件对象)]。
    返回 (补回数量, 仍缺失的名称列表)。补回数量含"只在会话里、没在显示集合里"
    被重新追加显示的部件。
    """
    restored = 0
    missing = []
    for name, path, part in records:
        if not part_still_loaded(session, part, path):
            if not path or not os.path.isfile(path):
                missing.append(name)
                continue
            try:
                open_part_file(session, path)
                restored += 1
                trace("REOPEN_OK", "%s <- %s" % (name, path))
            except Exception as exc:
                missing.append(name)
                trace("REOPEN_FAIL", "%s: %s" % (name, exc))
            continue
        if part_still_displayed(session, part, path):
            continue
        # 部件还在会话里，只是不显示：先试 SetDisplay，失败再退回重新打开
        if set_part_displayed(session, part):
            restored += 1
            trace("REDISPLAY_OK", "%s (SetDisplay)" % name)
            continue
        if not path or not os.path.isfile(path):
            continue
        try:
            open_part_file(session, path)
            restored += 1
            trace("REDISPLAY_OK", "%s <- %s" % (name, path))
        except Exception as exc:
            trace("REDISPLAY_FAIL", "%s: %s" % (name, exc))
    return restored, missing


def _dex_manager(session):
    """DexManager：优先会话级，其次当前工作部件"""
    manager = getattr(session, "DexManager", None)
    if manager is not None:
        return manager
    try:
        work = session.Parts.Work
    except Exception:
        work = None
    if work is not None:
        manager = getattr(work, "DexManager", None)
        if manager is not None:
            return manager
    raise RuntimeError("DexManager 不存在")


def create_importer(session, method_names):
    manager = _dex_manager(session)
    errors = []
    for method_name in method_names:
        factory = getattr(manager, method_name, None)
        if factory is None:
            errors.append("%s: missing" % method_name)
            continue
        try:
            return factory()
        except Exception as exc:
            errors.append("%s: %s" % (method_name, exc))
    raise RuntimeError("; ".join(errors))


def set_input_file(importer, path):
    for attr in ("InputFile", "FileName", "InputFileName"):
        if hasattr(importer, attr):
            try:
                setattr(importer, attr, path)
                return attr
            except Exception:
                continue
    setter = getattr(importer, "SetInputFile", None)
    if setter is not None:
        setter(path)
        return "SetInputFile"
    raise RuntimeError("导入器没有可用的输入文件属性")


def _import_to_enum(importer, member):
    """按嵌套枚举 / 扁平类名两种形式找 ImportTo 枚举成员。

    不同导入器的枚举类型不同（Step214Importer.ImportToOption、
    IgesImporter.ImportToEnum），先按导入器自身类名找；
    不能把 Step214 的枚举塞给 IGES，否则报 “Expecting ... type, found ...”。
    """
    cls = type(importer)
    owners = []
    for attr in ("ImportToOption", "ImportToEnum"):
        nested = getattr(cls, attr, None)
        if nested is not None:
            owners.append(nested)
    name = getattr(cls, "__name__", "")
    for suffix in ("ImportToEnum", "ImportToOption", "ImportTo"):
        flat = getattr(NXOpen, name + suffix, None)
        if flat is not None:
            owners.append(flat)
    if not owners:
        for attr in ("Step214ImporterImportToOption", "StepImporterImportToOption",
                     "IgesImporterImportToEnum", "IgesImporterImportToOption"):
            value = getattr(NXOpen, attr, None)
            if value is not None:
                owners.append(value)
    for owner in owners:
        value = getattr(owner, member, None)
        if value is not None:
            return value
    return None


def apply_import_to(importer):
    """三器统一口径：导入目标 = 工作部件。
    STEP / IGES 显式设 ImportTo=WorkPart；Parasolid 无 ImportTo（本就只导工作部件）。
    失败时抛错，由调用方记日志并继续。"""
    if not hasattr(importer, "ImportTo"):
        return "default(work part)"
    value = _import_to_enum(importer, "WorkPart")
    if value is None:
        raise RuntimeError("ImportTo 枚举里没有 WorkPart")
    importer.ImportTo = value
    return "ImportTo=WorkPart"


def translator_settings_file(label):
    r"""translator 默认设置文件：STEP 用 STEP214UG\step214ug.def，
    IGES 用 IGES\igesimport.def（与原生“文件→导入→STEP214 / IGES”同一份 def）。"""
    if label == "STEP":
        envs = (("STEP214UG_DIR", "step214ug.def"), ("UGII_STEP214_DIR", "step214ug.def"))
        relative = os.path.join("STEP214UG", "step214ug.def")
    elif label == "IGES":
        envs = (("IGESUG_DIR", "igesimport.def"), ("UGII_IGES_DIR", "igesimport.def"))
        relative = os.path.join("IGES", "igesimport.def")
    else:
        return ""
    for var, filename in envs:
        value = (os.environ.get(var) or "").strip()
        if value:
            path = os.path.join(value, filename)
            if os.path.isfile(path):
                return path
    base = (os.environ.get("UGII_BASE_DIR") or "").strip()
    if base:
        path = os.path.join(base, relative)
        if os.path.isfile(path):
            return path
    return ""


def translator_output_file(name):
    """translator 的输出 prt 路径：必须可写，且目标不能已存在。

    用 %TEMP%\\nx_batch_import 而不是 part.FullPath —— GUI 里 NX 的 cwd 可能是
    只读安装目录，写不了就没法翻译；文件名带 pid + 毫秒，避免目标已存在时
    translator 改写 _1.prt 而导入效果不确定。
    该 prt 只是 translator 的中间产物，NX 加载进工作部件后不留在磁盘上。
    仅 STEP / IGES 需要（两者都是“先翻成 prt 再加载”）。
    """
    folder = os.path.join(tempfile.gettempdir(), "nx_batch_import")
    os.makedirs(folder, exist_ok=True)
    stem = _ILLEGAL_CHARS.sub("_", name or "imported")[:MAX_NAME_LEN]
    return os.path.join(folder, "%s_%d_%d.prt" % (stem, os.getpid(),
                                                  int(time.time() * 1000)))


def import_temp_dir():
    """translator 中间产物目录：临时 prt 与 .log 都写在这里"""
    return os.path.join(tempfile.gettempdir(), "nx_batch_import")


def _logs_since(folder, since):
    """列出目录里修改时间 >= since 的 .log（即本次导入产生的转换器日志）

    用严格 >=（NTFS 的 mtime 精度足够），只清理本次导入之后产生的日志，不动更早的残留。
    """
    result = []
    try:
        for name in os.listdir(folder):
            if not name.lower().endswith(".log"):
                continue
            path = os.path.join(folder, name)
            try:
                if os.path.getmtime(path) >= since:
                    result.append(path)
            except Exception:
                continue
    except Exception:
        pass
    return result


def cleanup_run_logs(since):
    """导入结束（成功或失败）后删除本次产生的 translator .log

    删除失败不影响换算结果，只在日志里记一条；返回实际删掉的个数。
    """
    count = 0
    stuck = []
    for path in _logs_since(import_temp_dir(), since):
        try:
            os.remove(path)
            count += 1
        except FileNotFoundError:
            continue
        except Exception as exc:
            stuck.append("%s (%s)" % (os.path.basename(path), exc))
    if count:
        imp_log(imp_T("imp.log.temp_clean", count, import_temp_dir()))
    if stuck:
        imp_log(imp_T("imp.log.temp_clean_fail", len(stuck), "; ".join(stuck)))
    return count


def cleanup_old_logs(days=TEMP_LOG_DAYS):
    """会话启动时清理 N 天前的 translator .log；失败静默，不影响导入"""
    cutoff = time.time() - days * 86400
    count = 0
    try:
        for name in os.listdir(import_temp_dir()):
            if not name.lower().endswith(".log"):
                continue
            path = os.path.join(import_temp_dir(), name)
            try:
                if os.path.getmtime(path) < cutoff:
                    os.remove(path)
                    count += 1
            except Exception:
                continue
    except Exception:
        pass
    return count


def prepare_importer(importer, label, part, name="", sew_surfaces=SEW_SURFACES):
    """补齐 Dex 导入器必填项：SettingsFile / OutputFile / FileOpenFlag。

    STEP / IGES 不是“直接读进内存”，而是 translator 先写 prt、NX 再加载；
    DexManager 的 Commit 走同一条链，所以必须告诉它输出到哪个 prt：
    OutputFile 必须指向可写目录且文件尚不存在（见 translator_output_file），
    否则 translator 报 “Unable to import selected file to work part”，
    部件既没有几何也没有转换器日志——表现为“卡住/失败但没有产物”。
    FileOpenFlag=False 表示这是导入而不是打开文件（Siemens 官方样例口径）。
    Parasolid 不经过 translator，也没有这些属性，直接导工作部件。
    属性按 hasattr 逐个补，缺的跳过（不同 NX 版本类定义有差异）。
    """
    if label == "Parasolid":
        return "n/a(parasolid)"
    applied = []
    settings = translator_settings_file(label)
    if settings:
        if hasattr(importer, "SettingsFile"):
            importer.SettingsFile = settings
            applied.append("SettingsFile")
    else:
        imp_log(imp_T("imp.log.settings_missing"))
    output = ""
    try:
        output = translator_output_file(name)
    except Exception:
        try:
            output = part.FullPath or ""
        except Exception:
            output = ""
    if output and hasattr(importer, "OutputFile"):
        importer.OutputFile = output
        applied.append("OutputFile")
    if hasattr(importer, "FileOpenFlag"):
        importer.FileOpenFlag = False
        applied.append("FileOpenFlag=False")
    # 自动缝合：默认关。只有显式传 True 才写属性，False 时仅探查一次当前默认值，
    # 保证默认路径下导入结果与改动前完全一致（属性不存在只记日志，不抛错）。
    if label in ("STEP", "IGES"):
        if hasattr(importer, "SewSurfaces"):
            if label not in _sew_probed:
                _sew_probed.add(label)
                try:
                    imp_log(imp_T("imp.log.sew_probe", label, bool(importer.SewSurfaces)))
                except Exception as exc:
                    imp_log(imp_T("imp.log.sew_probe_fail", label, exc))
            if sew_surfaces:
                try:
                    importer.SewSurfaces = True
                    applied.append("SewSurfaces=True")
                    imp_log(imp_T("imp.log.sew_set", label))
                except Exception as exc:
                    imp_log(imp_T("imp.log.sew_set_fail", label, exc))
        else:
            imp_log(imp_T("imp.log.sew_unavailable", label))
    return ",".join(applied) if applied else "none"


def fit_view(part):
    try:
        view = part.Views.WorkView
        if view is not None:
            view.Fit()
            return True
    except Exception:
        pass
    try:
        part.Views.FitAll()
        return True
    except Exception as exc:
        imp_log(imp_T("imp.log.fit_fail", exc))
        return False


# ================= 文件名 / 部件名 =================

_ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_part_name(stem, used):
    """文件名主干 -> 合法部件名：清非法字符、截断 80、同名追加 _2/_3"""
    name = _ILLEGAL_CHARS.sub("_", stem or "").strip().strip(".")
    if not name:
        name = "imported"
    name = name[:MAX_NAME_LEN]
    base = name
    index = 2
    while name.lower() in used:
        suffix = "_%d" % index
        name = base[:MAX_NAME_LEN - len(suffix)] + suffix
        index += 1
    used.add(name.lower())
    return name


def plan_target(path, prefix="", suffix="", outdir="", collision="rename", reserved=None):
    """按同名处理与输出目录算出落盘目标，返回 (目标路径, 处理方式)。

    处理方式：new（新建）/ rename（同名改名 _2/_3）/ skip（跳过）/ overwrite（覆盖）。
    reserved 是本次批次已预定的绝对路径集合（normcase），避免两件抢同一个名字。
    前缀/后缀只作用于另存出来的文件名（.prt 源不另存，因此不受影响）。
    """
    folder = outdir or os.path.dirname(os.path.abspath(path))
    stem = imp_normalize_name(os.path.splitext(os.path.basename(path))[0])
    base = imp_safe_name_part(prefix) + stem + imp_safe_name_part(suffix)
    first = os.path.join(folder, base + ".prt")

    def taken(candidate):
        if os.path.exists(candidate):
            return True
        if reserved is not None:
            key = os.path.normcase(os.path.abspath(candidate))
            if key in reserved:
                return True
        return False

    if collision == "skip":
        if taken(first):
            return first, "skip"
        if reserved is not None:
            reserved.add(os.path.normcase(os.path.abspath(first)))
        return first, "new"

    if collision == "overwrite":
        mode = "overwrite" if taken(first) else "new"
        if reserved is not None:
            reserved.add(os.path.normcase(os.path.abspath(first)))
        return first, mode

    target = first
    if taken(target):
        serial = 2
        while True:
            target = os.path.join(folder, "%s_%d.prt" % (base, serial))
            serial += 1
            if not taken(target):
                break
        mode = "rename"
    else:
        mode = "new"
    if reserved is not None:
        reserved.add(os.path.normcase(os.path.abspath(target)))
    return target, mode


def save_target(path, prefix="", suffix="", outdir="", collision="rename"):
    """转换文件的落盘目标：outdir 为空时用源文件同目录。

    返回 (目标路径, 处理方式)；处理方式 = new / rename / skip / overwrite。
    """
    return plan_target(path, prefix, suffix, outdir, collision, None)


def file_written(path):
    """落盘校验：文件存在且非 0 字节（SaveAs 偶有“不报错但没写盘”）。"""
    try:
        return os.path.isfile(path) and os.path.getsize(path) > 0
    except Exception:
        return False


def existing_part_names(session):
    used = set()
    try:
        for part in list(session.Parts):
            try:
                leaf = part.LeafName or ""
            except Exception:
                leaf = ""
            stem = os.path.splitext(leaf)[0]
            if stem:
                used.add(stem.lower())
    except Exception:
        pass
    return used


def dedupe(paths):
    """按 normcase(abspath) 去重，保留用户选择顺序；返回 (队列, 跳过项)"""
    queue = []
    skipped = []
    seen = set()
    for path in paths:
        full = os.path.abspath(path)
        key = os.path.normcase(full)
        if key in seen:
            skipped.append((full, "dup"))
            continue
        seen.add(key)
        if not os.path.isfile(full):
            skipped.append((full, "missing"))
            continue
        queue.append(full)
    return queue, skipped


# ================= 单个文件 =================

def imp_backup_original(path):
    """覆盖前把原文件备份为同名 .bak（已存在则覆盖该 .bak）。

    备份失败必须中止覆盖并保留原文件；成功后 .bak 保留，由用户自行核对/清理。
    """
    bak = path + ".bak"
    try:
        shutil.copy2(path, bak)
    except Exception as e:
        raise RuntimeError(imp_T("imp.err.backup", path, bak, e))
    return bak


def import_one(session, path, used, save_as=True, prefix="", suffix="",
               outdir="", collision="rename", target=""):
    """返回 (状态, 部件名, 说明, 落盘状态, 部件)。

    状态 = ok / fail；落盘状态 = saved（已写入 .prt）/ mem（几何只在内存）/
    n/a（.prt 输入，不做另存）。
    save_as=False：只导入几何、不另存（供导出侧链式转换复用，调用方负责关闭部件）。
    prefix/suffix：只影响另存出来的 .prt 文件名，.prt 输入不受影响。
    target：调用方给定的落盘绝对路径（→prt 方向用计划算好的名字）；给定时
    不再用 outdir/collision 重算，已存在则先挪开、成功后再删旧件。
    """
    ext = os.path.splitext(path)[1].lower()
    started = time.time()

    if ext in SUPPORTED_PRT:
        imp_log(imp_T("imp.log.prt_open", path))
        part = imp_find_open_part(session, path)
        already = part is not None
        if already:
            imp_log(imp_T("imp.log.prt_already"))
            try:
                how = activate_display(session, part)
                imp_log(imp_T("imp.log.keep_open", how))
            except Exception as exc:
                raise RuntimeError(imp_T("imp.reason.open", exc))
            status = None
        else:
            part, status, how = open_part_file(session, path)
            imp_log(imp_T("imp.log.open_mode", how))
            trace("OPEN_MODE", how)
        name = ""
        try:
            name = os.path.splitext(part.LeafName or "")[0]
        except Exception:
            name = os.path.basename(path)
        if status is not None:
            text = _describe_load_status(status)
            if text:
                imp_log(imp_T("imp.log.load_status", text))
        imp_log(imp_T("imp.log.prt_done", time.time() - started, name))
        return "ok", name, "already_open" if already else "opened", "n/a", part

    if ext in SUPPORTED_STEP:
        methods = ("CreateStep214Importer", "CreateStepImporter")
        label = "STEP"
    elif ext in SUPPORTED_IGES:
        methods = ("CreateIgesImporter", "CreateIGESImporter")
        label = "IGES"
    elif ext in SUPPORTED_PARASOLID:
        methods = ("CreateParasolidImporter", "CreateXTImporter")
        label = "Parasolid"
    else:
        raise RuntimeError(imp_T("imp.reason.ext"))

    importer = None
    part = None
    previous = None
    done = False
    part_keep = False
    name = ""
    trace("IMPORT_START", "%s ext=%s" % (path, ext))
    try:
        # 先定落盘目标，再用目标主干命名新部件：
        # NX 的 NewDisplay 会拒绝与磁盘上已有 .prt 同名的部件（"File already exists"），
        # 必须在建部件之前就把 _2/_3 定下来，否则同名再导时 save_target 永远轮不到执行。
        if save_as:
            if target:
                # 显式目标：沿用计划算好的名字（可能已含 _2/_3），不重算
                save_path = os.path.abspath(target)
                target_mode = "overwrite" if os.path.exists(save_path) else "new"
            else:
                save_path, target_mode = save_target(path, prefix, suffix, outdir, collision)
            if target_mode == "skip":
                imp_log(imp_T("imp.log.skip_same", save_path))
                return "skip", "", save_path, "n/a", None
            stem = os.path.splitext(os.path.basename(save_path))[0]
            renamed = (target_mode == "rename")
        else:
            save_path, renamed = "", False
            stem = os.path.splitext(os.path.basename(path))[0]
        name = sanitize_part_name(stem, used)
        try:
            previous = session.Parts.Display
        except Exception:
            previous = None
        trace("DISPLAY_BEFORE", repr(previous))
        try:
            part, _status = new_display_part(session, name)
        except Exception as exc:
            raise RuntimeError(imp_T("imp.reason.no_part", exc))
        try:
            trace("NEW_DISPLAY_OK", "%s full=%r" % (name, part.FullPath))
        except Exception:
            trace("NEW_DISPLAY_OK", name)
        imp_log(imp_T("imp.log.new_part", name))
        try:
            how = set_work_part(session, part)
            imp_log(imp_T("imp.log.work_part", how))
            trace("SET_WORK_OK", how)
        except Exception as exc:
            raise RuntimeError(imp_T("imp.reason.work_part", exc))
        # 新部件图层默认状态对齐「NX 直接打开 STEP」，避免实体落在隐藏层导致图形区空白。
        # 只在 →prt 方向（save_as=True）做；链式转换的临时部件用完即关，导出行为保持原样。
        if save_as:
            try:
                align_new_part_layers(part)
            except Exception as exc:
                trace("LAYER_ALIGN_FAIL", exc)
        try:
            importer = create_importer(session, methods)
        except Exception as exc:
            raise RuntimeError(imp_T("imp.reason.import", exc))
        trace("IMPORTER_OK", type(importer).__name__)
        imp_log(imp_T("imp.log.import_begin", label, name))

        try:
            attr = set_input_file(importer, path)
        except Exception as exc:
            raise RuntimeError(imp_T("imp.reason.import", exc))
        trace("INPUT_FILE_OK", attr)
        try:
            prep = prepare_importer(importer, label, part, name)
            imp_log(imp_T("imp.log.prep", prep, getattr(importer, "OutputFile", "")))
            trace("PREP_OK", prep)
        except Exception as exc:
            imp_log(imp_T("imp.log.prep_fail", exc))
            trace("PREP_FAIL", exc)
        try:
            target = apply_import_to(importer)
            trace("IMPORT_TO_OK", target)
        except Exception as exc:
            imp_log(imp_T("imp.log.import_to", exc))
            target = "default"
            trace("IMPORT_TO_FAIL", exc)
        trace("COMMIT_BEFORE", "")
        try:
            status = importer.Commit()
        except Exception as exc:
            trace("COMMIT_EXC", exc)
            raise RuntimeError(imp_T("imp.reason.import", exc))
        trace("COMMIT_AFTER", "%.2fs" % (time.time() - started))
        # 几何已进内存：此后即使落盘失败也不再走“关闭空部件”路径
        part_keep = True
        text = _describe_load_status(status)
        if text:
            imp_log(imp_T("imp.log.load_status", text))
        try:
            trace("FIT_BEFORE", "")
            fit_view(part)
            trace("FIT_AFTER", "%d bodies" % len(list(part.Bodies)))
        except Exception as exc:
            trace("FIT_FAIL", exc)
        # 与 NX 打开 STEP 一致：产物直接落在源文件同级目录，同名加 _2/_3
        if not save_as:
            # 链式转换：几何留在内存，由调用方导出目标格式后关闭部件
            done = True
            imp_log(imp_T("imp.log.import_done", time.time() - started))
            return "ok", name, label + "/" + attr + "/" + target, "mem", part
        if renamed:
            imp_log(imp_T("imp.log.saveas_rename", save_path))
        # NX 的 SaveAs 不接受已存在的目标（"File already exists"）：覆盖场景先把
        # 旧件备份为同名 .bak（备份失败即中止），再移除旧件让 SaveAs 落盘；
        # 保存失败用 .bak 还原，成功则保留 .bak 供用户核对/清理。
        backup = ""
        imp_log(imp_T("imp.log.saveas_start", save_path))
        if os.path.exists(save_path):
            try:
                backup = imp_backup_original(save_path)
                os.remove(save_path)
            except Exception as exc:
                reason = str(exc).strip()
                imp_log(reason)
                return "fail", name, reason, "mem", part
            imp_log(imp_T("imp.log.backup", backup))
        trace("SAVEAS_BEFORE", save_path)
        try:
            part.SaveAs(save_path)
        except Exception as exc:
            trace("SAVEAS_EXC", exc)
            if backup:
                try:
                    shutil.copy2(backup, save_path)
                except Exception as back_exc:
                    imp_log("[prt] restore_failed %s: %s" % (save_path, back_exc))
            reason = imp_T("imp.err.saveas", exc, save_path).strip()
            imp_log(reason)
            return "fail", name, reason, "mem", part
        if not file_written(save_path):
            trace("SAVEAS_MISSING", save_path)
            if backup:
                try:
                    shutil.copy2(backup, save_path)
                except Exception as back_exc:
                    imp_log("[prt] restore_failed %s: %s" % (save_path, back_exc))
            reason = imp_T("imp.err.saveas_missing", save_path).strip()
            imp_log(reason)
            return "fail", name, reason, "mem", part
        # 成功时不删 .bak：留给用户核对/清理
        trace("SAVEAS_AFTER", save_path)
        done = True
        imp_log(imp_T("imp.log.import_done", time.time() - started))
        imp_log(imp_T("imp.log.saveas_ok", save_path))
        return "ok", name, label + "/" + attr + "/" + target, "saved", part
    finally:
        trace("IMPORT_END", "done=%r %.2fs" % (done, time.time() - started))
        if importer is not None:
            try:
                importer.Destroy()
            except Exception as exc:
                imp_log(imp_T("imp.log.destroy_fail", exc))
        # 失败清理：先切回原显示部件，再关掉空部件；关不掉就保留并记日志
        # part_keep=True 表示几何已导入但落盘失败，部件要留在内存并保持可见
        if not done and part is not None and not part_keep:
            if previous is not None and previous is not part:
                try:
                    session.Parts.SetDisplay(previous, False, True)
                except Exception as exc:
                    imp_log(imp_T("imp.log.fail", exc))
            if close_part(session, part):
                imp_log(imp_T("imp.log.close_empty", name))
            else:
                imp_log(imp_T("imp.log.close_fail", name))
        # 成功/失败都清掉本次 translator 产生的 .log，失败件不再残留；
        # 清理内部只记日志，异常不影响已得出的换算结果
        cleanup_run_logs(started)


# ================= 会话内直读（导入后打开） =================

def _folder_writable(folder):
    """目录能否落盘：写一个探针文件再删掉；不存在/只读/无权限都返回 False"""
    probe = ""
    try:
        if not os.path.isdir(folder):
            return False
        import tempfile
        handle = tempfile.NamedTemporaryFile(prefix=".nx_import_probe_",
                                             dir=folder, delete=False)
        probe = handle.name
        handle.write(b"0")
        handle.close()
        return True
    except Exception:
        return False
    finally:
        if probe:
            try:
                os.remove(probe)
            except Exception:
                pass


def _target_locked(path):
    """覆盖前的占用预检：只读先清掉；仍打不开 = 被占用或无权限"""
    if not os.path.exists(path):
        return False
    try:
        os.chmod(path, 0o600)
    except Exception:
        pass
    try:
        with open(path, "r+b"):
            return False
    except Exception:
        return True


def _snapshot_names(folder):
    try:
        return {name.lower() for name in os.listdir(folder)}
    except Exception:
        return set()


def _new_side_files(folder, before):
    """本次新增的 NX 中间文件：<名>_step.prt / <名>_igs.prt / <名>.log"""
    names = []
    try:
        for name in os.listdir(folder):
            low = name.lower()
            if low in before:
                continue
            if low.endswith(("_step.prt", "_stp.prt", "_igs.prt", "_x_t.prt",
                             "_xt.prt", ".log")):
                names.append(name)
    except Exception:
        pass
    return names


def _cleanup_side_files(folder, names, keep=()):
    """删掉翻译留下的中间 prt 与 .log，keep 里的名字（最终目标）不动"""
    keep_set = {str(item).lower() for item in keep}
    for name in names:
        if name.lower() in keep_set:
            continue
        path = os.path.join(folder, name)
        # SaveAs 之后中间件可能还被 NX 短暂占用，删不掉就重试几轮
        for attempt in range(6):
            try:
                os.remove(path)
                break
            except FileNotFoundError:
                break
            except Exception:
                try:
                    os.chmod(path, 0o666)
                except Exception:
                    pass
                time.sleep(0.15 * (attempt + 1))


def _parse_translation_log(log_path, offset=0):
    """只解析本次追加段，避免把上一件的成功当成这一件的结果。

    成功件里有 `Filing Part` 与 `Summary of Unigraphics Part Entities ... Total N`；
    坏件里有 `! ERROR-` 且没有 Filing Part（它也有一个 STEP 实体 Total，不能用）。
    """
    info = {"filed": False, "errors": [], "ug_total": 0, "found": False}
    try:
        size = os.path.getsize(log_path)
    except Exception:
        return info
    start = offset if 0 <= offset <= size else 0
    try:
        with open(log_path, "rb") as handle:
            handle.seek(start)
            data = handle.read()
    except Exception:
        return info
    text = data.decode("utf-8", "ignore")
    if not text.strip():
        return info
    if start == 0:
        # offset 失效时（文件被重写/截断）只取最后一段翻译
        for marker in ("STEP to UG", "IGES to UG"):
            pos = text.rfind(marker)
            if pos > 0:
                text = text[pos:]
                break
    info["found"] = True
    lines = text.splitlines()
    for line in lines:
        low = line.lower()
        if "filing part" in low:
            info["filed"] = True
        elif "! error-" in low:
            info["errors"].append(line.strip())
    ug_section = False
    for line in lines:
        low = line.lower()
        if "summary of unigraphics part entities" in low:
            ug_section = True
            continue
        if ug_section:
            match = re.search(r"\btotal\s+(\d+)\s*$", low)
            if match:
                info["ug_total"] = int(match.group(1))
                ug_section = False
    return info


def _redraw_nx():
    """让图形区立刻刷新：只重绘本进程窗口，不置顶、不抢焦点"""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        pid = os.getpid()
        flags = 0x0001 | 0x0080 | 0x0100 | 0x0004  # INVALIDATE|ALLCHILDREN|UPDATENOW|FRAME
        handles = []

        def _collect(hwnd, _lparam):
            owner = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if owner.value == pid and user32.IsWindowVisible(hwnd):
                handles.append(hwnd)
            return True

        callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(_collect)
        user32.EnumWindows(callback, 0)
        for hwnd in handles:
            user32.RedrawWindow(hwnd, None, None, flags)
            user32.UpdateWindow(hwnd)
    except Exception:
        pass


def _cancel_requested(cancel_flag):
    if cancel_flag and os.path.exists(cancel_flag):
        return True
    try:
        import ctypes
        return bool(ctypes.WinDLL("user32").GetAsyncKeyState(0x1B) & 0x8000)
    except Exception:
        return False


def _open_outcome(status, path, target, started, reason="", detail="", opened=False):
    return {"status": status, "src": path, "name": os.path.basename(path),
            "target": target, "reason": reason or "", "detail": detail,
            "seconds": time.time() - started, "opened": opened}


def open_one_in_session(session, path, prefix="", suffix="", outdir="", collision="skip",
                        reserved=None):
    """在当前会话直读一件：读完立即显示，再按同名/输出目录策略落盘。

    返回 dict：status=ok/skip/fail、detail、reason、target、seconds。
    失败会恢复原来的显示部件并清掉本次新增的中间 prt / .log。
    """
    started = time.time()
    reserved = set() if reserved is None else reserved
    ext = os.path.splitext(path)[1].lower()
    src_dir = os.path.dirname(os.path.abspath(path))
    name = os.path.basename(path)

    try:
        previous = session.Parts.Display
    except Exception:
        previous = None
    prev_path = ""
    if previous is not None:
        try:
            prev_path = os.path.normcase(os.path.abspath(previous.FullPath or ""))
        except Exception:
            prev_path = ""

    # .prt 源：只打开，不另存；输出目录与同名处理对它不生效
    if ext in SUPPORTED_PRT:
        existing = imp_find_open_part(session, path)
        if existing is not None:
            try:
                activate_display(session, existing)
            except Exception as exc:
                return _open_outcome("fail", path, path, started, str(exc), opened=True)
            fit_view(existing)
            return _open_outcome("ok", path, path, started, detail="already_open", opened=True)
        try:
            part, _status, method = open_part_file(session, path)
        except Exception as exc:
            return _open_outcome("fail", path, path, started, str(exc))
        try:
            shown = session.Parts.Display
        except Exception:
            shown = part
        if shown is not part or not part_file_on_disk(part):
            try:
                close_part(session, part)
            except Exception:
                pass
            return _open_outcome("fail", path, path, started, imp_T("imp.err.not_displayed"))
        trace("SESSION_OPEN", "%s %s %.2fs" % (name, method, time.time() - started))
        fit_view(part)
        return _open_outcome("ok", path, path, started, detail="opened", opened=True)

    target, mode = plan_target(path, prefix, suffix, outdir, collision, reserved)
    if mode == "skip":
        return _open_outcome("skip", path, target, started, detail="skip")

    if mode == "overwrite" and _target_locked(target):
        return _open_outcome("fail", path, target, started, imp_T("imp.err.locked", target))
    if not _folder_writable(src_dir):
        return _open_outcome("fail", path, target, started, imp_T("imp.err.src_readonly"))
    out_dir = os.path.dirname(os.path.abspath(target))
    if os.path.normcase(out_dir) != os.path.normcase(os.path.abspath(src_dir)) and not _folder_writable(out_dir):
        return _open_outcome("fail", path, target, started, imp_T("imp.err.out_readonly"))

    log_src = os.path.join(src_dir, os.path.splitext(name)[0] + ".log")
    try:
        log_before = os.path.getsize(log_src)
    except Exception:
        log_before = 0
    before = _snapshot_names(src_dir)
    reserved.add(os.path.normcase(os.path.abspath(target)))

    t0 = time.time()
    try:
        part, _status, method = open_part_file(session, path)
    except Exception as exc:
        _cleanup_side_files(src_dir, _new_side_files(src_dir, before))
        return _open_outcome("fail", path, target, started, str(exc))
    elapsed = time.time() - t0

    try:
        shown = session.Parts.Display
    except Exception:
        shown = part
    shown_path = ""
    if shown is not None:
        try:
            shown_path = os.path.normcase(os.path.abspath(shown.FullPath or ""))
        except Exception:
            shown_path = ""
    changed = shown is not None and shown_path not in ("", prev_path)
    info = _parse_translation_log(log_src, log_before)
    new_files = _new_side_files(src_dir, before)
    new_prt = [item for item in new_files if item.lower().endswith(".prt")]
    bodies = 0
    if shown is not None:
        try:
            bodies = len(list(shown.Bodies))
        except Exception:
            bodies = 0

    if not changed:
        ok, reason = False, imp_T("imp.err.not_displayed")
    elif info["errors"]:
        ok, reason = False, info["errors"][0]
    elif info["filed"] and info["ug_total"] > 0:
        ok, reason = True, ""
    elif info["found"] and (bodies > 0 or new_prt):
        ok, reason = True, ""
    elif not info["found"]:
        ok, reason = (bool(new_prt) or bodies > 0), imp_T("imp.err.no_geometry")
        if ok:
            reason = ""
    else:
        ok, reason = False, imp_T("imp.err.no_geometry")

    if not ok:
        if changed and shown is not None:
            try:
                close_part(session, shown)
            except Exception:
                pass
            if previous is not None:
                try:
                    activate_display(session, previous)
                except Exception:
                    pass
        _cleanup_side_files(src_dir, new_files)
        return _open_outcome("fail", path, target, started, reason)

    # 落盘：NX 已在源旁生成 <名>_step.prt；需要改名/换目录时 SaveAs
    current = ""
    try:
        current = os.path.abspath(shown.FullPath or "")
    except Exception:
        current = ""
    backup = ""
    if os.path.normcase(current) != os.path.normcase(os.path.abspath(target)):
        if mode == "overwrite" and os.path.exists(target):
            # 覆盖前备份为同名 .bak（已存在则覆盖）；备份失败即中止，原文件不动
            try:
                backup = imp_backup_original(target)
                os.remove(target)
            except Exception as exc:
                _cleanup_side_files(src_dir, new_files)
                return _open_outcome("fail", path, target, started, str(exc).strip())
            imp_log(imp_T("imp.log.backup", backup))
        try:
            shown.SaveAs(target)
            saved = file_written(target)
            save_reason = "" if saved else imp_T("imp.err.saveas_missing", target)
        except Exception as exc:
            saved = False
            save_reason = imp_T("imp.err.saveas", exc, target)
        if not saved:
            if backup:
                # 失败回滚：用 .bak 还原目标，.bak 本身保留不删
                try:
                    shutil.copy2(backup, target)
                except Exception:
                    pass
            _cleanup_side_files(src_dir, new_files)
            return _open_outcome("fail", path, target, started, save_reason)
        # 成功时不删 .bak：留给用户核对/清理

    # 中间 prt 可能仍被会话占用，先关掉再删，否则源目录会留下 <名>_stp.prt
    keep_name = os.path.basename(target)
    for extra in new_files:
        if not extra.lower().endswith((".prt",)):
            continue
        if extra.lower() == keep_name.lower():
            continue
        extra_path = os.path.join(src_dir, extra)
        try:
            stale = imp_find_open_part(session, extra_path)
        except Exception:
            stale = None
        if stale is not None and stale is not shown:
            try:
                close_part(session, stale)
            except Exception:
                pass
    _cleanup_side_files(src_dir, new_files, keep=(keep_name,))
    trace("SESSION_OPEN", "%s %s %.2fs -> %s" % (name, method, elapsed, target))
    fit_view(shown)
    return _open_outcome("ok", path, target, started, detail="saved", opened=True)


INTERACTIVE_PROGRESS_PS = r'''
param(
    [Parameter(Mandatory=$true)][string]$StateFile,
    [Parameter(Mandatory=$true)][string]$CancelFile,
    [Parameter(Mandatory=$true)][string]$CloseFile,
    [Parameter(Mandatory=$true)][string]$LabelsFile
)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$labels = [System.IO.File]::ReadAllText($LabelsFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json

$code = @"
using System;
using System.Windows.Forms;
public class NxNoActivateForm : Form {
    protected override bool ShowWithoutActivation { get { return true; } }
    protected override CreateParams CreateParams {
        get {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x08000000;
            cp.ExStyle |= 0x00000080;
            return cp;
        }
    }
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies @("System.Windows.Forms", "System.Drawing")

$form = New-Object NxNoActivateForm
$form.Text = [string]$labels.title
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.Size = New-Object System.Drawing.Size(380, 140)
$form.TopMost = $false
$form.ShowInTaskbar = $false
$area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(($area.Right - 400), ($area.Bottom - 170))

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Location = New-Object System.Drawing.Point(14, 12)
$lblTitle.Size = New-Object System.Drawing.Size(340, 20)
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblTitle.Text = [string]$labels.title
$form.Controls.Add($lblTitle)

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Location = New-Object System.Drawing.Point(14, 38)
$lblFile.Size = New-Object System.Drawing.Size(340, 20)
$lblFile.AutoEllipsis = $true
$form.Controls.Add($lblFile)

$lblStage = New-Object System.Windows.Forms.Label
$lblStage.Location = New-Object System.Drawing.Point(14, 60)
$lblStage.Size = New-Object System.Drawing.Size(240, 20)
$lblStage.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblStage)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Location = New-Object System.Drawing.Point(270, 84)
$btnCancel.Size = New-Object System.Drawing.Size(86, 26)
$btnCancel.Text = [string]$labels.cancel
$form.Controls.Add($btnCancel)

$btnCancel.add_Click({
    try { [System.IO.File]::WriteAllText($CancelFile, "1") } catch { }
    $btnCancel.Enabled = $false
    $btnCancel.Text = [string]$labels.cancelling
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 150
$timer.add_Tick({
    try {
        if ([System.IO.File]::Exists($CloseFile)) {
            $timer.Stop()
            $form.Close()
            return
        }
        if ([System.IO.File]::Exists($StateFile)) {
            $st = [System.IO.File]::ReadAllText($StateFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $lblTitle.Text = [string]$st.title
            $lblFile.Text = [string]$st.file
            $lblStage.Text = [string]$st.stage
        }
    } catch { }
})
$timer.Start()
$form.add_FormClosed({ try { $timer.Stop() } catch { } })
[void]$form.ShowDialog()
'''


def _start_interactive_progress():
    """起一个 WS_EX_NOACTIVATE 的进度小窗；失败返回 None（只写日志，不阻断导入）"""
    import json as _json
    import subprocess
    import tempfile
    try:
        powershell = imp__powershell_path()
        if powershell is None:
            return None
        folder = os.path.join(tempfile.gettempdir(),
                               "nx_import_open_%d_%d" % (os.getpid(), int(time.time())))
        os.makedirs(folder)
        ps_path = os.path.join(folder, "progress.ps1")
        labels_path = os.path.join(folder, "labels.json")
        state_path = os.path.join(folder, "state.json")
        cancel_path = os.path.join(folder, "cancel.flag")
        close_path = os.path.join(folder, "close.flag")
        labels = {
            "title": imp_T("imp.prog.title"),
            "cancel": imp_T("imp.prog.cancel"),
            "cancelling": imp_T("imp.prog.cancelling"),
        }
        with open(ps_path, "w", encoding="utf-8-sig") as handle:
            handle.write(INTERACTIVE_PROGRESS_PS)
        with open(labels_path, "w", encoding="utf-8") as handle:
            handle.write(_json.dumps(labels, ensure_ascii=False))
        with open(state_path, "w", encoding="utf-8") as handle:
            handle.write(_json.dumps({"title": labels["title"], "file": "", "stage": ""},
                                     ensure_ascii=False))
        command = [powershell, "-NoProfile", "-NonInteractive", "-STA",
                   "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
                   "-File", ps_path,
                   "-StateFile", state_path, "-CancelFile", cancel_path,
                   "-CloseFile", close_path, "-LabelsFile", labels_path]
        process = subprocess.Popen(command, stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL,
                                   creationflags=0x08000000)
        return {"proc": process, "dir": folder, "state": state_path,
                "cancel": cancel_path, "close": close_path, "labels": labels}
    except Exception as exc:
        trace("PROGRESS_START_FAIL", exc)
        return None


def _progress_write(progress, file_label, stage):
    if not progress:
        return
    import json as _json
    try:
        payload = {"title": progress["labels"]["title"],
                   "file": file_label or "", "stage": stage or ""}
        with open(progress["state"], "w", encoding="utf-8") as handle:
            handle.write(_json.dumps(payload, ensure_ascii=False))
    except Exception:
        pass


def _progress_stop(progress):
    if not progress:
        return
    try:
        with open(progress["close"], "w", encoding="utf-8") as handle:
            handle.write("1")
    except Exception:
        pass
    process = progress.get("proc")
    if process is not None:
        try:
            process.wait(timeout=2)
        except Exception:
            try:
                process.terminate()
            except Exception:
                pass


def run_interactive_import(session, jobs, log_path, listing, prefix="", suffix="",
                           outdir="", collision="skip"):
    """勾选「导入后打开」：当前会话逐件直读，边转边显示，最后汇总。

    返回 (成功数, 失败数, 跳过数, 取消剩余数)。
    """
    total = len(jobs)
    writable_cache = {}

    def _dir_ok(folder):
        key = os.path.normcase(os.path.abspath(folder))
        if key not in writable_cache:
            writable_cache[key] = _folder_writable(folder)
        return writable_cache[key]

    # 预检：不可写的目录在打开前就判失败，避免跑到一半才报错
    blocked = {}
    for item in jobs:
        path = item["src"]
        if os.path.splitext(path)[1].lower() in SUPPORTED_PRT:
            continue
        src_dir = os.path.dirname(os.path.abspath(path))
        out_dir = outdir or src_dir
        if not _dir_ok(src_dir):
            blocked[path] = imp_T("imp.err.src_readonly")
        elif not _dir_ok(out_dir):
            blocked[path] = imp_T("imp.err.out_readonly")
    if blocked:
        folders = sorted({os.path.dirname(os.path.abspath(path)) for path in blocked})
        body = imp_T("imp.err.readonly_body", "\n".join(folders))
        imp_log(body)
        message_box("Warning", imp_T("imp.err.readonly_title"), body)

    progress = _start_interactive_progress()
    reserved = set()
    ok_count = 0
    fail_count = len(blocked)
    skip_count = 0
    failures = [(path, reason) for path, reason in blocked.items()]
    started_all = time.time()
    canceled = False
    remaining = 0

    try:
        for index, item in enumerate(jobs, 1):
            path = item["src"]
            name = os.path.basename(path)
            _progress_write(progress, "%d/%d  %s" % (index, total, name),
                            imp_T("imp.prog.stage_convert"))
            if _cancel_requested(progress["cancel"] if progress else ""):
                canceled = True
                remaining = total - index + 1
                break
            if path in blocked:
                listing_write(listing, imp_T("imp.sum.fail_item", name, blocked[path]))
                continue
            try:
                outcome = open_one_in_session(
                    session, path, prefix=item.get("prefix", prefix),
                    suffix=item.get("suffix", suffix), outdir=outdir,
                    collision=collision, reserved=reserved)
            except Exception as exc:
                outcome = _open_outcome("fail", path, path, time.time(),
                                        str(exc) or traceback_text())
            _redraw_nx()
            if outcome["status"] == "ok":
                ok_count += 1
                if outcome["detail"] == "already_open":
                    imp_log(imp_T("imp.log.session_already", path))
                else:
                    imp_log(imp_T("imp.log.session_ok", path, outcome["seconds"]))
                listing_write(listing, imp_T("imp.log.session_ok", name, outcome["seconds"]))
                _progress_write(progress, "%d/%d  %s" % (index, total, name),
                                imp_T("imp.prog.stage_done"))
            elif outcome["status"] == "skip":
                skip_count += 1
                imp_log(imp_T("imp.log.session_skip", path))
                listing_write(listing, imp_T("imp.log.session_skip", name))
                _progress_write(progress, "%d/%d  %s" % (index, total, name),
                                imp_T("imp.prog.stage_skip"))
            else:
                fail_count += 1
                failures.append((path, outcome["reason"]))
                imp_log(imp_T("imp.log.fail", "%s: %s" % (path, outcome["reason"])))
                listing_write(listing, imp_T("imp.sum.fail_item", name, outcome["reason"]))
                _progress_write(progress, "%d/%d  %s" % (index, total, name),
                                imp_T("imp.prog.stage_fail"))
    finally:
        _progress_stop(progress)

    if canceled:
        imp_log(imp_T("imp.log.cancelled", remaining))

    lines = []
    for path, reason in failures[:10]:
        lines.append(imp_T("imp.sum.fail_item", os.path.basename(path), reason))
    if len(failures) > 10:
        lines.append(imp_T("imp.sum.more_fail", len(failures) - 10))
    if skip_count:
        lines.append(imp_T("imp.sum.skipped", skip_count))
    if not failures:
        lines.append(imp_T("imp.sum.no_fail"))
    lines.append(imp_T("imp.sum.interactive", ok_count, time.time() - started_all))
    if canceled:
        lines.append(imp_T("imp.sum.cancelled_note", remaining))
    summary = imp_T("imp.sum.body", ok_count, fail_count, skip_count,
                "\n".join(lines), log_path)
    message_box("Information", imp_T("imp.sum.title"), summary)
    return ok_count, fail_count, skip_count, remaining


def imp__load_runner():
    """公共后台执行器接口（原 nx_batch_runner.py 已并入本文件）"""
    return _runner_api()


imp__SHARED_RUNNER = None


def imp__shared():
    """公共执行器模块（缓存）：文件名清洗规则与 PS 片段都在那里，两窗共用一份"""
    imp__SHARED_RUNNER
    if imp__SHARED_RUNNER is None:
        imp__SHARED_RUNNER = imp__load_runner()
    return imp__SHARED_RUNNER


def imp_normalize_name(name):
    """把文件名中的全角括号等字符替换为英文符号（规则与导出窗同一份）"""
    return imp__shared().normalize_name(name)


def imp_safe_name_part(text):
    """前缀/后缀里 Windows 文件名不允许的字符替换成下划线（规则与导出窗同一份）"""
    return imp__shared().safe_name_part(text)


def imp__start_background(jobs, total, log_path, workers=2):
    """把清单交给独立 NX 批处理会话执行；失败返回 False，调用方回落前台循环"""
    try:
        runner = imp__load_runner()
        return runner.start("import", os.path.abspath(__file__), jobs, total,
                            None, log_path, imp_LANG,
                            extras={"workers": workers,
                                    "nx_pid": os.getpid()})
    except Exception as e:
        imp_log(imp_T("imp.log.background_fail", e))
        return False


LAST_IMPORT_PATH = os.path.join(tempfile.gettempdir(), "nx_batch_import_last.json")


def record_last_import(batch_id, paths):
    """把本次分片的产物合并进 %TEMP%\\nx_batch_import_last.json，供「打开上次导入结果」用"""
    import json

    if not paths:
        return
    lock_handle = None
    try:
        lock_handle = open(LAST_IMPORT_PATH + ".lock", "a+")
        try:
            import msvcrt
            msvcrt.locking(lock_handle.fileno(), msvcrt.LK_LOCK, 1)
        except Exception:
            pass
        data = {}
        try:
            with open(LAST_IMPORT_PATH, "r", encoding="utf-8-sig") as old:
                data = json.load(old) or {}
        except Exception:
            data = {}
        if str(data.get("batch") or "") == str(batch_id):
            files = list(dict.fromkeys(list(data.get("files") or []) + list(paths)))
        else:
            files = list(dict.fromkeys(paths))
        payload = {"batch": str(batch_id),
                   "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                   "count": len(files), "files": files}
        tmp = LAST_IMPORT_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as out:
            json.dump(payload, out, ensure_ascii=False, indent=2)
        os.replace(tmp, LAST_IMPORT_PATH)
        imp_log("[last] write=ok batch=%s count=%d" % (batch_id, len(files)))
    except Exception as exc:
        imp_log("[last] write=failed %s" % exc)
    finally:
        if lock_handle is not None:
            try:
                lock_handle.close()
            except Exception:
                pass


def imp__run_background_worker(plan_path):
    """独立 NX 会话里的执行入口：只跑清单，不弹任何窗"""
    imp__log_handle
    runner = imp__load_runner()
    plan = runner.load_plan(plan_path)
    try:
        imp__log_handle = open(plan.get("log_path"), "a", encoding="utf-8")
    except Exception:
        imp__log_handle = None
    session = NXOpen.Session.GetSession()
    ensure_multi_display(session)
    cleanup_old_logs()
    used = existing_part_names(session)
    total = plan.get("total") or len(plan.get("items") or [])
    imp_log(imp_T("imp.log.header", total, plan.get("log_path")))
    made = []

    def handler(item, report):
        index = item.get("index")
        path = item.get("src")
        imp_log("")
        imp_log(imp_T("imp.log.file", index, total, path))
        report("import")
        try:
            state, name, detail, _disk, part = import_one(
                session, path, used,
                prefix=str(item.get("prefix") or ""), suffix=str(item.get("suffix") or ""),
                outdir=str(item.get("outdir") or ""),
                collision=str(item.get("collision") or "rename"))
        except Exception as exc:
            reason = str(exc) or traceback_text()
            imp_log(imp_T("imp.log.fail", reason))
            return ("fail", reason)
        if state == "ok":
            try:
                keep_path = (path if os.path.splitext(path)[1].lower() in SUPPORTED_PRT
                             else part_file_on_disk(part))
                if keep_path and os.path.isfile(keep_path):
                    made.append(keep_path)
            except Exception:
                pass
            return ("ok", name)
        if state == "skip":
            return ("skip", detail)
        imp_log(imp_T("imp.log.fail", detail))
        return ("fail", detail)

    results, state = runner.execute(plan, handler, imp_log)
    failed = [r for r in results if r.get("status") == "fail"]
    imp_log("")
    imp_log(imp_T("imp.log.summary", state.get("ok", 0), len(failed), state.get("skip", 0)))
    batch_id = os.path.basename(os.path.dirname(os.path.abspath(plan_path)))
    record_last_import(batch_id, made)
    if imp__log_handle:
        try:
            imp__log_handle.close()
        except Exception:
            pass
        imp__log_handle = None


# ================= 主流程 =================

_STAGE_DIR = ""


def imp_main():
    """交互会话入口：菜单点完立即返回，选择/计划/起批处理都交给后台 run_journal 会话"""
    import json

    try:
        runner = imp__load_runner()
        if runner.activate_picker(imp_T("imp.ui.title")) == "ok":
            return
    except Exception:
        pass
    stage_dir = imp__stage_root()
    stage_path = os.path.join(stage_dir, "stage.json")
    payload = {"dir": stage_dir, "stage": stage_path,
               "lang": imp_LANG, "nx_pid": os.getpid()}
    try:
        with open(stage_path, "w", encoding="utf-8-sig") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
    except Exception as exc:
        _ui_error_log("stage write failed: %s" % exc)
        message_box("Error", imp_T("imp.confirm.title"), str(exc))
        return
    pid = imp__start_prepare_session(stage_path)
    if not pid:
        imp__notice_start_fail()
        return
    _ui_error_log("[async] prepare pid=%d stage=%s" % (pid, stage_dir))


def _import_flow():
    """后台准备会话内执行：选文件 → 生成计划 → 起批处理；不占交互前台"""
    # 日志句柄是模块级全局；这里必须声明，否则本函数内的 `_log_handle = None`
    # 会让它变成局部变量，收尾/后台分支读取时报 UnboundLocalError
    imp__log_handle

    trace("MAIN_START", "")
    session = NXOpen.Session.GetSession()
    log_path = open_log()
    trace("LOG_OPEN", log_path)
    if _STAGE_DIR:
        imp_log("[async] prepare start pid=%d stage=%s" % (os.getpid(), _STAGE_DIR))
    multi_status, multi_action = ensure_multi_display(session)
    trace("MULTI_DISPLAY", "status=%s action=%s after=%s"
          % (multi_status, multi_action, getattr(session.Parts, "AllowMultipleDisplayedParts", "?")))
    if "nx_batch_import_" in os.path.basename(log_path):
        imp_log(imp_T("imp.log.log_fallback", log_path))

    # 轻量清理：7 天前的 translator 日志；失败静默，绝不影响导入
    removed = cleanup_old_logs()
    if removed:
        imp_log(imp_T("imp.log.temp_clean_old", removed, TEMP_LOG_DAYS))

    listing = get_listing(session)
    trace("LISTING_OK", repr(listing))

    used_custom = False
    workers = read_workers()
    prefix, suffix = read_affixes()
    collision = read_collision()
    out_mode, outdir = read_output()
    try:
        trace("DIALOG_BEFORE", "")
        (picked, used_custom, workers, prefix, suffix,
         collision, out_mode, outdir) = pick_files_ui(
            imp_read_last_dir(), workers, prefix, suffix,
            collision, out_mode, outdir)
        trace("DIALOG_AFTER", "%d custom=%s workers=%d collision=%s out=%s"
              % (len(picked or []), used_custom, workers,
                 collision, out_mode))
    except Exception as exc:
        imp_log(imp_T("imp.err.dialog", exc))
        _ui_error_log("picker failed: %s" % exc)
        try:
            picked = select_files(imp_read_last_dir())
            used_custom = False
            trace("DIALOG_FALLBACK", "%d" % len(picked or []))
        except Exception as exc2:
            imp_log(imp_T("imp.err.dialog", exc2))
            message_box("Error", imp_T("imp.confirm.title"),
                        imp_T("imp.err.dialog", exc2))
            return
    if not picked:
        imp_log(imp_T("imp.msg.cancelled"))
        return

    try:
        # 记忆「最后一个加入的文件」所在目录与本次处理方式，下次打开回填
        imp_write_last_dir(os.path.dirname(os.path.abspath(picked[-1])))
        write_workers(workers)
        write_affixes(prefix, suffix)
        write_collision(collision)
        write_output(out_mode, outdir)
        trace("LAST_DIR_OK", "")
    except Exception:
        pass

    queue, skipped_all = dedupe(picked)
    for path, kind in skipped_all:
        if kind == "dup":
            imp_log(imp_T("imp.log.skip_dup", path))
        else:
            imp_log(imp_T("imp.log.skip_missing", path))
    if not queue:
        imp_log(imp_T("imp.msg.no_files"))
        message_box("Information", imp_T("imp.confirm.title"), imp_T("imp.msg.no_files"))
        return

    imp_log(imp_T("imp.log.header", len(queue), log_path))

    if used_custom:
        # 新选择窗里已选清单就是确认界面，不再二次弹框
        listing_write(listing, imp_T("imp.confirm.body", len(queue),
                                 "\n".join(os.path.basename(item) for item in queue)))
    else:
        preview = queue[:10]
        body = imp_T("imp.confirm.body", len(queue),
                 "\n".join(os.path.basename(item) for item in preview))
        if len(queue) > len(preview):
            body += imp_T("imp.confirm.more", len(queue) - len(preview))
        body += "\n\n" + imp_T("imp.confirm.proceed_log")

        listing_write(listing, imp_T("imp.confirm.body", len(queue),
                                "\n".join(preview)))
        if not confirm_yes(imp_T("imp.confirm.title"), body):
            imp_log(imp_T("imp.msg.cancelled"))
            return

    # 指定目录失效（旧记忆或窗口回落失败）时再用源目录
    if out_mode == "dir" and (not outdir or not os.path.isdir(outdir)):
        out_mode = "source"
        outdir = ""
    target_dir = outdir if out_mode == "dir" else ""

    jobs = [{"src": item, "index": index, "total": len(queue),
             "prefix": prefix, "suffix": suffix,
             "outdir": target_dir, "collision": collision}
            for index, item in enumerate(queue, 1)]

    # 转换一律交给独立 NX 批处理会话；转完不再自动显示，也不回落同步
    if imp__start_background(jobs, len(queue), log_path, workers):
        imp_log(imp_T("imp.log.background", len(queue)))
        imp_log("[async] background start total=%d" % len(queue))
        if imp__log_handle:
            try:
                imp__log_handle.close()
            except Exception:
                pass
            imp__log_handle = None
        return

    # 起不来就明确报错，绝不回落同步转换（同步会占住 NX 前台）
    imp_log("[async] background-failed stage=%s" % (_STAGE_DIR or "-"))
    if imp__log_handle:
        try:
            imp__log_handle.close()
        except Exception:
            pass
        imp__log_handle = None
    imp__notice_start_fail()


def imp__stage_root():
    path = os.path.join(
        tempfile.gettempdir(),
        "nx_batch_import_stage_%d_%d" % (os.getpid(), int(time.time() * 1000)))
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass
    return path


def imp__runner_exe():
    base = os.environ.get("UGII_BASE_DIR") or ""
    if base:
        exe = os.path.join(base, "NXBIN", "run_journal.exe")
        if os.path.isfile(exe):
            return exe
    return ""


def imp__start_prepare_session(stage_path):
    """起独立 NX 会话做选择/计划/起批处理；起不来返回 0（不回落同步）"""
    import subprocess

    exe = imp__runner_exe()
    if not exe:
        _ui_error_log("prepare spawn skipped: run_journal.exe not found")
        return 0
    env = os.environ.copy()
    env["NX_BATCH_STAGE"] = stage_path
    env.pop("NX_BATCH_PLAN", None)
    try:
        proc = subprocess.Popen(
            [exe, "-nx", os.path.abspath(__file__)], env=env,
            cwd=os.path.dirname(os.path.abspath(__file__)),
            creationflags=0x08000000)
    except Exception as exc:
        _ui_error_log("prepare spawn failed: %s" % exc)
        return 0
    return proc.pid


def imp__notice_start_fail():
    """后台会话起不来时给一次明确提示（不重试、不回落）"""
    text = imp_T("imp.err.background")
    _ui_error_log("[async] background start failed: %s" % text)
    try:
        ctypes.windll.user32.MessageBoxW(0, text, imp_T("imp.confirm.title"), 0x10)
    except Exception:
        try:
            message_box("Error", imp_T("imp.confirm.title"), text)
        except Exception:
            pass


def imp__run_prepare_stage(stage_path):
    """NX_BATCH_STAGE 会话入口：读 stage → 走原导入流程（选择窗/计划/起批处理）"""
    global _STAGE_DIR
    import json

    try:
        with open(stage_path, "r", encoding="utf-8-sig") as handle:
            stage = json.load(handle)
        _STAGE_DIR = stage.get("dir") or os.path.dirname(os.path.abspath(stage_path))
    except Exception:
        _STAGE_DIR = os.path.dirname(os.path.abspath(stage_path))
    _import_flow()


def traceback_text():
    try:
        import traceback
        return traceback.format_exc().strip().splitlines()[-1]
    except Exception:
        return "unknown error"


def _run_with_hider(entry):
    """所有入口统一：执行期间隐藏本进程的「工作进行中」框，结束/异常时恢复"""
    _waitdlg = None
    try:
        _waitdlg = imp__shared().start_waitdlg_hider(imp_log)
    except Exception:
        _waitdlg = None
    try:
        entry()
    finally:
        if _waitdlg is not None:
            try:
                _waitdlg.stop()
            except Exception:
                pass


def _import_api():
    """原 batch_import.py 的对外接口（合并后按命名空间暴露，调用点无需改动）"""
    return types.SimpleNamespace(
        existing_part_names=existing_part_names,
        import_one=import_one,
        close_part=close_part)

# ==============================================================================
# 段：批量导出 + 交互主流程（原 batch_export_step.py）
# ==============================================================================
# 批量把文件夹下所有 .prt 导出为 STEP (AP214)
# 
# 用法一（推荐，NX 界面内）：
#     NX 菜单栏左上角「菜单(Menu)」→ STEP Tools → Batch STEP Export...
#     （菜单位于本目录 startup\step_export.men，动作写的是裸文件名 format_convert.py；
#       NX 会把带盘符的绝对路径当成未知操作类型，所以脚本必须放在菜单所在的 startup 目录）
#     运行后弹出自定义选择窗口：左侧文件夹树、右侧 .prt 列表、下方已选清单，
#     右下角「确认导出」开始导出。窗口基于 PowerShell WinForms，打不开时自动回退系统对话框。
#     （也可用 NX 菜单 文件 → 执行 → NX Open 手动选中本脚本，效果相同）
# 用法二（命令行批处理，无界面）：
#     "C:\\Program Files\\Siemens\\NX 12.0\\NXBIN\\run_journal.exe" startup\format_convert.py
# 
# 注意：
#     1. 需要 STEP 转换器许可（Translator 授权），否则创建导出器会失败。
#     2. 为判断图层在当前视图里是否可见，脚本会临时打开并显示正在导出的部件；
#        导出结束只关闭脚本自己打开的部件，不会关闭你原本打开、正在编辑的部件，
#        并把显示部件恢复成你原来那个。
#     3. 每个 .prt 导出的 .step 就放在该 .prt 所在的那个文件夹里，
#        文件分散在不同子文件夹时，结果也分别留在各自文件夹。
#     4. 导出后会校验产物：文件存在、大于 0 字节且确实含几何才算成功；
#        为空或只有文件头会自动重试一次，仍无效则记为失败并删除该文件。
#        已存在但只有文件头（没有几何）的旧 .step 也会被重新导出，不会当成已完成跳过。
#     5. 只导出该部件当前工作视图中“可见”的图层（LayerManager.GetVisibilitiesInView），
#        隐藏图层上的实体会被排除。查询可见性要求部件是显示部件，因此脚本会临时
#        打开并显示该部件，导出后只关闭脚本自己打开的部件，并恢复你原本的显示部件。
#     6. NX 2506 起必须显式指定导出对象类型（实体/曲面/曲线）与图层范围，
#        否则 STEP 里只有文件头、没有几何，脚本已默认设置好。
# -*- coding: utf-8 -*-

import os
import glob
import json
import time
import stat
import shutil
import fnmatch
import traceback

import NXOpen

# ================= 多语言 =================
# UGII_LANG 的常见取值：simpl_chinese / trad_chinese / schinese / chinese / english
# NX 界面实测还会给出 schinese（简中）与 chinese（繁中），与 install.ps1 的映射保持一致。
# 未知或未设置时回落英文；三张表缺项时逐项回落英文。

_ZH_CN_ALIASES = frozenset((
    "simpl_chinese", "simplified_chinese", "zh_cn", "zh-cn", "chs",
    "schinese", "zh_hans",
))
_ZH_TW_ALIASES = frozenset((
    "trad_chinese", "traditional_chinese", "zh_tw", "zh-tw", "cht",
    "chinese", "zh_hant",
))


def _detect_lang():
    value = (os.environ.get("UGII_LANG") or "").strip().lower()
    # 精确成员判断：chinese 不能吞掉 simpl_chinese / trad_chinese
    if value in _ZH_CN_ALIASES:
        return "zh_cn"
    if value in _ZH_TW_ALIASES:
        return "zh_tw"
    return "en"


LANG = _detect_lang()

MSG = {
    "en": {
        "log.ignore_outside": "  Ignored (outside root): {0}",
        "log.already_included": "  Already included, skipped: {0}",
        "log.selected": "  Selected: {0}",
        "log.root_all": "<Root: all>",
        "log.rename_conflict": "Name conflict: {0} and {1} both map to {2}, using {3}",
        "log.export_from_fail": "      ExportFrom failed (ignored): {0}",
        "log.objecttypes_fail": "      Object types failed (ignored): {0}",
        "log.layermask_fail": "      Layer mask failed (ignored): {0}",
        "log.holdflag_fail": "      ProcessHoldFlag failed (ignored): {0}",
        "log.retry_deleted": "      Removed previous invalid output, retrying",
        "log.retry_delete_fail": "      Failed to remove invalid output (ignored): {0}",
        "log.retry2_ok": "      Attempt 2 succeeded, wrote {0} bytes",
        "log.invalid_attempt": "      Attempt {0} invalid: {1}",
        "log.cleanup_fail": "      Failed to clean invalid output (ignored): {0}",
        "log.folder_dialog_fail": "Folder dialog unavailable ({0})",
        "log.file_dialog_fail": "File dialog unavailable ({0})",
        "log.opening_picker": "Opening selection window...",
        "log.picker_fail": "Selection window unavailable ({0}), falling back to system dialogs",
        "log.picked_folders": "Selected folders: {0}",
        "log.picked_root": "Root folder selected: export everything.",
        "log.picked_files": "Selected {0} file(s):",
        "log.no_selection": "Nothing selected, cancelled, no file exported.",
        "log.no_prt_reread": "No .prt found in the selected scope, please choose again.",
        "log.preview_fail": "Preview window unavailable ({0}), exporting by the plan directly",
        "log.back": "Back to modify, reopening the selection window.",
        "log.preview_cancel": "Cancelled at the preview list, no file exported.",
        "log.src_dir": "Source : {0}",
        "log.src_any": "Source : any locations you selected (folders / files)",
        "log.output_line": "Output : {0} next to each .prt (in place)",
        "log.file_count": "Files  : {0}",
        "log.include_dirs": "Include folders: {0}",
        "log.exclude_dirs": "Exclude folders: {0}",
        "log.include_files": "Include files  : {0}",
        "log.exclude_files": "Exclude files  : {0}",
        "log.manual_files": "Manually selected: {0}",
        "log.no_files_end": "No matching .prt found, done.",
        "log.dryrun_hint": "(Hint: in interactive mode the preview list replaces DRY_RUN)",
        "log.dryrun_title": "Dry run: the following files would be exported (not executed)",
        "log.dryrun_skip_mark": "  [exists, would be skipped]",
        "log.modified_header": "{0} selected file(s) are open with unsaved changes:",
        "log.cancel_unsaved": "Export cancelled (nothing saved, no file exported).",
        "log.user_skip_mod": "  Skipped by user (unsaved changes): {0}",
        "log.saved_ok": "  Saved (user confirmed): {0}",
        "log.save_fail_skip": "  Save failed, skipping instead: {0} ({1})",
        "log.noninteractive_skip": "  Non-interactive mode, skipped by default: {0}",
        "reason.user_skip": "Skipped (unsaved changes, user chose to skip)",
        "reason.save_fail": "Skipped (unsaved changes, save failed)",
        "reason.noninteractive": "Skipped (unsaved changes, non-interactive default)",
        "log.skip_exists": "{0} Skipped (already exists): {1}",
        "log.no_visible": "      Failed: no visible layer in view {0}, not exported",
        "log.visible": "      Visible layers: {0} (view {1}, {2} layer(s))",
        "log.cover": "{0} Overwrite export (same name already exists): {1}",
        "log.backup": "{0} Original backed up as: {1}",
        "log.err_backup": "Backup failed, overwrite aborted: {0} -> {1} ({2})",
        "log.export_start": "{0} Export: {1}",
        "log.progress": "Exporting {0}/{1}: {2}",
        "log.done": "      Done {0:.1f}s -> {1} ({2} bytes)",
        "log.failed": "      Failed: {0}",
        "log.close_fail": "  Failed to close a temporarily opened part (ignored): {0}",
        "log.keep_modified": "  Kept open (unsaved changes): {0}",
        "log.display_restored": "  Original display part restored: {0}",
        "log.display_restore_fail": "  Failed to restore display part: {0}",
        "log.work_restored": "  Original work part restored: {0}",
        "log.work_restore_fail": "  Failed to restore work part: {0}",
        "log.summary": "Succeeded {0}, skipped {1}, failed {2}",
        "log.background": "Started {0} job(s) in background sessions; you can keep working in NX.",
        "log.background_fail": "Background session unavailable ({0}), running in the foreground.",
        "log.failed_item": "  Failed: {0}",
        "log.err_no_powershell": "powershell.exe not found",
        "log.err_no_result": "The window returned no result: {0}",
        "log.ui_hidden": "UI window '{0}' still invisible after 8s (pid {1}); it may have been launched with SW_HIDE.",
        "log.err_no_dir": "Folder does not exist: {0}",
        "log.err_creator": "Cannot create the STEP exporter (make sure the STEP translator is installed and licensed): {0}",
        "log.err_method_missing": "{0} not available",
        "log.err_method_fail": "{0} failed: {1}",
        "log.err_open_part": "Cannot open the part: {0} — {1}",
        "log.err_display_none": "Failed to switch the display part (NX returned none): {0}",
        "log.err_out_missing": "output file does not exist",
        "log.err_out_empty": "output file is 0 bytes",
        "log.err_out_header": "output file has only the header and no geometry ({0} bytes)",
        "log.err_export_invalid": "STEP export produced no valid content ({0})",
        "log.delete_old_fail": "      Failed to delete the old file (ignored): {0}",
        "log.err_remove_target": "Cannot overwrite {0}: the {2} file is read-only or open in another program (Excel/Notepad). Close it or clear read-only and retry ({1})",
        "log.missing_part": "File no longer exists (deleted or moved after selection): {0}",
        "log.busy": "Another export is already running; this run was not started (to avoid two windows changing the same files).",
        "log.stale_lock": "Removed the stale lock left by PID {0} and retried",
        "log.lock_unknown": "Cannot verify whether lock PID {0} still exists; treating the lock as busy",
        "log.log_fallback": "The log file is in use; this run writes its log to: {0}",
        "log.unsupported_name": "File name contains characters NX cannot handle (e.g., emoji); rename the file and retry: {0}",
        "log.too_long_path": "Path exceeds the Windows 260-character limit; shorten the folder depth and retry: {0}",
        "log.walk_fail": "Cannot access a folder (path too long or permission denied); skipped: {0}",
        "ui.folder_pick_title": "Select the folder to export as STEP (Cancel = finish)",
        "ui.file_pick_title": "Select the .prt / .stp / .igs files to export or convert",
        "ui.filter_prt": "NX parts and STEP/IGES (*.prt;*.stp;*.step;*.igs;*.iges)",
        "ui.filter_all": "All files (*.*)",
        "ui.picker_title": "Select folders / files to export or convert",
        "ui.picker_this_pc": "This PC",
        "ui.picker_quick_desktop": "Desktop",
        "ui.picker_quick_downloads": "Downloads",
        "ui.picker_quick_documents": "Documents",
        "ui.picker_quick_videos": "Videos",
        "ui.picker_tree_label": "Folders",
        "ui.picker_list_label": "Files (Ctrl/Shift multi-select, double-click to add)",
        "ui.picker_empty_dir": "(No files convertible in the current direction)",
        "ui.picker_denied": "Cannot access this folder (permission denied or invalid path):",
        "ui.picker_denied_hint": "Cannot access. Choose another folder or contact your administrator.",
        "ui.picker_col_name": "File name",
        "ui.picker_col_mtime": "Date modified",
        "ui.picker_col_size": "Size (KB)",
        "ui.picker_col_type": "Type",
        "ui.picker_col_path": "Path",
        "ui.type_dir": "Folder",
        "ui.type_file": "File",
        "ui.btn_add_folder": "+ Add folder",
        "ui.btn_add_file": "Add export files",
        "ui.btn_clear": "Clear",
        "ui.btn_del_rows": "Delete selected",
        "ui.menu_delete": "Delete selected",
        "ui.menu_select_all": "Select all",
        "ui.menu_clear": "Clear",
        "ui.btn_cancel": "Cancel",
        "ui.btn_ok": "OK",
        "ui.busy_title": "Export in progress",
        "ui.busy_text": "An export is already running. Try again later.",
        "ui.busy_ok": "OK",
        "ui.msg_remove_title": "Confirm removal",
        "ui.msg_remove_text": "Remove this item from the list?",
        "ui.msg_empty_title": "Notice",
        "ui.msg_empty_one": "Nothing has been added yet.",
        "ui.msg_empty_two": "Select a folder or file, then click Add.",
        "ui.preview_title": "Confirm files to convert",
        "ui.same_summary": "Same-name policy: {0} (set in the previous step; go back to change)",
        "ui.preview_hint": "Checked = replace existing files.",
        "ui.col_cover": "Overwrite",
        "ui.col_no": "No.",
        "ui.col_src": "Source",
        "ui.col_dst": "Output",
        "ui.col_state": "Status",
        "ui.state_new": "New",
        "ui.state_skip": "Skipped (exists)",
        "ui.state_overwrite": "Overwrite",
        "ui.chk_all": "Check all overwrite",
        "ui.btn_back": "Back",
        "ui.btn_start": "Start export",
        "ui.summary": "{0} file(s): new/overwrite {1}, skipped {2}",
        "ui.save_title": "Save current changes?",
        "ui.save_hint": "Save current changes?",
        "ui.save_item_prefix": "Save these changes: ",
        "ui.btn_cancel_export": "Cancel this export",
        "ui.btn_save_go": "Continue export",
        "ui.save_empty_confirm": "Nothing is checked. Skip unsaved changes in {0} file(s) and continue?",
        "ui.out_group": "Output folder",
        "ui.out_src": "Output to source folder",
        "ui.out_dir": "Output to a specified folder",
        "ui.out_browse": "Browse...",
        "ui.name_prefix": "Add prefix",
        "ui.name_suffix": "Add suffix",
        "ui.name_group": "Name and format",
        "ui.on_exists": "If the target exists",
        "ui.on_exists_skip": "Skip",
        "ui.on_exists_rename": "Rename (_2, _3 ...)",
        "ui.on_exists_overwrite": "Overwrite",
        "ui.format": "Format",
        "ui.format_step": "STEP (.step)",
        "ui.format_iges": "IGES (.igs)",
        "ui.state_rename": "Rename",
        "ui.state_self": "Same format (skipped)",
        "ui.dir_group": "Direction",
        "ui.dir_step": "to STEP (prt / igs sources)",
        "ui.dir_iges": "to IGES (prt / stp sources)",
        "ui.dir_prt_step": "prt→stp",
        "ui.dir_prt_iges": "prt→igs",
        "ui.dir_to_prt": "stp/igs→prt",
        "ui.dir_xt_prt": "x_t/x_b→prt",
        "ui.dir_any": "any→stp/igs",
        "ui.dir_xt_step": "x_t/x_b→stp",
        "ui.dir_xt_iges": "x_t/x_b→igs",
        "ui.right_title": "Converted name",
        "ui.col_file": "File name",
        "ui.left_empty": "(Double-click a file above to add)",
        "ui.right_empty": "(Shown after files are added)",
        "ui.folder_expand": "(Folder: export the .prt inside)",
        "ui.workers_label": "Concurrent processes",
        "ui.open_after": "Open converted files",
        "ui.preview_dup": " (duplicate, renamed)",
        "ui.preview_rename": " (renamed)",
        "ui.preview_exists": " (skipped)",
        "ui.preview_overwrite": " (overwrite)",
        "log.chain_same_fmt": "Source and target are the same format; skipped: {0}",
        "log.import_unavailable": "Import module unavailable: {0}",
        "log.import_fail": "Convert to PRT failed: {0} ({1})",
        "log.import_done": "Converted to PRT: {0} -> {1} {2}",
        "log.chain_begin": "Converting in memory: {0} -> {1}",
        "log.chain_import_fail": "Chain import failed: {0}",
        "log.chain_close_fail": "Could not close the temporary in-memory part: {0}",
        "log.chain_restore_fail": "Could not restore the previous display part: {0}",
        "log.err_step_from_part": "StepCreator does not support DisplayPart/WorkPart export; in-memory STEP export is unavailable",
        "log.err_iges_from_part": "IgesCreator does not support DisplayPart/WorkPart export; in-memory IGES export is unavailable",
        "log.out_dir_invalid": "Output folder is empty or does not exist, falling back to the source folder: {0}",
        "log.output_custom": "Output : {0} (flat, source folder structure not mirrored)",
        "log.format_line": "Format : {0}",
        "log.rename_policy": "Target exists, renamed: {0} -> {1}",
        "log.iges_exe_missing": "IGES translator iges.exe not found: {0}",
        "log.iges_def_fallback": "igesexport.def not found in the NX install, falling back to: {0}",
        "log.iges_timeout": "IGES translation timed out after {0}s and was terminated",
        "log.iges_exit": "IGES translation failed (exit code {0})",
        "log.iges_log_unread": "      (cannot read the IGES log: {0})",
        "log.iges_cmd_fail": "Cannot start the IGES translator: {0}",
        "log.err_iges_out": "IGES output is missing or 0 bytes: {0}",
        "log.err_iges_geom": "IGES output has no geometry (header and properties only): {0}",
        "log.err_target_locked": "Cannot replace the target file (open in another program, or still read-only): {0}",
    },
    "zh_cn": {
        "log.ignore_outside": "  已忽略（不在根目录内）: {0}",
        "log.already_included": "  已包含，跳过: {0}",
        "log.selected": "  已选择: {0}",
        "log.root_all": "<根目录：全部>",
        "log.rename_conflict": "重名冲突: {0} 与 {1} 转名后同为 {2}，改用 {3}",
        "log.export_from_fail": "      ExportFrom 设置失败（忽略）: {0}",
        "log.objecttypes_fail": "      对象类型设置失败（忽略）: {0}",
        "log.layermask_fail": "      图层范围设置失败（忽略）: {0}",
        "log.holdflag_fail": "      ProcessHoldFlag 设置失败（忽略）: {0}",
        "log.retry_deleted": "      已删除上次的无效文件，准备重试",
        "log.retry_delete_fail": "      删除无效文件失败（忽略）: {0}",
        "log.retry2_ok": "      第 2 次尝试成功，输出 {0} 字节",
        "log.invalid_attempt": "      第 {0} 次导出无效：{1}",
        "log.cleanup_fail": "      清理无效文件失败（忽略）: {0}",
        "log.folder_dialog_fail": "文件夹选择框不可用（{0}）",
        "log.file_dialog_fail": "文件选择框不可用（{0}）",
        "log.opening_picker": "正在打开选择窗口…",
        "log.picker_fail": "选择窗口不可用（{0}），回退到系统对话框",
        "log.picked_folders": "已选文件夹: {0}",
        "log.picked_root": "已选择根目录：导出全部。",
        "log.picked_files": "已选文件 {0} 个:",
        "log.no_selection": "未选择任何文件夹或文件，已取消，未导出任何文件。",
        "log.no_prt_reread": "所选范围里没有找到 .prt 文件，请重新选择。",
        "log.preview_fail": "预览窗口不可用（{0}），直接按清单导出",
        "log.back": "返回修改，重新打开选择窗口。",
        "log.preview_cancel": "已在预览清单处取消，未导出任何文件。",
        "log.src_dir": "源目录 : {0}",
        "log.src_any": "源目录 : 你选择的任意位置（文件夹 / 文件）",
        "log.output_line": "输出    : 与各 .prt 同目录（就地生成 {0}）",
        "log.file_count": "文件数  : {0}",
        "log.include_dirs": "仅包含文件夹: {0}",
        "log.exclude_dirs": "排除文件夹  : {0}",
        "log.include_files": "仅包含文件  : {0}",
        "log.exclude_files": "排除文件    : {0}",
        "log.manual_files": "手动指定文件: {0} 个",
        "log.no_files_end": "没有找到符合条件的 .prt 文件，结束。",
        "log.dryrun_hint": "（提示：交互模式下选完会先弹清单预览，不需要改 DRY_RUN）",
        "log.dryrun_title": "试运行模式：以下文件将被导出（未实际执行）",
        "log.dryrun_skip_mark": "  [已存在，会跳过]",
        "log.modified_header": "检测到 {0} 个已打开且有未保存修改的文件：",
        "log.cancel_unsaved": "已取消本次导出（未保存修改未保存，未导出任何文件）。",
        "log.user_skip_mod": "  用户选择跳过（未保存修改）: {0}",
        "log.saved_ok": "  已保存（用户确认）: {0}",
        "log.save_fail_skip": "  保存失败，改为跳过: {0} ({1})",
        "log.noninteractive_skip": "  非交互模式，按默认跳过（未保存修改）: {0}",
        "reason.user_skip": "跳过（未保存修改，用户选择跳过）",
        "reason.save_fail": "跳过（未保存修改，保存失败）",
        "reason.noninteractive": "跳过（未保存修改，非交互模式默认跳过）",
        "log.skip_exists": "{0} 跳过（已存在）: {1}",
        "log.no_visible": "      失败: 当前视图({0})没有可见图层，未导出",
        "log.visible": "      可见图层: {0}（视图 {1}，共 {2} 层）",
        "log.cover": "{0} 覆盖导出（已存在同名文件）: {1}",
        "log.backup": "{0} 原文件已备份为: {1}",
        "log.err_backup": "备份失败，已中止覆盖：{0} → {1}（{2}）",
        "log.export_start": "{0} 导出: {1}",
        "log.progress": "正在导出 {0}/{1}: {2}",
        "log.done": "      完成 {0:.1f}s -> {1} ({2} 字节)",
        "log.failed": "      失败: {0}",
        "log.close_fail": "  关闭临时打开的部件失败（忽略）: {0}",
        "log.keep_modified": "  保持打开（有未保存修改）: {0}",
        "log.display_restored": "  已切回原显示部件: {0}",
        "log.display_restore_fail": "  切回原显示部件失败: {0}",
        "log.work_restored": "  已恢复原工作部件: {0}",
        "log.work_restore_fail": "  恢复原工作部件失败: {0}",
        "log.summary": "成功 {0} 个，跳过 {1} 个，失败 {2} 个",
        "log.background": "已在后台启动 {0} 个任务，NX 可继续操作。",
        "log.background_fail": "后台会话不可用（{0}），改为前台执行。",
        "log.failed_item": "  失败: {0}",
        "log.err_no_powershell": "找不到 powershell.exe",
        "log.err_no_result": "窗口没有返回结果: {0}",
        "log.ui_hidden": "界面窗口「{0}」启动 8 秒仍不可见（pid {1}），疑似被 SW_HIDE 启动。",
        "log.err_no_dir": "目录不存在: {0}",
        "log.err_creator": "无法创建 STEP 导出器（请确认已安装 STEP 转换器并具备许可）: {0}",
        "log.err_method_missing": "{0} 不存在",
        "log.err_method_fail": "{0} 失败: {1}",
        "log.err_open_part": "无法打开部件: {0} — {1}",
        "log.err_display_none": "切换显示部件失败（NX 未返回显示部件）: {0}",
        "log.err_out_missing": "输出文件不存在",
        "log.err_out_empty": "输出文件为 0 字节",
        "log.err_out_header": "输出文件只有文件头、没有几何（{0} 字节）",
        "log.err_export_invalid": "STEP 导出没有产出有效内容（{0}）",
        "log.delete_old_fail": "      删除旧文件失败（忽略）: {0}",
        "log.err_remove_target": "无法覆盖 {0}：该 {2} 文件只读或正被其他程序占用（如 Excel/记事本），请关闭或取消只读后重试（{1}）",
        "log.missing_part": "文件不存在（可能在选择后被删除或移动）: {0}",
        "log.busy": "已有一次导出正在进行，本次未启动（避免两个窗口同时改同一批文件）。",
        "log.stale_lock": "已删除 PID {0} 遗留的陈旧锁，并重试本次导出",
        "log.lock_unknown": "无法确认锁内进程 {0} 是否仍存在，按占用处理",
        "log.log_fallback": "日志文件被占用，本次日志改写到: {0}",
        "log.unsupported_name": "文件名包含 NX 无法识别的特殊字符（如 emoji），请改名后重试: {0}",
        "log.too_long_path": "路径超过 Windows 260 字符限制，NX 无法访问；请缩短目录层级后重试: {0}",
        "log.walk_fail": "无法访问某个目录（路径过长或权限不足），已跳过: {0}",
        "ui.folder_pick_title": "选择要导出 STEP 的文件夹（取消 = 结束选择）",
        "ui.file_pick_title": "选择要导出/转换的 .prt / .stp / .igs 文件",
        "ui.filter_prt": "NX 部件与 STEP/IGES (*.prt;*.stp;*.step;*.igs;*.iges)",
        "ui.filter_all": "所有文件 (*.*)",
        "ui.picker_title": "选择要导出/转换的文件夹 / 文件",
        "ui.picker_this_pc": "此电脑",
        "ui.picker_quick_desktop": "桌面",
        "ui.picker_quick_downloads": "下载",
        "ui.picker_quick_documents": "文档",
        "ui.picker_quick_videos": "视频",
        "ui.picker_tree_label": "文件夹",
        "ui.picker_list_label": "文件（Ctrl/Shift 多选，双击加入）",
        "ui.picker_empty_dir": "（该目录下没有当前方向可转换的文件）",
        "ui.picker_denied": "无法访问该目录（权限不足或路径无效）：",
        "ui.picker_denied_hint": "无法访问，请换一个目录或联系管理员。",
        "ui.picker_col_name": "文件名称",
        "ui.picker_col_mtime": "修改日期",
        "ui.picker_col_size": "大小(KB)",
        "ui.picker_col_type": "类型",
        "ui.picker_col_path": "路径",
        "ui.type_dir": "文件夹",
        "ui.type_file": "文件",
        "ui.btn_add_folder": "＋ 添加文件夹",
        "ui.btn_add_file": "添加导出文件",
        "ui.btn_clear": "清空",
        "ui.btn_del_rows": "删除选中",
        "ui.menu_delete": "删除选中",
        "ui.menu_select_all": "全选",
        "ui.menu_clear": "清空",
        "ui.btn_cancel": "取消",
        "ui.btn_ok": "确定",
        "ui.busy_title": "导出正在进行",
        "ui.busy_text": "已有导出正在进行，请稍后再试。",
        "ui.busy_ok": "确定",
        "ui.msg_remove_title": "确认移除",
        "ui.msg_remove_text": "确定要从已选清单中移除吗？",
        "ui.msg_empty_title": "提示",
        "ui.msg_empty_one": "尚未添加任何文件。",
        "ui.msg_empty_two": "请先选中文件夹或文件，再点「添加」。",
        "ui.preview_title": "确认转换文件",
        "ui.same_summary": "同名策略：{0}（上一步已设定，修改请返回）",
        "ui.preview_hint": "勾选即替换同名文件。",
        "ui.col_cover": "覆盖",
        "ui.col_no": "序号",
        "ui.col_src": "源文件",
        "ui.col_dst": "输出",
        "ui.col_state": "状态",
        "ui.state_new": "新建",
        "ui.state_skip": "跳过（已存在）",
        "ui.state_overwrite": "覆盖",
        "ui.chk_all": "全选覆盖",
        "ui.btn_back": "返回",
        "ui.btn_start": "开始导出",
        "ui.summary": "{0} 个文件：新建/覆盖 {1}，跳过 {2}",
        "ui.save_title": "是否保存当前修改",
        "ui.save_hint": "是否保存当前修改？",
        "ui.save_item_prefix": "保存这些修改：",
        "ui.btn_cancel_export": "取消本次导出",
        "ui.btn_save_go": "继续导出",
        "ui.save_empty_confirm": "未勾选任何文件，将跳过 {0} 个文件的修改继续导出？",
        "ui.out_group": "输出目录",
        "ui.out_src": "输出到源文件目录",
        "ui.out_dir": "输出到指定目录",
        "ui.out_browse": "浏览…",
        "ui.name_prefix": "文件添加前缀",
        "ui.name_suffix": "文件添加后缀",
        "ui.name_group": "名称与格式",
        "ui.on_exists": "同名文件已存在时",
        "ui.on_exists_skip": "跳过",
        "ui.on_exists_rename": "改名（_2、_3…）",
        "ui.on_exists_overwrite": "覆盖",
        "ui.format": "格式",
        "ui.format_step": "STEP (.step)",
        "ui.format_iges": "IGES (.igs)",
        "ui.state_rename": "改名",
        "ui.state_self": "同格式（跳过）",
        "ui.dir_group": "转换方向",
        "ui.dir_step": "→ STEP（源 prt / igs）",
        "ui.dir_iges": "→ IGES（源 prt / stp）",
        "ui.dir_prt_step": "prt→stp",
        "ui.dir_prt_iges": "prt→igs",
        "ui.dir_to_prt": "stp/igs→prt",
        "ui.dir_xt_prt": "x_t/x_b→prt",
        "ui.dir_any": "任意→stp/igs",
        "ui.dir_xt_step": "x_t/x_b→stp",
        "ui.dir_xt_iges": "x_t/x_b→igs",
        "ui.right_title": "转换后名称",
        "ui.col_file": "文件名称",
        "ui.left_empty": "（双击上方文件加入）",
        "ui.right_empty": "（添加文件后显示）",
        "ui.folder_expand": "（文件夹：导出其中的 .prt）",
        "ui.workers_label": "同时执行数",
        "ui.open_after": "打开转换后的文件",
        "ui.preview_dup": "（重名，改名）",
        "ui.preview_rename": "（改名）",
        "ui.preview_exists": "（跳过）",
        "ui.preview_overwrite": "（覆盖）",
        "log.chain_same_fmt": "源与目标格式相同，已跳过：{0}",
        "log.import_unavailable": "导入模块不可用：{0}",
        "log.import_fail": "转 PRT 失败：{0}（{1}）",
        "log.import_done": "已转 PRT：{0} → {1} {2}",
        "log.chain_begin": "内存中转转换：{0} → {1}",
        "log.chain_import_fail": "链式导入失败：{0}",
        "log.chain_close_fail": "无法关闭临时内存部件：{0}",
        "log.chain_restore_fail": "无法恢复原来的显示部件：{0}",
        "log.err_step_from_part": "StepCreator 不支持“显示/工作部件”导出，无法做内存 STEP 导出",
        "log.err_iges_from_part": "IgesCreator 不支持“显示/工作部件”导出，无法做内存 IGES 导出",
        "log.out_dir_invalid": "指定输出目录为空或不存在，已回落到原文件所在路径: {0}",
        "log.output_custom": "输出    : {0}（扁平放置，不镜像原目录结构）",
        "log.format_line": "格式    : {0}",
        "log.rename_policy": "同名已存在，改名输出: {0} -> {1}",
        "log.iges_exe_missing": "找不到 IGES 转换器 iges.exe: {0}",
        "log.iges_def_fallback": "NX 安装目录下未找到 igesexport.def，回落到: {0}",
        "log.iges_timeout": "IGES 转换超时（{0} 秒），已终止",
        "log.iges_exit": "IGES 转换失败（exit code {0}）",
        "log.iges_log_unread": "      （无法读取 IGES 日志: {0}）",
        "log.iges_cmd_fail": "无法启动 IGES 转换器: {0}",
        "log.err_iges_out": "IGES 输出文件不存在或为 0 字节: {0}",
        "log.err_iges_geom": "IGES 输出不含几何（只有文件头与属性，没有曲线/曲面/实体）: {0}",
        "log.err_target_locked": "无法替换目标文件（被其他程序占用，或仍为只读）: {0}",
    },
    "zh_tw": {
        "log.ignore_outside": "  已忽略（不在根目錄內）: {0}",
        "log.already_included": "  已包含，略過: {0}",
        "log.selected": "  已選擇: {0}",
        "log.root_all": "<根目錄：全部>",
        "log.rename_conflict": "名稱衝突: {0} 與 {1} 轉名後同為 {2}，改用 {3}",
        "log.export_from_fail": "      ExportFrom 設定失敗（忽略）: {0}",
        "log.objecttypes_fail": "      物件類型設定失敗（忽略）: {0}",
        "log.layermask_fail": "      圖層範圍設定失敗（忽略）: {0}",
        "log.holdflag_fail": "      ProcessHoldFlag 設定失敗（忽略）: {0}",
        "log.retry_deleted": "      已刪除上次的無效檔案，準備重試",
        "log.retry_delete_fail": "      刪除無效檔案失敗（忽略）: {0}",
        "log.retry2_ok": "      第 2 次嘗試成功，輸出 {0} 位元組",
        "log.invalid_attempt": "      第 {0} 次匯出無效：{1}",
        "log.cleanup_fail": "      清理無效檔案失敗（忽略）: {0}",
        "log.folder_dialog_fail": "資料夾選擇框無法使用（{0}）",
        "log.file_dialog_fail": "檔案選擇框無法使用（{0}）",
        "log.opening_picker": "正在開啟選擇視窗…",
        "log.picker_fail": "選擇視窗無法使用（{0}），改回系統對話框",
        "log.picked_folders": "已選資料夾: {0}",
        "log.picked_root": "已選擇根目錄：全部匯出。",
        "log.picked_files": "已選檔案 {0} 個:",
        "log.no_selection": "未選擇任何資料夾或檔案，已取消，未匯出任何檔案。",
        "log.no_prt_reread": "所選範圍內找不到 .prt 檔案，請重新選擇。",
        "log.preview_fail": "預覽視窗無法使用（{0}），直接依清單匯出",
        "log.back": "返回修改，重新開啟選擇視窗。",
        "log.preview_cancel": "已在預覽清單處取消，未匯出任何檔案。",
        "log.src_dir": "來源目錄 : {0}",
        "log.src_any": "來源目錄 : 您選擇的任意位置（資料夾 / 檔案）",
        "log.output_line": "輸出    : 與各 .prt 同目錄（就地產生 {0}）",
        "log.file_count": "檔案數  : {0}",
        "log.include_dirs": "僅包含資料夾: {0}",
        "log.exclude_dirs": "排除資料夾  : {0}",
        "log.include_files": "僅包含檔案  : {0}",
        "log.exclude_files": "排除檔案    : {0}",
        "log.manual_files": "手動指定檔案: {0} 個",
        "log.no_files_end": "找不到符合條件的 .prt 檔案，結束。",
        "log.dryrun_hint": "（提示：互動模式下選完會先彈出清單預覽，不需要改 DRY_RUN）",
        "log.dryrun_title": "試執行模式：以下檔案將被匯出（未實際執行）",
        "log.dryrun_skip_mark": "  [已存在，會略過]",
        "log.modified_header": "偵測到 {0} 個已開啟且有未儲存修改的檔案：",
        "log.cancel_unsaved": "已取消本次匯出（未儲存修改未儲存，未匯出任何檔案）。",
        "log.user_skip_mod": "  使用者選擇略過（未儲存修改）: {0}",
        "log.saved_ok": "  已儲存（使用者確認）: {0}",
        "log.save_fail_skip": "  儲存失敗，改為略過: {0} ({1})",
        "log.noninteractive_skip": "  非互動模式，依預設略過（未儲存修改）: {0}",
        "reason.user_skip": "略過（未儲存修改，使用者選擇略過）",
        "reason.save_fail": "略過（未儲存修改，儲存失敗）",
        "reason.noninteractive": "略過（未儲存修改，非互動模式預設略過）",
        "log.skip_exists": "{0} 略過（已存在）: {1}",
        "log.no_visible": "      失敗: 目前視圖({0})沒有可見圖層，未匯出",
        "log.visible": "      可見圖層: {0}（視圖 {1}，共 {2} 層）",
        "log.cover": "{0} 覆蓋匯出（已存在同名檔案）: {1}",
        "log.backup": "{0} 原檔案已備份為: {1}",
        "log.err_backup": "備份失敗，已中止覆蓋：{0} → {1}（{2}）",
        "log.export_start": "{0} 匯出: {1}",
        "log.progress": "正在匯出 {0}/{1}: {2}",
        "log.done": "      完成 {0:.1f}s -> {1} ({2} 位元組)",
        "log.failed": "      失敗: {0}",
        "log.close_fail": "  關閉臨時開啟的部件失敗（忽略）: {0}",
        "log.keep_modified": "  保持開啟（有未儲存修改）: {0}",
        "log.display_restored": "  已切回原顯示部件: {0}",
        "log.display_restore_fail": "  切回原顯示部件失敗: {0}",
        "log.work_restored": "  已還原原工作部件: {0}",
        "log.work_restore_fail": "  還原原工作部件失敗: {0}",
        "log.summary": "成功 {0} 個，略過 {1} 個，失敗 {2} 個",
        "log.background": "已在背景啟動 {0} 個任務，NX 可繼續操作。",
        "log.background_fail": "背景工作階段不可用（{0}），改為前景執行。",
        "log.failed_item": "  失敗: {0}",
        "log.err_no_powershell": "找不到 powershell.exe",
        "log.err_no_result": "視窗沒有傳回結果: {0}",
        "log.ui_hidden": "介面視窗「{0}」啟動 8 秒仍不可見（pid {1}），疑似被 SW_HIDE 啟動。",
        "log.err_no_dir": "目錄不存在: {0}",
        "log.err_creator": "無法建立 STEP 匯出器（請確認已安裝 STEP 轉換器並具備授權）: {0}",
        "log.err_method_missing": "{0} 不存在",
        "log.err_method_fail": "{0} 失敗: {1}",
        "log.err_open_part": "無法開啟部件: {0} — {1}",
        "log.err_display_none": "切換顯示部件失敗（NX 未傳回顯示部件）: {0}",
        "log.err_out_missing": "輸出檔案不存在",
        "log.err_out_empty": "輸出檔案為 0 位元組",
        "log.err_out_header": "輸出檔案只有檔頭、沒有幾何（{0} 位元組）",
        "log.err_export_invalid": "STEP 匯出沒有產出有效內容（{0}）",
        "log.delete_old_fail": "      刪除舊檔案失敗（忽略）: {0}",
        "log.err_remove_target": "無法覆蓋 {0}：該 {2} 檔案唯讀或正被其他程式佔用（如 Excel/記事本），請關閉或取消唯讀後重試（{1}）",
        "log.missing_part": "檔案不存在（可能在選擇後被刪除或移動）: {0}",
        "log.busy": "已有一次匯出正在進行，本次未啟動（避免兩個視窗同時修改同一批檔案）。",
        "log.stale_lock": "已刪除 PID {0} 遺留的陳舊鎖，並重試本次匯出",
        "log.lock_unknown": "無法確認鎖內進程 {0} 是否仍存在，按佔用處理",
        "log.log_fallback": "日誌檔案被佔用，本次日誌改寫到: {0}",
        "log.unsupported_name": "檔案名稱包含 NX 無法辨識的特殊字元（如 emoji），請改名後重試: {0}",
        "log.too_long_path": "路徑超過 Windows 260 字元限制，NX 無法存取；請縮短目錄層級後重試: {0}",
        "log.walk_fail": "無法存取某個目錄（路徑過長或權限不足），已略過: {0}",
        "ui.folder_pick_title": "選擇要匯出 STEP 的資料夾（取消 = 結束選擇）",
        "ui.file_pick_title": "選擇要匯出/轉換的 .prt / .stp / .igs 檔案",
        "ui.filter_prt": "NX 部件與 STEP/IGES (*.prt;*.stp;*.step;*.igs;*.iges)",
        "ui.filter_all": "所有檔案 (*.*)",
        "ui.picker_title": "選擇要匯出/轉換的資料夾 / 檔案",
        "ui.picker_this_pc": "本機",
        "ui.picker_quick_desktop": "桌面",
        "ui.picker_quick_downloads": "下載",
        "ui.picker_quick_documents": "文件",
        "ui.picker_quick_videos": "影片",
        "ui.picker_tree_label": "資料夾",
        "ui.picker_list_label": "檔案（Ctrl/Shift 多選，雙擊加入）",
        "ui.picker_empty_dir": "（此目錄下沒有目前方向可轉換的檔案）",
        "ui.picker_denied": "無法存取此目錄（權限不足或路徑無效）：",
        "ui.picker_denied_hint": "無法存取，請換一個目錄或聯絡管理員。",
        "ui.picker_col_name": "檔案名稱",
        "ui.picker_col_mtime": "修改日期",
        "ui.picker_col_size": "大小(KB)",
        "ui.picker_col_type": "類型",
        "ui.picker_col_path": "路徑",
        "ui.type_dir": "資料夾",
        "ui.type_file": "檔案",
        "ui.btn_add_folder": "＋ 新增資料夾",
        "ui.btn_add_file": "新增匯出檔案",
        "ui.btn_clear": "清空",
        "ui.btn_del_rows": "刪除選取",
        "ui.menu_delete": "刪除選取",
        "ui.menu_select_all": "全選",
        "ui.menu_clear": "清空",
        "ui.btn_cancel": "取消",
        "ui.btn_ok": "確定",
        "ui.busy_title": "匯出正在進行",
        "ui.busy_text": "已有匯出正在進行，請稍後再試。",
        "ui.busy_ok": "確定",
        "ui.msg_remove_title": "確認移除",
        "ui.msg_remove_text": "確定要從已選清單中移除嗎？",
        "ui.msg_empty_title": "提示",
        "ui.msg_empty_one": "尚未新增任何檔案。",
        "ui.msg_empty_two": "請先選取資料夾或檔案，再點「新增」。",
        "ui.preview_title": "確認轉換檔案",
        "ui.same_summary": "同名策略：{0}（上一步已設定，修改請返回）",
        "ui.preview_hint": "勾選即取代同名檔案。",
        "ui.col_cover": "覆蓋",
        "ui.col_no": "序號",
        "ui.col_src": "來源檔案",
        "ui.col_dst": "輸出",
        "ui.col_state": "狀態",
        "ui.state_new": "新建",
        "ui.state_skip": "略過（已存在）",
        "ui.state_overwrite": "覆蓋",
        "ui.chk_all": "全選覆蓋",
        "ui.btn_back": "返回",
        "ui.btn_start": "開始匯出",
        "ui.summary": "{0} 個檔案：新建/覆蓋 {1}，略過 {2}",
        "ui.save_title": "是否儲存目前修改",
        "ui.save_hint": "是否儲存目前修改？",
        "ui.save_item_prefix": "儲存這些修改：",
        "ui.btn_cancel_export": "取消本次匯出",
        "ui.btn_save_go": "繼續匯出",
        "ui.save_empty_confirm": "未勾選任何檔案，將略過 {0} 個檔案的修改繼續匯出？",
        "ui.out_group": "輸出目錄",
        "ui.out_src": "輸出到來源檔案目錄",
        "ui.out_dir": "輸出到指定目錄",
        "ui.out_browse": "瀏覽…",
        "ui.name_prefix": "檔案新增前綴",
        "ui.name_suffix": "檔案新增後綴",
        "ui.name_group": "名稱與格式",
        "ui.on_exists": "同名檔案已存在時",
        "ui.on_exists_skip": "略過",
        "ui.on_exists_rename": "改名（_2、_3…）",
        "ui.on_exists_overwrite": "覆蓋",
        "ui.format": "格式",
        "ui.format_step": "STEP (.step)",
        "ui.format_iges": "IGES (.igs)",
        "ui.state_rename": "改名",
        "ui.state_self": "同格式（略過）",
        "ui.dir_group": "轉換方向",
        "ui.dir_step": "→ STEP（來源 prt / igs）",
        "ui.dir_iges": "→ IGES（來源 prt / stp）",
        "ui.dir_prt_step": "prt→stp",
        "ui.dir_prt_iges": "prt→igs",
        "ui.dir_to_prt": "stp/igs→prt",
        "ui.dir_xt_prt": "x_t/x_b→prt",
        "ui.dir_any": "任意→stp/igs",
        "ui.dir_xt_step": "x_t/x_b→stp",
        "ui.dir_xt_iges": "x_t/x_b→igs",
        "ui.right_title": "轉換後名稱",
        "ui.col_file": "檔案名稱",
        "ui.left_empty": "（雙擊上方檔案加入）",
        "ui.right_empty": "（新增檔案後顯示）",
        "ui.folder_expand": "（資料夾：匯出其中的 .prt）",
        "ui.workers_label": "同時執行數",
        "ui.open_after": "開啟轉換後的檔案",
        "ui.preview_dup": "（重名，改名）",
        "ui.preview_rename": "（改名）",
        "ui.preview_exists": "（略過）",
        "ui.preview_overwrite": "（覆蓋）",
        "log.chain_same_fmt": "來源與目標格式相同，已略過：{0}",
        "log.import_unavailable": "匯入模組不可用：{0}",
        "log.import_fail": "轉 PRT 失敗：{0}（{1}）",
        "log.import_done": "已轉 PRT：{0} → {1} {2}",
        "log.chain_begin": "記憶體中轉轉換：{0} → {1}",
        "log.chain_import_fail": "鏈式匯入失敗：{0}",
        "log.chain_close_fail": "無法關閉暫存記憶體部件：{0}",
        "log.chain_restore_fail": "無法還原原本的顯示部件：{0}",
        "log.err_step_from_part": "StepCreator 不支援「顯示/工作部件」匯出，無法做記憶體 STEP 匯出",
        "log.err_iges_from_part": "IgesCreator 不支援「顯示/工作部件」匯出，無法做記憶體 IGES 匯出",
        "log.out_dir_invalid": "指定輸出目錄為空或不存在，已改回原檔案所在路徑: {0}",
        "log.output_custom": "輸出    : {0}（扁平放置，不鏡像原目錄結構）",
        "log.format_line": "格式    : {0}",
        "log.rename_policy": "同名已存在，改名輸出: {0} -> {1}",
        "log.iges_exe_missing": "找不到 IGES 轉換器 iges.exe: {0}",
        "log.iges_def_fallback": "NX 安裝目錄下找不到 igesexport.def，改回: {0}",
        "log.iges_timeout": "IGES 轉換逾時（{0} 秒），已終止",
        "log.iges_exit": "IGES 轉換失敗（exit code {0}）",
        "log.iges_log_unread": "      （無法讀取 IGES 日誌: {0}）",
        "log.iges_cmd_fail": "無法啟動 IGES 轉換器: {0}",
        "log.err_iges_out": "IGES 輸出檔案不存在或為 0 位元組: {0}",
        "log.err_iges_geom": "IGES 輸出不含幾何（只有檔案頭與屬性，沒有曲線/曲面/實體）: {0}",
        "log.err_target_locked": "無法取代目標檔案（被其他程式佔用，或仍為唯讀）: {0}",
    },
}


def T(key, *args):
    """取当前语言文案；该语言缺项时回落英文。"""
    table = MSG.get(LANG) or {}
    text = table.get(key)
    if text is None:
        text = MSG["en"].get(key, key)
    if args:
        return text.format(*args)
    return text


def render_ps(template):
    """把 PowerShell 模板里的 [[key]] 占位符替换成当前语言文案（缺项回落英文）"""
    merged = dict(MSG["en"])
    merged.update(MSG.get(LANG) or {})
    for key, value in merged.items():
        template = template.replace("[[%s]]" % key, str(value))
    # 文件名清洗片段与导入窗共用一份（定义在公共执行器里）
    try:
        shared = _shared()
        template = template.replace("[[name_helpers]]", shared.NAME_HELPERS_PS)
        template = template.replace("[[left_list]]", shared.LEFT_LIST_PS)
        template = template.replace("[[front_helpers]]", shared.PICKER_FRONT_PS)
    except Exception:
        pass
    return template
# ==========================================

# 非 Unicode 控制台（英文版 cp437、繁中版 cp950 等）打印中文会抛 UnicodeEncodeError，
# 且异常发生在日志函数里，主循环的 except 再调日志会二次抛错、整个导出中断。
# 这里把 stdout 的编码错误处理降级为替换字符，让控制台输出永远不致命；
# 日志文件仍按 UTF-8 完整写入，不受影响。
try:
    import sys
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# 注意：不要在本文件里定义 startup(argc, args)。
# startup 目录下的 .py 会被 NX 当成“启动应用”导入；一旦它有 startup 函数，
# 导入就会成功并被注册，之后点菜单用 journal 播放同一个模块时，
# nxpythonldr 会直接崩（syslog 里是 signal 11，界面弹空白的 Journal execution error）。
# 没有 startup 时 NX 只会在 syslog 记一条“no attribute startup”的提示，功能正常。

# ================= 配置区 =================
# 已移除 BASE_DIR：不再有“项目根目录”概念。选择窗口从「此电脑」开始枚举所有盘符，
# 可以逐层展开任意目录，并在任意位置勾选文件夹或 .prt 加入清单。
RECURSIVE = True        # True = 递归处理所选文件夹的所有子文件夹（每个 .prt 各自导出到原处）
SKIP_EXISTING = True    # True = 已存在同名 .step 就跳过（只有文件头、没有几何的不算已存在）
VISIBLE_LAYERS_ONLY = True   # True = 只导出当前工作视图中可见的图层（隐藏层实体不进 STEP）
DRY_RUN = False         # 交互模式下选完会自动弹「清单预览」确认，无需改这里；
                        # 仅 INTERACTIVE=False（无界面批处理）时用来只看不导

# INTERACTIVE = True 时弹出自定义选择窗口：
#   左侧文件夹树（顶层为「此电脑」+ 所有盘符，单击选中、双击展开）、右侧该文件夹下的 .prt 列表
#   （可多选）、下方已选清单（点行尾 ✕ 确认后移除，其余行自动上移）、右下角「确认导出」。
# PICK_MODE 仅在窗口打不开、回退到系统对话框时决定弹哪些框。
PICK_MODE = "both"      # 仅影响回退方案："folder" / "file" / "both"
INTERACTIVE = True      # False = 不弹任何窗，完全按下面的配置规则跑

# ---- 筛选规则（不带 BASE_DIR，目录一律写绝对路径），支持 * 通配符，不区分大小写 ----
# 例：INCLUDE_DIRS = [r"D:\parts\box1", r"E:\proj"]   只处理这些文件夹（含其子文件夹）
#     EXCLUDE_DIRS = ["old*"]                         不处理名字匹配的文件夹（任意层级）
#     INCLUDE_FILES = ["顶盖.prt", "*上盖*"]           只处理文件名匹配这些的 prt
#     EXCLUDE_FILES = ["*draft*"]                     跳过文件名匹配的 prt
# 空列表 = 该条规则不生效；交互模式下以选择窗口勾选的结果为准
INCLUDE_DIRS = []
EXCLUDE_DIRS = []
INCLUDE_FILES = []
EXCLUDE_FILES = []
# =========================================

# IGES 分支：直调 NX 自带的 translator（iges.exe），不走 NXOpen 导出器。
# 仅当 UI 里把“格式”选成 IGES 时才会用到这些常量；默认 STEP 流程完全不受影响。
IGES_DEFAULT_DIR = r"D:\NX2506\IGES"   # 找不到 NX 安装目录时的回落位置
IGES_TIMEOUT = 600                     # 单件 IGES 转换超时（秒）
IGES_LOG_TAIL_LINES = 10               # 失败时日志尾部打印行数

_log_handle = None
_RUN_LOCK = None


def default_options():
    """导出选项默认值：不选新选项时与改动前的行为完全一致"""
    return {
        "out_mode": "source",                       # source = 原文档所在路径
        "out_dir": "",                              # 指定导出目录（out_mode=dir 时有效）
        "prefix": "",
        "suffix": "",
        # SKIP_EXISTING 是旧配置项：关闭它时旧行为是“已存在则覆盖”，这里保持一致
        "on_exists": "skip" if SKIP_EXISTING else "overwrite",
        "format": "step",                           # step / iges
        "direction": "prt_step",                    # prt_step/prt_iges/to_prt/xt_prt/any_step_igs
        "workers": 2,                               # 主窗并发（1~6），传给后台执行器
        "open_after": False,                        # →prt：勾选后转换完成即自动打开产物（默认不勾）
    }


DIRECTIONS = ("prt_step", "prt_iges", "to_prt", "xt_prt", "xt_step", "xt_iges", "any_step_igs")


def to_prt_direction(options):
    """当前是否为「→prt」方向（这两条走导入逻辑，产物是 .prt）"""
    return str((options or {}).get("direction") or "prt_step") in ("to_prt", "xt_prt")


def output_format(options):
    """规范化输出格式：→prt 返回 prt；其余取 format，缺省时按方向推（prt_iges→iges）"""
    if to_prt_direction(options):
        return "prt"
    value = str((options or {}).get("format") or "").strip().lower()
    if value in ("step", "iges"):
        return value
    if str((options or {}).get("direction") or "") in ("prt_iges", "xt_iges"):
        return "iges"
    return "step"


def output_ext(options):
    """输出扩展名（含点）：STEP -> .step，IGES -> .igs，→prt -> .prt"""
    fmt = output_format(options)
    if fmt == "prt":
        return ".prt"
    return ".igs" if fmt == "iges" else ".step"


def output_is_valid(path, options=None):
    """目标文件是否算“已存在”：PRT/IGES 只判存在且大小 > 0；STEP 沿用含几何判定"""
    if output_format(options) in ("prt", "iges"):
        return step_size(path) > 0
    return step_is_valid(path)


def resolve_output_options(options):
    """校验“指定导出目录”：为空或不存在时回落到原文档所在路径，并记一条日志"""
    result = dict(options or default_options())
    if result.get("out_mode") != "dir":
        result["out_mode"] = "source"
        result["out_dir"] = ""
        return result
    raw = str(result.get("out_dir") or "").strip()
    path = os.path.abspath(raw) if raw else ""
    if not raw or not os.path.isdir(path):
        log(T("log.out_dir_invalid", path or raw or "-"))
        result["out_mode"] = "source"
        result["out_dir"] = ""
        return result
    result["out_dir"] = path
    return result


_SHARED_RUNNER = None


def _shared():
    """公共执行器模块（缓存）：文件名规则与 PS 片段都在那里，两窗共用一份"""
    global _SHARED_RUNNER
    if _SHARED_RUNNER is None:
        _SHARED_RUNNER = _load_runner()
    return _SHARED_RUNNER


def normalize_name(name):
    """把文件名中的全角括号等字符替换为英文符号（规则在公共执行器里，与导入窗一致）"""
    return _shared().normalize_name(name)


def safe_name_part(text):
    """前缀/后缀里 Windows 文件名不允许的字符替换成下划线（规则在公共执行器里）"""
    return _shared().safe_name_part(text)


def force_remove(path):
    """删除文件；只读属性导致失败时先去掉只读再删（用于清理临时产物）"""
    try:
        os.remove(path)
    except PermissionError:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        os.remove(path)


def log(msg):
    """同时打印到 NX 信息窗口/控制台和日志文件"""
    text = str(msg)
    try:
        print(text)
    except Exception:
        # 控制台代码页无法表示这些字符（reconfigure 不可用时的兜底）：
        # 只放弃屏幕输出，日志文件仍照常写入
        pass
    if _log_handle:
        _log_handle.write(text + "\n")
        _log_handle.flush()


def plugin_dir():
    r"""插件根目录（脚本在 startup\ 下，取它的上一层；exe/被 exec 执行时回落当前目录）"""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()
    if os.path.basename(here).lower() == "startup":
        parent = os.path.dirname(here)
        if parent:
            return parent
    return here


# 选择窗「上次浏览目录」记忆：与导入侧 import_last_dir.txt 各存各的，互不影响
EXPORT_LAST_DIR_FILE = "export_last_dir.txt"


def _last_dir_path():
    return os.path.join(plugin_dir(), "logs", EXPORT_LAST_DIR_FILE)


def read_last_dir():
    try:
        with open(_last_dir_path(), "r", encoding="utf-8") as handle:
            value = handle.read().strip()
        if value and os.path.isdir(value):
            return value
    except Exception:
        pass
    return ""


def write_last_dir(folder):
    try:
        path = _last_dir_path()
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(folder)
    except Exception:
        pass


# 选择窗「上次用过的选项」记忆：与目录记忆分开存，损坏或值非法时静默回落默认值
EXPORT_OPTIONS_FILE = "export_options.json"


def _options_path():
    return os.path.join(plugin_dir(), "logs", EXPORT_OPTIONS_FILE)


def read_options():
    try:
        with open(_options_path(), "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except Exception:
        return {}
    return parse_picker_options(payload, {})


def write_options(options):
    try:
        path = _options_path()
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(dict(options or {}), handle, ensure_ascii=False, indent=1, sort_keys=True)
    except Exception:
        pass


SELECT_UI_PS = r'''
param(
    [string]$Root = "",
    [string]$OutMode = "source",
    [string]$OutDir = "",
    [string]$Prefix = "",
    [string]$Suffix = "",
    [string]$OnExists = "skip",
    [string]$Format = "step",
    [string]$Direction = "prt_step",
    [int]$Workers = 2,
    [string]$OpenAfter = "0",
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 弹出即抢前台（AttachThreadInput + BringWindowToTop + 瞬时 TopMost），片段在公共执行器里
[[front_helpers]]

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[ui.picker_title]]"
$form.ClientSize = New-Object System.Drawing.Size(944, 726)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.TopMost = $false
$uiFont = [System.Drawing.SystemFonts]::MessageBoxFont
try { $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9) } catch { }
$form.Font = $uiFont
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)

# ---- NX 2506 观感基准（实测值与来源见交付参数表）----
# 内容背景 #FFFFFF、分隔线 #DCDCDC、正文 #1E1E1E 取自 NX 2506 界面截图采样；
# 深色主题无本机 NX 截图可量，用 Windows 深色系统色近似；
# 环境变量 NX_PICKER_THEME=light|dark 可强制切换（仅用于渲染核对，功能无关）。
$script:isDark = $false
try {
    $themeKey = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction Stop
    if ($themeKey.AppsUseLightTheme -eq 0) { $script:isDark = $true }
} catch { }
if ($env:NX_PICKER_THEME -eq 'dark') { $script:isDark = $true }
elseif ($env:NX_PICKER_THEME -eq 'light') { $script:isDark = $false }
if ($script:isDark) {
    $uiBg = [System.Drawing.Color]::FromArgb(32, 32, 32)
    $uiField = [System.Drawing.Color]::FromArgb(24, 24, 24)
    $uiFg = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $uiLine = [System.Drawing.Color]::FromArgb(63, 63, 70)
    $uiBtn = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $uiFgDim = [System.Drawing.Color]::FromArgb(150, 150, 150)
} else {
    $uiBg = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $uiField = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $uiFg = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $uiLine = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $uiBtn = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $uiFgDim = [System.Drawing.Color]::FromArgb(120, 120, 120)
}
$form.BackColor = $uiBg
$form.ForeColor = $uiFg

function Add-Separator([int]$x, [int]$y, [int]$w) {
    $line = New-Object System.Windows.Forms.Panel
    $line.Location = New-Object System.Drawing.Point($x, $y)
    $line.Size = New-Object System.Drawing.Size($w, 1)
    $line.BackColor = $uiLine
    $form.Controls.Add($line)
}

function Apply-Theme([System.Windows.Forms.Control]$root) {
    foreach ($c in $root.Controls) {
        if ($c -is [System.Windows.Forms.GroupBox]) {
            $c.ForeColor = $uiFg
            $c.BackColor = $uiBg
        } elseif ($c -is [System.Windows.Forms.Label] -or $c -is [System.Windows.Forms.RadioButton]) {
            $c.ForeColor = $uiFg
            $c.BackColor = [System.Drawing.Color]::Transparent
        } elseif ($c -is [System.Windows.Forms.NumericUpDown]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
        } elseif ($c -is [System.Windows.Forms.ComboBox]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
            # 深色下标准 Combobox 会自绘白底，改 Flat 才吃 BackColor
            if ($script:isDark) { $c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
        } elseif ($c -is [System.Windows.Forms.TextBox]) {
            $c.BackColor = $uiField
            # 只读输入框用灰字（不依赖系统禁用态配色）
            if ($c.ReadOnly) { $c.ForeColor = $uiFgDim } else { $c.ForeColor = $uiFg }
        } elseif ($c -is [System.Windows.Forms.TreeView] -or $c -is [System.Windows.Forms.ListView]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
            # 列表网格线是系统色，深色下会变成刺眼白线，直接关掉
            if ($script:isDark -and $c -is [System.Windows.Forms.ListView]) { $c.GridLines = $false }
        } elseif ($c -is [System.Windows.Forms.Button]) {
            $c.BackColor = $uiBtn
            $c.ForeColor = $uiFg
            $c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $c.FlatAppearance.BorderColor = $uiLine
        }
        if ($c.HasChildren) { Apply-Theme $c }
    }
}

# ---- 顶部选项带：转换方向 / 同名处理 / 输出位置与路径 / 前后缀 ----
# 控件宽度按当前语言的实测文字宽度分配（TextRenderer.MeasureText）：
# 繁体/英文文案比简体长，沿用固定像素宽度会把文字顶到组边框上。
$uiRadioPad = 26      # 单选圆点 ~13px + 文字间距 + 右留白
$uiColGap = 8         # 同一行相邻控件间距
$uiBoxLeft = 14       # 组内左边距
$uiBoxRight = 12      # 组内右边距
$uiBoxH = 92          # 顶部三组统一高度
function Measure-UiText([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return 0 }
    $flags = [System.Windows.Forms.TextFormatFlags]::NoPadding
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($text, $form.Font, (New-Object System.Drawing.Size(4000, 200)), $flags)
    return $size.Width
}
function Get-RadioWidth([System.Windows.Forms.RadioButton]$rb) {
    return (Measure-UiText $rb.Text) + $uiRadioPad
}
function Place-Grid([object[]]$r1, [object[]]$r2, [object[]]$r3, [int]$top, [int]$pitch) {
    # 三行按列对齐：每列取该列最宽的控件，统一列起点与行距，避免标签互相挤压
    $rows = @($r1, $r2, $r3)
    $colW = @(0, 0, 0)
    foreach ($row in $rows) {
        for ($i = 0; $i -lt $row.Count; $i++) {
            $w = Get-RadioWidth $row[$i]
            if ($w -gt $colW[$i]) { $colW[$i] = $w }
        }
    }
    $colX = @(0, 0, 0)
    $x = $uiBoxLeft
    for ($i = 0; $i -lt 3; $i++) {
        $colX[$i] = $x
        $x += $colW[$i] + $uiColGap
    }
    $right = $uiBoxLeft
    $y = $top
    foreach ($row in $rows) {
        for ($i = 0; $i -lt $row.Count; $i++) {
            $rb = $row[$i]
            $rb.AutoSize = $false
            $rb.Location = New-Object System.Drawing.Point($colX[$i], $y)
            $rb.Size = New-Object System.Drawing.Size($colW[$i], 22)
            $edge = $colX[$i] + $colW[$i]
            if ($edge -gt $right) { $right = $edge }
        }
        $y += $pitch
    }
    return $right
}

$grpDir = New-Object System.Windows.Forms.GroupBox
$grpDir.Text = "[[ui.dir_group]]"
$grpDir.Location = New-Object System.Drawing.Point(8, 6)
$form.Controls.Add($grpDir)

$rbStep = New-Object System.Windows.Forms.RadioButton
$rbStep.Text = "[[ui.dir_prt_step]]"
$grpDir.Controls.Add($rbStep)

$rbIges = New-Object System.Windows.Forms.RadioButton
$rbIges.Text = "[[ui.dir_prt_iges]]"
$grpDir.Controls.Add($rbIges)

$rbToPrt = New-Object System.Windows.Forms.RadioButton
$rbToPrt.Text = "[[ui.dir_to_prt]]"
$grpDir.Controls.Add($rbToPrt)

$rbXtToPrt = New-Object System.Windows.Forms.RadioButton
$rbXtToPrt.Text = "[[ui.dir_xt_prt]]"
$grpDir.Controls.Add($rbXtToPrt)

$rbAny = New-Object System.Windows.Forms.RadioButton
$rbAny.Text = "[[ui.dir_any]]"
$grpDir.Controls.Add($rbAny)

$rbXtStep = New-Object System.Windows.Forms.RadioButton
$rbXtStep.Text = "[[ui.dir_xt_step]]"
$grpDir.Controls.Add($rbXtStep)

$rbXtIgs = New-Object System.Windows.Forms.RadioButton
$rbXtIgs.Text = "[[ui.dir_xt_iges]]"
$grpDir.Controls.Add($rbXtIgs)

# 第一行两项（日常互换 prt→stp / stp/igs→prt），
# 第二行两项（prt→igs / 任意→stp/igs），
# 第三行三项（x_t/x_b→prt / x_t/x_b→stp / x_t/x_b→igs）；三列按列对齐、行距统一
$dirRight = Place-Grid -r1 @($rbStep, $rbToPrt) -r2 @($rbIges, $rbAny) -r3 @($rbXtToPrt, $rbXtStep, $rbXtIgs) -top 18 -pitch 24
$dirContentRight = [Math]::Max($dirRight, ($uiBoxLeft + (Measure-UiText $grpDir.Text) + 2))
$grpDir.Size = New-Object System.Drawing.Size(($dirContentRight + $uiBoxRight), $uiBoxH)

$grpExists = New-Object System.Windows.Forms.GroupBox
$grpExists.Text = "[[ui.on_exists]]"
$form.Controls.Add($grpExists)

$rbNew = New-Object System.Windows.Forms.RadioButton
$rbNew.Text = "[[ui.on_exists_rename]]"
$grpExists.Controls.Add($rbNew)

$rbSkip = New-Object System.Windows.Forms.RadioButton
$rbSkip.Text = "[[ui.on_exists_skip]]"
$grpExists.Controls.Add($rbSkip)

$rbOverwrite = New-Object System.Windows.Forms.RadioButton
$rbOverwrite.Text = "[[ui.on_exists_overwrite]]"
$grpExists.Controls.Add($rbOverwrite)

# 三项纵向排列，宽度按最长项「改名（_2、_3…）」实测；组宽同时要容下组标题
$existsY = 18
foreach ($rb in @($rbNew, $rbSkip, $rbOverwrite)) {
    $rb.AutoSize = $false
    $rb.Location = New-Object System.Drawing.Point($uiBoxLeft, $existsY)
    $rb.Size = New-Object System.Drawing.Size((Get-RadioWidth $rb), 21)
    $existsY += 22
}
$existsContentRight = [Math]::Max(($uiBoxLeft + [Math]::Max((Get-RadioWidth $rbNew), [Math]::Max((Get-RadioWidth $rbSkip), (Get-RadioWidth $rbOverwrite)))), ($uiBoxLeft + (Measure-UiText $grpExists.Text) + 2))
$grpExists.Location = New-Object System.Drawing.Point(($grpDir.Left + $grpDir.Width + 8), 6)
$grpExists.Size = New-Object System.Drawing.Size(($existsContentRight + $uiBoxRight), $uiBoxH)

$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = "[[ui.out_group]]"
$form.Controls.Add($grpOut)

$rbOutSrc = New-Object System.Windows.Forms.RadioButton
$rbOutSrc.Text = "[[ui.out_src]]"
$rbOutSrc.TabIndex = 0
$grpOut.Controls.Add($rbOutSrc)

$rbOutDir = New-Object System.Windows.Forms.RadioButton
$rbOutDir.Text = "[[ui.out_dir]]"
$rbOutDir.TabIndex = 1
$grpOut.Controls.Add($rbOutDir)

$txtOutDir = New-Object System.Windows.Forms.TextBox
$txtOutDir.Text = $OutDir
$txtOutDir.TabIndex = 2
$grpOut.Controls.Add($txtOutDir)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "[[ui.out_browse]]"
$btnBrowse.TabIndex = 3
$grpOut.Controls.Add($btnBrowse)

# 输出组吃掉剩余宽度（窗口右边界固定 936）；两个单选并排放不下时（英文）改上下两行
$grpOut.Location = New-Object System.Drawing.Point(($grpExists.Left + $grpExists.Width + 8), 6)
$grpOut.Size = New-Object System.Drawing.Size((936 - $grpOut.Left), $uiBoxH)
$outSrcW = Get-RadioWidth $rbOutSrc
$outDstW = Get-RadioWidth $rbOutDir
$browseW = [Math]::Max(62, ((Measure-UiText $btnBrowse.Text) + 20))
$outInnerRight = $grpOut.Width - $uiBoxRight
$rbOutSrc.AutoSize = $false
$rbOutDir.AutoSize = $false
if (($uiBoxLeft + $outSrcW + $uiColGap + $outDstW) -le $outInnerRight) {
    $rbOutSrc.Location = New-Object System.Drawing.Point($uiBoxLeft, 18)
    $rbOutSrc.Size = New-Object System.Drawing.Size($outSrcW, 22)
    $rbOutDir.Location = New-Object System.Drawing.Point(($uiBoxLeft + $outSrcW + $uiColGap), 18)
    $rbOutDir.Size = New-Object System.Drawing.Size($outDstW, 22)
    $pathY = 48
} else {
    $rbOutSrc.Location = New-Object System.Drawing.Point($uiBoxLeft, 18)
    $rbOutSrc.Size = New-Object System.Drawing.Size($outSrcW, 22)
    $rbOutDir.Location = New-Object System.Drawing.Point($uiBoxLeft, 40)
    $rbOutDir.Size = New-Object System.Drawing.Size($outDstW, 22)
    $pathY = 64
}
$btnBrowse.Size = New-Object System.Drawing.Size($browseW, 26)
$btnBrowse.Location = New-Object System.Drawing.Point(($outInnerRight - $browseW), ($pathY - 1))
$txtOutDir.Location = New-Object System.Drawing.Point($uiBoxLeft, $pathY)
$txtOutDir.Size = New-Object System.Drawing.Size([Math]::Max(120, ($btnBrowse.Left - 6 - $uiBoxLeft)), 24)

# 浏览按钮：FolderBrowserDialog，选中的目录直接回填到文本框
$btnBrowse.add_Click({
    param($sender, $e)
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "[[ui.out_browse]]"
    if ($txtOutDir.Text -and [System.IO.Directory]::Exists($txtOutDir.Text)) {
        $dlg.SelectedPath = $txtOutDir.Text
    }
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutDir.Text = $dlg.SelectedPath
    }
})

switch ($Direction) {
    "prt_iges" { $rbIges.Checked = $true }
    "to_prt"   { $rbToPrt.Checked = $true }
    "xt_prt"   { $rbXtToPrt.Checked = $true }
    "xt_step"  { $rbXtStep.Checked = $true }
    "xt_iges"  { $rbXtIgs.Checked = $true }
    "any_step_igs" { $rbAny.Checked = $true }
    default    { $rbStep.Checked = $true }
}
if ($OnExists -eq "rename") { $rbNew.Checked = $true }
elseif ($OnExists -eq "overwrite") { $rbOverwrite.Checked = $true }
else { $rbSkip.Checked = $true }
if ($OutMode -eq "dir") { $rbOutDir.Checked = $true } else { $rbOutSrc.Checked = $true }

function Set-OutDirEnabled {
    # 路径行只在「输出到指定目录」时出现；选默认时整行隐藏（分组高度不变，不跳动）
    $on = $rbOutDir.Checked
    $txtOutDir.Visible = $on
    $btnBrowse.Visible = $on
}
Set-OutDirEnabled
$rbOutSrc.add_CheckedChanged({ param($sender, $e) Set-OutDirEnabled })
$rbOutDir.add_CheckedChanged({
    param($sender, $e)
    Set-OutDirEnabled
    # 切到指定目录且路径为空时直接聚焦输入框；切回默认不清空已输入的值
    if (-not $txtOutDir.Text.Trim()) { $txtOutDir.Focus() }
})

$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text = "[[ui.name_prefix]]"
$lblPrefix.Location = New-Object System.Drawing.Point(8, 102)
$lblPrefix.Size = New-Object System.Drawing.Size(86, 20)
$form.Controls.Add($lblPrefix)

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Text = $Prefix
$txtPrefix.Location = New-Object System.Drawing.Point(96, 100)
$txtPrefix.Size = New-Object System.Drawing.Size(140, 24)
$form.Controls.Add($txtPrefix)

$lblSuffix = New-Object System.Windows.Forms.Label
$lblSuffix.Text = "[[ui.name_suffix]]"
$lblSuffix.Location = New-Object System.Drawing.Point(248, 102)
$lblSuffix.Size = New-Object System.Drawing.Size(86, 20)
$form.Controls.Add($lblSuffix)

$txtSuffix = New-Object System.Windows.Forms.TextBox
$txtSuffix.Text = $Suffix
$txtSuffix.Location = New-Object System.Drawing.Point(336, 100)
$txtSuffix.Size = New-Object System.Drawing.Size(140, 24)
$form.Controls.Add($txtSuffix)

$lblTree = New-Object System.Windows.Forms.Label
$lblTree.Text = "[[ui.picker_tree_label]]"
$lblTree.Location = New-Object System.Drawing.Point(8, 132)
$lblTree.Size = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($lblTree)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Location = New-Object System.Drawing.Point(8, 152)
$tree.Size = New-Object System.Drawing.Size(340, 184)
$tree.HideSelection = $false
$tree.ItemHeight = 22
# NX 树形控件不用点状连线，展开箭头由系统绘制
$tree.ShowLines = $false
$tree.ShowRootLines = $false
$tree.Indent = 16
$tree.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($tree)

$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = "[[ui.picker_list_label]]"
$lblList.Location = New-Object System.Drawing.Point(364, 132)
$lblList.Size = New-Object System.Drawing.Size(568, 20)
$form.Controls.Add($lblList)

$list = New-Object System.Windows.Forms.ListView
$list.Location = New-Object System.Drawing.Point(364, 152)
$list.Size = New-Object System.Drawing.Size(568, 184)
$list.View = "Details"
$list.FullRowSelect = $true
$list.MultiSelect = $true
$list.GridLines = $true
[void]$list.Columns.Add("[[ui.picker_col_name]]", 290)
[void]$list.Columns.Add("[[ui.picker_col_mtime]]", 118)
[void]$list.Columns.Add("[[ui.picker_col_type]]", 64)
[void]$list.Columns.Add("[[ui.picker_col_size]]", 78)
$list.ShowItemToolTips = $true
$form.Controls.Add($list)

# ---- 已选清单（左）与转换后文件名对照（右）----
$btnAddFile = New-Object System.Windows.Forms.Button
$btnAddFile.Text = "[[ui.btn_add_file]]"
$btnAddFile.Location = New-Object System.Drawing.Point(8, 342)
$btnAddFile.Size = New-Object System.Drawing.Size(124, 26)
$form.Controls.Add($btnAddFile)

$btnAddFolder = New-Object System.Windows.Forms.Button
$btnAddFolder.Text = "[[ui.btn_add_folder]]"
$btnAddFolder.Location = New-Object System.Drawing.Point(136, 342)
$btnAddFolder.Size = New-Object System.Drawing.Size(110, 26)
$form.Controls.Add($btnAddFolder)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "[[ui.btn_clear]]"
$btnClear.Location = New-Object System.Drawing.Point(250, 342)
$btnClear.Size = New-Object System.Drawing.Size(62, 26)
$form.Controls.Add($btnClear)

# 多选删除：本按钮或左栏右键菜单，删完重编号并重算右栏
$btnDelRows = New-Object System.Windows.Forms.Button
$btnDelRows.Text = "[[ui.btn_del_rows]]"
$btnDelRows.Location = New-Object System.Drawing.Point(320, 342)
$btnDelRows.Size = New-Object System.Drawing.Size(148, 26)
$form.Controls.Add($btnDelRows)

$lblRightTitle = New-Object System.Windows.Forms.Label
$lblRightTitle.Text = "[[ui.right_title]]"
$lblRightTitle.Location = New-Object System.Drawing.Point(492, 348)
$lblRightTitle.Size = New-Object System.Drawing.Size(444, 20)
$form.Controls.Add($lblRightTitle)

$lblArrow = New-Object System.Windows.Forms.Label
$lblArrow.Text = "→"
$lblArrow.Location = New-Object System.Drawing.Point(470, 470)
$lblArrow.Size = New-Object System.Drawing.Size(20, 20)
$form.Controls.Add($lblArrow)

$left = New-Object System.Windows.Forms.ListView
$left.Location = New-Object System.Drawing.Point(8, 376)
$left.Size = New-Object System.Drawing.Size(460, 250)
$left.View = "Details"
$left.FullRowSelect = $true
$left.MultiSelect = $true
$left.GridLines = $true
$left.HideSelection = $false
[void]$left.Columns.Add("[[ui.col_no]]", 46)
[void]$left.Columns.Add("[[ui.col_file]]", 370)
[void]$left.Columns.Add("✕", 36)
$left.ShowItemToolTips = $true
$form.Controls.Add($left)

$right = New-Object System.Windows.Forms.ListView
$right.Location = New-Object System.Drawing.Point(492, 376)
$right.Size = New-Object System.Drawing.Size(444, 250)
$right.View = "Details"
$right.FullRowSelect = $true
$right.MultiSelect = $false
$right.GridLines = $true
[void]$right.Columns.Add("[[ui.col_no]]", 46)
[void]$right.Columns.Add("[[ui.col_file]]", 390)
$right.ShowItemToolTips = $true
$form.Controls.Add($right)

# 空表占位：列表本身没有占位文字，用同底色 Label 覆盖，列表有行时隐藏
$lblLeftEmpty = New-Object System.Windows.Forms.Label
$lblLeftEmpty.Text = "[[ui.left_empty]]"
$lblLeftEmpty.Location = New-Object System.Drawing.Point(16, 402)
$lblLeftEmpty.Size = New-Object System.Drawing.Size(440, 20)
$lblLeftEmpty.BackColor = $uiField
$lblLeftEmpty.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblLeftEmpty)

$lblRightEmpty = New-Object System.Windows.Forms.Label
$lblRightEmpty.Text = "[[ui.right_empty]]"
$lblRightEmpty.Location = New-Object System.Drawing.Point(500, 402)
$lblRightEmpty.Size = New-Object System.Drawing.Size(430, 20)
$lblRightEmpty.BackColor = $uiField
$lblRightEmpty.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblRightEmpty)

$lblWorkers = New-Object System.Windows.Forms.Label
$lblWorkers.Text = "[[ui.workers_label]]"
$lblWorkers.Location = New-Object System.Drawing.Point(8, 632)
$lblWorkers.Size = New-Object System.Drawing.Size(176, 22)
$form.Controls.Add($lblWorkers)

$numWorkers = New-Object System.Windows.Forms.NumericUpDown
$numWorkers.Location = New-Object System.Drawing.Point(188, 629)
$numWorkers.Size = New-Object System.Drawing.Size(60, 24)
$numWorkers.Minimum = 1
$numWorkers.Maximum = 6
$w = 2
try { $w = [int]$Workers } catch { }
if ($w -lt 1) { $w = 1 }
if ($w -gt 6) { $w = 6 }
$numWorkers.Value = [decimal]$w
$form.Controls.Add($numWorkers)

# 只有 →prt 两个方向才产生 .prt，其余方向不显示这个勾选（避免用户误以为会打开 step/igs）
$cbOpenAfter = New-Object System.Windows.Forms.CheckBox
$cbOpenAfter.Text = "[[ui.open_after]]"
$cbOpenAfter.AutoSize = $true
$cbOpenAfter.Location = New-Object System.Drawing.Point(274, 631)
$cbOpenAfter.Checked = ($OpenAfter -eq "1")
$form.Controls.Add($cbOpenAfter)

function Update-OpenAfter {
    $cbOpenAfter.Visible = ($rbToPrt.Checked -or $rbXtToPrt.Checked)
}
Update-OpenAfter

# 方向变了，文件列表的源扩展名跟着变；窗体尚未建好时跳过
function Update-DirSource {
    Update-OpenAfter
    if (Get-Command Show-DirFiles -ErrorAction SilentlyContinue) { Show-DirFiles $script:curDir }
}
Update-DirSource
$rbStep.add_CheckedChanged({ param($sender, $e) Update-DirSource })
$rbIges.add_CheckedChanged({ param($sender, $e) Update-DirSource })
$rbToPrt.add_CheckedChanged({ param($sender, $e) Update-DirSource })
$rbXtToPrt.add_CheckedChanged({ param($sender, $e) Update-DirSource })
$rbXtStep.add_CheckedChanged({ param($sender, $e) Update-DirSource })
$rbXtIgs.add_CheckedChanged({ param($sender, $e) Update-DirSource })
$rbAny.add_CheckedChanged({ param($sender, $e) Update-DirSource })

# 行高按 NX 实测行距（22px@125% ≈ 17.6 逻辑 + 行间留白）统一为 22：
# ListView 只能借 ImageList 撑高行，图标列宽 1px 不占位
$script:rowHeight = 22
$script:rowImages = New-Object System.Windows.Forms.ImageList
$script:rowImages.ImageSize = New-Object System.Drawing.Size(1, $script:rowHeight)
$script:rowImages.Images.Add((New-Object System.Drawing.Bitmap(1, $script:rowHeight)))
$list.SmallImageList = $script:rowImages
$left.SmallImageList = $script:rowImages
$right.SmallImageList = $script:rowImages

# ---- 结果名预览：按 build_plan 的命名规则在 PS 侧复算 ----
# Get-SafePart / Get-CleanStem 与导入窗共用，定义在 nx_batch_runner.NAME_HELPERS_PS
[[name_helpers]]

function Get-OutExt {
    # 方向决定产物扩展名：→prt 两个方向产物都是 .prt
    if ($rbToPrt.Checked -or $rbXtToPrt.Checked) { return ".prt" }
    if ($rbIges.Checked -or $rbXtIgs.Checked) { return ".igs" }
    return ".step"
}

function Get-SourceExts {
    # 与 Python 侧 source_exts 同口径：只列当前方向可转的源文件
    if ($rbToPrt.Checked) { return @(".stp", ".step", ".igs", ".iges") }
    if ($rbXtToPrt.Checked -or $rbXtStep.Checked -or $rbXtIgs.Checked) { return @(".x_t", ".x_b") }
    if ($rbAny.Checked) { return @(".prt", ".stp", ".step", ".igs", ".iges") }
    return @(".prt")
}

function Get-OutRoot {
    if ($rbOutDir.Checked) {
        $d = $txtOutDir.Text.Trim()
        if ($d -and [System.IO.Directory]::Exists($d)) { return $d }
    }
    return ""
}

# 左栏清单逻辑与导入窗共用，定义在 nx_batch_runner.LEFT_LIST_PS；
# 这里只提供文案与各窗自己的统一刷新入口 Refresh-List
$script:txtDelRows = "[[ui.btn_del_rows]]"
$script:txtMenuDelete = "[[ui.menu_delete]]"
$script:txtMenuSelectAll = "[[ui.menu_select_all]]"
$script:txtMenuClear = "[[ui.menu_clear]]"
[[left_list]]

function Update-Preview {
    $right.BeginUpdate()
    try {
        $right.Items.Clear()
        $ext = Get-OutExt
        $root = Get-OutRoot
        $prefix = Get-SafePart $txtPrefix.Text
        $suffix = Get-SafePart $txtSuffix.Text
        $used = @{}
        $no = 0
        foreach ($row in $left.Items) {
            $no++
            $src = [string]$row.Tag
            if ($row.Name -eq "DIR") {
                Add-RightRow $no "[[ui.folder_expand]]" $src $false
                continue
            }
            $srcName = [System.IO.Path]::GetFileName($src)
            $srcExt = [System.IO.Path]::GetExtension($src).ToLower()
            # .stp/.igs 与 .step/.iges 视为同一目标格式（与 build_plan 的 normalize 一致）
            if ($srcExt -eq ".stp") { $srcExt = ".step" }
            if ($srcExt -eq ".iges") { $srcExt = ".igs" }
            $stem = Get-CleanStem ([System.IO.Path]::GetFileNameWithoutExtension($srcName))
            $baseName = $prefix + $stem + $suffix
            $targetDir = $root
            if (-not $targetDir) { $targetDir = [System.IO.Path]::GetDirectoryName($src) }
            $outName = $baseName + $ext
            # 只有目标与源是同一个文件才算“同格式跳过”；换到别的目录允许真转换
            if ($srcExt -eq $ext -and
                ([System.IO.Path]::Combine($targetDir, $outName)).ToLower() -eq $src.ToLower()) {
                Add-RightRow $no ("[[ui.state_self]]") $src $true
                continue
            }
            $state = ""
            $key = ($targetDir + "|" + $outName).ToLower()
            if ($used.ContainsKey($key)) {
                $serial = 2
                while ($used.ContainsKey((($targetDir + "|" + ($baseName + "_" + $serial + $ext))).ToLower())) { $serial++ }
                $outName = $baseName + "_" + $serial + $ext
                $key = ($targetDir + "|" + $outName).ToLower()
                $state = "[[ui.preview_dup]]"
            } else {
                $full = [System.IO.Path]::Combine($targetDir, $outName)
                if ([System.IO.File]::Exists($full)) {
                    if ($rbNew.Checked) {
                        $serial = 2
                        while ($true) {
                            $cand = $baseName + "_" + $serial + $ext
                            $ck = ($targetDir + "|" + $cand).ToLower()
                            if (-not $used.ContainsKey($ck) -and -not [System.IO.File]::Exists([System.IO.Path]::Combine($targetDir, $cand))) { break }
                            $serial++
                        }
                        $outName = $baseName + "_" + $serial + $ext
                        $key = ($targetDir + "|" + $outName).ToLower()
                        $state = "[[ui.preview_rename]]"
                    } elseif ($rbSkip.Checked) {
                        $state = "[[ui.preview_exists]]"
                    } else {
                        $state = "[[ui.preview_overwrite]]"
                    }
                }
            }
            $used[$key] = $srcName
            Add-RightRow $no ($outName + $state) ([System.IO.Path]::Combine($targetDir, $outName)) ($state -ne "")
        }
    } finally {
        $right.EndUpdate()
        Update-EmptyHints
    }
}

# 各窗自己的统一刷新入口：清单变动后走这里（共享片段内部只认这个名字）
function Refresh-List {
    Renumber-Left
    Update-Preview
    Update-EmptyHints
}

Update-EmptyHints


# ---- 隐藏项跟随资源管理器设置：HKCU\...\Explorer\Advanced 的 Hidden / ShowSuperHidden ----
# Hidden=2 不显示隐藏项；ShowSuperHidden=0 不显示系统项。读不到就都显示，避免「文件突然不见」。
$script:showHidden = $true
$script:showSystem = $true
try {
    $adv = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -ErrorAction Stop
    if ($null -ne $adv.Hidden) { $script:showHidden = ([int]$adv.Hidden -ne 2) }
    if ($null -ne $adv.ShowSuperHidden) { $script:showSystem = ([int]$adv.ShowSuperHidden -ne 0) }
} catch { }

function Test-ItemVisible([System.IO.FileSystemInfo]$info) {
    # 属性读不到（权限/占位）时按可见处理，宁可多显示也不要漏
    try { $attr = $info.Attributes } catch { return $true }
    if (($attr -band [System.IO.FileAttributes]::Hidden) -and -not $script:showHidden) { return $false }
    if (($attr -band [System.IO.FileAttributes]::System) -and -not $script:showSystem) { return $false }
    return $true
}

$script:PH = "__placeholder__"
# 同一目录只枚举一次：重复点击/折叠再展开不再走磁盘
$script:dirCache = @{}

function New-DirNode($parent, [string]$path) {
    # $parent 既可能是 TreeView（顶层常用位置）也可能是 TreeNode；不加类型约束
    # 只加一层：子目录等到展开时才枚举，避免一上来递归整块硬盘
    $node = New-Object System.Windows.Forms.TreeNode([System.IO.Path]::GetFileName($path.TrimEnd('\')))
    if (-not $node.Text) { $node.Text = $path }
    $node.Tag = $path
    [void]$parent.Nodes.Add($node)
    $hasChild = $false
    try {
        foreach ($d in [System.IO.Directory]::GetDirectories($path)) {
            try { $di = New-Object System.IO.DirectoryInfo($d) } catch { continue }
            if (Test-ItemVisible $di) { $hasChild = $true; break }
        }
    } catch { }
    if ($hasChild) {
        $ph = New-Object System.Windows.Forms.TreeNode("...")
        $ph.Tag = $script:PH
        [void]$node.Nodes.Add($ph)
    }
    return $node
}

# 从「此电脑 → 盘符」沿路径逐级建节点并展开，返回目标目录节点（用于「上次位置」记忆）
function Expand-ToPath([string]$path) {
    if (-not $path) { return $null }
    $full = $path.TrimEnd('\')
    if (-not $full) { return $null }
    try {
        if (-not [System.IO.Directory]::Exists($full)) { return $null }
    } catch { return $null }
    $rootPart = [System.IO.Path]::GetPathRoot($full)
    if (-not $rootPart) { return $null }
    $node = $null
    # 盘符节点挂在「此电脑」下面，先查根节点本身，再查它的一级子节点
    foreach ($n in $tree.Nodes) {
        if (([string]$n.Tag).TrimEnd('\') -eq $rootPart.TrimEnd('\')) { $node = $n; break }
        foreach ($c in $n.Nodes) {
            if (([string]$c.Tag).TrimEnd('\') -eq $rootPart.TrimEnd('\')) { $node = $c; break }
        }
        if ($node) { break }
    }
    if (-not $node) { return $null }
    if ($node.Parent) { $node.Parent.Expand() }
    [void]$node.Expand()
    $rest = $full.Substring($rootPart.Length).Trim('\')
    if ($rest) {
        $walk = $rootPart
        foreach ($seg in $rest.Split('\')) {
            if (-not $seg) { continue }
            $walk = [System.IO.Path]::Combine($walk, $seg)
            $child = $null
            foreach ($c in $node.Nodes) {
                if (([string]$c.Tag).TrimEnd('\') -eq $walk.TrimEnd('\')) { $child = $c; break }
            }
            if (-not $child) { return $null }
            $node = $child
            [void]$node.Expand()
        }
    }
    return $node
}

# 顶层：常用位置（桌面/下载/文档/视频）→「此电脑」+ 各盘符；上次位置靠展开定位，不再单独挂节点
$quickSpecs = @(
    @("[[ui.picker_quick_desktop]]", [System.Environment]::GetFolderPath('Desktop')),
    @("[[ui.picker_quick_downloads]]", (Join-Path $env:USERPROFILE 'Downloads')),
    @("[[ui.picker_quick_documents]]", [System.Environment]::GetFolderPath('MyDocuments')),
    @("[[ui.picker_quick_videos]]", [System.Environment]::GetFolderPath('MyVideos'))
)
foreach ($qs in $quickSpecs) {
    $qp = [string]$qs[1]
    if (-not $qp) { continue }
    try { if (-not [System.IO.Directory]::Exists($qp)) { continue } } catch { continue }
    $qNode = New-DirNode $tree $qp
    $qNode.Text = [string]$qs[0]
    $qNode.ToolTipText = $qp
}
$pcNode = New-Object System.Windows.Forms.TreeNode("[[ui.picker_this_pc]]")
$pcNode.Tag = ""
try {
    foreach ($drv in [System.IO.DriveInfo]::GetDrives()) {
        try {
            if (-not $drv.IsReady) { continue }
            $label = $drv.Name
            if ($drv.VolumeLabel) { $label = "$($drv.Name)  $($drv.VolumeLabel)" }
            $node = New-DirNode $pcNode $drv.Name
            $node.Text = $label
        } catch { continue }
    }
} catch { }
[void]$tree.Nodes.Add($pcNode)
$pcNode.Expand()

$tree.add_BeforeExpand({
    param($sender, $e)
    $node = $e.Node
    if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Tag -eq $script:PH) {
        $node.Nodes.Clear()
        $path = $node.Tag
        if ($path -and [System.IO.Directory]::Exists($path)) {
            try {
                $dirKey = $path + "|dirs"
                if ($script:dirCache.ContainsKey($dirKey)) {
                    $dirs = $script:dirCache[$dirKey]
                } else {
                    $dirs = [System.IO.Directory]::GetDirectories($path)
                    $script:dirCache[$dirKey] = $dirs
                }
                foreach ($d in $dirs) {
                    try { $di = New-Object System.IO.DirectoryInfo($d) } catch { continue }
                    if (Test-ItemVisible $di) { [void](New-DirNode $node $d) }
                }
            } catch {
                $lblList.Text = "[[ui.picker_denied]]"
            }
        }
    }
})

function Show-DirFiles([string]$path) {
    $script:curDir = $path
    $list.Items.Clear()
    $lblList.Text = "[[ui.picker_list_label]]"
    if (-not $path) { return }
    # Directory.Exists 对“权限不足”的目录同样返回 False，这里必须显式提示，
    # 否则用户只会看到列表被清空，误以为该目录本来就没有 .prt。
    if (-not [System.IO.Directory]::Exists($path)) {
        $lblList.Text = "[[ui.picker_denied]]"
        # 界面上只给人话 + 路径；完整原因写日志，避免把 .NET 异常原文糊到用户脸上
        try {
            $errLog = Join-Path $env:TEMP "nx_step_export_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: cannot access '{1}' (Directory.Exists=False; missing or permission denied)" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path)
        } catch { }
        # 用 BeginInvoke 排到消息循环里再弹，并指定 owner：提示框始终压在窗体之上
        $hint = "[[ui.picker_denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[ui.picker_title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
        return
    }
    try {
        $dirKey = $path + "|files"
        if ($script:dirCache.ContainsKey($dirKey)) {
            $files = $script:dirCache[$dirKey]
        } else {
            # 缓存全部候选扩展名，方向切换时只做显示层过滤，不重复读盘
            $files = @()
            foreach ($f in [System.IO.Directory]::GetFiles($path, "*.*")) {
                $lf = $f.ToLower()
                if (-not ($lf.EndsWith(".prt") -or $lf.EndsWith(".stp") -or $lf.EndsWith(".step") -or
                    $lf.EndsWith(".igs") -or $lf.EndsWith(".iges") -or
                    $lf.EndsWith(".x_t") -or $lf.EndsWith(".x_b"))) { continue }
                try { $fi = New-Object System.IO.FileInfo($f) } catch { continue }
                if (Test-ItemVisible $fi) { $files += $f }
            }
            $script:dirCache[$dirKey] = $files
        }
        $allow = Get-SourceExts
        $arr = New-Object System.Collections.ArrayList
        foreach ($f in $files) {
            $lf = $f.ToLower()
            $hit = $false
            foreach ($x in $allow) { if ($lf.EndsWith($x)) { $hit = $true; break } }
            if (-not $hit) { continue }
            $fi = New-Object System.IO.FileInfo($f)
            $item = New-Object System.Windows.Forms.ListViewItem($fi.Name)
            $item.ImageIndex = 0
            [void]$item.SubItems.Add($fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
            $extText = $fi.Extension.TrimStart('.').ToUpper()
            if (-not $extText) { $extText = "[[ui.type_file]]" }
            [void]$item.SubItems.Add($extText)
            [void]$item.SubItems.Add([string]([math]::Round($fi.Length / 1KB, 1)))
            $item.Tag = $f
            $item.ToolTipText = $f
            [void]$arr.Add($item)
        }
        if ($arr.Count -gt 0) {
            # 批量添加，避免大目录逐行刷新
            $list.BeginUpdate()
            $list.Items.AddRange([System.Windows.Forms.ListViewItem[]]$arr.ToArray())
            $list.EndUpdate()
        }
        if ($arr.Count -eq 0) { $lblList.Text = "[[ui.picker_empty_dir]]" }
    } catch {
        $lblList.Text = "[[ui.picker_denied]]"
        # 异常原文只进日志（含类型与堆栈），弹窗只保留易懂的提示
        try {
            $errLog = Join-Path $env:TEMP "nx_step_export_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: failed to list .prt in '{1}'`r`n{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path, $_.Exception.ToString())
        } catch { }
        $hint = "[[ui.picker_denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[ui.picker_title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
    }
}

$tree.add_AfterSelect({
    param($sender, $e)
    Show-DirFiles ([string]$e.Node.Tag)
})

$list.add_DoubleClick({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-LeftRow $sel.Tag "FILE"
    }
})

$left.add_MouseClick({
    param($sender, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $hit = $left.HitTest($e.X, $e.Y)
    if (-not $hit.Item) { return }
    if (-not $hit.SubItem) { return }
    if ($hit.Item.SubItems.IndexOf($hit.SubItem) -ne 2) { return }
    $path = [string]$hit.Item.Tag
    $ans = [System.Windows.Forms.MessageBox]::Show("[[ui.msg_remove_text]]`n`n$path", "[[ui.msg_remove_title]]", "YesNo", "Question")
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        # 逐行删除会打乱撤销记录的原始位置，这里清空单步撤销
        $script:lastDeleted = $null
        $left.Items.RemoveAt($hit.Item.Index)
        Refresh-List
        Update-DeleteState
    }
})

$btnAddFolder.add_Click({
    param($sender, $e)
    if ($tree.SelectedNode -and $tree.SelectedNode.Tag) {
        Add-LeftRow $tree.SelectedNode.Tag "DIR"
    }
})

$btnAddFile.add_Click({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-LeftRow $sel.Tag "FILE"
    }
})

$btnClear.add_Click({
    param($sender, $e)
    Clear-LeftRows
})

# 多选删除的按钮 / 快捷键 / 右键菜单接线都在共享片段里：Initialize-LeftList
Initialize-LeftList

# 任一影响结果名的输入变化都重算右栏
$txtPrefix.add_TextChanged({ Update-Preview })
$txtSuffix.add_TextChanged({ Update-Preview })
$txtOutDir.add_TextChanged({ Update-Preview })
$txtOutDir.add_Leave({ Update-Preview })
$rbStep.add_CheckedChanged({ Update-Preview })
$rbIges.add_CheckedChanged({ Update-Preview })
$rbNew.add_CheckedChanged({ Update-Preview })
$rbSkip.add_CheckedChanged({ Update-Preview })
$rbOverwrite.add_CheckedChanged({ Update-Preview })
$rbOutSrc.add_CheckedChanged({ Update-Preview })
$rbOutDir.add_CheckedChanged({ Update-Preview })

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[ui.btn_cancel]]"
$btnCancel.Location = New-Object System.Drawing.Point(742, 660)
$btnCancel.Size = New-Object System.Drawing.Size(80, 28)
$form.Controls.Add($btnCancel)
$btnCancel.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btnOK.Text = "[[ui.btn_ok]]"
$btnOK.Location = New-Object System.Drawing.Point(832, 660)
$btnOK.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnOK)
$btnOK.add_Click({
    param($sender, $e)
    if ($left.Items.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("[[ui.msg_empty_one]]`n[[ui.msg_empty_two]]", "[[ui.msg_empty_title]]", "OK", "Warning")
        return
    }
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# 先建树句柄再定位：展开/选中会触发 BeforeExpand 建子节点、AfterSelect 填充右侧列表，
# 保证窗口显示时列表已是上次目录的内容（而不是空列表）
[void]$tree.Handle

# 上次位置：沿「此电脑 → 盘符 → 逐级子目录」自动展开并选中，右侧列表随即显示该目录
$targetNode = Expand-ToPath $Root
if ($targetNode) {
    try {
        $tree.SelectedNode = $targetNode
        $targetNode.EnsureVisible()
    } catch { }
} else {
    try { $tree.SelectedNode = $pcNode } catch { }
}

$form.add_Shown({
    # 弹出即在前台：只在显示那一刻抢一次，之后点回 NX 不收回、不关闭
    Show-PickerFront $form ($OutFile + ".front") | Out-Null
})

Apply-Theme $form
Update-EmptyHints
[void]$form.ShowDialog()

$folders = @()
$files = @()
$ok = ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
if ($ok) {
    foreach ($row in $left.Items) {
        if ($row.Name -eq "DIR") { $folders += [string]$row.Tag }
        else { $files += [string]$row.Tag }
    }
}

# 选项字段：取消时也照样回传（Python 侧只在 ok=True 时使用，便于“返回修改”保留上次选择）
$outModeValue = "source"
if ($rbOutDir.Checked) { $outModeValue = "dir" }
$onExistsValue = "skip"
if ($rbNew.Checked) { $onExistsValue = "rename" }
elseif ($rbOverwrite.Checked) { $onExistsValue = "overwrite" }
$formatValue = "step"
if ($rbIges.Checked -or $rbXtIgs.Checked) { $formatValue = "iges" }
$directionValue = "prt_step"
if ($rbIges.Checked) { $directionValue = "prt_iges" }
elseif ($rbToPrt.Checked) { $directionValue = "to_prt" }
elseif ($rbXtToPrt.Checked) { $directionValue = "xt_prt" }
elseif ($rbXtStep.Checked) { $directionValue = "xt_step" }
elseif ($rbXtIgs.Checked) { $directionValue = "xt_iges" }
elseif ($rbAny.Checked) { $directionValue = "any_step_igs" }
$workersValue = 2
try { $workersValue = [int]$numWorkers.Value } catch { }
if ($workersValue -lt 1) { $workersValue = 1 }
if ($workersValue -gt 6) { $workersValue = 6 }

$payload = [ordered]@{ ok = $ok; folders = $folders; files = $files;
    out_mode = $outModeValue; out_dir = $txtOutDir.Text.Trim();
    prefix = $txtPrefix.Text.Trim(); suffix = $txtSuffix.Text.Trim();
    on_exists = $onExistsValue; format = $formatValue; workers = $workersValue;
    # ShowDialog 返回后窗体已隐藏，子控件 Visible 恒为 false，不能用它判断；
    # 勾选态与方向单选在关闭后仍保留，只允许 →prt 两个方向回传 true
    open_after = [bool]($cbOpenAfter.Checked -and ($rbToPrt.Checked -or $rbXtToPrt.Checked));
    direction = $directionValue }
$json = $payload | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''


PREVIEW_PS = r'''
param(
    [Parameter(Mandatory=$true)][string]$PlanFile,
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 弹出即抢前台（AttachThreadInput + BringWindowToTop + 瞬时 TopMost），片段在公共执行器里
[[front_helpers]]

$plan = Get-Content -Raw -Encoding UTF8 $PlanFile | ConvertFrom-Json

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[ui.preview_title]]"
$form.ClientSize = New-Object System.Drawing.Size(960, 540)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.TopMost = $false
$uiFont = [System.Drawing.SystemFonts]::MessageBoxFont
try { $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9) } catch { }
$form.Font = $uiFont
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "[[ui.preview_hint]]"
$lbl.Location = New-Object System.Drawing.Point(12, 10)
$lbl.Size = New-Object System.Drawing.Size(936, 22)
$form.Controls.Add($lbl)

# 同名策略只回显一行只读摘要（在上一步选择窗已设定），这里不提供单选
$lblSame = New-Object System.Windows.Forms.Label
$lblSame.AutoSize = $false
$lblSame.ForeColor = [System.Drawing.SystemColors]::GrayText
$lblSame.Location = New-Object System.Drawing.Point(12, 34)
$lblSame.Size = New-Object System.Drawing.Size(936, 20)
$lblSame.Text = [string]$plan.same
$form.Controls.Add($lblSame)

$view = New-Object System.Windows.Forms.ListView
$view.Location = New-Object System.Drawing.Point(12, 58)
$view.Size = New-Object System.Drawing.Size(936, 374)
$view.View = "Details"
$view.FullRowSelect = $true
$view.MultiSelect = $false
$view.GridLines = $true
$view.CheckBoxes = $true
[void]$view.Columns.Add("[[ui.col_cover]]", 90)
[void]$view.Columns.Add("[[ui.col_no]]", 46)
[void]$view.Columns.Add("[[ui.col_src]]", 300)
[void]$view.Columns.Add("[[ui.col_dst]]", 300)
[void]$view.Columns.Add("[[ui.col_state]]", 150)
$view.ShowItemToolTips = $true
$form.Controls.Add($view)

# 只有“目标已存在”的行允许勾选；新建行点了也会被立刻取消，避免误勾
$view.add_ItemCheck({
    param($sender, $e)
    $tag = $view.Items[$e.Index].Tag
    if (-not $tag -or -not $tag.exists) {
        $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
    }
})

$items = New-Object System.Collections.ArrayList
foreach ($row in $plan.rows) {
    $stateText = "[[ui.state_new]]"
    if ($row.state -eq "SKIP") { $stateText = "[[ui.state_skip]]" }
    elseif ($row.state -eq "OVERWRITE") { $stateText = "[[ui.state_overwrite]]" }
    elseif ($row.state -eq "RENAME") { $stateText = "[[ui.state_rename]]" }
    elseif ($row.state -eq "SELF") { $stateText = "[[ui.state_self]]" }
    $item = New-Object System.Windows.Forms.ListViewItem("")
    [void]$item.SubItems.Add([string]$row.no)
    [void]$item.SubItems.Add($row.src)
    [void]$item.SubItems.Add($row.dst)
    [void]$item.SubItems.Add($stateText)
    # bulk = 真正的「覆盖」行（目标已存在且策略为覆盖）；新建/改名/跳过/同格式行只显示、不参与全选
    $item.Tag = @{ exists = [bool]$row.exists
                   bulk = ([bool]$row.exists -and $row.state -eq "OVERWRITE")
                   path = [string]$row.dst_abs }
    $item.ToolTipText = [string]$row.dst_abs
    if ($row.state -eq "NEW") { $item.ForeColor = [System.Drawing.Color]::Gray }
    if ($row.state -eq "SKIP") { $item.ForeColor = [System.Drawing.Color]::DimGray }
    if ($row.state -eq "SELF") { $item.ForeColor = [System.Drawing.Color]::DimGray }
    if ($row.state -eq "RENAME") { $item.ForeColor = [System.Drawing.Color]::DarkGreen }
    if ($row.exists -and $row.override) { $item.Checked = $true }
    [void]$items.Add($item)
}
$view.BeginUpdate()
$view.Items.AddRange([System.Windows.Forms.ListViewItem[]]$items.ToArray())
$view.EndUpdate()

$lblSum = New-Object System.Windows.Forms.Label
$lblSum.Text = $plan.summary
$lblSum.Location = New-Object System.Drawing.Point(12, 440)
$lblSum.Size = New-Object System.Drawing.Size(936, 22)
$form.Controls.Add($lblSum)

$chkAll = New-Object System.Windows.Forms.CheckBox
$chkAll.Text = "[[ui.chk_all]]"
$chkAll.Location = New-Object System.Drawing.Point(12, 466)
$chkAll.Size = New-Object System.Drawing.Size(360, 22)
$chkAll.ThreeState = $true
$form.Controls.Add($chkAll)

# 半选/全选只是「已存在且可覆盖」行的镜像：任一行变化都同步一次
$script:syncingChkAll = $false
function Sync-ChkAll {
    $total = 0
    $checked = 0
    foreach ($item in $view.Items) {
        if ($item.Tag -and $item.Tag.bulk) {
            $total++
            if ($item.Checked) { $checked++ }
        }
    }
    $script:syncingChkAll = $true
    if ($total -eq 0) {
        $chkAll.Enabled = $false
        $chkAll.CheckState = [System.Windows.Forms.CheckState]::Unchecked
    } else {
        $chkAll.Enabled = $true
        if ($checked -eq 0) { $chkAll.CheckState = [System.Windows.Forms.CheckState]::Unchecked }
        elseif ($checked -eq $total) { $chkAll.CheckState = [System.Windows.Forms.CheckState]::Checked }
        else { $chkAll.CheckState = [System.Windows.Forms.CheckState]::Indeterminate }
    }
    $script:syncingChkAll = $false
}

$chkAll.add_Click({
    param($sender, $e)
    if ($script:syncingChkAll) { return }
    $total = 0
    $checked = 0
    foreach ($item in $view.Items) {
        if ($item.Tag -and $item.Tag.bulk) {
            $total++
            if ($item.Checked) { $checked++ }
        }
    }
    if ($total -eq 0) { return }
    # 未全选 → 全选；已全选 → 全部取消；新建/跳过行不动
    $target = -not ($checked -eq $total)
    $script:syncingChkAll = $true
    foreach ($item in $view.Items) {
        if ($item.Tag -and $item.Tag.bulk) { $item.Checked = $target }
    }
    $script:syncingChkAll = $false
    Sync-ChkAll
})
$view.add_ItemChecked({
    param($sender, $e)
    if (-not $script:syncingChkAll) { Sync-ChkAll }
})
Sync-ChkAll

$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Text = "[[ui.btn_back]]"
$btnBack.Location = New-Object System.Drawing.Point(520, 494)
$btnBack.Size = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnBack)
$btnBack.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $form.Close()
})

# 显式的「取消」按钮：与关窗等价（返回取消，不导出、不写盘）
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[ui.btn_cancel]]"
$btnCancel.Location = New-Object System.Drawing.Point(630, 494)
$btnCancel.Size = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnCancel)
$btnCancel.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btnOK.Text = "[[ui.btn_start]]"
$btnOK.Location = New-Object System.Drawing.Point(740, 494)
$btnOK.Size = New-Object System.Drawing.Size(138, 32)
$form.Controls.Add($btnOK)
$btnOK.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

$form.add_Shown({
    # 弹出即在前台：只在显示那一刻抢一次，之后点回 NX 不收回、不关闭
    Show-PickerFront $form ($OutFile + ".front") | Out-Null
})

[void]$form.ShowDialog()

$action = "cancel"
if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) { $action = "ok" }
elseif ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::Retry) { $action = "back" }

$overrides = @()
if ($action -eq "ok") {
    foreach ($item in $view.Items) {
        if (-not ($item.Tag -and $item.Tag.exists)) { continue }
        $overrides += [ordered]@{ path = $item.Tag.path; override = [bool]$item.Checked }
    }
}

$payload = [ordered]@{ action = $action; overrides = $overrides }
$json = $payload | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''


SAVE_PS = r'''
param(
    [Parameter(Mandatory=$true)][string]$ItemsFile,
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$items = Get-Content -Raw -Encoding UTF8 $ItemsFile | ConvertFrom-Json

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[ui.save_title]]"
$form.ClientSize = New-Object System.Drawing.Size(760, 480)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.TopMost = $false
$uiFont = [System.Drawing.SystemFonts]::MessageBoxFont
try { $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9) } catch { }
$form.Font = $uiFont
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "[[ui.save_hint]]"
$lbl.Location = New-Object System.Drawing.Point(12, 10)
$lbl.Size = New-Object System.Drawing.Size(736, 40)
$form.Controls.Add($lbl)

$panel = New-Object System.Windows.Forms.FlowLayoutPanel
$panel.Location = New-Object System.Drawing.Point(12, 54)
$panel.Size = New-Object System.Drawing.Size(736, 356)
$panel.AutoScroll = $true
$panel.FlowDirection = "TopDown"
$panel.WrapContents = $false
$form.Controls.Add($panel)

foreach ($item in $items.items) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "[[ui.save_item_prefix]]" + $item.name
    $chk.Tag = $item.path
    $chk.Size = New-Object System.Drawing.Size(700, 26)
    $chk.Checked = $false
    $chk.ToolTipText = $item.path
    $panel.Controls.Add($chk)
}

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[ui.btn_cancel_export]]"
$btnCancel.Location = New-Object System.Drawing.Point(500, 442)
$btnCancel.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnCancel)
$btnCancel.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btnOK.Text = "[[ui.btn_save_go]]"
$btnOK.Location = New-Object System.Drawing.Point(628, 442)
$btnOK.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnOK)
# 零勾选时先确认，避免误按主按钮把「继续导出」变成静默全部跳过
$btnOK.add_Click({
    param($sender, $e)
    $total = 0
    $picked = 0
    foreach ($ctrl in $panel.Controls) {
        if ($ctrl -is [System.Windows.Forms.CheckBox]) {
            $total++
            if ($ctrl.Checked) { $picked++ }
        }
    }
    if ($picked -eq 0) {
        $msg = "[[ui.save_empty_confirm]]" -f $total
        $choice = [System.Windows.Forms.MessageBox]::Show(
            $msg,
            "[[ui.save_title]]",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})
$form.AcceptButton = $btnOK

[void]$form.ShowDialog()

$ok = ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
$saved = @()
if ($ok) {
    foreach ($ctrl in $panel.Controls) {
        if ($ctrl -is [System.Windows.Forms.CheckBox] -and $ctrl.Checked) {
            $saved += $ctrl.Tag
        }
    }
}

$payload = [ordered]@{ ok = $ok; save = $saved }
$json = $payload | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''
# 并发提示：短提示窗，几秒后自动关闭，避免 journal 结束后窗口一直挂着
BUSY_PS = r'''
param(
    [string]$Title = "",
    [string]$Text = "",
    [string]$OkText = "OK",
    [int]$Seconds = 8
)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = $Title
$form.ClientSize = New-Object System.Drawing.Size(420, 150)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.TopMost = $false
$uiFont = [System.Drawing.SystemFonts]::MessageBoxFont
try { $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9) } catch { }
$form.Font = $uiFont
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = $Text
$lbl.Location = New-Object System.Drawing.Point(16, 18)
$lbl.Size = New-Object System.Drawing.Size(388, 68)
$form.Controls.Add($lbl)

$btn = New-Object System.Windows.Forms.Button
$btn.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btn.Text = $OkText
$btn.Location = New-Object System.Drawing.Point(150, 104)
$btn.Size = New-Object System.Drawing.Size(120, 30)
$btn.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $btn
$form.Controls.Add($btn)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [Math]::Max(1, $Seconds) * 1000
$timer.add_Tick({
    try { $timer.Stop() } catch { }
    $form.Close()
})
$timer.Start()

[void]$form.ShowDialog()
'''


def _has_visible_window(pid):
    """目标进程有没有可见顶层窗口；仅用于诊断窗口被隐藏导致的卡死"""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        user32.EnumWindows.argtypes = (ctypes.c_void_p, wintypes.LPARAM)
        user32.GetWindowThreadProcessId.argtypes = (wintypes.HWND,
                                                    ctypes.POINTER(wintypes.DWORD))
        found = []

        def cb(hwnd, _):
            owner = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if (owner.value == int(pid) and user32.IsWindowVisible(hwnd)
                    and user32.GetWindowTextLengthW(hwnd) > 0):
                found.append(hwnd)
            return True

        user32.EnumWindows(proc(cb), 0)
        return bool(found)
    except Exception:
        return True


def _powershell_path():
    """返回 powershell.exe 的可执行路径；找不到时返回 None"""
    candidates = [
        os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                     "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
        "powershell",
    ]
    for candidate in candidates:
        if os.path.exists(candidate) or candidate == "powershell":
            return candidate
    return None


def run_powershell(script_text, arguments, tag):
    """把 PowerShell 脚本写到临时目录并执行，返回生成的结果 JSON（字典）"""
    import json
    import tempfile
    import subprocess

    temp_dir = tempfile.gettempdir()
    _rotate_log(os.path.join(temp_dir, "nx_step_export_ui_errors.log"))
    ps_path = os.path.join(temp_dir, "nx_step_export_%s.ps1" % tag)
    out_path = os.path.join(temp_dir, "nx_step_export_%s.json" % tag)
    front_path = out_path + ".front"

    with open(ps_path, "w", encoding="utf-8-sig") as handle:
        handle.write(render_ps(script_text))
    for stale in (out_path, front_path):
        try:
            if os.path.exists(stale):
                os.remove(stale)
        except Exception:
            pass

    powershell = _powershell_path()
    if powershell is None:
        raise RuntimeError(T("log.err_no_powershell"))

    command = [powershell, "-NoProfile", "-NonInteractive", "-STA",
               "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
               "-File", ps_path] + list(arguments) + ["-OutFile", out_path]

    kwargs = {"stdout": subprocess.PIPE, "stderr": subprocess.PIPE,
              "creationflags": 0x08000000}
    # 这里绝不能设 STARTUPINFO 的 wShowWindow=0/SW_HIDE：
    # WinForms 窗体第一次 ShowWindow 会被它覆盖，窗口永远不可见，
    # 调用方在 subprocess 上一直等 → NX 表现为「点开功能后卡死」。
    # 控制台黑框由 CREATE_NO_WINDOW 处理，无需再隐藏窗口。
    process = subprocess.Popen(command, **kwargs)
    started = time.time()
    flagged = False
    while process.poll() is None:
        if not flagged and time.time() - started > 8:
            flagged = True
            if not _has_visible_window(process.pid):
                log(T("log.ui_hidden", tag, process.pid))
        time.sleep(0.2)
    stderr = b""
    try:
        if process.stderr is not None:
            stderr = process.stderr.read()
    except Exception:
        pass

    # 抢前台验收结果（窗口自己写的 ok/failed），日志里可判断
    try:
        if os.path.exists(front_path):
            with open(front_path, "r", encoding="ascii") as handle:
                front_state = (handle.read() or "").strip() or "unknown"
            log("[dlg] front=%s tag=%s" % (front_state, tag))
    except Exception:
        pass

    if not os.path.exists(out_path):
        detail = ""
        try:
            detail = (stderr or b"").decode("utf-8", "ignore")[:300]
        except Exception:
            pass
        raise RuntimeError(T("log.err_no_result", detail))

    with open(out_path, "r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def _rotate_log(path, max_bytes=2 * 1024 * 1024, keep=3):
    """日志超过上限时轮转：当前 -> .1，旧的依次后移，最多保留 keep 份"""
    try:
        if (not os.path.exists(path)) or os.path.getsize(path) < max_bytes:
            return
        oldest = "%s.%d" % (path, keep)
        if os.path.exists(oldest):
            force_remove(oldest)
        for index in range(keep - 1, 0, -1):
            src = "%s.%d" % (path, index)
            if os.path.exists(src):
                os.replace(src, "%s.%d" % (path, index + 1))
        os.replace(path, path + ".1")
    except Exception:
        pass


def temp_target(out_path):
    """覆盖导出用的同目录临时文件名：保留扩展名，保证导出器能识别格式"""
    folder, name = os.path.split(out_path)
    stem, ext = os.path.splitext(name)
    return os.path.join(folder, "%s.nxtmp_%d%s" % (stem, os.getpid(), ext))


def temp_sidecar(path):
    """导出器会在产物旁写同名 .log；临时产物的日志同样要改名或清掉，不能留在用户目录"""
    return os.path.splitext(path)[0] + ".log"


def backup_original(path):
    """覆盖前把原产物备份为同名 .bak（已存在则覆盖该 .bak）。

    备份失败必须中止覆盖并保留原文件，不能"没备份就覆盖"。
    停止/回滚时不删除 .bak，只删本次新生成的产物。
    """
    bak = path + ".bak"
    try:
        clear_readonly(bak)
        shutil.copy2(path, bak)
    except Exception as e:
        raise RuntimeError(T("log.err_backup", path, bak, e))
    return bak


def clear_readonly(path):
    """覆盖前清掉目标文件的只读属性：只读目标也能替换，清不掉就让调用方失败"""
    try:
        attrs = os.stat(path).st_file_attributes
    except Exception:
        return
    if attrs & stat.FILE_ATTRIBUTE_READONLY:
        os.chmod(path, stat.S_IWRITE)


def build_jobs(plan, options, overrides, skip_unsaved, base_dir):
    """把计划的每一行解析成可执行任务，前台/后台两种执行方式共用"""
    jobs = []
    total = len(plan)
    for index, (prt_path, out_path, state) in enumerate(plan, 1):
        prefix = "[%d/%d]" % (index, total)
        relative = rel_display(prt_path, base_dir)
        job = {"src": prt_path, "dst": out_path, "action": "do", "msg": "",
               "prefix": prefix, "index": index, "total": total}
        prt_key = os.path.normcase(os.path.abspath(prt_path))
        if prt_key in skip_unsaved:
            job["action"] = "skip"
            job["msg"] = "%s %s: %s" % (prefix, T(skip_unsaved[prt_key]), relative)
        elif state == "SELF":
            job["action"] = "self"
            job["msg"] = T("log.chain_same_fmt", relative)
        else:
            override = overrides.get(os.path.normcase(os.path.abspath(out_path)))
            existing_valid = output_is_valid(out_path, options)
            if (state == "SKIP" and existing_valid and override is not True) or \
                    (state == "OVERWRITE" and override is False):
                job["action"] = "skip"
                job["msg"] = T("log.skip_exists", prefix,
                               rel_display(out_path, base_dir))
        jobs.append(job)
    return jobs


def _import_to_prt(session, src_path, out_path, options, relative, base_dir):
    """→prt 方向：复用导入逻辑把源件转成 .prt，产物名严格用计划算好的路径"""
    try:
        BI = batch_import_module()
    except Exception as e:
        log(T("log.import_unavailable", str(e)))
        return "fail"
    used = set()
    try:
        used.update(BI.existing_part_names(session))
    except Exception:
        pass
    try:
        status, _name, note, _saved, _part = BI.import_one(
            session, src_path, used, save_as=True,
            prefix=safe_name_part(options.get("prefix")),
            suffix=safe_name_part(options.get("suffix")),
            target=out_path)
    except Exception as e:
        log(T("log.import_fail", relative, e))
        return "fail"
    if status == "ok" and output_is_valid(out_path, options):
        log(T("log.import_done", relative,
              rel_display(out_path, base_dir), note or ""))
        return "ok"
    if status == "skip":
        log(T("log.skip_exists", "", rel_display(out_path, base_dir)))
        return "skip"
    log(T("log.import_fail", relative, note or status))
    return "fail"


def export_one(session, job, options, chain_used, ctx, report=None):
    """导出单个文件（前台/后台共用）；返回 ok / fail"""
    if report is None:
        def report(stage):
            return None
    prt_path, out_path = job["src"], job["dst"]
    base_dir = ctx.get("base_dir")
    relative = rel_display(prt_path, base_dir)
    prefix = job.get("prefix") or ""
    src_ext = os.path.splitext(prt_path)[1].lower()
    fmt = output_format(options)
    item_parts = []

    # 选择/预览与真正导出之间有时间差，文件可能已被删除或移走
    if len(prt_path) > 259:
        log(T("log.too_long_path", relative))
        return "fail"
    if not os.path.isfile(prt_path):
        log(T("log.missing_part", relative))
        return "fail"
    # NX 无法打开含 emoji（非 BMP 字符）的文件路径，提前给出明确提示
    if any(ord(ch) > 0xFFFF for ch in relative):
        log(T("log.unsupported_name", relative))
        return "fail"

    start = time.time()
    log(T("log.progress", job.get("index"), job.get("total"), relative))
    if fmt == "prt":
        # →prt：整条链路交给导入侧（STEP/IGES/Parasolid → 新 .prt），不走导出器
        return _import_to_prt(session, prt_path, out_path, options, relative, base_dir)
    layer_mask = "1-256"
    try:
        if fmt == "step" and VISIBLE_LAYERS_ONLY and src_ext == ".prt":
            # 图层可见性只能在显示部件上查询；脚本自己打开的部件会在结束时关闭
            report("open")
            part = ensure_displayed(session, prt_path, item_parts)
            layer_mask, view_name, visible_count = visible_layer_mask(part)
            if not layer_mask:
                log(T("log.no_visible", view_name))
                return "fail"
            log(T("log.visible", layer_mask, view_name, visible_count))

        # 覆盖导出：先把原产物备份为同名 .bak，再写同目录临时文件、成功后原子替换
        write_path = out_path
        if os.path.exists(out_path):
            bak = backup_original(out_path)
            log(T("log.backup", prefix, rel_display(bak, base_dir)))
            if output_is_valid(out_path, options):
                log(T("log.cover", prefix, relative))
                write_path = temp_target(out_path)
            else:
                # 目标存在但不是有效产物（只有文件头等）：先移除，避免转换器拒绝写入
                try:
                    os.remove(out_path)
                except Exception as e:
                    raise RuntimeError(T("log.err_remove_target", out_path, e,
                                         os.path.splitext(out_path)[1] or "PART"))
        report("export")
        try:
            if src_ext in CHAIN_SOURCE_EXTS:
                # 链式转换：内存建件 → 导入几何 → 直接导出目标格式（不落中间 .prt）
                size = chain_convert(session, prt_path, write_path, fmt,
                                     chain_used, layer_mask)
            elif fmt == "iges":
                size = export_iges(prt_path, write_path)
            else:
                size = export_step(session, prt_path, write_path, layer_mask)
            if write_path != out_path:
                # 同盘 rename 是原子的：要么旧文件完好，要么新文件已完整落盘
                # 只读目标先清属性再替换；替换失败时旧文件内容仍保持完好
                clear_readonly(out_path)
                try:
                    os.replace(write_path, out_path)
                except PermissionError:
                    # 被占用/仍只读：给用户一句能照做的提示，别只丢 WinError 原文
                    raise RuntimeError(T("log.err_target_locked",
                                         rel_display(out_path, base_dir)))
                # 导出器按临时文件名写了 .log；改名回正常名字，别把 .nxtmp_* 留在用户目录
                stray = temp_sidecar(write_path)
                if os.path.exists(stray):
                    try:
                        os.replace(stray, temp_sidecar(out_path))
                    except Exception:
                        try:
                            os.remove(stray)
                        except Exception:
                            pass
        except Exception:
            if write_path != out_path:
                for path in (write_path, temp_sidecar(write_path)):
                    try:
                        force_remove(path)
                    except Exception:
                        pass
            raise
        report("verify")
        log(T("log.done", time.time() - start,
              rel_display(out_path, base_dir), size))
        return "ok"
    except Exception as e:
        log(T("log.failed", e))
        for line in traceback.format_exc().splitlines():
            log("      " + line)
        return "fail"
    finally:
        if item_parts:
            # 先切回用户原状态，再逐件关闭“本件由脚本新打开”的部件
            _restore_state(session, ctx.get("original_display_part"),
                           ctx.get("original_work_part"))
            close_opened_parts(session, item_parts,
                               ctx.get("original_keys") or set())
            del item_parts[:]


def _load_runner():
    """公共后台执行器接口（原 nx_batch_runner.py 已并入本文件）"""
    return _runner_api()


def _start_background(jobs, total, options, log_path, base_dir, nx_pid=0):
    """把清单交给独立 NX 批处理会话执行；失败返回 False，调用方回落前台循环"""
    auto_open = (bool((options or {}).get("open_after", False))
                 and to_prt_direction(options))
    log("[async] start direction=%s open_after=%d auto_open=%d workers=%s jobs=%d"
        % ((options or {}).get("direction"), 1 if (options or {}).get("open_after") else 0,
           1 if auto_open else 0, (options or {}).get("workers"), len(jobs or [])))
    try:
        runner = _load_runner()
        return runner.start("export", os.path.abspath(__file__), jobs, total,
                            options, log_path, LANG, _lock_path(), base_dir,
                            extras={"workers": (options or {}).get("workers"),
                                    "nx_pid": int(nx_pid or 0),
                                    "auto_open": auto_open})
    except Exception as e:
        log(T("log.background_fail", e))
        return False


def _hand_off_lock():
    """后台运行期间把锁交给进度窗接管（由它改写 PID 并在结束时删除）"""
    global _RUN_LOCK
    _RUN_LOCK = None


def _products_path(plan_path):
    """每个后台分片把本片产物清单写到 plan 同目录的 products_N.json

    按分片落盘，进度窗收尾时汇总，不会与其它批次/分片互相覆盖。
    """
    name = os.path.basename(plan_path)
    if name.startswith("plan_"):
        name = "products_" + name[5:]
    else:
        name = "products_" + name
    return os.path.join(os.path.dirname(plan_path), name)


def _run_background_worker(plan_path):
    """独立 NX 会话里的执行入口：只跑清单，不弹任何窗"""
    global _log_handle
    runner = _load_runner()
    plan = runner.load_plan(plan_path)
    try:
        _log_handle = open(plan.get("log_path"), "a", encoding="utf-8")
    except Exception:
        _log_handle = None
    session = NXOpen.Session.GetSession()
    options = plan.get("options") or default_options()
    chain_used = set()
    ctx = {"base_dir": plan.get("base_dir"), "original_display_part": None,
           "original_work_part": None, "original_keys": set()}

    # 本片产物台账：每处理一项就落盘。停止/回滚时即使进程被强杀，也能按这份清单逐项清理。
    # 只登记 action=do 的任务；跳过项的目标文件不是本批次产物，永远不进清单。
    fmt = output_format(options)
    records = []
    by_dst = {}
    for item in plan.get("items") or []:
        if item.get("action") != "do":
            continue
        rec = {"src": item.get("src"), "dst": item.get("dst"), "status": "pending",
               "preexisting": False, "tmp": []}
        records.append(rec)
        try:
            by_dst[os.path.normcase(os.path.abspath(rec["dst"]))] = rec
        except Exception:
            pass

    def snapshot():
        try:
            runner.write_json(_products_path(plan_path), {"fmt": fmt, "items": records})
        except Exception:
            pass

    def handler(item, report):
        if item.get("action") != "do":
            if item.get("msg"):
                log(item["msg"])
            return ("skip", "")
        rec = None
        try:
            rec = by_dst.get(os.path.normcase(os.path.abspath(item.get("dst"))))
        except Exception:
            rec = None
        if rec is not None:
            # 处理前先记录目标是否已存在：覆盖模式下这就是“原有文件已被替换”，回滚时必须保留
            rec["preexisting"] = os.path.exists(rec["dst"])
            tmp = temp_target(rec["dst"])
            rec["tmp"] = [tmp, temp_sidecar(tmp)]
            rec["status"] = "running"
            snapshot()
        status = export_one(session, item, options, chain_used, ctx, report)
        if rec is not None:
            rec["status"] = status
            if status == "ok":
                if fmt == "prt":
                    last_ok.append(rec["dst"])
                elif not rec["preexisting"]:
                    # 导出器在产物旁写同名 .log，同样属于本批次产物；覆盖模式保留原文件时不追加
                    rec["tmp"].append(temp_sidecar(rec["dst"]))
            snapshot()
        elif status == "ok" and fmt == "prt":
            last_ok.append(item.get("dst"))
        return (status, "")

    last_ok = []
    snapshot()
    results, state = runner.execute(plan, handler, log)
    snapshot()
    failed = [rel_display(r.get("src"), ctx["base_dir"])
              for r in results if r.get("status") == "fail"]
    log("=" * 60)
    log(T("log.summary", state.get("ok", 0), state.get("skip", 0), len(failed)))
    for name in failed:
        log(T("log.failed_item", name))
    log("=" * 60)
    if _log_handle is not None:
        _log_handle.close()
        _log_handle = None


def spawn_powershell(script_text, arguments, tag):
    """启动 PowerShell 但不等待退出（用于短提示窗：journal 结束窗口仍在，定时自动关闭）"""
    import tempfile
    import subprocess

    temp_dir = tempfile.gettempdir()
    _rotate_log(os.path.join(temp_dir, "nx_step_export_ui_errors.log"))
    ps_path = os.path.join(temp_dir, "nx_step_export_%s.ps1" % tag)
    with open(ps_path, "w", encoding="utf-8-sig") as handle:
        handle.write(render_ps(script_text))
    powershell = _powershell_path()
    if powershell is None:
        return False
    command = [powershell, "-NoProfile", "-NonInteractive", "-STA",
               "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
               "-File", ps_path] + list(arguments)
    try:
        process = subprocess.Popen(command, stdin=subprocess.DEVNULL,
                                   stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL,
                                   creationflags=0x08000000)  # CREATE_NO_WINDOW
    except TypeError:
        process = subprocess.Popen(command, stdin=subprocess.DEVNULL,
                                   stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL)
    return True


def notify_busy():
    """并发冲突时在 NX 里给一个短提示（8 秒自动关，可手动点确定）；失败静默"""
    try:
        spawn_powershell(BUSY_PS, [
            "-Title", T("ui.busy_title"),
            "-Text", T("ui.busy_text"),
            "-OkText", T("ui.busy_ok"),
            "-Seconds", "8",
        ], "busy")
    except Exception:
        pass


def _skip_preview(plan, options):
    """→prt 方向只有「同名=覆盖且目标已存在」时才需要确认；其余方向保持原有预览"""
    if not to_prt_direction(options):
        return False
    return not any(state == "OVERWRITE" and os.path.exists(out_path)
                   for _src, out_path, state in plan)


def preview_plan(plan, base_dir, options=None):
    """导出前弹出清单确认窗，返回 (action, overrides)

    plan 的每一项是 (源文件, 输出路径, 行状态)，状态在 build_plan 里按同名策略算好：
    NEW / SKIP / RENAME / OVERWRITE。

    action: "ok" / "back" / "cancel"
    overrides: {规范化绝对输出路径: True/False}，只包含窗口里可勾选的行（目标已存在）；
               没出现在字典里的行沿用该行状态。
    """
    import json
    import tempfile

    rows = []
    new_count = 0
    skip_count = 0
    for index, (prt_path, out_path, state) in enumerate(plan, 1):
        exists = os.path.exists(out_path)
        if state in ("SKIP", "SELF"):
            skip_count += 1
        else:
            new_count += 1
        rows.append({
            "no": index,
            "src": rel_display(prt_path, base_dir),
            # 预览里显示的是拼接前缀/后缀、并按同名策略算好的最终文件名
            "dst": rel_display(out_path, base_dir),
            "dst_abs": os.path.abspath(out_path),
            "state": state,
            "exists": exists,
            # 初始勾选状态：只有“覆盖”策略的行默认勾上，其余保持不勾
            "override": bool(exists and state == "OVERWRITE"),
        })

    summary = T("ui.summary", len(rows), new_count, skip_count)

    # 同名策略已在上一步选择窗定过：这里只回显一行只读摘要，不再给单选
    policy_key = {"rename": "ui.on_exists_rename",
                  "skip": "ui.on_exists_skip",
                  "overwrite": "ui.on_exists_overwrite"}.get(
                      str((options or {}).get("on_exists") or "skip"), "ui.on_exists_skip")
    same_text = T("ui.same_summary", T(policy_key))

    plan_path = os.path.join(tempfile.gettempdir(), "nx_step_export_plan.json")
    with open(plan_path, "w", encoding="utf-8-sig") as handle:
        json.dump({"rows": rows, "summary": summary, "same": same_text},
                  handle, ensure_ascii=False)

    data = run_powershell(PREVIEW_PS, ["-PlanFile", plan_path], "preview") or {}
    overrides = {}
    for item in data.get("overrides") or []:
        try:
            key = os.path.normcase(os.path.abspath(item.get("path")))
            overrides[key] = bool(item.get("override"))
        except Exception:
            continue
    return data.get("action", "cancel"), overrides


def collect_modified_open(session, plan):
    """返回计划里“已在当前会话打开且 IsModified=True”的项

    判定只看内存中 Part.IsModified，不看文件时间戳：
    用户可能只改了图层、显示状态，时间戳没变，但导出前仍应让用户决定是否保存。
    """
    open_parts = {}
    try:
        for part in session.Parts:
            try:
                full_path = part.FullPath
            except Exception:
                continue
            if full_path:
                open_parts[os.path.normcase(os.path.abspath(full_path))] = part
    except Exception:
        return []

    items = []
    seen = set()
    for prt_path, _out_path, _state in plan:
        key = os.path.normcase(os.path.abspath(prt_path))
        if key in seen or key not in open_parts:
            continue
        part = open_parts[key]
        try:
            modified = bool(part.IsModified)
        except Exception:
            modified = False
        if modified:
            seen.add(key)
            items.append({"path": os.path.abspath(prt_path),
                          "name": os.path.basename(prt_path),
                          "part": part})
    return items


def collect_modified_all(session):
    """启动前的轻量检查：当前会话里所有“已打开且有未保存修改”的部件

    A1 起，导出在后台会话里跑，此时还没选文件、拿不到计划，
    所以先按“全部已打开部件”检查一次并给一次提示，避免静默丢失修改。
    """
    items = []
    try:
        parts = list(session.Parts)
    except Exception:
        return items
    for part in parts:
        try:
            if not bool(part.IsModified):
                continue
            full_path = part.FullPath or ""
        except Exception:
            continue
        if not full_path:
            continue
        items.append({"path": os.path.abspath(full_path),
                      "name": os.path.basename(full_path),
                      "part": part})
    return items


def ask_save_modified(items):
    """弹一次窗口列出所有“已打开且有未保存修改”的文件，返回 (是否继续, 保存集合)

    默认全部不勾选（= 跳过、不写盘）；用户勾选的文件才保存。
    关闭窗口按“取消本次导出”处理。
    """
    import json
    import tempfile

    payload = {"items": [{"path": item["path"], "name": item["name"]}
                         for item in items]}
    # 注意：这里必须是独立文件名。run_powershell 启动前会删除“输出文件”
    # （%TEMP%\nx_step_export_unsaved.json），若输入清单也叫这个名字，
    # 会在 PowerShell 读取前被删掉，对话框就一个勾选框都建不出来，
    # 用户勾了“保存”也只会返回空列表，进而被当成“跳过”。
    items_path = os.path.join(tempfile.gettempdir(),
                              "nx_step_export_unsaved_items.json")
    with open(items_path, "w", encoding="utf-8-sig") as handle:
        json.dump(payload, handle, ensure_ascii=False)

    data = run_powershell(SAVE_PS, ["-ItemsFile", items_path], "unsaved") or {}
    if not data.get("ok"):
        return False, set()
    save_keys = set()
    for path in data.get("save") or []:
        try:
            save_keys.add(os.path.normcase(os.path.abspath(path)))
        except Exception:
            continue
    return True, save_keys


def parse_picker_options(payload, previous=None):
    """把选择窗口回传的导出选项规范化；缺字段或非法值沿用上次值/默认值"""
    result = dict(previous or default_options())
    if not isinstance(payload, dict):
        return result
    mode = str(payload.get("out_mode") or "").strip().lower()
    if mode in ("source", "dir"):
        result["out_mode"] = mode
    result["out_dir"] = str(payload.get("out_dir") or "").strip()
    result["prefix"] = str(payload.get("prefix") or "")
    result["suffix"] = str(payload.get("suffix") or "")
    policy = str(payload.get("on_exists") or "").strip().lower()
    if policy in ("skip", "rename", "overwrite"):
        result["on_exists"] = policy
    fmt = str(payload.get("format") or "").strip().lower()
    if fmt in ("step", "iges"):
        result["format"] = fmt
    direction = str(payload.get("direction") or "").strip().lower()
    if direction in DIRECTIONS:
        result["direction"] = direction
    try:
        workers = int(payload.get("workers") or result.get("workers") or 2)
    except Exception:
        workers = 2
    result["workers"] = max(1, min(6, workers))
    open_after = payload.get("open_after")
    if isinstance(open_after, bool):
        result["open_after"] = open_after
    return result


def _picker_options(options=None):
    """选择窗参数的来源：默认值 < 记忆文件 < 调用方传入"""
    opts = dict(default_options())
    opts.update(read_options())
    opts.update(options or {})
    return opts


def _picker_args(base_dir, opts):
    """拼选择窗命令行参数（同步/异步两条路共用，避免两处漂移）"""
    args = []
    if base_dir:
        args += ["-Root", base_dir]
    # 把上次的选项回传，用户点「返回修改」时窗口里保留上次填写的内容
    # 注意：PowerShell 会丢弃空字符串参数（会报 “Missing an argument”），
    # 所以值为空时干脆不传，让脚本里的默认值（空串）生效。
    args += ["-OutMode", opts["out_mode"],
             "-OnExists", opts["on_exists"],
             "-Format", opts["format"],
             "-Direction", opts.get("direction") or "prt_step",
             "-Workers", str(opts.get("workers") or 2),
             "-OpenAfter", "1" if opts.get("open_after", False) else "0"]
    for flag, key in (("-OutDir", "out_dir"),
                      ("-Prefix", "prefix"),
                      ("-Suffix", "suffix")):
        value = str(opts.get(key) or "")
        if value:
            args += [flag, value]
    return args


def _apply_picker_result(data, opts):
    """把选择窗结果写进记忆文件，返回 (folders, files, options)

    folders: None = 没选文件夹；[绝对路径...] = 勾选的文件夹
    """
    folders = data.get("folders") or []
    files = data.get("files") or []
    if isinstance(folders, str):
        folders = [folders] if folders else []
    if isinstance(files, str):
        files = [files] if files else []

    chosen_dirs = []
    for folder in folders:
        path = os.path.abspath(folder)
        if os.path.isdir(path):
            chosen_dirs.append(path)
        else:
            log(T("log.ignore_outside", folder))

    options = parse_picker_options(data, opts)
    # 确认后记住本次选项，并把「最后一个目标文件所在目录」作为下次展开位置
    # （没有文件时退回最后一个文件夹）；取消时上面已提前返回，不会写
    write_options(options)
    picked_files = [os.path.abspath(p) for p in files if p]
    if picked_files:
        last_dir = os.path.dirname(picked_files[-1])
    elif chosen_dirs:
        last_dir = chosen_dirs[-1]
    else:
        last_dir = ""
    if last_dir:
        write_last_dir(last_dir)
    if chosen_dirs:
        return chosen_dirs, files, options
    return None, files, options


def choose_targets(base_dir=None, options=None):
    """弹出 PowerShell 选择窗口（默认从「此电脑」+ 所有盘符开始）

    返回 (folders, files, options)

    folders: None = 没选文件夹；[绝对路径...] = 勾选的文件夹
    files:   绝对路径列表
    options: 导出选项字典（out_mode / out_dir / prefix / suffix / on_exists / format）
    """
    opts = _picker_options(options)
    data = run_powershell(SELECT_UI_PS, _picker_args(base_dir, opts), "picker")

    if not data or not data.get("ok"):
        return None, [], opts

    return _apply_picker_result(data, opts)


def rel_display(path, base_dir=None):
    """日志里显示的路径：有根目录时相对显示，否则显示绝对路径"""
    if base_dir:
        try:
            return os.path.relpath(path, base_dir)
        except ValueError:
            return path
    return os.path.abspath(path) if path else path


def pick_folder(title):
    """弹出 Windows 文件夹选择框，返回所选路径；用户取消返回 None"""
    import ctypes
    from ctypes import wintypes

    class BROWSEINFOW(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND),
            ("pidlRoot", ctypes.c_void_p),
            ("pszDisplayName", wintypes.LPWSTR),
            ("lpszTitle", wintypes.LPCWSTR),
            ("ulFlags", wintypes.UINT),
            ("lpfn", ctypes.c_void_p),
            ("lParam", ctypes.c_void_p),
            ("iImage", ctypes.c_int),
        ]

    shell32 = ctypes.windll.shell32
    user32 = ctypes.windll.user32
    shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
    shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFOW)]
    shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR]

    display_name = ctypes.create_unicode_buffer(260)
    info = BROWSEINFOW()
    info.hwndOwner = user32.GetForegroundWindow()
    info.pszDisplayName = ctypes.cast(display_name, wintypes.LPWSTR)
    info.lpszTitle = title
    info.ulFlags = 0x00000001 | 0x00000040   # 只列文件系统目录 + 新式可调整对话框

    pidl = shell32.SHBrowseForFolderW(ctypes.byref(info))
    if not pidl:
        return None
    try:
        path_buffer = ctypes.create_unicode_buffer(260)
        if not shell32.SHGetPathFromIDListW(pidl, path_buffer):
            return None
        return path_buffer.value
    finally:
        ctypes.windll.ole32.CoTaskMemFree(ctypes.c_void_p(pidl))


def pick_files(title, initial_dir):
    """弹出 Windows 多选文件对话框，返回选中的文件绝对路径列表；取消返回 []"""
    import ctypes
    from ctypes import wintypes

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD),
            ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE),
            ("lpstrFilter", wintypes.LPCWSTR),
            ("lpstrCustomFilter", wintypes.LPWSTR),
            ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD),
            ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD),
            ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD),
            ("lpstrInitialDir", wintypes.LPCWSTR),
            ("lpstrTitle", wintypes.LPCWSTR),
            ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD),
            ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", wintypes.LPCWSTR),
            ("lCustData", wintypes.LPARAM),
            ("lpfnHook", ctypes.c_void_p),
            ("lpTemplateName", wintypes.LPCWSTR),
            ("pvReserved", ctypes.c_void_p),
            ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buffer_size = 65536
    file_buffer = ctypes.create_unicode_buffer(buffer_size)
    dialog = OPENFILENAMEW()
    dialog.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    dialog.hwndOwner = ctypes.windll.user32.GetForegroundWindow()
    dialog.lpstrFilter = u"%s\0*.prt;*.stp;*.step;*.igs;*.iges\0%s\0*.*\0\0" % (
        T("ui.filter_prt"), T("ui.filter_all"))
    dialog.lpstrFile = ctypes.cast(file_buffer, wintypes.LPWSTR)
    dialog.nMaxFile = buffer_size
    dialog.lpstrTitle = title
    dialog.lpstrInitialDir = initial_dir
    dialog.lpstrDefExt = u"prt"
    # 允许多选 + 资源管理器风格 + 文件必须存在 + 隐藏只读勾选
    dialog.Flags = 0x00000200 | 0x00080000 | 0x00001000 | 0x00000800 | 0x00000004

    if not ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(dialog)):
        return []

    # 单选时返回完整路径；多选时返回 "目录\0文件1\0文件2\0\0"
    raw = ctypes.wstring_at(ctypes.addressof(file_buffer), buffer_size)
    parts = [p for p in raw.split(u"\0") if p]
    if not parts:
        return []
    if len(parts) == 1:
        return [parts[0]]
    folder = parts[0]
    return [os.path.join(folder, name) for name in parts[1:]]


def to_rel_dir(folder, base_dir=None):
    """无根目录时返回绝对路径；有根目录时返回相对路径，超出根目录返回 None"""
    folder = os.path.abspath(folder)
    if not base_dir:
        return folder
    try:
        rel = os.path.relpath(folder, base_dir)
    except ValueError:
        return None
    if rel == ".." or rel.startswith(".." + os.sep):
        return None
    return rel


def select_folders(base_dir=None):
    """弹窗逐个选择文件夹（系统对话框，仅在选择窗口打不开时用），取消即结束

    返回绝对路径列表；一个都没选则返回 None
    """
    import ctypes
    ole32 = ctypes.windll.ole32
    com_ready = ole32.CoInitialize(None) in (0, 1)   # S_OK / S_FALSE

    chosen = []
    try:
        while True:
            folder = pick_folder(T("ui.folder_pick_title"))
            if not folder:
                break

            rel = to_rel_dir(folder, base_dir)
            if rel is None:
                log(T("log.ignore_outside", folder))
                continue

            if rel != ".":
                # 已选目录的子目录不再重复添加
                if any(rel == c or rel.startswith(c + os.sep) for c in chosen):
                    log(T("log.already_included", rel))
                    continue
                # 选了上级目录，把此前选的子目录收掉
                chosen = [c for c in chosen if not c.startswith(rel + os.sep)]

            if rel in chosen:
                continue
            chosen.append(rel)
            log(T("log.selected", T("log.root_all") if rel == "." else rel))
    finally:
        if com_ready:
            ole32.CoUninitialize()

    if not chosen:
        return None
    return chosen


def match_any(text, patterns):
    """text 命中任意一个通配符规则即返回 True"""
    text = text.lower().replace("/", os.sep)
    for pattern in patterns:
        if fnmatch.fnmatch(text, pattern.lower().replace("/", os.sep)):
            return True
    return False


def dir_match(rel_dir, patterns):
    """rel_dir 自身、它的任一路径层级、或其中任一文件夹名命中规则即算匹配

    例如 INCLUDE_DIRS=["box1"] 能覆盖 box1、box1/sub；
    EXCLUDE_DIRS=["old*"] 能排除根下的 old 和 box2/old_data。
    """
    if not patterns:
        return False
    candidates = ["."]
    if rel_dir not in ("", "."):
        parts = rel_dir.split(os.sep)
        for i in range(len(parts)):
            candidates.append(os.sep.join(parts[:i + 1]))
        candidates.extend(parts)   # 单独文件夹名也算，便于匹配任意层级的同名目录
    return any(match_any(c, patterns) for c in candidates)


# 链式转换的源扩展名：非 .prt 源在“批量导出”里自动走“内存导入 → 直接导出目标格式”
# x_t/x_b 是 Parasolid 文本/二进制，导入链路与 stp/igs 一致（由 batch_import.import_one 分流）
CHAIN_SOURCE_EXTS = (".stp", ".step", ".igs", ".iges", ".x_t", ".x_b")
# 文件对话框/目录列表可选的源类型（.prt 走原有导出，其余走链式转换）
SELECTABLE_EXTS = (".prt",) + CHAIN_SOURCE_EXTS

# 方向 → 允许的源扩展名（x_t/x_b→stp/igs 已接线，走内存导入后直接导出）
SOURCE_EXTS = {
    "prt_step": (".prt",),
    "prt_iges": (".prt",),
    "to_prt": (".stp", ".step", ".igs", ".iges"),
    "xt_prt": (".x_t", ".x_b"),
    "xt_step": (".x_t", ".x_b"),
    "xt_iges": (".x_t", ".x_b"),
    "any_step_igs": (".prt", ".stp", ".step", ".igs", ".iges"),
}


def source_exts(options):
    direction = str((options or {}).get("direction") or "prt_step")
    return SOURCE_EXTS.get(direction, (".prt",))


def find_prt_files(base_dir=None, include_dirs=None, picked_files=None, scan_dirs=True,
                   exts=None):
    """收集 .prt 并按 EXCLUDE_DIRS / INCLUDE_FILES / EXCLUDE_FILES 筛选

    base_dir     : 已废弃，保留参数只为兼容旧调用；现在为 None
    include_dirs : None = 用配置里的 INCLUDE_DIRS；[绝对路径] = 只收这些目录（含子目录）
    picked_files : 手动点选的文件（绝对路径），不受文件夹/文件名规则限制
    scan_dirs    : False = 不遍历目录，只导出 picked_files
    """
    include_dirs = list(INCLUDE_DIRS if include_dirs is None else include_dirs)
    picked_files = [os.path.abspath(p) for p in (picked_files or [])]
    exts = tuple(exts or (".prt",))

    def _walk_error(err):
        # 路径过长/权限不足时 os.walk 会跳过该目录，这里显式记一条，避免静默漏导
        try:
            log(T("log.walk_fail", getattr(err, "filename", "") or ""))
        except Exception:
            pass

    found = []
    if not scan_dirs:
        pass   # 只导手动点选的文件，不扫目录
    else:
        roots = include_dirs if include_dirs else ([base_dir] if base_dir else [])
        for root_dir in roots:
            if RECURSIVE:
                for root, _dirs, names in os.walk(root_dir, onerror=_walk_error):
                    for name in names:
                        if name.lower().endswith(exts):
                            found.append(os.path.join(root, name))
            else:
                for ext in exts:
                    for path in glob.glob(os.path.join(root_dir, "*" + ext)):
                        if os.path.isfile(path):
                            found.append(path)

    selected = []
    for path in found:
        rel_path = os.path.relpath(path, base_dir) if base_dir else os.path.abspath(path)
        rel_dir = os.path.dirname(rel_path)
        file_name = os.path.basename(path)

        # 文件夹过滤：排除优先，其次必须落在包含范围内（空 = 不限）
        if dir_match(rel_dir, EXCLUDE_DIRS):
            continue
        if include_dirs and not dir_match(rel_dir, include_dirs):
            continue

        # 文件过滤：文件名或相对路径命中即可
        if EXCLUDE_FILES and (match_any(file_name, EXCLUDE_FILES)
                              or match_any(rel_path, EXCLUDE_FILES)):
            continue
        if INCLUDE_FILES and not (match_any(file_name, INCLUDE_FILES)
                                  or match_any(rel_path, INCLUDE_FILES)):
            continue

        selected.append(path)

    # 手动点选的具体文件优先，不受文件夹/文件名规则限制；只收当前方向允许的源类型
    for path in picked_files:
        if path.lower().endswith(exts) and os.path.isfile(path) \
                and path not in selected:
            selected.append(path)

    selected.sort(key=lambda p: p.lower())
    return selected


def build_plan(files, options=None):
    """生成 (源文件, 输出路径, 行状态) 列表

    - 输出目录：options["out_mode"]=="dir" 时统一落到 options["out_dir"]（扁平放置，
      不镜像原目录结构）；否则仍是各 .prt 所在目录（默认，与改动前一致）
    - 文件名：前缀 + 原名 + 后缀 + 扩展名（STEP -> .step，IGES -> .igs）
    - 同名策略 options["on_exists"]：
        skip      = 已存在有效产物则标记 SKIP（只有文件头、没有几何的 STEP 不算已存在）
        rename    = 已存在就追加 _2/_3…直到不冲突，标记 RENAME
        overwrite = 照原名输出，标记 OVERWRITE（导出时写临时文件再原子替换，不备份）
    - 同一批次内两件原始文件撞到同一输出名时，仍按原逻辑追加 _2/_3（标记 RENAME）
    """
    options = options or {}
    ext = output_ext(options)
    out_mode = options.get("out_mode") or "source"
    out_dir = options.get("out_dir") or ""
    prefix = safe_name_part(options.get("prefix"))
    suffix = safe_name_part(options.get("suffix"))
    on_exists = options.get("on_exists") or "skip"

    plan = []
    used = {}   # (目录小写, 输出名小写) -> 来源文件名
    for prt_path in files:
        src_dir = os.path.dirname(prt_path)
        target_dir = out_dir if (out_mode == "dir" and out_dir) else src_dir
        src_name = os.path.basename(prt_path)
        src_ext = os.path.splitext(src_name)[1].lower()
        stem = normalize_name(os.path.splitext(src_name)[0])
        base_name = prefix + stem + suffix
        out_name = base_name + ext
        # 只有“源与目标落在同一个文件”时才跳过（同目录、同格式、同名）：
        # 例如 igs→igs 且未指定其它输出目录。指定了不同输出目录时允许真转换，
        # 否则“igs 转 igs 到别处”会被误判为 SELF 而永远没有产物。
        if src_ext == ext and os.path.normcase(
                os.path.abspath(os.path.join(target_dir, out_name))) == \
                os.path.normcase(os.path.abspath(prt_path)):
            log(T("log.chain_same_fmt", src_name))
            plan.append((prt_path, prt_path, "SELF"))
            continue
        state = "NEW"
        key = (target_dir.lower(), out_name.lower())
        if key in used:
            # 同一批次内的撞名：沿用原来的“追加 _2/_3”，避免同目录内互相覆盖
            serial = 2
            while (target_dir.lower(),
                   ("%s_%d%s" % (base_name, serial, ext)).lower()) in used:
                serial += 1
            new_name = "%s_%d%s" % (base_name, serial, ext)
            log(T("log.rename_conflict", src_name, used[key], out_name, new_name))
            out_name = new_name
            state = "RENAME"
        elif os.path.exists(os.path.join(target_dir, out_name)):
            target_path = os.path.join(target_dir, out_name)
            if on_exists == "rename":
                serial = 2
                while True:
                    candidate = "%s_%d%s" % (base_name, serial, ext)
                    if (target_dir.lower(), candidate.lower()) not in used and \
                            not os.path.exists(os.path.join(target_dir, candidate)):
                        break
                    serial += 1
                log(T("log.rename_policy", out_name, candidate))
                out_name = candidate
                state = "RENAME"
            elif on_exists == "overwrite":
                state = "OVERWRITE"
            elif output_is_valid(target_path, options):
                state = "SKIP"

        used[(target_dir.lower(), out_name.lower())] = src_name
        plan.append((prt_path, os.path.join(target_dir, out_name), state))
    return plan


def create_step_creator(session):
    """按 NX 版本依次尝试创建 STEP 导出器"""
    dm = session.DexManager
    errors = []
    for method_name in ("CreateStep214Creator", "CreateStepCreator"):
        factory = getattr(dm, method_name, None)
        if factory is None:
            errors.append(T("log.err_method_missing", method_name))
            continue
        try:
            return factory()
        except Exception as e:
            errors.append(T("log.err_method_fail", method_name, e))
    raise RuntimeError(T("log.err_creator", "; ".join(errors)))


def configure_step_creator(creator, prt_path, step_path, layer_mask, source_part=None):
    """把导出器配置成“从磁盘现有部件导出当前可见图层的几何”

    NX 2506 的导出器默认既不指定对象类型、也不指定图层范围，
    若沿用默认值，生成的 STEP 只有文件头、没有几何，所以这里全部显式设置。
    source_part 非空时改为“从显示/工作部件导出”，供链式转换使用内存部件（无磁盘路径）。
    """
    creator.OutputFile = step_path

    # 从磁盘读取，而不是从“显示部件”读取：导出内容不依赖当前显示状态
    option = getattr(type(creator), "ExportFromOption", None)
    if source_part is None:
        creator.InputFile = prt_path
        if option is not None and hasattr(option, "ExistingPart"):
            try:
                creator.ExportFrom = option.ExistingPart
            except Exception as e:
                log(T("log.export_from_fail", e))
    else:
        applied = False
        for flag in ("DisplayPart", "WorkPart"):
            value = getattr(option, flag, None) if option is not None else None
            if value is None:
                continue
            try:
                creator.ExportFrom = value
                applied = True
                break
            except Exception as e:
                log(T("log.export_from_fail", e))
        if not applied:
            raise RuntimeError(T("log.err_step_from_part"))

    # 对象类型：实体 + 曲面 + 曲线（与 STEP 转换器模块掩码一致）
    types = getattr(creator, "ObjectTypes", None)
    if types is not None:
        try:
            for flag in ("Solids", "Surfaces", "Curves"):
                if hasattr(type(types), flag):
                    setattr(types, flag, True)
        except Exception as e:
            log(T("log.objecttypes_fail", e))

    # 图层范围：只写当前视图里可见的层；空掩码等于不导任何对象，调用方会先拦截
    if hasattr(type(creator), "LayerMask"):
        try:
            creator.LayerMask = layer_mask
        except Exception as e:
            log(T("log.layermask_fail", e))

    # 等待转换真正结束再返回，否则可能读到还没写完的文件
    if hasattr(type(creator), "ProcessHoldFlag"):
        try:
            creator.ProcessHoldFlag = True
        except Exception as e:
            log(T("log.holdflag_fail", e))


def step_size(step_path):
    """返回 STEP 文件大小；不存在返回 -1"""
    try:
        return os.path.getsize(step_path) if os.path.exists(step_path) else -1
    except Exception:
        return -1


def step_has_geometry(step_path):
    """判断 STEP 是否真的含有几何（而不只是文件头 + 产品结构）

    只输出文件头的 STEP 也有 2KB 左右、大小非 0，单看文件大小会误判成功。
    这里统计几类几何实体；只有 1 个 CARTESIAN_POINT 的情况属于“空壳”，判为无效。
    """
    try:
        faces = edges = points = 0
        with open(step_path, "rb") as handle:
            while True:
                chunk = handle.read(1048576)
                if not chunk:
                    break
                faces += chunk.count(b"ADVANCED_FACE")
                edges += chunk.count(b"EDGE_CURVE")
                points += chunk.count(b"CARTESIAN_POINT")
        return faces > 0 or edges > 0 or points > 1
    except Exception:
        return False


def iges_has_geometry(iges_path):
    """判断 IGES 是否真的含几何（空壳里只有 124 变换、406 属性、314 颜色等）

    实机确认：导出器不设 ObjectTypes/LayerMask 时会写出 3KB 左右、
    只有 124+406 的文件，大小非 0 但没有任何曲线/曲面/实体。
    """
    try:
        with open(iges_path, "r", encoding="latin-1", errors="ignore") as handle:
            for line in handle:
                if len(line) < 73 or line[72] != "D":
                    continue
                try:
                    kind = int(line[:8])
                except ValueError:
                    continue
                # 100-199 曲线/曲面，500-599 拓扑（顶点/边/环/面/壳）
                if 100 <= kind <= 199 or 500 <= kind <= 599:
                    return True
        return False
    except Exception:
        return False


def step_is_valid(step_path):
    """已存在的 STEP 是否可以直接跳过：必须非空且确实含几何"""
    return step_size(step_path) > 0 and step_has_geometry(step_path)


def format_layer_mask(layers):
    """把层号列表压缩成 1-5,8,10-12 形式；无可见层返回空串"""
    if not layers:
        return ""
    layers = sorted(set(layers))
    result = []
    start = prev = layers[0]
    for number in layers[1:]:
        if number == prev + 1:
            prev = number
            continue
        result.append("%d" % start if start == prev else "%d-%d" % (start, prev))
        start = prev = number
    result.append("%d" % start if start == prev else "%d-%d" % (start, prev))
    return ",".join(result)


def find_open_part(session, path):
    """在会话里按完整路径查找已打开的部件；找不到返回 None

    完整路径规范化后仍不一致时（大小写、映射盘符、8.3 短名、NAS 路径写法差异），
    再按文件名回退匹配：NX 对同一文件只保留一个部件，同名即可视为同一件。
    """
    target = os.path.normcase(os.path.abspath(path))
    target_name = os.path.normcase(os.path.basename(path))
    try:
        parts = list(session.Parts)
    except Exception:
        return None
    by_name = None
    for part in parts:
        try:
            current = part.FullPath
        except Exception:
            continue
        if not current:
            continue
        if os.path.normcase(os.path.abspath(current)) == target:
            return part
        if by_name is None and os.path.normcase(os.path.basename(current)) == target_name:
            by_name = part
    return by_name


def _part_key(part):
    try:
        return os.path.normcase(os.path.abspath(part.FullPath))
    except Exception:
        return None


def _session_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def _open_failure_detail(exc, status):
    """把 NXOpen 原始异常与 PartLoadStatus 的真实原因拼成可判读的说明

    NX 2506 的 PartLoadStatus 只有 Dispose / GetPartName / GetStatus /
    GetStatusDescription / NumberUnloadedParts（没有 GetPartLoadStatus），
    所以要按序号逐个取状态码与描述，否则日志里只剩一句笼统的“打不开”。
    """
    items = []
    if exc is not None:
        items.append("%s: %s" % (type(exc).__name__, exc))
    if status is not None:
        try:
            count = int(getattr(status, "NumberUnloadedParts", 0) or 0)
        except Exception:
            count = 0
        for index in range(max(count, 1)):
            row = []
            try:
                name = str(status.GetPartName(index) or "")
            except Exception:
                name = ""
            for getter in ("GetStatus", "GetStatusDescription"):
                try:
                    value = str(getattr(status, getter)(index) or "")
                except Exception:
                    value = ""
                if value:
                    row.append(value)
            if row:
                items.append("%s%s" % ((name + " ") if name else "", " / ".join(row)))
    if not items:
        items.append("no detail")
    return "; ".join(items)


def ensure_displayed(session, prt_path, opened_parts):
    """确保部件已加载并成为显示部件（图层可见性查询要求显示部件）

    返回显示部件；只有脚本自己新打开的部件（含装配加载出来的组件）才记入 opened_parts，
    调用方在每件导出后立即关闭这些部件。
    """
    before = set(k for k in (_part_key(p) for p in _session_parts(session)) if k)
    part = find_open_part(session, prt_path)
    try:
        display = session.Parts.Display
        display_path = display.FullPath if display else None
    except Exception:
        display_path = None

    if part is None:
        if not os.path.isfile(prt_path):
            raise RuntimeError(T("log.missing_part", prt_path))
        status = None
        try:
            part, status = session.Parts.OpenBaseDisplay(prt_path)
        except Exception as exc:
            # NXOpen 抛异常时也要把原始 message 带进日志，别只留一句笼统的“打不开”
            raise RuntimeError(T("log.err_open_part", prt_path,
                                 _open_failure_detail(exc, None)))
        if part is None:
            # 文件被其他进程独占、损坏、版本不符等情况下 NX 返回 None 而不是抛异常；
            # PartLoadStatus 里已带真实原因，必须一并写进日志
            raise RuntimeError(T("log.err_open_part", prt_path,
                                 _open_failure_detail(None, status)))
    elif not display_path or os.path.normcase(os.path.abspath(display_path)) != \
            os.path.normcase(os.path.abspath(prt_path)):
        # 已打开但不是显示部件：用“切换活动显示部件”而不是再次打开文件，
        # 后者对已打开部件会被 NX 拒绝（文件已存在）
        session.Parts.SetActiveDisplay(
            part,
            NXOpen.DisplayPartOption.ReplaceExisting,
            NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)

    # 记录本次新加载的部件（装配的组件也会被 OpenBaseDisplay 一起加载）
    for candidate in _session_parts(session):
        key = _part_key(candidate)
        if key and key not in before and candidate not in opened_parts:
            opened_parts.append(candidate)

    display = session.Parts.Display
    if display is None:
        raise RuntimeError(T("log.err_display_none", prt_path))
    return display


def _is_assembly(part):
    try:
        return part.ComponentAssembly.RootComponent is not None
    except Exception:
        return False


def _restore_state(session, display_part, work_part):
    """把显示/工作部件切回导出前的状态（用于关闭临时部件前）"""
    if display_part is not None:
        try:
            session.Parts.SetActiveDisplay(
                display_part,
                NXOpen.DisplayPartOption.ReplaceExisting,
                NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
        except Exception:
            pass
    if work_part is not None:
        try:
            session.Parts.SetWork(work_part)
        except Exception:
            pass


def close_opened_parts(session, parts, original_keys):
    """关闭脚本自己打开、且无未保存修改的部件：先组件后装配，逐个释放"""
    ordered = sorted(parts, key=lambda p: 1 if _is_assembly(p) else 0)
    for part in ordered:
        key = _part_key(part)
        if key and key in original_keys:
            continue
        try:
            if bool(part.IsModified):
                log(T("log.keep_modified", key or part.FullPath))
                continue
        except Exception:
            pass
        try:
            part.Close(NXOpen.BasePart.CloseWholeTree.FalseValue,
                       NXOpen.BasePart.CloseModified.CloseModified, None)
        except Exception as e:
            log(T("log.close_fail", e))


def visible_layer_mask(part):
    """按当前工作视图的图层可见性生成 LayerMask

    采用 NX 的“视图内可见”判定（LayerManager.GetVisibilitiesInView），
    不是“图层可显示/可隐藏”的全局设置。工作层永远可见，用它的状态值作为
    “可见”的参照，避免依赖枚举名字（不同语言/版本命名可能不同）。
    """
    view = part.Views.WorkView
    infos = part.Layers.GetVisibilitiesInView(view)

    visible_value = None
    try:
        work_layer = int(part.Layers.WorkLayer)
    except Exception:
        work_layer = None
    if work_layer and 1 <= work_layer <= len(infos):
        visible_value = getattr(infos[work_layer - 1].State, "value", None)

    layers = []
    for index, info in enumerate(infos):
        state = getattr(info.State, "value", None)
        if visible_value is None or state == visible_value:
            layers.append(index + 1)
    return format_layer_mask(layers), view.Name, len(layers)


OUTPUT_NAME_EXTS = (".step", ".stp", ".igs", ".iges")


def output_candidates(out_path):
    """同一主干的其它扩展名写法：导出器可能把 .stp 规范化成 .step、把 .igs 规范化成 .iges"""
    stem = os.path.splitext(out_path)[0]
    names = []
    for ext in OUTPUT_NAME_EXTS:
        candidate = stem + ext
        if os.path.normcase(candidate) != os.path.normcase(out_path):
            names.append(candidate)
    return names


def snapshot_candidates(out_path):
    """导出前记录已存在的同名其它扩展名文件，避免把用户原有文件当成本次产物"""
    return set(os.path.normcase(p) for p in output_candidates(out_path) if os.path.exists(p))


def reclaim_output(out_path, before):
    """保证预览名 == 落盘名：导出器若换了扩展名写法，就把本次新产物改回计划里的名字"""
    if os.path.exists(out_path):
        return out_path
    for candidate in output_candidates(out_path):
        if os.path.normcase(candidate) in before or not os.path.exists(candidate):
            continue
        try:
            os.replace(candidate, out_path)
            log("[name] %s -> %s" % (os.path.basename(candidate), os.path.basename(out_path)))
        except Exception as e:
            log("[name] %s: %s" % (candidate, e))
        return out_path
    return out_path


def export_step(session, prt_path, step_path, layer_mask, source_part=None):
    """导出单个文件；产出为空或无几何时重试一次，仍无效则报错

    source_part 非空时从该内存部件导出（链式转换），此时 prt_path 传空串。
    """
    last_error = ""
    before = snapshot_candidates(step_path)
    for attempt in (1, 2):
        if attempt == 2:
            # 先清掉上次留下的无效文件（含导出器换了扩展名写出的），避免转换器因“目标已存在”而拒绝覆盖
            stale = [step_path] + [p for p in output_candidates(step_path)
                                   if os.path.normcase(p) not in before]
            for path in stale:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                        log(T("log.retry_deleted"))
                except Exception as e:
                    log(T("log.retry_delete_fail", e))
        creator = create_step_creator(session)
        try:
            configure_step_creator(creator, prt_path, step_path, layer_mask, source_part)
            creator.Commit()
        finally:
            try:
                creator.Destroy()
            except Exception:
                pass

        reclaim_output(step_path, before)
        size = step_size(step_path)
        if size > 0 and step_has_geometry(step_path):
            if attempt == 2:
                log(T("log.retry2_ok", size))
            return size

        if size < 0:
            last_error = T("log.err_out_missing")
        elif size == 0:
            last_error = T("log.err_out_empty")
        else:
            last_error = T("log.err_out_header", size)
        log(T("log.invalid_attempt", attempt, last_error))
        time.sleep(0.4)

    # 无效产物不留在原目录，否则下次会被“已存在”规则当成已完成
    try:
        if os.path.exists(step_path):
            os.remove(step_path)
    except Exception as e:
        log(T("log.cleanup_fail", e))

    raise RuntimeError(T("log.err_export_invalid", last_error))


def _iges_tool_path(filename):
    """在 IGES 安装目录里找 iges.exe / igesexport.def；找不到返回 None"""
    for env_name in ("IGESUG_DIR", "UGII_IGES_DIR"):
        base = (os.environ.get(env_name) or "").strip()
        if not base:
            continue
        candidate = os.path.join(base, filename)
        if os.path.isfile(candidate):
            return candidate
    base = (os.environ.get("UGII_BASE_DIR") or "").strip()
    if base:
        candidate = os.path.join(base, "IGES", filename)
        if os.path.isfile(candidate):
            return candidate
    return None


def iges_exe_path():
    """IGES 转换器路径；优先环境变量，其次 NX 安装目录，最后回落到默认目录"""
    path = _iges_tool_path("iges.exe")
    if path:
        return path
    fallback = os.path.join(IGES_DEFAULT_DIR, "iges.exe")
    if os.path.isfile(fallback):
        return fallback
    raise RuntimeError(T("log.iges_exe_missing", fallback))


def iges_def_path():
    """IGES 转换配置 igesexport.def；回落到默认目录时记一条警告"""
    path = _iges_tool_path("igesexport.def")
    if path:
        return path
    fallback = os.path.join(IGES_DEFAULT_DIR, "igesexport.def")
    log(T("log.iges_def_fallback", fallback))
    return fallback


def _iges_log_tail(log_path, lines=None):
    """失败时返回日志尾部若干行；读不到就说明读不到"""
    import io
    lines = IGES_LOG_TAIL_LINES if lines is None else lines
    try:
        with io.open(log_path, "r", encoding="utf-8", errors="replace") as handle:
            text = handle.read()
    except Exception as e:
        return "\n" + T("log.iges_log_unread", e)
    tail = [line for line in text.splitlines() if line.strip()][-lines:]
    if not tail:
        return "\n" + T("log.iges_log_unread", log_path)
    return "\n" + "\n".join("      " + line for line in tail)


def export_iges(prt_path, igs_path):
    """直调 NX 的 IGES translator 导出单个文件，返回产物字节数

    命令形如：iges.exe <prt> "out=<igs>" "log=<log>" "def=<def>"
    日志写到 %TEMP%\\nx_batch_export_iges\\，不落在用户文档目录。
    IGES 不做“是否含几何”校验，只判存在且大小 > 0。
    """
    import subprocess
    import tempfile

    exe = iges_exe_path()
    def_file = iges_def_path()

    log_dir = os.path.join(tempfile.gettempdir(), "nx_batch_export_iges")
    try:
        if not os.path.isdir(log_dir):
            os.makedirs(log_dir)
    except Exception:
        pass
    stem = os.path.splitext(os.path.basename(prt_path))[0]
    log_path = os.path.join(log_dir, "%s_%d_%d.log" % (
        stem, os.getpid(), int(time.time() * 1000)))

    command = [exe, prt_path,
               "out=%s" % igs_path,
               "log=%s" % log_path,
               "def=%s" % def_file]
    before = snapshot_candidates(igs_path)
    try:
        process = subprocess.run(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=IGES_TIMEOUT,
            creationflags=0x08000000)   # CREATE_NO_WINDOW
    except TypeError:
        # 老版本 Python 不认 creationflags 关键字
        process = subprocess.run(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=IGES_TIMEOUT)
    except Exception as e:
        # TimeoutExpired 在不同 Python 版本里可能是不同名字，按类型名判断更稳
        if type(e).__name__ == "TimeoutExpired":
            raise RuntimeError(T("log.iges_timeout", IGES_TIMEOUT))
        raise RuntimeError(T("log.iges_cmd_fail", e))

    if process.returncode != 0:
        raise RuntimeError(T("log.iges_exit", process.returncode)
                           + _iges_log_tail(log_path))

    reclaim_output(igs_path, before)
    size = step_size(igs_path)
    if size <= 0:
        raise RuntimeError(T("log.err_iges_out", igs_path)
                           + _iges_log_tail(log_path))
    if not iges_has_geometry(igs_path):
        # 只有文件头/属性的 IGES 不能算成功，否则用户拿到一个打不开的空文件
        raise RuntimeError(T("log.err_iges_geom", igs_path)
                           + _iges_log_tail(log_path))
    return size


def create_iges_creator(session):
    """按 NX 版本依次尝试创建 IGES 导出器（内存部件导出用）"""
    dm = session.DexManager
    errors = []
    for method_name in ("CreateIgesCreator", "CreateIges214Creator"):
        factory = getattr(dm, method_name, None)
        if factory is None:
            errors.append(T("log.err_method_missing", method_name))
            continue
        try:
            return factory()
        except Exception as e:
            errors.append(T("log.err_method_fail", method_name, e))
    raise RuntimeError(T("log.err_creator", "; ".join(errors)))


def export_iges_mem(session, part, igs_path, layer_mask):
    """把内存部件直接导成 IGES（不落中间 .prt，不调用 iges.exe）

    实机确认：必须显式设置 ObjectTypes 与 LayerMask，否则只写出 124/406 的空壳；
    失败时抛异常，由 chain_convert 关闭部件。
    """
    last_error = ""
    before = snapshot_candidates(igs_path)
    for attempt in (1, 2):
        if attempt == 2:
            try:
                if os.path.exists(igs_path):
                    os.remove(igs_path)
                    log(T("log.retry_deleted"))
            except Exception as e:
                log(T("log.retry_delete_fail", e))
        creator = create_iges_creator(session)
        try:
            creator.OutputFile = igs_path
            option = getattr(type(creator), "ExportFromOption", None)
            applied = False
            for flag in ("DisplayPart", "WorkPart"):
                value = getattr(option, flag, None) if option is not None else None
                if value is None:
                    continue
                try:
                    creator.ExportFrom = value
                    applied = True
                    break
                except Exception as e:
                    log(T("log.export_from_fail", e))
            if not applied:
                raise RuntimeError(T("log.err_iges_from_part"))
            # 对象类型必须显式勾上：NX 2506 的 IGES 导出器默认不选任何类型，
            # 实测只写出 124/406 的空壳（约 3KB，大小非 0 但没有几何）
            types = getattr(creator, "ObjectTypes", None)
            if types is not None:
                try:
                    for flag in ("Solids", "Surfaces", "Curves"):
                        if hasattr(type(types), flag):
                            setattr(types, flag, True)
                except Exception as e:
                    log(T("log.objecttypes_fail", e))
            # 图层掩码同样必须显式给：留空导出器不导任何对象
            if hasattr(type(creator), "LayerMask"):
                try:
                    creator.LayerMask = layer_mask or "1-256"
                except Exception as e:
                    log(T("log.layermask_fail", e))
            if hasattr(type(creator), "ProcessHoldFlag"):
                try:
                    creator.ProcessHoldFlag = True
                except Exception as e:
                    log(T("log.holdflag_fail", e))
            creator.Commit()
        finally:
            try:
                creator.Destroy()
            except Exception:
                pass

        reclaim_output(igs_path, before)
        size = step_size(igs_path)
        if size > 0 and iges_has_geometry(igs_path):
            if attempt == 2:
                log(T("log.retry2_ok", size))
            return size
        if size > 0:
            last_error = T("log.err_iges_geom", igs_path)
        else:
            last_error = T("log.err_iges_out", igs_path)
        log(T("log.invalid_attempt", attempt, last_error))
        time.sleep(0.4)

    try:
        if os.path.exists(igs_path):
            os.remove(igs_path)
    except Exception as e:
        log(T("log.cleanup_fail", e))
    raise RuntimeError(T("log.err_export_invalid", last_error))


def batch_import_module():
    """导入接口（原 batch_import.py 已并入本文件）"""
    return _import_api()


def chain_convert(session, src_path, out_path, fmt, used, layer_mask):
    """stp/igs/x_t/x_b -> stp/igs：内存建件导入后直接导出，不落中间 .prt、不删源文件

    返回产物字节数；失败抛出异常。finally 关闭本次新建的部件，
    避免导入成功但导出失败时留下内存垃圾。源文件在整个过程中不被修改。
    """
    BI = batch_import_module()
    if not used:
        used.update(BI.existing_part_names(session))
    try:
        previous = session.Parts.Display
    except Exception:
        previous = None
    state, name, detail, _disk, part = BI.import_one(session, src_path, used, save_as=False)
    if state != "ok" or part is None:
        raise RuntimeError(T("log.chain_import_fail", detail or state))
    try:
        log(T("log.chain_begin", name, fmt.upper()))
        if fmt == "iges":
            return export_iges_mem(session, part, out_path, layer_mask)
        return export_step(session, "", out_path, layer_mask, source_part=part)
    finally:
        try:
            BI.close_part(session, part)
        except Exception as e:
            log(T("log.chain_close_fail", e))
        if previous is not None:
            try:
                session.Parts.SetDisplay(previous, False, True)
            except Exception as e:
                log(T("log.chain_restore_fail", e))


def legacy_pick(base_dir=None):
    """选择窗口不可用时的后备方案：分步弹出系统文件夹/文件对话框"""
    folders = None
    files = []
    if PICK_MODE in ("folder", "both"):
        try:
            folders = select_folders(base_dir)
        except Exception as e:
            log(T("log.folder_dialog_fail", e))
    if PICK_MODE in ("file", "both"):
        try:
            files = pick_files(T("ui.file_pick_title"), base_dir or "")
        except Exception as e:
            files = []
            log(T("log.file_dialog_fail", e))
    return folders, files


def _lock_path():
    import tempfile
    return os.path.join(tempfile.gettempdir(), "nx_step_export.lock")


def _proc_stamp(pid):
    """进程创建时间（FILETIME 字符串）：用来识别 PID 被系统复用，取不到返回空串"""
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(0x1000, False, int(pid))  # QUERY_LIMITED_INFORMATION
        if not handle:
            return ""
        try:
            created = wintypes.FILETIME()
            exited = wintypes.FILETIME()
            kern = wintypes.FILETIME()
            user = wintypes.FILETIME()
            if not kernel32.GetProcessTimes(handle, ctypes.byref(created),
                                            ctypes.byref(exited),
                                            ctypes.byref(kern), ctypes.byref(user)):
                return ""
            return str((created.dwHighDateTime << 32) | created.dwLowDateTime)
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return ""


def _proc_parent_map():
    """进程快照：pid -> (父 pid, 可执行文件名)，用于往上找发起本会话的 NX 主进程"""
    try:
        import ctypes
        from ctypes import wintypes

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_char * 260),
            ]

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        kernel32.CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
        if not snapshot or snapshot == wintypes.HANDLE(-1).value:
            return {}
        table = {}
        try:
            entry = PROCESSENTRY32()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
            ok = kernel32.Process32First(snapshot, ctypes.byref(entry))
            while ok:
                name = entry.szExeFile.decode("mbcs", "ignore")
                table[int(entry.th32ProcessID)] = (int(entry.th32ParentProcessID), name)
                ok = kernel32.Process32Next(snapshot, ctypes.byref(entry))
        finally:
            kernel32.CloseHandle(snapshot)
        return table
    except Exception:
        return {}


def _nx_owner_pid(start_pid=None):
    """往上找发起本会话的 NX 主程序（ugraf.exe）PID；找不到返回 0。

    NX 菜单里的 python 宿主是 run_journal.exe，它只活到脚本返回，不能当作
    「发起批处理的 NX」；必须沿父进程链找到 ugraf.exe，否则进度窗要么不跟随
    NX 关闭、要么被误判成已退出。
    """
    pid = int(start_pid or os.getpid())
    try:
        table = _proc_parent_map()
    except Exception:
        return 0
    seen = set()
    for _ in range(12):
        if pid <= 0 or pid in seen:
            break
        seen.add(pid)
        info = table.get(pid)
        if not info:
            break
        parent, name = info
        if str(name).strip().lower() in ("ugraf.exe", "ugraf"):
            return pid
        pid = parent
    return 0


def _pid_alive(pid, stamp=""):
    """判断锁文件里的 PID 是否仍在运行。

    返回 True=进程存在，False=进程已不存在（陈旧锁），None=无法判断（按占用处理）。
    stamp 是写锁时记下的进程创建时间：PID 被系统复用后两者不一致，按陈旧锁处理。
    """
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(0x1000, False, int(pid))  # QUERY_LIMITED_INFORMATION
        if handle:
            kernel32.CloseHandle(handle)
            if stamp:
                now = _proc_stamp(pid)
                if now and now != stamp:
                    return False
            return True
        # 打不开句柄时区分“PID 不存在”(ERROR_INVALID_PARAMETER=87) 与“权限不足”
        last = kernel32.GetLastError()
        if last == 87:
            return False
        return None
    except Exception:
        return None


def _lock_event(message):
    """锁事件落盘：此时主日志可能尚未打开，先写控制台 + %TEMP% 下的锁日志"""
    log(message)
    try:
        import tempfile
        path = os.path.join(tempfile.gettempdir(), "nx_step_export_busy.log")
        _rotate_log(path)
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(message + "\n")
    except Exception:
        pass


def _acquire_run_lock():
    """单实例锁：防止连点两次菜单/重复触发时两个导出同时改同一批文件"""
    global _RUN_LOCK
    path = _lock_path()
    for attempt in (1, 2):
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            # 锁内容为「PID|进程创建时间」：只记 PID 时，进程被杀后 PID 被系统复用，
            # 新进程会让这把陈旧锁永远判「忙」（真机现象：每次启动都提示已被占用）。
            os.write(fd, ("%d|%s" % (os.getpid(), _proc_stamp(os.getpid()))).encode("ascii"))
            os.close(fd)
            _RUN_LOCK = path
            return True
        except FileExistsError:
            pid = None
            stamp = ""
            try:
                with open(path, "r", encoding="ascii") as handle:
                    head, _, tail = (handle.read() or "").strip().partition("|")
                pid = int(head or "0")
                stamp = tail.strip()
            except Exception:
                pid = None
                stamp = ""
            # 先验证锁里的进程是否真的存在：已退出 → 删陈旧锁并重试；
            # 存在 → busy；无法判断（权限/异常）→ 记日志后按 busy 处理（安全优先）
            state = _pid_alive(pid, stamp) if pid else None
            if attempt == 1 and state is False:
                try:
                    os.remove(path)
                    _lock_event(T("log.stale_lock", pid))
                except Exception:
                    pass
                continue
            if attempt == 1 and state is None:
                _lock_event(T("log.lock_unknown", pid if pid else "?"))
            message = T("log.busy")
            try:
                print(message)
            except Exception:
                pass
            busy_log = None
            try:
                import tempfile
                busy_log = os.path.join(tempfile.gettempdir(), "nx_step_export_busy.log")
                _rotate_log(busy_log)
                with open(busy_log, "a", encoding="utf-8") as handle:
                    handle.write(message + "\n")
            except Exception:
                pass
            # 上一次点击的选择窗还开着：把它提到前台，不叠第二个窗、也不再弹繁忙提示
            front_state = "none"
            try:
                front_state = _load_runner().activate_picker(T("ui.picker_title"))
            except Exception:
                front_state = "none"
            if front_state != "ok":
                notify_busy()
            try:
                if busy_log:
                    with open(busy_log, "a", encoding="utf-8") as handle:
                        handle.write("[dlg] front=%s\n" % front_state)
            except Exception:
                pass
            return False
        except Exception:
            # 锁机制本身不可用时，不要阻止正常导出
            return True
    return False


def _release_run_lock():
    global _RUN_LOCK
    if _RUN_LOCK:
        try:
            os.remove(_RUN_LOCK)
        except Exception:
            pass
        _RUN_LOCK = None


ENV_STAGE = "NX_BATCH_STAGE"

# 后台会话启动失败时的提示文案（不落在 lang 资源里，避免三份语言文件一起改）
_NOTICE_START_FAIL = {
    "zh_cn": ("格式转换", "后台转换会话启动失败，本次导出未执行。\n详情见日志：\n%s"),
    "zh_tw": ("格式轉換", "後台轉換工作階段啟動失敗，本次匯出未執行。\n詳情見日誌：\n%s"),
    "en": ("Format conversion",
           "Failed to start the background session; nothing was exported.\n"
           "See log:\n%s"),
}


def _notice_start_fail(log_path):
    """后台会话起不来时给一次可见提示（只提示，不回落同步）"""
    title, body = _NOTICE_START_FAIL.get(LANG, _NOTICE_START_FAIL["zh_cn"])
    try:
        spawn_powershell(BUSY_PS, [
            "-Title", title,
            "-Text", body % (log_path or ""),
            "-OkText", T("ui.busy_ok"),
            "-Seconds", "20",
        ], "asyncfail")
    except Exception:
        pass


def _open_main_log():
    """打开主日志（路径与同步流程一致），返回实际日志路径"""
    global _log_handle
    log_path = None
    try:
        log_dir = os.path.join(plugin_dir(), "logs")
        if not os.path.isdir(log_dir):
            os.makedirs(log_dir)
        log_path = os.path.join(log_dir, "batch_export_step.log")
        _rotate_log(log_path)
        _log_handle = open(log_path, "w", encoding="utf-8-sig")
    except Exception:
        import tempfile
        fallback = os.path.join(tempfile.gettempdir(),
                                "nx_step_export_%d.log" % os.getpid())
        log_path = fallback
        _log_handle = open(fallback, "w", encoding="utf-8-sig")
        log(T("log.log_fallback", fallback))
    return log_path


def _close_main_log():
    global _log_handle
    try:
        if _log_handle is not None:
            _log_handle.close()
    except Exception:
        pass
    _log_handle = None


def _runner_exe():
    """NX 自带的后台宿主 run_journal.exe（不依赖用户机器上的 Python）"""
    return os.path.join(os.environ.get("UGII_BASE_DIR") or "",
                        "NXBIN", "run_journal.exe")


def _stage_root():
    import tempfile
    return os.path.join(tempfile.gettempdir(),
                        "nx_batch_export_stage_%d_%d"
                        % (os.getpid(), int(time.time() * 1000)))


def _spawn_picker_async(stage):
    """非阻塞起选择窗：立即返回，不等窗口关闭（结果由后台会话接手）"""
    import subprocess

    ps_path = os.path.join(stage["dir"], "select.ps1")
    with open(ps_path, "w", encoding="utf-8-sig") as handle:
        handle.write(render_ps(SELECT_UI_PS))
    for stale in (stage["select_out"], stage["select_front"]):
        try:
            if os.path.exists(stale):
                os.remove(stale)
        except Exception:
            pass
    powershell = _powershell_path()
    if powershell is None:
        raise RuntimeError(T("log.err_no_powershell"))
    options = _picker_options(None)
    arguments = _picker_args(read_last_dir(), options) + \
        ["-OutFile", stage["select_out"]]
    command = [powershell, "-NoProfile", "-NonInteractive", "-STA",
               "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
               "-File", ps_path] + arguments
    subprocess.Popen(command, stdin=subprocess.DEVNULL,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                     creationflags=0x08000000)


def _wait_picker_result(out_path, timeout=6 * 3600):
    """后台会话里等选择窗结果（普通窗，用户可能挑很久）"""
    started = time.time()
    while time.time() - started < timeout:
        if os.path.exists(out_path):
            try:
                with open(out_path, "r", encoding="utf-8-sig") as handle:
                    return json.load(handle)
            except Exception:
                pass
        time.sleep(0.2)
    return None


def _read_front_mark(path):
    try:
        with open(path, "r", encoding="ascii") as handle:
            return (handle.read() or "").strip() or "unknown"
    except Exception:
        return "unknown"


def _write_lock_pid(pid):
    """把单实例锁改挂到后台会话（父进程退出后锁仍有效）"""
    if not _RUN_LOCK:
        return
    try:
        with open(_RUN_LOCK, "w", encoding="ascii") as handle:
            handle.write("%d|%s" % (int(pid), _proc_stamp(int(pid))))
    except Exception:
        pass


def _drop_stage_lock(lock_path):
    """后台会话提前结束（取消/失败）时清掉自己这段的锁"""
    try:
        if lock_path and os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


def _start_prepare_session(stage):
    """起后台 NX 会话（run_journal.exe）；成功返回 PID，失败返回 0"""
    import subprocess

    runner = _runner_exe()
    if not os.path.isfile(runner):
        log("[async] prepare-missing-runner: %s" % runner)
        return 0
    env = dict(os.environ)
    env[ENV_STAGE] = stage["stage"]
    try:
        process = subprocess.Popen(
            [runner, "-nx", os.path.abspath(__file__)], env=env,
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL, creationflags=0x08000000)
    except Exception as e:
        log("[async] prepare-start-failed: %s" % e)
        return 0
    return int(process.pid)


def _run_async():
    """A1：菜单点完立即返回，选择/预览/计划/起批处理全部交给后台 NX 会话

    任何一步失败都只记日志 + 一次提示，绝不回落到前台同步等待。
    """
    session = NXOpen.Session.GetSession()
    log_path = _open_main_log()

    # 启动前轻量检查：此时还没选文件，按“全部已打开且有未保存修改”的部件检查一次
    skip_keys = {}
    try:
        modified = collect_modified_all(session)
    except Exception as e:
        modified = []
        log("[async] modified-check-failed: %s" % e)
    if modified:
        log(T("log.modified_header", len(modified)))
        for item in modified:
            log("  %s" % rel_display(item["path"], None))
        confirmed, save_keys = ask_save_modified(modified)
        if not confirmed:
            log(T("log.cancel_unsaved"))
            _close_main_log()
            return
        for item in modified:
            item_key = os.path.normcase(os.path.abspath(item["path"]))
            if item_key not in save_keys:
                log(T("log.user_skip_mod", item["name"]))
                skip_keys[item_key] = "reason.user_skip"
                continue
            try:
                item["part"].Save(NXOpen.BasePartSaveComponents.TrueValue,
                                  NXOpen.BasePartCloseAfterSave.FalseValue)
                log(T("log.saved_ok", rel_display(item["path"], None)))
            except Exception as e:
                log(T("log.save_fail_skip", item["name"], e))
                skip_keys[item_key] = "reason.save_fail"

    try:
        stage_dir = _stage_root()
        os.makedirs(stage_dir)
        stage = {
            "dir": stage_dir,
            "stage": os.path.join(stage_dir, "stage.json"),
            "select_out": os.path.join(stage_dir, "select.json"),
            "select_front": os.path.join(stage_dir, "select.json.front"),
            "log_path": log_path,
            "lang": LANG,
            "lock": _lock_path(),
            "skip_keys": sorted(skip_keys),
            # 发起本会话的 NX 主程序（ugraf.exe）。菜单 python 宿主 run_journal 只活到
            # 脚本返回，不能拿它当 NX：进度窗必须跟着真正的主程序退出而关闭。
            "nx_pid": _nx_owner_pid(),
        }
        with open(stage["stage"], "w", encoding="utf-8") as handle:
            json.dump(stage, handle, ensure_ascii=False)
        _spawn_picker_async(stage)
    except Exception as e:
        log("[async] picker-spawn-failed: %s" % e)
        _close_main_log()
        return

    pid = _start_prepare_session(stage)
    if not pid:
        _notice_start_fail(log_path)
        _close_main_log()
        return
    _write_lock_pid(pid)
    log("[async] prepare pid=%d nx_pid=%d stage=%s" % (pid, stage.get("nx_pid") or 0, stage_dir))
    _hand_off_lock()
    _close_main_log()


def _run_prepare_stage(stage_path):
    """后台会话入口：等选择结果 → 计划 → 预览 → 起批处理（不占用交互 NX）"""
    global _log_handle
    try:
        with open(stage_path, "r", encoding="utf-8") as handle:
            stage = json.load(handle)
    except Exception:
        return
    try:
        _log_handle = open(stage.get("log_path"), "a", encoding="utf-8")
    except Exception:
        _log_handle = None
    log("[async] prepare start pid=%d" % os.getpid())
    lock_path = stage.get("lock") or _lock_path()
    try:
        with open(lock_path, "w", encoding="ascii") as handle:
            # 与 _acquire_run_lock 一致：PID 会被系统复用，锁里必须带进程创建时间
            handle.write("%d|%s" % (os.getpid(), _proc_stamp(os.getpid())))
    except Exception:
        pass

    log(T("log.opening_picker"))
    data = _wait_picker_result(stage["select_out"])
    if not data or not data.get("ok"):
        log(T("log.preview_cancel"))
        _drop_stage_lock(lock_path)
        _close_main_log()
        return
    log("[dlg] front=%s tag=prepare" % _read_front_mark(stage["select_front"]))

    base_dir = None
    options = _picker_options(None)
    try:
        folders, files, options = _apply_picker_result(data, options)
    except Exception as e:
        log(T("log.picker_fail", e))
        _drop_stage_lock(lock_path)
        _close_main_log()
        return
    overrides = {}
    plan = None
    while True:
        options = resolve_output_options(options)
        include_dirs = folders
        picked_files = [os.path.abspath(p) for p in files if p]
        scan_dirs = not (picked_files and folders is None)
        if folders is None and not picked_files:
            log(T("log.no_selection"))
            _drop_stage_lock(lock_path)
            _close_main_log()
            return
        if folders:
            log(T("log.picked_folders", ", ".join(folders)))
        elif folders is not None:
            log(T("log.picked_root"))
        if picked_files:
            log(T("log.picked_files", len(picked_files)))
            for path in picked_files:
                log("  %s" % rel_display(path, base_dir))

        plan = build_plan(find_prt_files(base_dir, include_dirs,
                                         picked_files, scan_dirs,
                                         exts=source_exts(options)), options)
        if not plan:
            log(T("log.no_prt_reread"))
        else:
            try:
                action, overrides = ("ok", {}) if _skip_preview(plan, options) \
                    else preview_plan(plan, base_dir, options)
            except Exception as e:
                log(T("log.preview_fail", e))
                action, overrides = "ok", {}
            if action == "ok":
                break
            if action != "back":
                log(T("log.preview_cancel"))
                _drop_stage_lock(lock_path)
                _close_main_log()
                return
            log(T("log.back"))
        # 重新选择：在后台会话里同步弹窗，不占用交互 NX
        try:
            folders, files, options = choose_targets(
                base_dir=read_last_dir(), options=options)
        except Exception as e:
            log(T("log.picker_fail", e))
            _drop_stage_lock(lock_path)
            _close_main_log()
            return

    total = len(plan)
    log(T("log.file_count", total))
    skip_unsaved = {}
    for key in stage.get("skip_keys") or []:
        skip_unsaved[key] = "reason.user_skip"
    jobs = build_jobs(plan, options, overrides, skip_unsaved, base_dir)

    if _start_background(jobs, total, options, stage.get("log_path"), base_dir,
                         int(stage.get("nx_pid") or 0)):
        log(T("log.background", len(jobs)))
        _hand_off_lock()
        _close_main_log()
        return
    log("[async] prepare-background-failed")
    _notice_start_fail(stage.get("log_path"))
    _drop_stage_lock(lock_path)
    _close_main_log()


def main():
    if not _acquire_run_lock():
        return
    try:
        # A1：交互模式下点完菜单立即返回，剩下的交给后台 NX 会话
        if INTERACTIVE and not DRY_RUN:
            _run_async()
        else:
            _run_locked()
    finally:
        _release_run_lock()


def _run_locked():
    global _log_handle

    session = NXOpen.Session.GetSession()
    # 不再有“项目根目录”；base_dir 恒为 None，仅作为兼容参数传给旧函数
    base_dir = None

    # 日志放到插件目录下的 logs\；插件目录不可写时回落 %TEMP%
    # utf-8-sig：带 BOM 的 UTF-8，英文版/繁中版 Windows 下用 Excel 直接打开不乱码；
    # 本脚本自身从不回读该日志，所以 BOM 不会影响任何逻辑
    log_path = None
    try:
        log_dir = os.path.join(plugin_dir(), "logs")
        if not os.path.isdir(log_dir):
            os.makedirs(log_dir)
        log_path = os.path.join(log_dir, "batch_export_step.log")
        _rotate_log(log_path)
        _log_handle = open(log_path, "w", encoding="utf-8-sig")
    except Exception:
        # 插件目录只读/被占用时退到 %TEMP%，保证导出仍可继续
        import tempfile
        fallback = os.path.join(tempfile.gettempdir(),
                                "nx_step_export_%d.log" % os.getpid())
        log_path = fallback
        _log_handle = open(fallback, "w", encoding="utf-8-sig")
        log(T("log.log_fallback", fallback))

    # 交互选择结果：include_dirs=None 表示沿用配置里的 INCLUDE_DIRS
    include_dirs = None
    picked_files = []
    scan_dirs = True

    plan = None
    overrides = {}
    # 交互模式首次不预置选项：由 choose_targets 从记忆文件读取上次选择，
    # 避免默认值把记忆值覆盖掉；非交互模式在下面统一回落到默认值
    options = None if INTERACTIVE else default_options()
    if INTERACTIVE:
        while True:
            folders = None
            picked = []
            try:
                log(T("log.opening_picker"))
                folders, picked, picked_options = choose_targets(
                    base_dir=read_last_dir(), options=options)
                options = picked_options
            except Exception as e:
                log(T("log.picker_fail", e))
                folders, picked = legacy_pick()
            # 指定导出目录为空/不存在时回落到原文档所在路径（交互模式下当次即生效）
            options = resolve_output_options(options)

            include_dirs = None
            picked_files = []
            scan_dirs = True

            if folders is not None:
                include_dirs = folders
                if folders:
                    log(T("log.picked_folders", ", ".join(folders)))
                else:
                    log(T("log.picked_root"))
            if picked:
                picked_files = picked
                log(T("log.picked_files", len(picked_files)))
                for path in picked_files:
                    log("  %s" % rel_display(path, base_dir))

            # 只选了文件、没选文件夹时，不再扫描目录
            scan_dirs = not (picked_files and folders is None)

            if folders is None and not picked_files:
                log(T("log.no_selection"))
                _log_handle.close()
                _log_handle = None
                return

            plan = build_plan(find_prt_files(base_dir, include_dirs,
                                             picked_files, scan_dirs,
                                             exts=source_exts(options)), options)
            if not plan:
                log(T("log.no_prt_reread"))
                continue

            try:
                action, overrides = ("ok", {}) if _skip_preview(plan, options) \
                    else preview_plan(plan, base_dir, options)
            except Exception as e:
                log(T("log.preview_fail", e))
                action = "ok"
                overrides = {}

            if action == "back":
                log(T("log.back"))
                continue
            if action != "ok":
                log(T("log.preview_cancel"))
                _log_handle.close()
                _log_handle = None
                return
            break

    if plan is None:
        options = resolve_output_options(options)
        plan = build_plan(find_prt_files(base_dir, include_dirs,
                                         picked_files, scan_dirs,
                                         exts=source_exts(options)), options)

    total = len(plan)
    log("=" * 60)
    log(T("log.src_any") if base_dir is None else T("log.src_dir", base_dir))
    if options.get("out_mode") == "dir":
        log(T("log.output_custom", options.get("out_dir") or ""))
    else:
        log(T("log.output_line", output_ext(options)))
    log(T("log.format_line", output_format(options).upper()))
    log(T("log.file_count", total))
    if scan_dirs:
        effective_dirs = list(INCLUDE_DIRS if include_dirs is None else include_dirs)
        if effective_dirs:
            log(T("log.include_dirs", ", ".join(effective_dirs)))
    if EXCLUDE_DIRS:
        log(T("log.exclude_dirs", ", ".join(EXCLUDE_DIRS)))
    if INCLUDE_FILES:
        log(T("log.include_files", ", ".join(INCLUDE_FILES)))
    if EXCLUDE_FILES:
        log(T("log.exclude_files", ", ".join(EXCLUDE_FILES)))
    if picked_files:
        log(T("log.manual_files", len(picked_files)))
    log("=" * 60)

    if total == 0:
        log(T("log.no_files_end"))
        log(T("log.dryrun_hint"))
        _log_handle.close()
        _log_handle = None
        return

    if DRY_RUN:
        log(T("log.dryrun_title"))
        for index, (prt_path, out_path, state) in enumerate(plan, 1):
            mark = T("log.dryrun_skip_mark") if state == "SKIP" else ""
            log("  %3d. %s  ->  %s%s" % (
                index,
                rel_display(prt_path, base_dir),
                rel_display(out_path, base_dir),
                mark))
        log("=" * 60)
        _log_handle.close()
        _log_handle = None
        return

    # 导出前检查：计划里哪些文件已经打开且有未保存修改（只按内存 IsModified 判断）
    modified_items = collect_modified_open(session, plan)
    skip_unsaved = {}
    if modified_items:
        log(T("log.modified_header", len(modified_items)))
        for item in modified_items:
            log("  %s" % rel_display(item["path"], base_dir))
        if INTERACTIVE:
            confirmed, save_keys = ask_save_modified(modified_items)
            if not confirmed:
                log(T("log.cancel_unsaved"))
                _log_handle.close()
                _log_handle = None
                return
            for item in modified_items:
                item_key = os.path.normcase(os.path.abspath(item["path"]))
                if item_key not in save_keys:
                    log(T("log.user_skip_mod", item["name"]))
                    skip_unsaved[item_key] = "reason.user_skip"
                    continue
                try:
                    item["part"].Save(NXOpen.BasePartSaveComponents.TrueValue,
                                      NXOpen.BasePartCloseAfterSave.FalseValue)
                    log(T("log.saved_ok", rel_display(item["path"], base_dir)))
                except Exception as e:
                    log(T("log.save_fail_skip", item["name"], e))
                    skip_unsaved[item_key] = "reason.save_fail"
        else:
            for item in modified_items:
                log(T("log.noninteractive_skip", item["name"]))
                skip_unsaved[os.path.normcase(os.path.abspath(item["path"]))] = \
                    "reason.noninteractive"

    ok, skipped, failed = 0, 0, []
    original_display_part = None
    original_display_path = None
    original_work_part = None
    try:
        original_display_part = session.Parts.Display
        if original_display_part is not None:
            original_display_path = original_display_part.FullPath
        original_work_part = session.Parts.Work
    except Exception:
        original_display_part = None
        original_work_part = None
    original_keys = set(k for k in (_part_key(p) for p in _session_parts(session)) if k)

    chain_used = set()   # 链式转换新建的部件名，跨文件复用避免重名
    # 计划 → 可执行任务（同名策略/未保存修改在这里一次性定下来，前后台共用）
    jobs = build_jobs(plan, options, overrides, skip_unsaved, base_dir)
    ctx = {"base_dir": base_dir,
           "original_display_part": original_display_part,
           "original_work_part": original_work_part,
           "original_keys": original_keys}

    # 优先交给独立 NX 批处理会话：本会话立即返回，前台不再被占用
    if _start_background(jobs, total, options, log_path, base_dir):
        log(T("log.background", len(jobs)))
        _hand_off_lock()
        _log_handle.close()
        _log_handle = None
        return

    # 后台不可用时回落到原来的前台执行（同一套单文件逻辑）
    for job in jobs:
        if job["action"] != "do":
            if job["msg"]:
                log(job["msg"])
            skipped += 1
            continue
        if export_one(session, job, options, chain_used, ctx) == "ok":
            ok += 1
        else:
            failed.append(rel_display(job["src"], base_dir))

    # 恢复用户原本的显示部件
    if original_display_part is not None:
        try:
            # 用“切回活动显示部件”的方式恢复：对已打开的部件再调 OpenBaseDisplay 会被 NX 拒绝，
            # 导致显示区变空；SetActiveDisplay 不会重复打开文件。
            _status, _load = session.Parts.SetActiveDisplay(
                original_display_part,
                NXOpen.DisplayPartOption.ReplaceExisting,
                NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
            log(T("log.display_restored", original_display_path))
        except Exception as e:
            log(T("log.display_restore_fail", e))
    # 工作部件也还原成执行前那个（切显示部件会把工作部件跟着改掉）
    if original_work_part is not None:
        try:
            current_work = session.Parts.Work
            if current_work is None or \
                    os.path.normcase(os.path.abspath(current_work.FullPath)) != \
                    os.path.normcase(os.path.abspath(original_work_part.FullPath)):
                session.Parts.SetWork(original_work_part)
                log(T("log.work_restored", original_work_part.FullPath))
        except Exception as e:
            log(T("log.work_restore_fail", e))

    log("=" * 60)
    log(T("log.summary", ok, skipped, len(failed)))
    for name in failed:
        log(T("log.failed_item", name))
    log("=" * 60)

    _log_handle.close()
    _log_handle = None


# ==============================================================================
# 入口：交互主流程 / 后台准备会话 / 独立批处理会话
# ==============================================================================
if __name__ == "__main__":
    # 独立批处理会话（run_journal.exe）：只跑计划清单，不弹任何窗
    _batch_plan = os.environ.get(ENV_PLAN) or ""
    # 后台准备会话（run_journal.exe -nx）：等选择窗结果 → 预览 → 起批处理
    _batch_stage = os.environ.get(ENV_STAGE) or ""


    def _stage_kind(path):
        """准备会话的 stage.json 归属：新文件带 kind，旧文件按字段判别"""
        try:
            with open(path, "r", encoding="utf-8-sig") as _h:
                _data = json.load(_h)
        except Exception:
            return "export"
        _k = str(_data.get("kind") or "")
        if _k in ("export", "import"):
            return _k
        return "export" if "select_out" in _data else "import"


    # plan 优先：worker 由进度窗（继承自准备会话）启动，环境里可能同时残留
    # NX_BATCH_STAGE。若 stage 先判，worker 会重入准备阶段，弹出第二个确认框。
    if _batch_plan:
        try:
            _plan_kind = str(load_plan(_batch_plan).get("kind") or "export")
        except Exception:
            _plan_kind = "export"
        if _plan_kind == "import":
            _run_with_hider(lambda: imp__run_background_worker(_batch_plan))
        else:
            _run_background_worker(_batch_plan)
    elif _batch_stage:
        if _stage_kind(_batch_stage) == "import":
            _run_with_hider(lambda: imp__run_prepare_stage(_batch_stage))
        else:
            _run_prepare_stage(_batch_stage)
    else:
        # 过渡：执行期间把 NX 自己的「工作进行中」框藏起来（脚本结束/异常时恢复）
        _waitdlg = None
        try:
            _waitdlg = _load_runner().start_waitdlg_hider(log)
        except Exception:
            _waitdlg = None
        try:
            main()
        finally:
            if _waitdlg is not None:
                _waitdlg.stop()
