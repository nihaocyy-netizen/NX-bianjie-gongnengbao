# NX 批量工具 —— 卸载脚本（custom_dirs.dat 外科式删除）
# 作者：Wong、Htet、Limulu
# 只删除本插件那一行；其他行、换行、只读属性原样保留；不整份还原备份。
# 并把安装期设置的 UGII_USER_DIR 恢复到安装前的值（记录在 install_env_state.json）。
# 不删除 .prt / .step / 日志 / 本插件文件夹。
# 用法：NX便捷功能包.exe --uninstall --nx-dir "D:\NX2506"，或 powershell -File uninstall.ps1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$NxDir = "",
    [string]$InstallDir = "",
    [string]$CustomDirs = "",
    [ValidateSet("User", "Process")][string]$Scope = "User",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$BundleName = "NX工具"
$BackupSuffix = ".nxbat.bak"

# custom_dirs.dat 一律以 UTF-8 无 BOM 写回；读取时逐行兼容 UTF-8 / GBK 旧行
$Script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$Script:StrictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$Script:StrictGbk = [System.Text.Encoding]::GetEncoding(936,
    [System.Text.EncoderFallback]::ExceptionFallback, [System.Text.DecoderFallback]::ExceptionFallback)
$Script:Latin1 = [System.Text.Encoding]::GetEncoding(28591)

function Fail([string]$msg) {
    Write-Host ""
    Write-Host "[失败] $msg" -ForegroundColor Red
    exit 1
}

function Normalize-Dir([string]$path) {
    if ($path -and $path.Length -gt 3) { return $path.TrimEnd('\') }
    return $path
}

function Get-FileState([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) {
        return @{ Exists = $false; Bytes = 0; Sha = "(none)" }
    }
    $item = Get-Item -LiteralPath $path -Force
    $sha = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    return @{ Exists = $true; Bytes = $item.Length; Sha = $sha }
}

function Test-IsReadOnly([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    $attrs = (Get-Item -LiteralPath $path -Force).Attributes
    return [bool]($attrs -band [System.IO.FileAttributes]::ReadOnly)
}

function Set-ReadOnly([string]$path, [bool]$on) {
    $attrs = (Get-Item -LiteralPath $path -Force).Attributes
    if ($on) { $attrs = $attrs -bor [System.IO.FileAttributes]::ReadOnly }
    else { $attrs = $attrs -band (-bnot [System.IO.FileAttributes]::ReadOnly) }
    (Get-Item -LiteralPath $path -Force).Attributes = $attrs
}

function Decode-Line([byte[]]$bytes) {
    # custom_dirs.dat 可能混用编码（逐行尝试 严格 UTF-8 / 严格 GBK / Latin-1）
    foreach ($enc in @($Script:StrictUtf8, $Script:StrictGbk, $Script:Latin1)) {
        try { return $enc.GetString($bytes) } catch { }
    }
    return ""
}

function Split-RawLines([byte[]]$raw) {
    $lines = @()
    $start = 0
    for ($i = 0; $i -lt $raw.Length; $i++) {
        if ($raw[$i] -eq 10) {
            $lines += ,([byte[]]($raw[$start..$i]))
            $start = $i + 1
        }
    }
    if ($start -lt $raw.Length) {
        $lines += ,([byte[]]($raw[$start..($raw.Length - 1)]))
    }
    # 逗号阻止返回时展开：只有一行时 $lines 会被拆成 byte[]，调用方 foreach 会按字节迭代，
    # 导致整行被拆成单字符行（实测：custom_dirs.dat 只有一行时被写成 "# / 空格 / f / a / k / e"）
    return ,$lines
}

function Test-PluginLineText([string]$text, [string]$bundle) {
    $t = $text.Trim()
    if (-not $t) { return $false }
    if ($t.StartsWith("#")) { return $false }
    try { $full = [System.IO.Path]::GetFullPath($t) } catch { return $false }
    return ((Normalize-Dir $full).ToLower() -eq (Normalize-Dir $bundle).ToLower())
}

# 读出「逐行解码后的文本 + 换行风格 + 末尾是否有换行 + 只读属性」
function Read-CustomDirsState([string]$path) {
    $state = @{ Exists = $false; Lines = @(); Nl = "`r`n"; Trailing = $true; ReadOnly = $false }
    if (-not (Test-Path -LiteralPath $path)) { return $state }
    $state.Exists = $true
    $state.ReadOnly = Test-IsReadOnly $path
    $raw = [System.IO.File]::ReadAllBytes($path)
    if ($raw.Length -eq 0) { return $state }
    $hasCr = $false
    foreach ($b in $raw) { if ($b -eq 13) { $hasCr = $true; break } }
    if (-not $hasCr) { $state.Nl = "`n" }
    $state.Trailing = ($raw[$raw.Length - 1] -eq 10)
    $lines = @()
    $first = $true
    foreach ($seg in (Split-RawLines $raw)) {
        $b = [byte[]]$seg
        $n = $b.Length
        while ($n -gt 0 -and ($b[$n - 1] -eq 10 -or $b[$n - 1] -eq 13)) { $n-- }
        $text = ""
        if ($n -gt 0) { $text = Decode-Line ([byte[]]($b[0..($n - 1)])) }
        if ($first) {
            $first = $false
            if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
        }
        $lines += $text
    }
    $state.Lines = $lines
    return $state
}

# 整文件写回：UTF-8 无 BOM；写前自检、临时文件 + 原子替换；失败即抛错停下
function Write-CustomDirsFile([string]$path, [string[]]$lines, [string]$nl, [bool]$trailing) {
    $text = ($lines -join $nl)
    if ($trailing -and $lines.Length -gt 0) { $text += $nl }
    $bytes = $Script:Utf8NoBom.GetBytes($text)
    $check = ""
    try { $check = $Script:StrictUtf8.GetString($bytes) }
    catch { throw "写入自检失败：生成的字节不是合法 UTF-8（$path），已中止，原文件未改动" }
    if ($check -ne $text -or $check.IndexOf([char]0xFFFD) -ge 0) {
        throw "写入自检失败：存在无法无损转为 UTF-8 的字符（$path），已中止，原文件未改动"
    }
    $dir = Split-Path -Parent $path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = $path + ".nxfull.tmp"
    [System.IO.File]::WriteAllBytes($tmp, $bytes)
    try {
        if (Test-Path -LiteralPath $path) {
            try { [System.IO.File]::Replace($tmp, $path, [NullString]::Value) }
            catch [System.IO.IOException] { [System.IO.File]::WriteAllBytes($path, $bytes); [System.IO.File]::Delete($tmp) }
        }
        else { [System.IO.File]::Move($tmp, $path) }
    }
    finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Read-EnvState([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $raw = [System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($false)))
        return ($raw | ConvertFrom-Json)
    }
    catch { return $null }
}

function Restore-UserDirEnv([string]$bundle, [string]$statePath, [bool]$preview) {
    $out = @{ Action = "none"; Message = ""; Target = "" }
    $cur = [Environment]::GetEnvironmentVariable("UGII_USER_DIR", "User")
    $st = Read-EnvState $statePath
    $prev = $null
    if (($null -ne $st) -and ($st.PSObject.Properties.Name -contains "prevUser")) { $prev = $st.prevUser }
    $curTrim = Normalize-Dir $cur
    $curIsOurs = ($curTrim -and $curTrim.ToLower() -eq $bundle.ToLower())
    if ($curTrim -and -not $curIsOurs -and ($null -ne $prev) -and (Normalize-Dir $prev).ToLower() -ne $curTrim.ToLower()) {
        $out.Action = "skip"
        $out.Message = "当前 UGII_USER_DIR 指向其他目录（$cur），未改动"
        return $out
    }
    if (-not $curTrim -and ($null -eq $st)) {
        $out.Message = "当前为空且无安装记录，无需处理"
        return $out
    }
    # 没有安装记录且当前值不属于本插件：说明不是本插件设置的，绝不能清空
    if (-not $curIsOurs -and ($null -eq $st)) {
        $out.Action = "skip"
        $out.Message = "无安装记录，且 UGII_USER_DIR 指向其他目录（$cur），未改动"
        return $out
    }
    if ($preview) { $out.Action = "preview"; $out.Message = "预览：将恢复到安装前的值"; return $out }
    if ($prev -and (Normalize-Dir $prev).Length -gt 0) {
        [Environment]::SetEnvironmentVariable("UGII_USER_DIR", (Normalize-Dir $prev), "User")
        $out.Action = "restored"
        $out.Target = (Normalize-Dir $prev)
        $out.Message = "已恢复到安装前的值：$($out.Target)"
    }
    else {
        [Environment]::SetEnvironmentVariable("UGII_USER_DIR", $null, "User")
        $out.Action = "cleared"
        $out.Message = "已删除（安装前该变量为空）"
    }
    if (Test-Path -LiteralPath $statePath) { Remove-Item -LiteralPath $statePath -Force }
    return $out
}

# 旧模块化安装（v3.3.5 及更早）留下的 Htet工具 / 格式转换 两行：一并摘掉，目录不删
function Remove-LegacyLines([string]$path, [string[]]$legacyDirs) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $state = Read-CustomDirsState $path
    if (-not $state.Exists) { return @() }
    $legacyNorm = @()
    foreach ($d in $legacyDirs) { $legacyNorm += (Normalize-Dir $d).ToLower() }
    $kept = @(); $removed = @()
    foreach ($line in $state.Lines) {
        $t = $line.Trim()
        $isLegacy = $false
        if ($t -and -not $t.StartsWith('#')) {
            try { $full = (Normalize-Dir ([System.IO.Path]::GetFullPath($t))).ToLower() } catch { $full = "" }
            if ($full -and ($legacyNorm -contains $full)) { $isLegacy = $true }
        }
        if ($isLegacy) { $removed += $t } else { $kept += $line }
    }
    if ($removed.Count -eq 0) { return @() }
    $bak = $path + $BackupSuffix
    if (-not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $path -Destination $bak -Force }
    $wasReadOnly = $state.ReadOnly
    if ($wasReadOnly) { Set-ReadOnly $path $false }
    try { Write-CustomDirsFile $path ([string[]]$kept) $state.Nl $state.Trailing }
    finally { if ($wasReadOnly -and (Test-Path -LiteralPath $path)) { Set-ReadOnly $path $true } }
    return $removed
}

# 入口：-InstallDir 指「包根」，插件目录固定为 <包根>\NX工具。
if (-not $InstallDir) { $InstallDir = $PSScriptRoot }
if (-not (Test-Path -LiteralPath $InstallDir -PathType Container)) {
    Fail "InstallDir 不存在或不是文件夹: $InstallDir"
}
$packageRoot = Normalize-Dir ((Resolve-Path -LiteralPath $InstallDir).Path)
$bundle = Join-Path $packageRoot $BundleName
$LegacyDirs = @((Join-Path $packageRoot "Htet工具"), (Join-Path $packageRoot "格式转换"))

Write-Host "=============================================="
Write-Host "  NX便捷功能包 —— 卸载"
Write-Host "=============================================="
Write-Host "插件目录 : $bundle"

if (-not $CustomDirs) {
    if (-not $NxDir) {
        Write-Host ""
        Write-Host "请选择 NX 安装目录（包含 UGII 文件夹的那一层，例如 D:\NX2506）"
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "选择 NX 安装目录（包含 UGII 文件夹）"
        $dialog.ShowNewFolderButton = $false
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "已取消，未做任何修改。"
            exit 0
        }
        $NxDir = $dialog.SelectedPath
    }
    if (-not (Test-Path -LiteralPath $NxDir)) { Fail "NX 安装目录不存在: $NxDir" }
    $menuDir = Join-Path (Join-Path (Resolve-Path -LiteralPath $NxDir).Path "UGII") "menus"
    $customDirs = Join-Path $menuDir "custom_dirs.dat"
} else {
    $customDirs = [System.IO.Path]::GetFullPath($CustomDirs)
}
Write-Host "注册文件 : $customDirs"

if (-not $Force) {
    Write-Host ""
    Write-Host "将执行："
    Write-Host "  1) 从 custom_dirs.dat 精确删除本插件行（其他行、换行、只读属性保持不变）"
    Write-Host "  2) 把 UGII_USER_DIR 恢复到安装前的值（安装时记录在 install_env_state.json；
     安装前为空则删除该项）"
    Write-Host "  3) 仅当存在安装备份（*.nxbat.bak）时，把 application\ 下的 3 份 .dlx 原样回填"
    Write-Host "     （不会删除用户原有的 dlx；没有备份的一份一律保持现状）"
    Write-Host "  4) 只删除 startup\ 内与 application\ 同哈希的脚本副本（用户改过的副本保留）"
    Write-Host "不会删除任何 .prt / .step / 日志文件，也不会删除本文件夹。"
    $answer = Read-Host "继续？(Y/N)"
    if ($answer -notmatch '^[Yy]') {
        Write-Host "已取消，未做任何修改。"
        exit 0
    }
}

$before = Get-FileState $customDirs
Write-Host ("写前状态 : {0} 字节  SHA256 {1}" -f $before.Bytes, $before.Sha)

$removed = $false
if (-not (Test-Path -LiteralPath $customDirs)) {
    Write-Host "注册文件 : 不存在，无需卸载"
}
else {
    $state = Read-CustomDirsState $customDirs
    $has = $false
    foreach ($line in $state.Lines) { if (Test-PluginLineText $line $bundle) { $has = $true; break } }
    if (-not $has) {
        Write-Host "注册文件 : 没有本插件行，无需卸载"
    }
    elseif ($PSCmdlet.ShouldProcess($customDirs, "删除本插件行")) {
        $bak = $customDirs + $BackupSuffix
        if (-not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $customDirs -Destination $bak -Force }
        $wasReadOnly = $state.ReadOnly
        if ($wasReadOnly) { Set-ReadOnly $customDirs $false }
        try {
            # 字节级外科删除：只去掉本插件行，其他行（含 GBK 行）原字节、原换行一律不动
            $raw = [System.IO.File]::ReadAllBytes($customDirs)
            $ms = New-Object System.IO.MemoryStream
            $hit = $false
            $trimSep = $false
            $first = $true
            # 注意：不能写成 @(Split-RawLines $raw)——函数用 return ,$lines 返回，
            # @() 会把整个行数组再包一层，导致 $segs[0] 是「行数组」而不是单行
            $segs = Split-RawLines $raw
            try {
                for ($si = 0; $si -lt $segs.Count; $si++) {
                    $b = [byte[]]$segs[$si]
                    $n = $b.Length
                    while ($n -gt 0 -and ($b[$n - 1] -eq 10 -or $b[$n - 1] -eq 13)) { $n-- }
                    $text = ""
                    if ($n -gt 0) { $text = Decode-Line ([byte[]]($b[0..($n - 1)])) }
                    if ($first -and $text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
                    $first = $false
                    if (Test-PluginLineText $text $bundle) {
                        $hit = $true
                        # 末行且本行无换行：安装时补的分隔符需要一并去掉，才能逐字节还原
                        if ($si -eq ($segs.Count - 1) -and $n -eq $b.Length) { $trimSep = $true }
                        continue
                    }
                    $ms.Write($b, 0, $b.Length)
                }
                $out = $ms.ToArray()
                if ($trimSep -and $out.Length -gt 0) {
                    $len = $out.Length
                    if ($len -ge 2 -and $out[$len - 2] -eq 13 -and $out[$len - 1] -eq 10) { $len -= 2 }
                    elseif ($out[$len - 1] -eq 10) { $len -= 1 }
                    $trimmed = New-Object byte[] $len
                    if ($len -gt 0) { [Array]::Copy($out, 0, $trimmed, 0, $len) }
                    $out = $trimmed
                }
            }
            finally { $ms.Dispose() }
            if ($hit) {
                $tmp = $customDirs + ".nxfull.tmp"
                [System.IO.File]::WriteAllBytes($tmp, $out)
                [System.IO.File]::Replace($tmp, $customDirs, [NullString]::Value)
            }
        }
        finally { if ($wasReadOnly -and (Test-Path -LiteralPath $customDirs)) { Set-ReadOnly $customDirs $true } }
        $removed = $true
        Write-Host "注册结果 : 已删除本插件行（其他行原始字节与换行保持不变）"
    }
    else {
        Write-Host "注册结果 : 未改动（-WhatIf 预览）"
    }
}

$after = Get-FileState $customDirs
Write-Host ("写后状态 : {0} 字节  SHA256 {1}" -f $after.Bytes, $after.Sha)

# 旧模块化安装残留：再摘一次旧的两行（不删目录）
$legacyRemoved = @()
if ($PSCmdlet.ShouldProcess($customDirs, "移除旧模块化注册行（Htet工具 / 格式转换）")) {
    $legacyRemoved = Remove-LegacyLines $customDirs $LegacyDirs
}
if ($legacyRemoved.Count -gt 0) {
    Write-Host ("旧版清理 : 已移除旧模块化注册行 {0} 条（旧目录保留）" -f $legacyRemoved.Count)
    foreach ($l in $legacyRemoved) { Write-Host ("  - {0}" -f $l) }
    $after = Get-FileState $customDirs
    Write-Host ("写后状态 : {0} 字节  SHA256 {1}" -f $after.Bytes, $after.Sha)
}

# 恢复安装期写入的对话框资源（BlockStyler .dlx：Htet_auto_drawing / export_body01 / select_bodies）
# 说明：只做“有安装备份才回填”；没有 .nxbat.bak 的一律保持现状，绝不删除用户原有的 dlx。
$dlxNames = @("Htet_auto_drawing.dlx", "Htet_export_body01.dlx", "Htet_select_bodies.dlx")
$dlxRestored = 0
foreach ($name in $dlxNames) {
    $dlxDst = Join-Path (Join-Path $bundle "application") $name
    $dlxBak = $dlxDst + $BackupSuffix
    if (-not (Test-Path -LiteralPath $dlxBak)) {
        Write-Host ("dlx 处理 : {0} 无安装备份，保持现状（未删除任何文件）" -f $name)
        continue
    }
    if ($PSCmdlet.ShouldProcess($dlxDst, "恢复安装前的 dlx 并删除备份")) {
        Copy-Item -LiteralPath $dlxBak -Destination $dlxDst -Force
        Remove-Item -LiteralPath $dlxBak -Force
        $dlxRestored++
        Write-Host ("dlx 处理 : {0} 已恢复安装前文件，备份已删除（无残余）" -f $name)
    }
    else {
        Write-Host ("dlx 处理 : {0} 未改动（-WhatIf 预览）" -f $name)
    }
}

# 移除安装期部署到 startup\ 的脚本副本（仅删除与 application\ 内容一致的副本）
$scriptRemoved = 0
$startupDir = Join-Path $bundle "startup"
    foreach ($py in (Get-ChildItem -LiteralPath (Join-Path $bundle "application") -Filter "*.py" -File)) {
    if ($py.Name -eq "format_convert.py") { continue }   # 母本在 startup\，不是部署副本
    $dst = Join-Path $startupDir $py.Name
    if (-not (Test-Path -LiteralPath $dst)) { continue }
    if ((Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $py.FullName -Algorithm SHA256).Hash) { continue }
    if ($PSCmdlet.ShouldProcess($dst, "移除 startup 脚本副本")) { Remove-Item -LiteralPath $dst -Force; $scriptRemoved++ }
}
Write-Host "脚本清理 : 已移除 startup 内脚本副本 $scriptRemoved 个"

# 移除安装期部署到 startup\ 的对话框副本（仅删除与 application\ 母本一致的副本，不删用户自备文件）
foreach ($dlxName in @("Quickhole.dlx", "nx_screw_insert.dlx")) {
    $dlxDst = Join-Path $startupDir $dlxName
    $dlxSrc = Join-Path (Join-Path $bundle "application") $dlxName
    if ((Test-Path -LiteralPath $dlxDst) -and (Test-Path -LiteralPath $dlxSrc) -and
        ((Get-FileHash -LiteralPath $dlxDst -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $dlxSrc -Algorithm SHA256).Hash)) {
        if ($PSCmdlet.ShouldProcess($dlxDst, "移除 startup $dlxName 副本")) {
            Remove-Item -LiteralPath $dlxDst -Force
            Write-Host "对话框清理 : 已移除 startup\$dlxName"
        }
    }
}

# 移除安装期部署到 startup\screw_icons\ 的位图副本（逐个比对母本，删完空目录才移除；
# 用户自建的图标不会被删，目录里只要还有别的文件就保留）。
$iconSrcDir = Join-Path (Join-Path $bundle "application") "screw_icons"
$iconDstDir = Join-Path $startupDir "screw_icons"
$iconRemoved = 0
if ((Test-Path -LiteralPath $iconSrcDir -PathType Container) -and (Test-Path -LiteralPath $iconDstDir -PathType Container)) {
    foreach ($icon in (Get-ChildItem -LiteralPath $iconSrcDir -Filter "*.bmp" -File)) {
        $dst = Join-Path $iconDstDir $icon.Name
        if (-not (Test-Path -LiteralPath $dst)) { continue }
        if ((Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $icon.FullName -Algorithm SHA256).Hash) { continue }
        if ($PSCmdlet.ShouldProcess($dst, "移除 startup 螺钉图标副本")) { Remove-Item -LiteralPath $dst -Force; $iconRemoved++ }
    }
    if (@(Get-ChildItem -LiteralPath $iconDstDir -Force).Count -eq 0) {
        if ($PSCmdlet.ShouldProcess($iconDstDir, "移除空的 startup\screw_icons")) { Remove-Item -LiteralPath $iconDstDir -Force }
    }
}
Write-Host "图标清理 : 已移除 startup\screw_icons 副本 $iconRemoved 个"

$envStatePath = Join-Path $bundle "install_env_state.json"
$envResult = $null
if ($PSCmdlet.ShouldProcess("UGII_USER_DIR", "恢复安装前的值")) {
    $envResult = Restore-UserDirEnv $bundle $envStatePath ([bool]$WhatIfPreference)
}
else {
    $envResult = @{ Action = "preview"; Message = "预览：将恢复到安装前的值"; Target = "" }
}
if ($envResult.Action -eq "skip") { Write-Host ("环境变量 : " + $envResult.Message) -ForegroundColor Yellow }
else { Write-Host ("环境变量 : " + $envResult.Message) }

$installLog = Join-Path $env:TEMP "nx_batch_tools_install.log"
$logLine = "{0}  UNINSTALL  bundle={1}  custom_dirs={2}  before={3}B/{4}  after={5}B/{6}  removed={7}  dlx_restored={8}  env_action={9}  env_target={10}" -f `
    (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $bundle, $customDirs,
    $before.Bytes, $before.Sha, $after.Bytes, $after.Sha, $removed, $dlxRestored,
    $envResult.Action, $envResult.Target
[System.IO.File]::AppendAllText($installLog, $logLine + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "安装日志 : $installLog"

Write-Host ""
Write-Host "卸载完成。重启 NX 后菜单将不再出现。"
Write-Host "如需彻底清理，请手动删除整个文件夹：$bundle"
Write-Host "（脚本不会删除自身所在目录，避免误删其他内容）"
Write-Host ""
