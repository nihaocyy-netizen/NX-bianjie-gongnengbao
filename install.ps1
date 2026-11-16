# NX便捷功能包 v3.3.6 —— 合并模块安装脚本（NX工具 单一目录，custom_dirs.dat 一行）
# 作者：Wong、Htet、Limulu
# 作用：1) 只把 <包根>\NX工具 追加到 <NX>\UGII\menus\custom_dirs.dat 的一行；
#          并移除 v3.3.5 及更早模块化安装留下的 Htet工具 / 格式转换 两行（不删目录）
#       2) 按界面语言把 lang\ 下的菜单/工具条/Ribbon 母本转码写入 startup\
#          把 lang\ 下三份 BlockStyler 对话框（.dlx）覆盖到 application\
#          并把 application\ 下的 .py / .vb 与 Quickhole.dlx、nx_screw_insert.dlx、
#          screw_icons\ 部署到 startup\
#       3) 设置用户级 UGII_USER_DIR 指向插件根（NX 据此搜索 startup\ 与 application\ 下的 .py；
#          写入前记录原值到 install_env_state.json，卸载时按原值恢复/删除）
#          若已存在指向其他目录的 UGII_USER_DIR，则不改动并告警（避免带偏其他工具）
# 用法：NX便捷功能包.exe，或 powershell -File install.ps1 -NxDir "D:\NX2506"
# 测试参数：-CustomDirs <副本路径>（不写真实 NX 配置）

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$NxDir = "",
    [string]$InstallDir = "",
    [ValidateScript({
        if ($_ -notmatch '^(simpl_chinese|simplified_chinese|schinese|zh_cn|zh-cn|chs|zh_hans|trad_chinese|traditional_chinese|chinese|zh_tw|zh-tw|cht|zh_hant|english|en|en_us|en-us)$') {
            throw "不支持的 -Lang 值: $_"
        }
        return $true
    })]
    [string]$Lang = "",
    [string]$CustomDirs = "",
    [ValidateSet("User", "Process")][string]$Scope = "User",
    # 写盘码页（系统默认 locale）表示不了某个字符时：默认严格报错并列出缺字；显式加此开关则按字回退（替换为 ?）。
    [switch]$AllowCodePageFallback
)

$ErrorActionPreference = "Stop"
$BundleName = "NX工具"
$BackupSuffix = ".nxbat.bak"

# custom_dirs.dat 必须写成 UTF-8 无 BOM：NX 2506 按 UTF-8 读该文件，
# 而 PowerShell 的默认编码在本机是 GBK(936)，含中文的插件路径会被写成乱码，
# 整行被 NX 判为无效目录而丢弃（实测：只剩 Htet工具 一个子菜单）。
$Script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$Script:StrictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$Script:StrictGbk = [System.Text.Encoding]::GetEncoding(936,
    [System.Text.EncoderFallback]::ExceptionFallback, [System.Text.DecoderFallback]::ExceptionFallback)
$Script:Latin1 = [System.Text.Encoding]::GetEncoding(28591)
$Script:AllowCodePageFallback = [bool]$AllowCodePageFallback

function Fail([string]$msg) {
    Write-Host ""
    Write-Host "[失败] $msg" -ForegroundColor Red
    exit 1
}

# SHA256 短摘要显示：Get-FileState 在文件不存在时返回 "(none)"（6 字符），
# 直接 .Substring(0,8) 会抛「索引和长度必须引用该字符串内的位置」，这里统一加保护。
function Short-Sha([string]$sha) {
    if ([string]::IsNullOrEmpty($sha)) { return "(none)" }
    if ($sha.Length -le 8) { return $sha }
    return $sha.Substring(0, 8)
}

function Normalize-Dir([string]$path) {
    if ($path -and $path.Length -gt 3) { return $path.TrimEnd('\') }
    return $path
}

function Resolve-LangCode([string]$value) {
    switch -Regex ($value.Trim().ToLower()) {
        '^(simpl_chinese|simplified_chinese|schinese|zh_cn|zh-cn|chs|zh_hans)$' { return 'zh_cn' }
        '^(trad_chinese|traditional_chinese|chinese|zh_tw|zh-tw|cht|zh_hant)$' { return 'zh_tw' }
        '^(english|en|en_us|en-us)$' { return 'en' }
        default { return 'en' }
    }
}

function Get-LangName([string]$code) {
    switch ($code) {
        'zh_cn' { return '简体中文' }
        'zh_tw' { return '繁體中文' }
        default { return 'English' }
    }
}

function Get-FileState([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) {
        return @{ Exists = $false; Bytes = 0; Sha = "(none)" }
    }
    $item = Get-Item -LiteralPath $path -Force
    $sha = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    return @{ Exists = $true; Bytes = $item.Length; Sha = $sha }
}

# 落盘码页跟随「系统默认 locale」：三组同机实测（CP950 轮）表明 NX 解 .men 用的是
# GetSystemDefaultLCID 对应的 ANSI 码页，且该 locale 在系统区域变更后需重启才刷新——
# 未重启时它可能既不是用户区域、也不是即时的 GetACP()。
function Get-SystemAnsiCodePage {
    # 系统 ANSI 代码页（等价 GetACP），用于日志对照与系统默认 locale 取不到时回退
    $cp = 0
    try { $cp = [int](Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage' -Name ACP -ErrorAction Stop).ACP } catch { }
    if ($cp -le 0) { $cp = [System.Text.Encoding]::Default.CodePage }
    return $cp
}

function Get-SystemDefaultLcid {
    # 系统默认 locale 的 LCID（P/Invoke GetSystemDefaultLCID）。系统区域变更后需重启才刷新。
    try {
        if (-not ("NXCodePage.Native" -as [type])) {
            Add-Type -Namespace NXCodePage -Name Native -MemberDefinition '[System.Runtime.InteropServices.DllImport("kernel32.dll")] public static extern uint GetSystemDefaultLCID();'
        }
        return [int][NXCodePage.Native]::GetSystemDefaultLCID()
    } catch { return 0 }
}

function Get-SystemDefaultLocaleCodePage {
    # 系统默认 locale（GetSystemDefaultLCID）的 LCID / 名称 / ANSI 码页，供写盘与日志核对
    $lcid = Get-SystemDefaultLcid
    $name = ''
    $cp = 0
    if ($lcid -gt 0) {
        try {
            $ci = New-Object System.Globalization.CultureInfo($lcid)
            $name = [string]$ci.Name
            $cp = [int]$ci.TextInfo.ANSICodePage
        } catch { $name = ''; $cp = 0 }
    }
    return @{ Lcid = $lcid; Name = $name; CodePage = $cp }
}

function Get-TargetCodePage([string]$langCode) {
    # 优先系统默认 locale 的 ANSI 码页（NX 实际按它解 .men）；取不到再回退系统 ACP。
    # $langCode 只用于选择母本语言，不参与码页判定。
    $sysLocale = Get-SystemDefaultLocaleCodePage
    if ($sysLocale.CodePage -gt 0) { return [int]$sysLocale.CodePage }
    return Get-SystemAnsiCodePage
}

function Get-StrictEncoding([int]$codePage) {
    return [System.Text.Encoding]::GetEncoding($codePage,
        [System.Text.EncoderFallback]::ExceptionFallback,
        [System.Text.DecoderFallback]::ExceptionFallback)
}

function Get-UnmappableChars([string]$text, [int]$codePage) {
    $enc = Get-StrictEncoding $codePage
    $bad = @()
    foreach ($ch in $text.ToCharArray()) {
        try { [void]$enc.GetBytes([string]$ch) }
        catch { if ($bad -notcontains [string]$ch) { $bad += [string]$ch } }
    }
    return $bad
}

function Convert-MenuResource([string]$srcFile, [string]$dstFile, [string]$langCode) {
    # 母本只读（P0-2）：srcFile 恒为 UTF-8 无 BOM，绝不就地改写；每次从母本重新转码。
    # 写完回读校验（P0-3）：按写盘码页（系统默认 locale 的 ANSI）解码回来必须与母本文本完全相同，否则抛错。
    $cp = Get-TargetCodePage $langCode
    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    try { $text = $strictUtf8.GetString([System.IO.File]::ReadAllBytes($srcFile)) }
    catch { throw "母本不是严格 UTF-8：$srcFile（lang\ 下应是未转码原文，请勿就地转码）" }

    $usedFallback = $false
    $enc = Get-StrictEncoding $cp
    try { $bytes = $enc.GetBytes($text) }
    catch [System.Text.EncoderFallbackException] {
        $bad = Get-UnmappableChars $text $cp
        if (-not $Script:AllowCodePageFallback) {
            throw ("CP{0} 无法表示以下字符：{1}（母本 {2}）。若确需安装，可加 -AllowCodePageFallback 按字回退。" -f `
                $cp, ($bad -join ' '), $srcFile)
        }
        $bytes = [System.Text.Encoding]::GetEncoding($cp).GetBytes($text)
        $usedFallback = $true
    }
    [System.IO.File]::WriteAllBytes($dstFile, $bytes)

    $back = (Get-StrictEncoding $cp).GetString([System.IO.File]::ReadAllBytes($dstFile))
    # 回退模式下必然有字符被替换，此时已在上方列出缺字，不再做逐字相等校验
    if ((-not $usedFallback) -and ($back -ne $text)) {
        throw ("回读校验失败：{0} 按 CP{1} 解码后与母本不一致（写入 {2} 字节）" -f $dstFile, $cp, $bytes.Length)
    }
    return @{ Encoding = "CP$cp"; CodePage = $cp; Fallback = $usedFallback; Bytes = $bytes.Length }
}

function Write-LangResource([string]$srcFile, [string]$dstFile, [string]$langCode) {
    # 母本只读：从 lang\ 原文直接转码写到 startup\（NX 只读取 startup\ 里的生效文件）
    return (Convert-MenuResource $srcFile $dstFile $langCode)
}

# ---------- custom_dirs.dat 行级读写 ----------
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
    # custom_dirs.dat 可能混用编码：逐行依次尝试 严格 UTF-8 -> 严格 GBK(936) -> Latin-1
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

# 移除 v3.3.5 及更早模块化安装留下的两行（Htet工具 / 格式转换）。只摘注册行，不删目录、不记入卸载清单。
function Remove-LegacyPluginLines([string]$path, [string[]]$legacyDirs) {
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
    [void](Backup-Once $path)
    $wasReadOnly = $state.ReadOnly
    if ($wasReadOnly) { Set-ReadOnly $path $false }
    try { Write-CustomDirsFile $path ([string[]]$kept) $state.Nl $state.Trailing }
    finally { if ($wasReadOnly -and (Test-Path -LiteralPath $path)) { Set-ReadOnly $path $true } }
    return $removed
}

$PluginBundle = $null

# 读出「逐行解码后的文本 + 换行风格 + 末尾是否有换行 + 只读属性 + 整文件是否已是合法 UTF-8」
function Read-CustomDirsState([string]$path) {
    $state = @{ Exists = $false; HasLine = $false; Lines = @(); Nl = "`r`n"; Trailing = $true; ReadOnly = $false; IsUtf8 = $true }
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
    try { [void]$Script:StrictUtf8.GetString($raw); $state.IsUtf8 = $true } catch { $state.IsUtf8 = $false }
    foreach ($line in $lines) {
        if (Test-PluginLineText $line $PluginBundle) { $state.HasLine = $true; break }
    }
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

function Backup-Once([string]$path) {
    $bak = $path + $BackupSuffix
    if (-not (Test-Path -LiteralPath $bak) -and (Test-Path -LiteralPath $path)) {
        Copy-Item -LiteralPath $path -Destination $bak -Force
        return $bak
    }
    if (Test-Path -LiteralPath $bak) { return $bak }
    return ""
}

function Add-PluginLine([string]$path) {
    $state = Read-CustomDirsState $path
    # 已有本行：整份文件一个字节都不动
    if ($state.HasLine) { return $false }
    Backup-Once $path | Out-Null
    $wasReadOnly = $state.ReadOnly
    if ($wasReadOnly) { Set-ReadOnly $path $false }
    try {
        $ms = New-Object System.IO.MemoryStream
        try {
            $noTrailing = $false
            if ($state.Exists) {
                $raw = [System.IO.File]::ReadAllBytes($path)
                if ($raw.Length -gt 0) {
                    $ms.Write($raw, 0, $raw.Length)
                    # 末行无换行时先补一个分隔符，保证新行独立；已有字节原样保留
                    if ($raw[$raw.Length - 1] -ne 10) {
                        $eol = $Script:Utf8NoBom.GetBytes($state.Nl)
                        $ms.Write($eol, 0, $eol.Length)
                        # 原文件末尾无换行 → 本行也不补末尾换行，卸载时才能逐字节还原
                        $noTrailing = $true
                    }
                }
            }
            # 只追加本行（UTF-8 无 BOM），绝不整文件重写、绝不转码其他行
            $lineText = (Normalize-Dir $PluginBundle)
            if (-not $noTrailing) { $lineText += $state.Nl }
            $lineBytes = $Script:Utf8NoBom.GetBytes($lineText)
            $ms.Write($lineBytes, 0, $lineBytes.Length)
            $out = $ms.ToArray()
        }
        finally { $ms.Dispose() }
        $tmp = $path + ".nxfull.tmp"
        $dir = Split-Path -Parent $path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [System.IO.File]::WriteAllBytes($tmp, $out)
        if (Test-Path -LiteralPath $path) {
            try { [System.IO.File]::Replace($tmp, $path, [NullString]::Value) }
            catch [System.IO.IOException] { [System.IO.File]::WriteAllBytes($path, $out); [System.IO.File]::Delete($tmp) }
        }
        else { [System.IO.File]::Move($tmp, $path) }
    }
    finally {
        if (Test-Path -LiteralPath ($path + ".nxfull.tmp")) { Remove-Item -LiteralPath ($path + ".nxfull.tmp") -Force -ErrorAction SilentlyContinue }
        if ($wasReadOnly -and (Test-Path -LiteralPath $path)) { Set-ReadOnly $path $true }
    }
    Write-Host "注册结果 : 已追加本插件行（其他行原始字节未改动）" -ForegroundColor Green
    return $true
}

# ---------- 用户级环境变量 UGII_USER_DIR（脚本搜索路径） ----------
function Read-EnvState([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $raw = [System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($false)))
        return ($raw | ConvertFrom-Json)
    }
    catch { return $null }
}

function Write-EnvState([string]$path, $prevUser, [string]$written, [string]$machineValue) {
    $obj = [ordered]@{
        schema          = "nx-batch-tools/env-state/1"
        bundle          = $PluginBundle
        prevUser        = $prevUser
        writtenUserDir  = $written
        machineUserDir  = $machineValue
        updated         = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
    $json = $obj | ConvertTo-Json -Depth 4
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Set-UserDirEnv([string]$bundle, [string]$statePath, [bool]$preview) {
    $result = @{ Action = "none"; Message = ""; Prev = $null; Current = "" }
    # -Scope Process：只在本次进程内设置（自动化测试用），不写用户级、不写状态文件，进程结束即消失。
    $envScope = if ($Scope -eq 'Process') { 'Process' } else { 'User' }
    $persist = ($envScope -eq 'User')
    $curUser = [Environment]::GetEnvironmentVariable("UGII_USER_DIR", $envScope)
    $curMachine = [Environment]::GetEnvironmentVariable("UGII_USER_DIR", "Machine")
    $st = Read-EnvState $statePath
    $prev = $null
    if (($null -ne $st) -and ($st.PSObject.Properties.Name -contains "prevUser")) { $prev = $st.prevUser }
    $result.Current = $curUser
    $result.Prev = $prev
    $curTrim = Normalize-Dir $curUser
    if ($curTrim -and $curTrim.ToLower() -ne $bundle.ToLower()) {
        # P1-2：历史版本常把 UGII_USER_DIR 指向源码/工作区副本（含 startup\Htet_menu.men），
        # 会让 NX 加载到未转码的旧菜单。识别为「本插件的历史副本」时接管，并记录原值供卸载恢复。
        $isLegacyCopy = (Test-Path -LiteralPath (Join-Path $curTrim 'startup\Htet_menu.men')) -or
                        (Test-Path -LiteralPath (Join-Path $curTrim 'startup\format_convert.py'))
        if (-not $isLegacyCopy) {
            $result.Action = "skip"
            $result.Message = "已存在指向其他目录的 UGII_USER_DIR（$curUser），未覆盖；如本插件脚本仍找不到，请人工确认该目录归属"
            return $result
        }
        if ($preview) { $result.Action = "preview"; $result.Message = "预览：将把 UGII_USER_DIR 从本插件历史副本 $curUser 改指到 $bundle"; return $result }
        [Environment]::SetEnvironmentVariable("UGII_USER_DIR", $bundle, $envScope)
        $verify = [Environment]::GetEnvironmentVariable("UGII_USER_DIR", $envScope)
        if ($persist) { Write-EnvState $statePath $curUser $bundle $curMachine }
        $result.Action = "takeover"
        $result.Prev = $curUser
        $result.Message = if ($persist) {
            "UGII_USER_DIR 原指向本插件历史副本 $curUser，已改指 $verify（卸载时恢复原值）"
        } else {
            "UGII_USER_DIR 原指向本插件历史副本 $curUser，已在进程范围改指 $verify（测试用，未写用户环境）"
        }
        return $result
    }
    if ($curTrim -and $curTrim.ToLower() -eq $bundle.ToLower()) {
        if (-not $persist) {
            $result.Action = "unchanged"
            $result.Message = "UGII_USER_DIR 已指向本插件（进程范围，未写用户环境）"
            return $result
        }
        if (-not (Test-Path -LiteralPath $statePath)) {
            Write-EnvState $statePath $null $bundle $curMachine
            $result.Message = "UGII_USER_DIR 已指向本插件；已补写状态文件（卸载时按空值删除）"
        } else {
            $result.Message = "UGII_USER_DIR 已指向本插件（幂等，未重复写入）"
        }
        $result.Action = "unchanged"
        return $result
    }
    if ($preview) { $result.Action = "preview"; $result.Message = "预览：将设置 UGII_USER_DIR=$bundle"; return $result }
    [Environment]::SetEnvironmentVariable("UGII_USER_DIR", $bundle, $envScope)
    $verify = [Environment]::GetEnvironmentVariable("UGII_USER_DIR", $envScope)
    if ($persist) { Write-EnvState $statePath $null $bundle $curMachine }
    $result.Action = "set"
    $result.Message = if ($persist) {
        "已设置用户级 UGII_USER_DIR=$verify（原值为空，卸载时删除该项）"
    } else {
        "已在进程范围设置 UGII_USER_DIR=$verify（测试用，未写用户环境）"
    }
    return $result
}

# ---------- 入口 ----------
# -InstallDir 现在指「包根」（install.ps1 所在目录），实际插件目录固定为 <包根>\NX工具。
if (-not $InstallDir) { $InstallDir = $PSScriptRoot }
if (-not (Test-Path -LiteralPath $InstallDir -PathType Container)) {
    Fail "InstallDir 不存在或不是文件夹: $InstallDir"
}
$packageRoot = Normalize-Dir ((Resolve-Path -LiteralPath $InstallDir).Path)
$bundle = Join-Path $packageRoot $BundleName
if (-not (Test-Path -LiteralPath $bundle -PathType Container)) {
    Fail "找不到插件目录 $bundle；请确认压缩包已完整解压（包根应含 install.ps1 与 $BundleName\）"
}
$PluginBundle = $bundle
$LegacyDirs = @((Join-Path $packageRoot "Htet工具"), (Join-Path $packageRoot "格式转换"))

Write-Host "=============================================="
Write-Host "  NX便捷功能包 —— 安装（NX工具）"
Write-Host "=============================================="
Write-Host "插件目录 : $bundle"

$langSource = ""
if ($Lang) { $langCode = Resolve-LangCode $Lang; $langSource = "命令行 -Lang=$Lang" }
else {
    $rawLang = $env:UGII_LANG
    if ($rawLang) {
        $langCode = Resolve-LangCode $rawLang
        $langSource = "UGII_LANG=$rawLang"
    } else {
        $ui = [System.Globalization.CultureInfo]::InstalledUICulture.Name.ToLower()
        if ($ui -like 'zh-cn*' -or $ui -like 'zh-hans*') { $langCode = 'zh_cn'; $langSource = "系统 UI 语言 $ui" }
        elseif ($ui -like 'zh-tw*' -or $ui -like 'zh-hk*' -or $ui -like 'zh-hant*') { $langCode = 'zh_tw'; $langSource = "系统 UI 语言 $ui" }
        else { $langCode = 'en'; $langSource = "系统 UI 语言 $ui（非中文，回落英文）" }
    }
}
Write-Host "界面语言 : $(Get-LangName $langCode)（$langSource）"

# 完整性校验
$startupDir = Join-Path $bundle "startup"
$needStartup = @("Htet_menu.men", "Htet_menu.tbr", "step_export.men", "step_export.tbr",
    "Htet_red.bmp", "Htet_black.bmp", "Htet_yellow.bmp", "Htet_random.bmp",
    "Htet_blue.bmp", "Htet_brown.bmp", "Htet_gray.bmp", "Htet_green.bmp",
    "Htet_orange.bmp", "Htet_purple.bmp", "Htet_notransparent.bmp",
    "Htet_Hidealllayers.bmp", "Htet_Showalllayers.bmp",
    "Htet_Nextlayer.bmp", "Htet_Previouslayer.bmp",
    "Htet_Movetolayer255.bmp", "Htet_Openfilefolder.bmp",
    "Htet_displayonly.bmp", "Htet_transparent.bmp", "Htet_layers.bmp",
    "Htet_worklayer.bmp", "Htet_batch.bmp", "Htet_reset.bmp", "Htet_hole.bmp",
    "Htet_Sniptool.bmp", "Htet_Toleranceexe.bmp", "Htet_Calculatoropen.bmp",
    "Htet_Switch200worklayer.bmp")
foreach ($name in $needStartup) {
    if (-not (Test-Path (Join-Path $startupDir $name))) { Fail "缺少文件 startup\$name，请确认整个 $BundleName 文件夹完整" }
}
foreach ($code in @("zh_cn", "zh_tw", "en")) {
    foreach ($name in @("Htet_menu.men", "Htet_menu.tbr", "Htet_tools.rtb", "Htet_export_body01.dlx", "Htet_select_bodies.dlx", "step_export.men", "step_export.tbr", "Format_convert.rtb")) {
        if (-not (Test-Path (Join-Path $bundle "lang\$code\$name"))) { Fail "缺少文件 lang\$code\$name，请确认整个 $BundleName 文件夹完整" }
    }
}
# Htet_*.dlx 不随包放在 application\：安装时由 lang\<code>\ 的母本写入（见下方“对话框资源”段）。
# Quickhole.dlx 不区分语言，母本留在 application\，由下方随 Quickhole.py 一并部署到 startup\。
# 因此这里校验 .py 与 Quickhole.dlx；lang\ 下三语母本已在上面逐语言校验。
foreach ($name in @("Htet_autodwg.py", "Htet_common.py", "Htet_lang.py", "Quickhole.py", "Htet_Sniptool.vb",
        "format_convert.py", "Quickhole.dlx", "QuickholeFeature.py", "QuickholeParametricUI.py",
        "nx_screw_insert_ui.py", "nx_screw_native.py", "nx_screw_insert.dlx",
        "CodexScrewNativeCore.py", "CodexScrewUI.py", "CustomFeatureConfiguration.xml")) {
    if (-not (Test-Path (Join-Path $bundle "application\$name"))) { Fail "缺少文件 application\$name，请确认整个 $BundleName 文件夹完整" }
}
# 螺钉对话框按自身目录读取 screw_icons\*.bmp；菜单 BITMAP 指向 ..\bitmap\nx_screw_insert_ui.bmp。
if (-not (Test-Path -LiteralPath (Join-Path $bundle "application\screw_icons") -PathType Container) -or
    @(Get-ChildItem -LiteralPath (Join-Path $bundle "application\screw_icons") -Filter "*.bmp" -File).Count -eq 0) {
    Fail "缺少文件 application\screw_icons\*.bmp，请确认整个 $BundleName 文件夹完整"
}
if (-not (Test-Path (Join-Path $bundle "bitmap\nx_screw_insert_ui.bmp"))) {
    Fail "缺少文件 bitmap\nx_screw_insert_ui.bmp，请确认整个 $BundleName 文件夹完整"
}
# 公差查询依赖随包程序；缺失时明确报错，不做静默回落。
if (-not (Test-Path (Join-Path $bundle "exefile\GongCha.exe"))) {
    Fail "缺少文件 exefile\GongCha.exe，请确认整个 $BundleName 文件夹完整"
}

# format_convert.py 的母本放在 startup\（历史原因）：先按母本同步到 application\，再参与下面的统一部署。
$fcStartup = Join-Path $startupDir "format_convert.py"
$fcApp = Join-Path (Join-Path $bundle "application") "format_convert.py"
if (Test-Path -LiteralPath $fcStartup) {
    $fcSame = (Test-Path -LiteralPath $fcApp) -and ((Get-FileHash -LiteralPath $fcApp -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $fcStartup -Algorithm SHA256).Hash)
    if (-not $fcSame) {
        if ($PSCmdlet.ShouldProcess($fcApp, "同步 format_convert.py 母本")) { Copy-Item -LiteralPath $fcStartup -Destination $fcApp -Force }
        Write-Host "脚本部署 : application\format_convert.py 已按 startup 母本同步"
    }
}

# NX 解析菜单 ACTIONS 里的 .py 时只查各定制目录的 startup\（不查 application\），
# 因此把 application\ 下的脚本部署一份到 startup\（内容一致则跳过，幂等）
$scriptDeployed = 0
$scriptSame = 0
foreach ($py in (Get-ChildItem -LiteralPath (Join-Path $bundle "application") -Filter "*.py" -File)) {
    $dst = Join-Path $startupDir $py.Name
    if (Test-Path -LiteralPath $dst) {
        if ((Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $py.FullName -Algorithm SHA256).Hash) {
            $scriptSame++
            continue
        }
    }
    if ($PSCmdlet.ShouldProcess($dst, "部署脚本")) { Copy-Item -LiteralPath $py.FullName -Destination $dst -Force; $scriptDeployed++ }
}
Write-Host "脚本部署 : startup 内 .py 就绪（新写 $scriptDeployed 个，已一致 $scriptSame 个）"

# 截图工具是 VB 脚本（菜单 ACTIONS 指向 .vb），同样只从 startup\ 解析，需一并部署。
$vbDeployed = 0
$vbSame = 0
foreach ($vb in (Get-ChildItem -LiteralPath (Join-Path $bundle "application") -Filter "*.vb" -File)) {
    $dst = Join-Path $startupDir $vb.Name
    if (Test-Path -LiteralPath $dst) {
        if ((Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $vb.FullName -Algorithm SHA256).Hash) {
            $vbSame++
            continue
        }
    }
    if ($PSCmdlet.ShouldProcess($dst, "部署脚本")) { Copy-Item -LiteralPath $vb.FullName -Destination $dst -Force; $vbDeployed++ }
}
Write-Host "脚本部署 : startup 内 .vb 就绪（新写 $vbDeployed 个，已一致 $vbSame 个）"

# Quickhole / 螺钉对话框都按脚本所在目录加载同名 .dlx；脚本执行于 startup\，
# 因此把 application\ 下的母本逐个部署到 startup\（内容一致则跳过，幂等）。
foreach ($dlxName in @("Quickhole.dlx", "nx_screw_insert.dlx")) {
    $dlxSrc = Join-Path (Join-Path $bundle "application") $dlxName
    $dlxDst = Join-Path $startupDir $dlxName
    $dlxState = "已一致"
    if (-not (Test-Path -LiteralPath $dlxDst)) {
        if ($PSCmdlet.ShouldProcess($dlxDst, "部署 $dlxName")) {
            Copy-Item -LiteralPath $dlxSrc -Destination $dlxDst -Force
            $dlxState = "新写"
        }
    }
    elseif ((Get-FileHash -LiteralPath $dlxDst -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $dlxSrc -Algorithm SHA256).Hash) {
        if ($PSCmdlet.ShouldProcess($dlxDst, "更新 $dlxName")) {
            Copy-Item -LiteralPath $dlxSrc -Destination $dlxDst -Force
            $dlxState = "已更新"
        }
    }
    Write-Host "对话框部署 : startup\$dlxName $dlxState"
}

# 螺钉对话框从自身目录读取 screw_icons\*.bmp，脚本执行于 startup\，
# 因此把 application\screw_icons\ 下的位图逐个部署到 startup\screw_icons\（幂等）。
$iconSrc = Join-Path (Join-Path $bundle "application") "screw_icons"
$iconDst = Join-Path $startupDir "screw_icons"
if ($PSCmdlet.ShouldProcess($iconDst, "部署 screw_icons")) {
    New-Item -ItemType Directory -Force -Path $iconDst | Out-Null
}
$iconDeployed = 0
$iconSame = 0
foreach ($icon in (Get-ChildItem -LiteralPath $iconSrc -Filter "*.bmp" -File)) {
    $dst = Join-Path $iconDst $icon.Name
    if (Test-Path -LiteralPath $dst) {
        if ((Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $icon.FullName -Algorithm SHA256).Hash) {
            $iconSame++
            continue
        }
    }
    if ($PSCmdlet.ShouldProcess($dst, "部署螺钉图标")) { Copy-Item -LiteralPath $icon.FullName -Destination $dst -Force; $iconDeployed++ }
}
Write-Host "图标部署 : startup\screw_icons 内位图就绪（新写 $iconDeployed 个，已一致 $iconSame 个）"

# 自定义特征（添加螺钉 / 腰型孔参数化）由 NX 从插件根的 application\ 读取，
# 安装目录就是包目录本身，无需搬运；这里显式确认一次，避免手工拷贝时漏掉这个 XML。
$featureCfg = Join-Path (Join-Path $bundle "application") "CustomFeatureConfiguration.xml"
if (-not (Test-Path -LiteralPath $featureCfg)) {
    Fail "缺少文件 application\CustomFeatureConfiguration.xml，请确认整个 $BundleName 文件夹完整"
}
Write-Host "特征配置 : application\CustomFeatureConfiguration.xml 已就位（NX 直接读取插件根）"

# 定位 NX
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
    $NxDir = (Resolve-Path -LiteralPath $NxDir).Path
    if ((Split-Path -Leaf $NxDir).ToLower() -eq "ugii") { $NxDir = Split-Path -Parent $NxDir }
    if (-not (Test-Path (Join-Path $NxDir "UGII"))) { Fail "所选目录下没有 UGII 文件夹: $NxDir" }
} else {
    if (-not $NxDir) { $NxDir = "" }
}

if ($CustomDirs) { $customDirs = [System.IO.Path]::GetFullPath($CustomDirs) }
else {
    $menuDir = Join-Path (Join-Path $NxDir "UGII") "menus"
    if (-not (Test-Path -LiteralPath $menuDir)) { Fail "找不到 $menuDir，请确认 NX 安装完整" }
    $customDirs = Join-Path $menuDir "custom_dirs.dat"
}
$nxShown = if ($NxDir) { $NxDir } else { "（使用测试副本，未涉及 NX 安装目录）" }
Write-Host "NX 目录  : $nxShown"
Write-Host "注册文件 : $customDirs"

$before = Get-FileState $customDirs
Write-Host ("写前状态 : {0} 字节  SHA256 {1}" -f $before.Bytes, $before.Sha)

$EnvStateFile = Join-Path $bundle "install_env_state.json"
$envResult = $null
if ($PSCmdlet.ShouldProcess("UGII_USER_DIR", "设置用户级环境变量指向插件根")) {
    $envResult = Set-UserDirEnv $bundle $EnvStateFile ([bool]$WhatIfPreference)
}
else {
    $envResult = @{ Action = "preview"; Message = "预览：将设置用户级 UGII_USER_DIR=$bundle"; Prev = $null; Current = "" }
}
if ($envResult.Action -eq "skip") { Write-Host ("脚本路径 : " + $envResult.Message) -ForegroundColor Yellow }
else { Write-Host ("脚本路径 : " + $envResult.Message) }
$state = Read-CustomDirsState $customDirs
if ($state.HasLine) {
    if ($state.IsUtf8) { Write-Host "旧版清理 : custom_dirs.dat 里已有本插件行，保持不变（只补不改）" }
    else { Write-Host "旧版清理 : 已有本插件行；文件含非 UTF-8 行，这些行按原字节保留（不转码）" }
}
else { Write-Host "旧版清理 : 未发现本插件行" }

if ($PSCmdlet.ShouldProcess($customDirs, "写入插件目录")) {
    if (Add-PluginLine $customDirs) { Write-Host "注册结果 : 已写入 custom_dirs.dat（只追加本插件一行）" }
    else { Write-Host "注册结果 : 已存在，未重复写入（幂等）" }
} else {
    Write-Host "注册结果 : 未写入（-WhatIf 预览）"
}
# 旧模块化安装（v3.3.5 及更早）的两行：摘掉，目录原地保留不删
$legacyRemoved = @()
if ($PSCmdlet.ShouldProcess($customDirs, "移除旧模块化安装的 Htet工具 / 格式转换 两行（不删目录）")) {
    $legacyRemoved = Remove-LegacyPluginLines $customDirs $LegacyDirs
}
if ($legacyRemoved.Count -gt 0) {
    Write-Host ("旧版清理 : 已移除旧模块化注册行 {0} 条（旧目录保留，可手动移走）" -f $legacyRemoved.Count) -ForegroundColor Yellow
    foreach ($l in $legacyRemoved) { Write-Host ("  - {0}" -f $l) -ForegroundColor Yellow }
} else { Write-Host "旧版清理 : 无旧模块化注册行" }
$after = Get-FileState $customDirs
Write-Host ("写后状态 : {0} 字节  SHA256 {1}" -f $after.Bytes, $after.Sha)

# 其它注册目录：只提示、绝不自动删除（用户明确要求不升级为删除）
$otherDirs = @()
try {
    $rawCd = [System.IO.File]::ReadAllBytes($customDirs)
    $cdText = $null
    foreach ($e in @($Script:StrictUtf8, $Script:StrictGbk, $Script:Latin1)) {
        try { $cdText = $e.GetString($rawCd); break } catch { }
    }
    if ($null -ne $cdText) {
        foreach ($line in ($cdText -split "`r?`n")) {
            $t = $line.Trim()
            if (-not $t -or $t.StartsWith('#')) { continue }
            if ((Normalize-Dir $t).ToLower() -ne (Normalize-Dir $bundle).ToLower()) { $otherDirs += $t }
        }
    }
}
catch { }
if ($otherDirs.Count -gt 0) {
    Write-Host ("其它注册目录 : 检测到 {0} 条其它注册目录，可能同时加载多份菜单互相覆盖（未改动）：" -f $otherDirs.Count) -ForegroundColor Yellow
    foreach ($d in $otherDirs) { Write-Host ("  - {0}" -f $d) -ForegroundColor Yellow }
}
else { Write-Host "其它注册目录 : 无" }

# 界面资源（菜单 / 工具条）
# 母本只读（P0-2）：lang\ 下三语原文恒为 UTF-8 无 BOM，不再就地转码；
# 只把当前界面语言那一份从母本转出、按写盘码页（系统默认 locale 的 ANSI）写入 startup\（P0-1）。
$langDir = Join-Path $bundle "lang\$langCode"
$encInfo = $null
$resInfo = @{}
foreach ($name in @("Htet_menu.men", "Htet_menu.tbr", "step_export.men", "step_export.tbr")) {
    $srcFile = Join-Path $langDir $name
    $dstFile = Join-Path $startupDir $name
    if (-not (Test-Path -LiteralPath $srcFile)) { Fail "缺少界面资源母本：$srcFile" }
    if ($PSCmdlet.ShouldProcess($dstFile, "从母本转码 $langCode 界面资源（CP$(Get-TargetCodePage $langCode)）")) {
        $beforeRes = Get-FileState $dstFile
        try { $encInfo = Write-LangResource $srcFile $dstFile $langCode }
        catch { Fail $_.Exception.Message }
        # 逐文件记录转码结果：两个 .men/.tbr 的回退状态可能不同，不能共用一个变量
        $resInfo[$name] = $encInfo
        if ($encInfo.Fallback) {
            $srcText = [System.IO.File]::ReadAllText($srcFile, (New-Object System.Text.UTF8Encoding($false, $true)))
            $bad = Get-UnmappableChars $srcText $encInfo.CodePage
            Write-Host ("资源 {0}: CP{1} 无法表示 {2} 个字符（{3}），已按字回退写入" -f $name, $encInfo.CodePage, $bad.Count, ($bad -join ' ')) -ForegroundColor Yellow
        }
        $afterRes = Get-FileState $dstFile
        Write-Host ("资源 {0}: {1} 字节 {2} -> {3} 字节 {4}（CP{5}）" -f $name, $beforeRes.Bytes, (Short-Sha $beforeRes.Sha), $afterRes.Bytes, (Short-Sha $afterRes.Sha), $encInfo.CodePage)
    }
}
if ($encInfo) {
    $shown = if ($encInfo.Fallback) { "$($encInfo.Encoding)（含回退字符）" } else { "$($encInfo.Encoding)" }
    Write-Host "界面资源 : 已写入（$shown）"
}

# 功能区选项卡（.rtb，v3.3.2）：NX 2506 只扫描 <自定义目录>\application\profiles\All\ 下的 .rtb
# （美雷 D:\ngll_tool\application\profiles\All\MeiLeiCommon_*.rtb 即此写法，user.mtx 里有对应 ribbontab 登记）；
# 放在 application\ 根下不会被扫描，放在 startup\ 的 .tbr 只生成传统工具条，ribbon 里都不会出现选项卡。
# 母本在 lang\<code>\Htet_tools.rtb（UTF-8），按同一写盘码页转出到 application\profiles\All\。
# 两个模块的组必须写在同一个 rtb 里：实测同 TITLE 的两个 rtb 不会合并，会出现两个同名选项卡。
$rbnName = "NX_tools.rtb"
$rbnSrc = Join-Path $langDir "Htet_tools.rtb"
$rbnDir = Join-Path $bundle "application\profiles\All"
$rbnDst = Join-Path $rbnDir $rbnName
$rbnInfo = $null
if (-not (Test-Path -LiteralPath $rbnSrc)) { Fail "缺少界面资源母本：$rbnSrc" }
if (-not (Test-Path -LiteralPath $rbnDir)) { New-Item -ItemType Directory -Path $rbnDir -Force | Out-Null }
foreach ($stale in @((Join-Path $bundle "application\Htet_tools.rtb"), (Join-Path $rbnDir "Htet_tools.rtb"), (Join-Path $rbnDir "Format_convert.rtb"))) {
    if (Test-Path -LiteralPath $stale) { Remove-Item -LiteralPath $stale -Force; Write-Host ("功能区   : 清理旧文件 {0}" -f $stale.Substring($bundle.Length).TrimStart('\')) }
}
$rbnUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$rbnText = $rbnUtf8.GetString([System.IO.File]::ReadAllBytes($rbnSrc))
$rbnSib = Join-Path $langDir "Format_convert.rtb"
if (Test-Path -LiteralPath $rbnSib) {
    $sibBody = ($rbnUtf8.GetString([System.IO.File]::ReadAllBytes($rbnSib)) -split "`r?`n" | Where-Object { $_ -notmatch '^\s*(TITLE|VERSION)\b' }) -join "`r`n"
    $rbnText = $rbnText.TrimEnd() + "`r`n`r`n" + $sibBody.Trim() + "`r`n"
    Write-Host "功能区   : 合并为单一选项卡（Htet 五组 + 导入导出）"
}
$rbnTmp = Join-Path $env:TEMP ("nx_rbn_" + [guid]::NewGuid().ToString('N') + ".rtb")
[System.IO.File]::WriteAllText($rbnTmp, $rbnText, (New-Object System.Text.UTF8Encoding($false)))
$rbnBefore = Get-FileState $rbnDst
if ($PSCmdlet.ShouldProcess($rbnDst, "生成 $langCode 功能区定义")) {
    try { $rbnInfo = Write-LangResource $rbnTmp $rbnDst $langCode }
    catch { Fail $_.Exception.Message }
    $rbnAfter = Get-FileState $rbnDst
    Write-Host ("功能区   : {0} {1} 字节 {2} -> {3} 字节 {4}（CP{5}）" -f $rbnName, $rbnBefore.Bytes, (Short-Sha $rbnBefore.Sha), $rbnAfter.Bytes, (Short-Sha $rbnAfter.Sha), (Get-TargetCodePage $langCode))
}
Remove-Item -LiteralPath $rbnTmp -Force -ErrorAction SilentlyContinue
if (-not $WhatIfPreference) {
    try {
        $rbnDstText = (Get-StrictEncoding (Get-TargetCodePage $langCode)).GetString([System.IO.File]::ReadAllBytes($rbnDst))
        if ($rbnDstText -eq $rbnText) { Write-Host ("功能区自检: CP{0} 内容一致" -f (Get-TargetCodePage $langCode)) -ForegroundColor Green }
        else { Write-Host "功能区自检: 内容不一致（请重装）" -ForegroundColor Red }
    }
    catch { Write-Host "功能区自检: 无法读回校验（请重装）" -ForegroundColor Yellow }
}

# 安装后自检（P0-4）：按写盘码页（系统默认 locale 的 ANSI）读回 startup\ 里真正生效的文件，逐字比对母本。
# 不一致即以非 0 退出码失败，不再静默"装完了但还是乱码"。
$verifyResults = @()
$verifyOk = $true
$verifyCp = Get-TargetCodePage $langCode
$verifyUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
foreach ($name in @("Htet_menu.men", "Htet_menu.tbr", "step_export.men", "step_export.tbr")) {
    $srcFile = Join-Path $langDir $name
    $dstFile = Join-Path $startupDir $name
    if (-not (Test-Path -LiteralPath $dstFile)) { $verifyResults += "$name=missing"; $verifyOk = $false; continue }
    if ($resInfo.ContainsKey($name) -and $resInfo[$name].Fallback) { $verifyResults += "$name=fallback"; continue }
    if (-not (Test-Path -LiteralPath $srcFile)) { $verifyResults += "$name=nosrc"; $verifyOk = $false; continue }
    try { $srcText = $verifyUtf8.GetString([System.IO.File]::ReadAllBytes($srcFile)) }
    catch { $verifyResults += "$name=src_not_utf8"; $verifyOk = $false; continue }
    try { $dstText = (Get-StrictEncoding $verifyCp).GetString([System.IO.File]::ReadAllBytes($dstFile)) }
    catch { $verifyResults += "$name=decode_fail"; $verifyOk = $false; continue }
    if ($dstText -eq $srcText) { $verifyResults += "$name=ok" }
    else {
        # 指出第一处差异，便于定位是哪个字转坏了
        $pos = 0
        while (($pos -lt $srcText.Length) -and ($pos -lt $dstText.Length) -and ($srcText[$pos] -eq $dstText[$pos])) { $pos++ }
        $verifyResults += ("$name=mismatch@$pos(src='" + $srcText.Substring($pos, [Math]::Min(6, $srcText.Length - $pos)) + "')")
        $verifyOk = $false
    }
}
if ($WhatIfPreference) { $verifyOk = $true; $verifyResults = @("whatif") }
$verifyText = $verifyResults -join ","
if ($verifyOk) { Write-Host "自检通过 : CP$verifyCp  $verifyText" -ForegroundColor Green }
else { Write-Host "自检失败 : CP$verifyCp  $verifyText" -ForegroundColor Red }

# 对话框资源（BlockStyler .dlx：UTF-8 原文，不转码）
# 两份都按界面语言覆盖到 application\；lang\ 缺该文件时保留现有文件并记日志（不报错）
$dlxDir = Join-Path $bundle "application"
$dlxNames = @("Htet_export_body01.dlx", "Htet_select_bodies.dlx")
$dlxDeployed = 0
$dlxMissing = @()
$dlxAfterMain = $null
foreach ($name in $dlxNames) {
    $dlxSrc = Join-Path $langDir $name
    $dlxDst = Join-Path $dlxDir $name
    if (-not (Test-Path -LiteralPath $dlxSrc)) {
        $dlxMissing += $name
        Write-Host ("dlx 语言 : {0} 缺少 {1}，保留现有文件（未改）" -f $langCode, $name) -ForegroundColor Yellow
        continue
    }
    $dlxBak = $dlxDst + $BackupSuffix
    $dlxBefore = Get-FileState $dlxDst
    if ($PSCmdlet.ShouldProcess($dlxDst, "写入 $langCode 对话框资源")) {
        if ((Test-Path -LiteralPath $dlxDst) -and (-not (Test-Path -LiteralPath $dlxBak))) {
            Copy-Item -LiteralPath $dlxDst -Destination $dlxBak -Force
        }
        Copy-Item -LiteralPath $dlxSrc -Destination $dlxDst -Force
        $dlxDeployed++
    }
    $dlxAfter = Get-FileState $dlxDst
    if ($name -eq "Htet_export_body01.dlx") { $dlxAfterMain = $dlxAfter }
    Write-Host ("dlx 语言 : {0}  {1}（{2} 字节 {3} -> {4} 字节 {5}）" -f $langCode, $name, $dlxBefore.Bytes, (Short-Sha $dlxBefore.Sha), $dlxAfter.Bytes, (Short-Sha $dlxAfter.Sha))
}

$installLog = Join-Path $env:TEMP "nx_batch_tools_install.log"
$fallbackNames = @($resInfo.Keys | Where-Object { $resInfo[$_].Fallback })
$localeCpInfo = Get-SystemDefaultLocaleCodePage
$sysCp = Get-SystemAnsiCodePage
$finalCp = Get-TargetCodePage $langCode
$cpReason = if ($localeCpInfo.CodePage -gt 0) { 'system_lcid' } else { 'system_acp_fallback' }
if ($localeCpInfo.CodePage -gt 0 -and $localeCpInfo.CodePage -ne $sysCp) {
    Write-Host ("码页口径 : 系统默认 locale {0} (LCID {1}) → CP{2}；GetACP → CP{3}；两者不一致（系统区域改动后未重启，LCID 未刷新），按系统默认 locale 写盘" -f $localeCpInfo.Name, $localeCpInfo.Lcid, $localeCpInfo.CodePage, $sysCp) -ForegroundColor Yellow
}
$logLine = "{0}  lang={1}  source={2}  bundle={3}  nx={4}  custom_dirs={5}  before={6}B/{7}  after={8}B/{9}  env_action={10}  env_user_dir={11}  codepage={12}  dlx_lang={13}  dlx={14}B/{15}  dlx_deployed={16}  dlx_missing={17}  startup={18}  verify={19}  verify_ok={20}  fallback_files={21}  other_dirs={22}  sys_default_lcid={23}  lcid_codepage={24}  lcid_locale={25}  sys_codepage={26}  cp_reason={27}" -f `
    (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $langCode, $langSource, $bundle, $NxDir, $customDirs,
    $before.Bytes, $before.Sha, $after.Bytes, $after.Sha, $envResult.Action,
    [Environment]::GetEnvironmentVariable("UGII_USER_DIR", "User"),
    $finalCp,
    $langCode,
    $(if ($dlxAfterMain) { $dlxAfterMain.Bytes } else { 0 }),
    $(if ($dlxAfterMain) { $dlxAfterMain.Sha } else { "(none)" }),
    $dlxDeployed,
    $(if ($dlxMissing.Count -gt 0) { $dlxMissing -join "," } else { "-" }),
    $startupDir, $verifyText,
    $(if ($verifyOk) { "true" } else { "false" }),
    $(if ($fallbackNames.Count -gt 0) { $fallbackNames -join "," } else { "-" }),
    $(if ($otherDirs.Count -gt 0) { $otherDirs -join "|" } else { "-" }),
    $localeCpInfo.Lcid, $localeCpInfo.CodePage,
    $(if ($localeCpInfo.Name) { $localeCpInfo.Name } else { "-" }), $sysCp, $cpReason
[System.IO.File]::AppendAllText($installLog, $logLine + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "安装日志 : $installLog"

# 自检失败 -> 非 0 退出（P0-4），避免"装完显示成功、菜单其实还是坏的"
if ((-not $WhatIfPreference) -and (-not $verifyOk)) {
    Fail "安装后自检失败（CP$verifyCp）：$verifyText。已写入的文件见 $startupDir；请把本行日志一并反馈。"
}

Write-Host ""
if ($WhatIfPreference) {
    Write-Host "以上为 -WhatIf 预览，未做任何实际改动。"
}
else {
    Write-Host "安装完成（界面语言：$(Get-LangName $langCode)）。重启 NX 后功能区出现「NX 工具」选项卡（English 界面为 NX tools），菜单栏「NX 工具 → 颜色 / 显示 / 图层 / 出图 / 常用工具」同样可用。"
    Write-Host "若选项卡未出现：右键功能区空白处 → 定制功能区，勾选「NX 工具」，或重启 NX；仍无则确认 application\profiles\All\NX_tools.rtb 存在。"
    Write-Host "（UGII_USER_DIR 仅在 NX 新启动时生效；已打开的 NX 需重启。）"
    Write-Host "卸载：运行 uninstall.bat（或 uninstall.ps1 -NxDir ""$NxDir""），会自动把 UGII_USER_DIR 恢复到安装前的值。"
}
Write-Host ""
