# VBAdminx: 查找已安装软件注册的 COM 类型库。
#
# 只读取注册表 HKCR\TypeLib 的公开登记信息，返回库名称、GUID、版本与文件路径。
# 不注册类型库、不加载、不执行任何组件。
#
# 用法：
#   powershell -File discover-com.ps1 -Keywords "Excel;Word;ADODB" -Bitness 64
# 输出：JSON 数组（写到标准输出；诊断信息写标准错误）。

param(
    [Parameter(Mandatory = $true)][string]$Keywords,
    [Parameter(Mandatory = $true)][ValidateSet('32', '64')][string]$Bitness
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$keywordList = @($Keywords -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($keywordList.Count -eq 0) {
    Write-Output '[]'
    exit 0
}

$preferredPlatform = if ($Bitness -eq '64') { 'win64' } else { 'win32' }
$fallbackPlatform = if ($Bitness -eq '64') { 'win32' } else { 'win64' }
$allowedExtensions = @('.tlb', '.olb', '.dll', '.ocx', '.exe')
$results = New-Object System.Collections.ArrayList
$seen = @{}

function Add-Result {
    param([string]$LibraryName, [string]$Guid, [string]$Version, [string]$Platform, [string]$FilePath)
    $key = ($Guid + '|' + $Version + '|' + $Platform + '|' + $FilePath.ToLowerInvariant())
    if ($seen.ContainsKey($key)) { return }
    $seen[$key] = $true
    [void]$results.Add([ordered]@{
        libraryName = $LibraryName
        guid        = $Guid
        version     = $Version
        platform    = $Platform
        path        = $FilePath
    })
}

function Find-PlatformPaths {
    param($VersionKey)
    # 兼容两种登记层级：{ver}\win32 与 {ver}\{lcid}\win32。
    $found = @()
    foreach ($childName in $VersionKey.GetSubKeyNames()) {
        if ($childName -in @('win32', 'win64', 'win16')) {
            $found += @{ Platform = $childName; Key = $VersionKey.OpenSubKey($childName) }
            continue
        }
        $child = $VersionKey.OpenSubKey($childName)
        if ($null -eq $child) { continue }
        foreach ($grandName in $child.GetSubKeyNames()) {
            if ($grandName -in @('win32', 'win64', 'win16')) {
                $found += @{ Platform = $grandName; Key = $child.OpenSubKey($grandName) }
            }
        }
    }
    return $found
}

try {
    $typeLibRoot = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey('TypeLib')
    if ($null -eq $typeLibRoot) {
        Write-Output '[]'
        exit 0
    }

    foreach ($guidName in $typeLibRoot.GetSubKeyNames()) {
        if ($guidName -notmatch '^\{[0-9A-Fa-f-]{36}\}$') { continue }
        $guidKey = $typeLibRoot.OpenSubKey($guidName)
        if ($null -eq $guidKey) { continue }

        foreach ($versionName in $guidKey.GetSubKeyNames()) {
            $versionKey = $guidKey.OpenSubKey($versionName)
            if ($null -eq $versionKey) { continue }

            $libraryName = [string]$versionKey.GetValue('')
            if ([string]::IsNullOrWhiteSpace($libraryName)) { continue }

            $matched = $false
            foreach ($keyword in $keywordList) {
                if ($libraryName.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $matched = $true
                    break
                }
            }
            if (-not $matched) { continue }

            $paths = @(Find-PlatformPaths -VersionKey $versionKey)
            if ($paths.Count -eq 0) { continue }

            $chosen = $paths | Where-Object { $_.Platform -eq $preferredPlatform } | Select-Object -First 1
            if ($null -eq $chosen) {
                $chosen = $paths | Where-Object { $_.Platform -eq $fallbackPlatform } | Select-Object -First 1
            }
            if ($null -eq $chosen -or $null -eq $chosen.Key) { continue }

            $filePath = [string]$chosen.Key.GetValue('')
            if ([string]::IsNullOrWhiteSpace($filePath)) { continue }
            $filePath = $filePath.Trim('"').Trim()
            if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { continue }
            $extension = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
            if ($allowedExtensions -notcontains $extension) { continue }

            $versionLabel = $versionName
            Add-Result -LibraryName $libraryName -Guid $guidName -Version $versionLabel -Platform $chosen.Platform -FilePath $filePath
        }
    }
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    Write-Output '[]'
    exit 0
}

$sorted = @($results | Sort-Object @{ Expression = { $_.libraryName } }, @{ Expression = { $_.version } })
$json = ConvertTo-Json -InputObject $sorted -Depth 4
if ($sorted.Count -eq 0) { $json = '[]' }
elseif (-not $json.TrimStart().StartsWith('[')) { $json = '[' + $json + ']' }
Write-Output $json
