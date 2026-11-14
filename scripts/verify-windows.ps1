#Requires -Version 5.1
<#
.SYNOPSIS
Verifies the OfficeProductivitySuite build on Windows: prerequisites, static checks, build,
artifacts, core tests, and an optional per-user MSI install/uninstall round trip.
.DESCRIPTION
The script never changes Office security or trust settings and never downloads anything.
Every check reports PASS, FAIL, SKIP, or WARN. The process exits with code 1 when any FAIL
is recorded, so it can be used directly in CI.
.PARAMETER Configuration
The Visual Studio configuration that is built and verified.
.PARAMETER Platform
The MSI architecture (x86 or x64). The managed solution itself stays Any CPU.
.PARAMETER MinimumTests
The minimum number of passing core tests. Defaults to 47.
.PARAMETER SkipStatic
Skips scripts\verify-static.py when it cannot run or is intentionally out of scope.
.PARAMETER SkipBuild
Reuses an existing build instead of invoking scripts\build-windows.ps1.
.PARAMETER SkipTests
Skips executing tests\Suite.Core.Tests\bin\<Configuration>\Suite.Core.Tests.exe.
.PARAMETER InstallMsi
Installs the MSI per-user, verifies files and HKCU add-in registration, and uninstalls again.
.PARAMETER KeepInstalled
Keeps the add-in installed after a successful MSI round trip instead of uninstalling it.
.PARAMETER MsiPath
Overrides the MSI path. Defaults to installer\Suite.Installer\bin\<Configuration>\<Platform>.
.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts\verify-windows.ps1
.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts\verify-windows.ps1 -InstallMsi
.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts\verify-windows.ps1 -SkipBuild -SkipTests
#>
[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64",

    [ValidateRange(0, 10000)]
    [int]$MinimumTests = 47,

    [switch]$SkipStatic,
    [switch]$SkipBuild,
    [switch]$SkipTests,
    [switch]$InstallMsi,
    [switch]$KeepInstalled,

    [string]$MsiPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourceDir = Join-Path $repoRoot "src"
$installerDir = Join-Path $repoRoot "installer\Suite.Installer"
$hostNames = @("Word", "Excel", "PowerPoint")
$checks = New-Object System.Collections.Generic.List[object]
$logLines = New-Object System.Collections.Generic.List[string]

function Add-Check {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$Detail = ""
    )

    $checks.Add([pscustomobject]@{ Name = $Name; Status = $Status; Detail = $Detail })
    $line = "[{0}] {1}" -f $Status, $Name
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        $line = $line + " - " + $Detail
    }
    $logLines.Add($line)
    Write-Host $line
}

function Get-RegistryValue {
    param(
        [Parameter(Mandatory = $true)][string]$Hive,
        [Parameter(Mandatory = $true)][string]$View,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Name = ""
    )

    $hiveKind = [Microsoft.Win32.RegistryHive]::LocalMachine
    if ($Hive -eq "CurrentUser") {
        $hiveKind = [Microsoft.Win32.RegistryHive]::CurrentUser
    }

    $viewKind = [Microsoft.Win32.RegistryView]::Default
    if ($View -eq "Registry32") {
        $viewKind = [Microsoft.Win32.RegistryView]::Registry32
    }
    elseif ($View -eq "Registry64") {
        $viewKind = [Microsoft.Win32.RegistryView]::Registry64
    }

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hiveKind, $viewKind)
        try {
            $key = $base.OpenSubKey($Path, $false)
            if (-not $key) {
                return $null
            }
            try {
                return $key.GetValue($Name, $null)
            }
            finally {
                $key.Dispose()
            }
        }
        finally {
            $base.Dispose()
        }
    }
    catch {
        return $null
    }
}

function Test-RegistryKey {
    param(
        [Parameter(Mandatory = $true)][string]$Hive,
        [Parameter(Mandatory = $true)][string]$View,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $hiveKind = [Microsoft.Win32.RegistryHive]::LocalMachine
    if ($Hive -eq "CurrentUser") {
        $hiveKind = [Microsoft.Win32.RegistryHive]::CurrentUser
    }

    $viewKind = [Microsoft.Win32.RegistryView]::Default
    if ($View -eq "Registry32") {
        $viewKind = [Microsoft.Win32.RegistryView]::Registry32
    }
    elseif ($View -eq "Registry64") {
        $viewKind = [Microsoft.Win32.RegistryView]::Registry64
    }

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hiveKind, $viewKind)
        try {
            $key = $base.OpenSubKey($Path, $false)
            if (-not $key) {
                return $false
            }
            $key.Dispose()
            return $true
        }
        finally {
            $base.Dispose()
        }
    }
    catch {
        return $false
    }
}

function Get-CommandPath {
    param([Parameter(Mandatory = $true)][string]$Name)

    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $command) {
        return $null
    }
    return $command.Source
}

function Get-AppPath {
    param([Parameter(Mandatory = $true)][string]$ExecutableName)

    foreach ($view in @("Registry64", "Registry32")) {
        $path = Get-RegistryValue -Hive "LocalMachine" -View $view -Path ("SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + $ExecutableName) -Name ""
        if (-not [string]::IsNullOrWhiteSpace([string]$path)) {
            return [string]$path
        }
    }
    return $null
}

function Get-OfficePlatform {
    foreach ($view in @("Registry64", "Registry32")) {
        $value = Get-RegistryValue -Hive "LocalMachine" -View $view -Path "SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "Platform"
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    $office64 = Get-RegistryValue -Hive "LocalMachine" -View "Registry64" -Path "SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot" -Name "Path"
    $office32 = Get-RegistryValue -Hive "LocalMachine" -View "Registry32" -Path "SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot" -Name "Path"
    if (-not [string]::IsNullOrWhiteSpace([string]$office64)) {
        return "x64"
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$office32)) {
        return "x86"
    }
    return $null
}

function Get-OfficeVersion {
    foreach ($view in @("Registry64", "Registry32")) {
        $value = Get-RegistryValue -Hive "LocalMachine" -View $view -Path "SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "VersionToReport"
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    foreach ($version in @("16.0", "15.0", "14.0")) {
        $root = Get-RegistryValue -Hive "LocalMachine" -View "Registry64" -Path ("SOFTWARE\Microsoft\Office\" + $version + "\Word\InstallRoot") -Name "Path"
        if ([string]::IsNullOrWhiteSpace([string]$root)) {
            $root = Get-RegistryValue -Hive "LocalMachine" -View "Registry32" -Path ("SOFTWARE\Microsoft\Office\" + $version + "\Word\InstallRoot") -Name "Path"
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$root)) {
            return $version
        }
    }
    return $null
}

function Find-MSBuild {
    $programFilesX86 = ${env:ProgramFiles(x86)}
    if (-not $programFilesX86) {
        return $null
    }

    $vswhere = Join-Path $programFilesX86 "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere -PathType Leaf)) {
        $command = Get-CommandPath "vswhere.exe"
        if ($command) {
            $vswhere = $command
        }
    }
    if (-not (Test-Path $vswhere -PathType Leaf)) {
        return $null
    }

    $arguments = @("-version", "[17.0,18.0)", "-latest", "-products", "*", "-requires", "Microsoft.Component.MSBuild", "-property", "installationPath")
    $installPath = & $vswhere @arguments | Select-Object -First 1
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace([string]$installPath)) {
        return $null
    }

    $installPath = ([string]$installPath).Trim()
    $msbuild = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
    if (-not (Test-Path $msbuild -PathType Leaf)) {
        return $null
    }

    return [pscustomobject]@{ InstallPath = $installPath; MSBuild = $msbuild }
}

function Find-VstoTargets {
    param([Parameter(Mandatory = $true)][string]$InstallPath)

    $direct = Join-Path $InstallPath "MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools\Microsoft.VisualStudio.OfficeTools.targets"
    if (Test-Path $direct -PathType Leaf) {
        return $direct
    }

    $root = Join-Path $InstallPath "MSBuild\Microsoft\VisualStudio"
    if (-not (Test-Path $root -PathType Container)) {
        return $null
    }
    $candidate = Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue |
        Sort-Object -Property Name -Descending |
        ForEach-Object {
            $path = Join-Path $_.FullName "OfficeTools\Microsoft.VisualStudio.OfficeTools.targets"
            if (Test-Path $path -PathType Leaf) { $path }
        } |
        Select-Object -First 1
    return $candidate
}

function Find-WiXBin {
    $candidates = @()
    if ($env:WIX) {
        $candidates += (Join-Path $env:WIX "bin")
        $candidates += $env:WIX
    }

    foreach ($candidate in $candidates) {
        if ((Test-Path (Join-Path $candidate "candle.exe") -PathType Leaf) -and
            (Test-Path (Join-Path $candidate "light.exe") -PathType Leaf)) {
            return (Resolve-Path $candidate).Path
        }
    }

    $command = Get-CommandPath "candle.exe"
    if ($command) {
        $candidate = Split-Path -Parent $command
        if (Test-Path (Join-Path $candidate "light.exe") -PathType Leaf) {
            return $candidate
        }
    }

    $programFilesX86 = ${env:ProgramFiles(x86)}
    if ($programFilesX86) {
        $roots = @(Get-ChildItem -Path $programFilesX86 -Directory -Filter "WiX Toolset v3.*" -ErrorAction SilentlyContinue |
            Sort-Object -Property Name -Descending)
        foreach ($root in $roots) {
            $candidate = Join-Path $root.FullName "bin"
            if ((Test-Path (Join-Path $candidate "candle.exe") -PathType Leaf) -and
                (Test-Path (Join-Path $candidate "light.exe") -PathType Leaf)) {
                return $candidate
            }
        }
    }

    return $null
}

function Wait-Until {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Condition,
        [int]$TimeoutSeconds = 90
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (& $Condition) {
            return $true
        }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

function Invoke-MsiExec {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
    return $process.ExitCode
}

function Get-AddinKeyPath {
    param([Parameter(Mandatory = $true)][string]$HostName)

    return "Software\Microsoft\Office\" + $HostName + "\Addins\Suite." + $HostName
}

function Test-RibbonResource {
    param(
        [Parameter(Mandatory = $true)][string]$AssemblyPath,
        [Parameter(Mandatory = $true)][string]$ExpectedName
    )

    try {
        $assembly = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($AssemblyPath)
        $names = $assembly.GetManifestResourceNames()
        if ($names -notcontains $ExpectedName) {
            return "缺少嵌入资源 " + $ExpectedName
        }
        return $null
    }
    catch {
        return "无法读取程序集资源：" + $_.Exception.Message
    }
}

function Get-Summary {
    $passed = @($checks | Where-Object { $_.Status -eq "PASS" }).Count
    $failed = @($checks | Where-Object { $_.Status -eq "FAIL" }).Count
    $skipped = @($checks | Where-Object { $_.Status -eq "SKIP" }).Count
    $warned = @($checks | Where-Object { $_.Status -eq "WARN" }).Count
    return [pscustomobject]@{ Passed = $passed; Failed = $failed; Skipped = $skipped; Warned = $warned }
}

$artifactsDir = Join-Path $repoRoot "artifacts"
$logPath = Join-Path $artifactsDir ("verify-windows-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
$runtime = $null

try {
    Write-Host "Office Productivity Suite verification"
    Write-Host "Repository: $repoRoot"
    Write-Host "Configuration: $Configuration; MSI platform: $Platform"
    Write-Host ""

    if ($env:OS -ne "Windows_NT") {
        Add-Check "Windows 环境" "FAIL" "当前系统不是 Windows，验证脚本只能在 Windows 上运行。"
        return
    }
    $osCaption = $null
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($os) {
        $osCaption = [string]$os.Caption
    }
    Add-Check "Windows 环境" "PASS" $osCaption

    $is64BitOs = [Environment]::Is64BitOperatingSystem
    Add-Check "64 位操作系统" ($(if ($is64BitOs) { "PASS" } else { "FAIL" })) ("当前进程位宽：" + [Environment]::Is64BitProcess)

    $psVersion = $PSVersionTable.PSVersion
    Add-Check "PowerShell 版本" ($(if ($psVersion.Major -ge 5) { "PASS" } else { "FAIL" })) ([string]$psVersion)

    $netRelease = Get-RegistryValue -Hive "LocalMachine" -View "Registry64" -Path "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
    if ($null -eq $netRelease) {
        $netRelease = Get-RegistryValue -Hive "LocalMachine" -View "Registry32" -Path "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
    }
    if ($null -eq $netRelease) {
        Add-Check ".NET Framework 4.8" "FAIL" "未找到 v4 Full 的 Release 值。"
    }
    elseif ([int]$netRelease -ge 528040) {
        Add-Check ".NET Framework 4.8" "PASS" ("Release=" + $netRelease)
    }
    else {
        Add-Check ".NET Framework 4.8" "FAIL" ("Release=" + $netRelease + "，需要 528040 或更高。")
    }

    $officePlatform = Get-OfficePlatform
    $officeVersion = Get-OfficeVersion
    if ($officePlatform) {
        Add-Check "Office 位数" "PASS" ([string]$officePlatform)
    }
    else {
        Add-Check "Office 位数" "WARN" "未检测到 ClickToRun/InstallRoot 信息。"
    }
    if (-not $officeVersion) {
        Add-Check "Office 版本" "WARN" "未检测到 Office 版本注册信息。"
    }
    elseif ($officeVersion -like "16.*") {
        Add-Check "Office 版本" "PASS" ([string]$officeVersion)
    }
    else {
        Add-Check "Office 版本" "WARN" ("检测到 " + $officeVersion + "，产品目标为 Office 2016 及以上（16.0）。")
    }

    $vsto64 = Get-RegistryValue -Hive "LocalMachine" -View "Registry64" -Path "SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" -Name "Version"
    $vsto32 = Get-RegistryValue -Hive "LocalMachine" -View "Registry32" -Path "SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" -Name "Version"
    $vstoDetail = "64 位：" + $(if ($vsto64) { [string]$vsto64 } else { "未安装" }) + "；32 位：" + $(if ($vsto32) { [string]$vsto32 } else { "未安装" })
    if ($officePlatform -eq "x64" -and $vsto64) {
        Add-Check "VSTO Runtime" "PASS" $vstoDetail
    }
    elseif ($officePlatform -eq "x86" -and $vsto32) {
        Add-Check "VSTO Runtime" "PASS" $vstoDetail
    }
    elseif ($vsto64 -or $vsto32) {
        Add-Check "VSTO Runtime" "WARN" ($vstoDetail + "；请确认与 Office 位数一致。")
    }
    else {
        Add-Check "VSTO Runtime" "FAIL" $vstoDetail
    }

    $piaRoots = @(
        (Join-Path $env:WINDIR "Microsoft.NET\assembly\GAC_MSIL"),
        (Join-Path $env:WINDIR "assembly\GAC_MSIL")
    )
    foreach ($hostName in $hostNames) {
        $found = $false
        foreach ($root in $piaRoots) {
            if (Test-Path (Join-Path $root ("Microsoft.Office.Interop." + $hostName)) -PathType Container) {
                $found = $true
                break
            }
        }
        if ($found) {
            Add-Check ("Office PIA - " + $hostName) "PASS" "已安装互操作程序集。"
        }
        elseif ($officePlatform) {
            Add-Check ("Office PIA - " + $hostName) "FAIL" "GAC 中缺少 Microsoft.Office.Interop.$hostName，MSBuild 无法解析引用。"
        }
        else {
            Add-Check ("Office PIA - " + $hostName) "SKIP" "未检测到 Office，跳过 PIA 检查。"
        }
    }

    foreach ($executable in @(@("Word", "WINWORD.EXE"), @("Excel", "EXCEL.EXE"), @("PowerPoint", "POWERPNT.EXE"))) {
        $path = Get-AppPath -ExecutableName $executable[1]
        if ($path) {
            Add-Check ("Office 应用 - " + $executable[0]) "PASS" $path
        }
        else {
            Add-Check ("Office 应用 - " + $executable[0]) "WARN" "未找到 " + $executable[1] + "，无法进行 UI 级加载验证。"
        }
    }

    $runtime = Find-MSBuild
    if ($runtime) {
        Add-Check "Visual Studio 2022 与 MSBuild" "PASS" $runtime.MSBuild
        $vstoTargets = Find-VstoTargets -InstallPath $runtime.InstallPath
        if ($vstoTargets) {
            Add-Check "VSTO 构建目标" "PASS" $vstoTargets
        }
        else {
            Add-Check "VSTO 构建目标" "FAIL" "缺少 OfficeTools targets，请安装 Visual Studio 的 Office/SharePoint 开发工作负载。"
        }
    }
    else {
        Add-Check "Visual Studio 2022 与 MSBuild" "FAIL" "未找到 MSBuild。请安装 Visual Studio 2022 与 Office/SharePoint 开发工作负载。"
    }

    $wixBin = Find-WiXBin
    if ($wixBin) {
        Add-Check "WiX Toolset 3.x" "PASS" $wixBin
    }
    else {
        Add-Check "WiX Toolset 3.x" "FAIL" "未找到 candle.exe 与 light.exe。"
    }

    if (-not $SkipStatic) {
        $python = Get-CommandPath "python"
        if (-not $python) {
            $python = Get-CommandPath "py"
        }
        if (-not $python) {
            Add-Check "静态检查" "SKIP" "未找到 Python，无法运行 scripts\verify-static.py。"
        }
        else {
            $staticScript = Join-Path $PSScriptRoot "verify-static.py"
            & $python $staticScript
            $staticExit = $LASTEXITCODE
            if ($staticExit -eq 0) {
                Add-Check "静态检查" "PASS" "verify-static.py 通过。"
            }
            else {
                Add-Check "静态检查" "FAIL" ("verify-static.py 退出码 " + $staticExit)
            }
        }
    }
    else {
        Add-Check "静态检查" "SKIP" "已通过 -SkipStatic 跳过。"
    }

    if (-not $SkipBuild) {
        if (-not $runtime -or -not $wixBin) {
            Add-Check "解决方案构建" "SKIP" "缺少 MSBuild 或 WiX，跳过构建。"
        }
        else {
            $buildScript = Join-Path $PSScriptRoot "build-windows.ps1"
            try {
                & $buildScript -Configuration $Configuration -Platform $Platform
                if ($LASTEXITCODE -eq 0) {
                    Add-Check "解决方案构建与 MSI" "PASS" "build-windows.ps1 完成。"
                }
                else {
                    Add-Check "解决方案构建与 MSI" "FAIL" ("build-windows.ps1 退出码 " + $LASTEXITCODE)
                }
            }
            catch {
                Add-Check "解决方案构建与 MSI" "FAIL" $_.Exception.Message
            }
        }
    }
    else {
        Add-Check "解决方案构建" "SKIP" "已通过 -SkipBuild 跳过。"
    }

    $requiredOutputs = @()
    foreach ($project in @("Suite.Core", "Suite.Infrastructure", "Suite.Windowing")) {
        $requiredOutputs += Join-Path $sourceDir (Join-Path $project ("bin\" + $Configuration + "\" + $project + ".dll"))
    }
    foreach ($hostName in $hostNames) {
        foreach ($extension in @("dll", "vsto", "dll.manifest")) {
            $requiredOutputs += Join-Path $sourceDir (Join-Path ("Suite." + $hostName) ("bin\" + $Configuration + "\Suite." + $hostName + "." + $extension))
        }
    }

    $missingOutputs = @()
    $emptyOutputs = @()
    foreach ($output in $requiredOutputs) {
        if (-not (Test-Path $output -PathType Leaf)) {
            $missingOutputs += $output
        }
        elseif ((Get-Item $output).Length -le 0) {
            $emptyOutputs += $output
        }
    }
    if ($missingOutputs.Count -gt 0) {
        Add-Check "构建产物" "FAIL" ("缺少 " + $missingOutputs.Count + " 个文件，首个：" + $missingOutputs[0])
    }
    elseif ($emptyOutputs.Count -gt 0) {
        Add-Check "构建产物" "FAIL" ("存在 0 字节文件，首个：" + $emptyOutputs[0])
    }
    else {
        Add-Check "构建产物" "PASS" ($requiredOutputs.Count.ToString() + " 个 DLL、VSTO 与清单文件齐全且非空。")
    }

    $manifestErrors = @()
    foreach ($hostName in $hostNames) {
        foreach ($extension in @("vsto", "dll.manifest")) {
            $path = Join-Path $sourceDir (Join-Path ("Suite." + $hostName) ("bin\" + $Configuration + "\Suite." + $hostName + "." + $extension))
            if (-not (Test-Path $path -PathType Leaf)) {
                continue
            }
            try {
                [xml](Get-Content -Path $path -Raw) | Out-Null
            }
            catch {
                $manifestErrors += ($path + "：" + $_.Exception.Message)
            }
        }
    }
    if ($manifestErrors.Count -eq 0) {
        Add-Check "VSTO 清单 XML" "PASS" "6 个 .vsto/.dll.manifest 文件均可解析。"
    }
    else {
        Add-Check "VSTO 清单 XML" "FAIL" ($manifestErrors -join "；")
    }

    $ribbonResources = @{
        "Suite.Word" = "Suite.Word.WordRibbon.xml"
        "Suite.Excel" = "Suite.Excel.Ribbon.SuiteRibbon.xml"
        "Suite.PowerPoint" = "Suite.PowerPoint.Ribbon.xml"
    }
    if ($PSVersionTable.PSEdition -eq "Core") {
        Add-Check "Ribbon 资源" "SKIP" "PowerShell 7 无法反射加载 .NET Framework VSTO 程序集，请改用 Windows PowerShell 5.1 复核。"
    }
    else {
        foreach ($assemblyName in $ribbonResources.Keys) {
            $assemblyPath = Join-Path $sourceDir (Join-Path $assemblyName ("bin\" + $Configuration + "\" + $assemblyName + ".dll"))
            if (-not (Test-Path $assemblyPath -PathType Leaf)) {
                Add-Check ("Ribbon 资源 - " + $assemblyName) "SKIP" "程序集未生成。"
                continue
            }
            $resourceError = Test-RibbonResource -AssemblyPath $assemblyPath -ExpectedName $ribbonResources[$assemblyName]
            if (-not $resourceError) {
                Add-Check ("Ribbon 资源 - " + $assemblyName) "PASS" $ribbonResources[$assemblyName]
            }
            else {
                Add-Check ("Ribbon 资源 - " + $assemblyName) "FAIL" $resourceError
            }
        }
    }

    $testExe = Join-Path $repoRoot (Join-Path "tests\Suite.Core.Tests" ("bin\" + $Configuration + "\Suite.Core.Tests.exe"))
    if ($SkipTests) {
        Add-Check "核心测试" "SKIP" "已通过 -SkipTests 跳过。"
    }
    elseif (-not (Test-Path $testExe -PathType Leaf)) {
        Add-Check "核心测试" "FAIL" ("测试程序不存在：" + $testExe)
    }
    else {
        $rawOutput = & $testExe 2>&1
        $testExit = $LASTEXITCODE
        $outputLines = @($rawOutput | ForEach-Object { [string]$_ })
        $passCount = @($outputLines | Where-Object { $_ -match '^\[PASS\]' }).Count
        $failCount = @($outputLines | Where-Object { $_ -match '^\[FAIL\]' }).Count
        $logLines.Add("----- Suite.Core.Tests output -----")
        foreach ($outputLine in $outputLines) {
            $logLines.Add($outputLine)
        }
        $logLines.Add("----- end of test output -----")
        $testDetail = "通过 " + $passCount + "，失败 " + $failCount + "，退出码 " + $testExit
        if ($testExit -eq 0 -and $failCount -eq 0 -and $passCount -ge $MinimumTests) {
            Add-Check "核心测试" "PASS" $testDetail
        }
        elseif ($testExit -eq 0 -and $passCount -lt $MinimumTests) {
            Add-Check "核心测试" "FAIL" ($testDetail + "，低于最低要求 " + $MinimumTests + " 项。")
        }
        else {
            Add-Check "核心测试" "FAIL" $testDetail
        }
    }

    if ($InstallMsi) {
        $resolvedMsi = $MsiPath
        if ([string]::IsNullOrWhiteSpace($resolvedMsi)) {
            $resolvedMsi = Join-Path $installerDir ("bin\" + $Configuration + "\" + $Platform + "\OfficeProductivitySuite-" + $Configuration + "-" + $Platform + ".msi")
        }
        if (-not (Test-Path $resolvedMsi -PathType Leaf)) {
            Add-Check "MSI 安装验证" "SKIP" ("未找到 MSI：" + $resolvedMsi)
        }
        else {
            $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $identity
            if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Add-Check "MSI 运行身份" "WARN" "当前为管理员会话；按用户安装将写入该账户的 LOCALAPPDATA。"
            }

            $installDir = Join-Path $env:LOCALAPPDATA "OfficeProductivitySuite"
            $expectedFiles = @("Suite.Core.dll", "Suite.Infrastructure.dll", "Suite.Windowing.dll")
            foreach ($hostName in $hostNames) {
                $expectedFiles += ("Suite." + $hostName + ".dll")
                $expectedFiles += ("Suite." + $hostName + ".vsto")
                $expectedFiles += ("Suite." + $hostName + ".dll.manifest")
            }

            $preexisting = Test-RegistryKey -Hive "CurrentUser" -View "Default" -Path (Get-AddinKeyPath -HostName "Word")
            if ($preexisting) {
                Add-Check "MSI 预检查" "WARN" "检测到已注册的加载项；本轮将覆盖安装并在结束时卸载。"
            }

            $installExit = Invoke-MsiExec -Arguments @("/i", ('"' + $resolvedMsi + '"'), "/qn", "/norestart")
            if ($installExit -ne 0 -and $installExit -ne 3010) {
                Add-Check "MSI 安装" "FAIL" ("msiexec 退出码 " + $installExit)
            }
            else {
                $installed = Wait-Until -Condition { Test-Path (Join-Path $installDir "Suite.Core.dll") -PathType Leaf } -TimeoutSeconds 90
                if ($installed) {
                    Add-Check "MSI 安装" "PASS" ("退出码 " + $installExit + "，安装目录 " + $installDir)
                }
                else {
                    Add-Check "MSI 安装" "FAIL" "安装完成后未找到 Suite.Core.dll。"
                }

                $missingInstalled = @()
                foreach ($file in $expectedFiles) {
                    $path = Join-Path $installDir $file
                    if (-not (Test-Path $path -PathType Leaf)) {
                        $missingInstalled += $file
                    }
                    elseif ((Get-Item $path).Length -le 0) {
                        $missingInstalled += ($file + " (0 字节)")
                    }
                }
                if ($missingInstalled.Count -eq 0) {
                    Add-Check "已安装文件" "PASS" ($expectedFiles.Count.ToString() + " 个文件存在且非空。")
                }
                else {
                    Add-Check "已安装文件" "FAIL" ("缺少或为空：" + ($missingInstalled -join "，"))
                }

                foreach ($hostName in $hostNames) {
                    $keyPath = Get-AddinKeyPath -HostName $hostName
                    $loadBehavior = Get-RegistryValue -Hive "CurrentUser" -View "Default" -Path $keyPath -Name "LoadBehavior"
                    $manifest = Get-RegistryValue -Hive "CurrentUser" -View "Default" -Path $keyPath -Name "Manifest"
                    $expectedManifest = "file:///" + (Join-Path $installDir ("Suite." + $hostName + ".vsto"))
                    if ($null -eq $loadBehavior -or $null -eq $manifest) {
                        Add-Check ("加载项注册 - " + $hostName) "FAIL" ("缺少 HKCU 注册：" + $keyPath)
                    }
                    elseif ([int]$loadBehavior -ne 3) {
                        Add-Check ("加载项注册 - " + $hostName) "FAIL" ("LoadBehavior=" + $loadBehavior + "，应为 3。")
                    }
                    elseif ([string]$manifest -ne $expectedManifest) {
                        Add-Check ("加载项注册 - " + $hostName) "FAIL" ("Manifest=" + $manifest + "，应为 " + $expectedManifest)
                    }
                    else {
                        Add-Check ("加载项注册 - " + $hostName) "PASS" ("LoadBehavior=3，Manifest=" + $manifest)
                    }
                }

                if ($KeepInstalled) {
                    Add-Check "MSI 卸载" "SKIP" "已通过 -KeepInstalled 保留安装。"
                }
                else {
                    $uninstallExit = Invoke-MsiExec -Arguments @("/x", ('"' + $resolvedMsi + '"'), "/qn", "/norestart")
                    $removed = Wait-Until -Condition { -not (Test-Path (Join-Path $installDir "Suite.Core.dll") -PathType Leaf) } -TimeoutSeconds 90
                    if (($uninstallExit -eq 0 -or $uninstallExit -eq 3010) -and $removed) {
                        Add-Check "MSI 卸载" "PASS" ("退出码 " + $uninstallExit + "，已移除安装文件。")
                    }
                    elseif ($uninstallExit -ne 0 -and $uninstallExit -ne 3010) {
                        Add-Check "MSI 卸载" "FAIL" ("msiexec 退出码 " + $uninstallExit)
                    }
                    else {
                        Add-Check "MSI 卸载" "FAIL" "卸载完成后仍存在 Suite.Core.dll。"
                    }

                    $remainingKeys = @()
                    foreach ($hostName in $hostNames) {
                        if (Test-RegistryKey -Hive "CurrentUser" -View "Default" -Path (Get-AddinKeyPath -HostName $hostName)) {
                            $remainingKeys += $hostName
                        }
                    }
                    if ($remainingKeys.Count -eq 0) {
                        Add-Check "卸载后注册清理" "PASS" "HKCU 加载项键已移除。"
                    }
                    else {
                        Add-Check "卸载后注册清理" "FAIL" ("仍存在：" + ($remainingKeys -join "，"))
                    }
                }
            }
        }
    }
    else {
        Add-Check "MSI 安装验证" "SKIP" "未指定 -InstallMsi，跳过安装/卸载往返验证。"
    }

    if ($officePlatform -and $psVersion.Major -ge 5) {
        Add-Check "Office UI 加载验证" "SKIP" "需人工在 Word/Excel/PowerPoint 中确认 Ribbon 按钮、撤销、备份与拆分版式。"
    }
}
catch {
    Add-Check "未处理的验证错误" "FAIL" $_.Exception.Message
}
finally {
    $summary = Get-Summary
    Write-Host ""
    Write-Host "===== 验证结果 ====="
    $checks | Format-Table -AutoSize | Out-String | Write-Host
    Write-Host ("通过 " + $summary.Passed + "，失败 " + $summary.Failed + "，跳过 " + $summary.Skipped + "，警告 " + $summary.Warned)

    try {
        if (-not (Test-Path $artifactsDir -PathType Container)) {
            New-Item -ItemType Directory -Path $artifactsDir -Force | Out-Null
        }
        $logLines | Set-Content -Path $logPath -Encoding UTF8
        Write-Host ("日志：" + $logPath)
    }
    catch {
        Write-Warning ("无法写入日志 " + $logPath + "：" + $_.Exception.Message)
    }

    if ($summary.Failed -gt 0) {
        exit 1
    }
    exit 0
}
