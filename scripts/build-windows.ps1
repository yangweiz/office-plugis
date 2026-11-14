#Requires -Version 5.1
<#
.SYNOPSIS
Builds the OfficeProductivitySuite solution and its WiX 3.x MSI on Windows.
.PARAMETER Configuration
The Visual Studio configuration to build.
.PARAMETER Platform
The MSI and VSTO registry architecture. The managed solution remains Any CPU.
.PARAMETER SignPath
Optional path to a PFX/P12 Authenticode certificate. If it has a password, provide
it through the SUITE_SIGNING_PASSWORD environment variable. No signature or
signature verification is attempted when this parameter is omitted.
#>
[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64",

    [string]$SignPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    Write-Host "Running: $FilePath"
    & $FilePath @Arguments
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "$FilePath failed with exit code $exitCode."
    }
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

    $candleCommand = Get-Command "candle.exe" -ErrorAction SilentlyContinue
    if ($candleCommand) {
        $candidate = Split-Path -Parent $candleCommand.Source
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

function Find-SignTool {
    $command = Get-Command "signtool.exe" -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $programFilesX86 = ${env:ProgramFiles(x86)}
    if ($programFilesX86) {
        $sdkBin = Join-Path $programFilesX86 "Windows Kits\10\bin"
        if (Test-Path $sdkBin -PathType Container) {
            $tool = Get-ChildItem -Path $sdkBin -Filter "signtool.exe" -File -Recurse -ErrorAction SilentlyContinue |
                Sort-Object -Property FullName -Descending |
                Select-Object -First 1
            if ($tool) {
                return $tool.FullName
            }
        }
    }

    return $null
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$solution = Join-Path $repoRoot "OfficeProductivitySuite.sln"
$installerDir = Join-Path $repoRoot "installer\Suite.Installer"
$productWxs = Join-Path $installerDir "Product.wxs"
$sourceDir = Join-Path $repoRoot "src"

$programFilesX86 = ${env:ProgramFiles(x86)}
if (-not $programFilesX86) {
    throw "A 64-bit Windows installation with Visual Studio 2022 is required."
}

$vswhere = Join-Path $programFilesX86 "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere -PathType Leaf)) {
    $vswhereCommand = Get-Command "vswhere.exe" -ErrorAction SilentlyContinue
    if ($vswhereCommand) {
        $vswhere = $vswhereCommand.Source
    }
}
if (-not (Test-Path $vswhere -PathType Leaf)) {
    throw "vswhere.exe was not found. Install Visual Studio 2022 with the MSBuild and VSTO development workloads."
}

$vswhereArguments = @(
    "-version", "[17.0,18.0)",
    "-latest",
    "-products", "*",
    "-requires", "Microsoft.Component.MSBuild",
    "-property", "installationPath"
)
$vsInstall = & $vswhere @vswhereArguments | Select-Object -First 1
if ($LASTEXITCODE -ne 0) {
    throw "vswhere.exe failed with exit code $LASTEXITCODE."
}
if ([string]::IsNullOrWhiteSpace([string]$vsInstall)) {
    throw "Visual Studio 2022 with MSBuild was not found."
}
$vsInstall = ([string]$vsInstall).Trim()
$msbuild = Join-Path $vsInstall "MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild -PathType Leaf)) {
    throw "MSBuild was not found under the selected Visual Studio 2022 installation: $vsInstall"
}

$wixBin = Find-WiXBin
if (-not $wixBin) {
    throw "WiX Toolset 3.x (candle.exe and light.exe) was not found. Install WiX 3.x before building the MSI."
}
$candle = Join-Path $wixBin "candle.exe"
$light = Join-Path $wixBin "light.exe"

$signTool = $null
$certificatePath = $null
if (-not [string]::IsNullOrWhiteSpace($SignPath)) {
    if (-not (Test-Path $SignPath -PathType Leaf)) {
        throw "The Authenticode certificate specified by SignPath does not exist: $SignPath"
    }
    $certificatePath = (Resolve-Path $SignPath).Path
    $signTool = Find-SignTool
    if (-not $signTool) {
        throw "SignPath was supplied, but signtool.exe was not found. Install a Windows 10/11 SDK signing tool."
    }
}

Write-Host "Building with Visual Studio: $vsInstall"
Write-Host "MSI architecture: $Platform; solution platform: Any CPU"
Write-Host "Installation prerequisites: .NET Framework 4.8 and the VSTO Runtime matching Office bitness. Office PIA and VSTO Runtime files are not bundled."

$restoreArguments = @(
    $solution,
    "/t:Restore",
    "/p:RestorePackagesConfig=true",
    "/p:Configuration=$Configuration",
    "/p:Platform=Any CPU",
    "/nologo",
    "/verbosity:minimal"
)
Invoke-Native -FilePath $msbuild -Arguments $restoreArguments

$buildArguments = @(
    $solution,
    "/t:Build",
    "/m",
    "/p:Configuration=$Configuration",
    "/p:Platform=Any CPU",
    "/nologo",
    "/verbosity:minimal"
)
$importedCertificate = $null
$certificateStore = $null
$certificateWasAdded = $false
try {
if ($certificatePath) {
    $securePassword = New-Object System.Security.SecureString
    if ($env:SUITE_SIGNING_PASSWORD) {
        $securePassword = ConvertTo-SecureString $env:SUITE_SIGNING_PASSWORD -AsPlainText -Force
    }
    $importedCertificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $importedCertificate.Import($certificatePath, $securePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)
    if (-not $importedCertificate.HasPrivateKey) {
        throw "The signing certificate does not contain a private key: $certificatePath"
    }
    $certificateStore = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Store -ArgumentList @("My", "CurrentUser")
    $certificateStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $existingCertificates = $certificateStore.Certificates.Find([System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $importedCertificate.Thumbprint, $false)
    $certificateWasAdded = $existingCertificates.Count -eq 0
    if ($certificateWasAdded) {
        $certificateStore.Add($importedCertificate)
    }
    $certificateStore.Close()
    $buildArguments += @(
        "/p:SignManifests=true",
        "/p:ManifestCertificateThumbprint=$($importedCertificate.Thumbprint)"
    )
}
Invoke-Native -FilePath $msbuild -Arguments $buildArguments
} finally {
    if ($certificateStore -and $importedCertificate -and $certificateWasAdded) {
        $certificateStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $certificateStore.Remove($importedCertificate)
        $certificateStore.Close()
    }
    if ($importedCertificate) {
        $importedCertificate.Dispose()
    }
}

$requiredOutputs = @()
foreach ($project in @("Suite.Core", "Suite.Infrastructure", "Suite.Windowing")) {
    $requiredOutputs += Join-Path $sourceDir "$project\bin\$Configuration\$project.dll"
}
foreach ($hostProject in @("Suite.Word", "Suite.Excel", "Suite.PowerPoint")) {
    foreach ($extension in @("dll", "vsto", "dll.manifest")) {
        $requiredOutputs += Join-Path $sourceDir "$hostProject\bin\$Configuration\$hostProject.$extension"
    }
}
$missingOutputs = @($requiredOutputs | Where-Object { -not (Test-Path $_ -PathType Leaf) })
if ($missingOutputs.Count -gt 0) {
    $missingList = $missingOutputs -join [Environment]::NewLine
    throw "Required installer inputs are missing. Build the VSTO projects with the Office/SharePoint development workload so .vsto and .dll.manifest files are generated:$([Environment]::NewLine)$missingList"
}

$coreTestExe = Join-Path $repoRoot "tests\Suite.Core.Tests\bin\$Configuration\Suite.Core.Tests.exe"
if (-not (Test-Path $coreTestExe -PathType Leaf)) {
    throw "The core test executable was not built: $coreTestExe"
}
Invoke-Native -FilePath $coreTestExe -Arguments @()

$outputDir = Join-Path $installerDir "bin\$Configuration\$Platform"
$objectDir = Join-Path $installerDir "obj\$Configuration\$Platform"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $objectDir -Force | Out-Null
$wixObject = Join-Path $objectDir "Product.wixobj"
$msiPath = Join-Path $outputDir "OfficeProductivitySuite-$Configuration-$Platform.msi"

$candleArguments = @(
    "-nologo",
    "-arch", $Platform,
    "-dConfiguration=$Configuration",
    "-dPlatform=$Platform",
    "-dProductVersion=1.0.0",
    "-dSuiteSourceDir=$sourceDir",
    "-out", $wixObject,
    $productWxs
)
Invoke-Native -FilePath $candle -Arguments $candleArguments

$lightArguments = @(
    "-nologo",
    "-out", $msiPath,
    $wixObject
)
Invoke-Native -FilePath $light -Arguments $lightArguments

if ($certificatePath) {
    $signArguments = @("sign", "/fd", "SHA256", "/f", $certificatePath)
    if ($env:SUITE_SIGNING_PASSWORD) {
        $signArguments += @("/p", $env:SUITE_SIGNING_PASSWORD)
    }
    $signArguments += @("/v", $msiPath)
    Invoke-Native -FilePath $signTool -Arguments $signArguments
    Invoke-Native -FilePath $signTool -Arguments @("verify", "/pa", "/v", $msiPath)
    Write-Host "Authenticode signature verification passed: $msiPath"
} else {
    Write-Warning "No SignPath certificate was supplied; the MSI is unsigned and no Authenticode verification was performed."
}

Write-Host "MSI created successfully: $msiPath"
