<#
.SYNOPSIS
  추가 기능 3개와 수신기를 빌드하고 사용자별 MSI 를 만든다. (Windows, Visual Studio 2022 + Office 개발 워크로드 필요)

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File installer\build.ps1 -Version 1.2.0
  → artifacts\OfficeToMarkdown-1.2.0.msi
#>
param(
    [string]$Version = "1.0.0",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$artifacts = Join-Path $root "artifacts"

function Invoke-Checked {
    param([string]$Exe, [string[]]$Arguments)
    & $Exe @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Exe 실패 (종료 코드 $LASTEXITCODE)" }
}

# 1. VSTO 추가 기능: dotnet build 로는 빌드되지 않으므로 Visual Studio 의 MSBuild 를 쓴다. (매니페스트 서명 포함)
$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) { throw "Visual Studio 를 찾을 수 없습니다: $vswhere" }
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
if (-not $msbuild) { throw "MSBuild 를 찾을 수 없습니다." }
Write-Host "== 추가 기능 빌드 ($Configuration)"
Invoke-Checked $msbuild @((Join-Path $root "OfficeToMarkdown.sln"), "/restore", "/m", "/v:minimal", "/p:Configuration=$Configuration")

# 2. 수신기: .NET 런타임이 없는 PC 에서도 실행되도록 self-contained 단일 exe 로 게시한다.
Write-Host "== 수신기 게시"
$listenerOut = Join-Path $artifacts "listener"
if (Test-Path $listenerOut) { Remove-Item $listenerOut -Recurse -Force }
Invoke-Checked dotnet @(
    "publish", (Join-Path $root "samples\ClipboardListener\ClipboardListener.csproj"),
    "-c", "Release", "-r", "win-x64", "--self-contained",
    "-p:PublishSingleFile=true", "-p:IncludeNativeLibrariesForSelfExtract=true", "-p:EnableCompressionInSingleFile=true",
    "-o", $listenerOut)

# 3. MSI
Write-Host "== MSI 빌드"
Invoke-Checked dotnet @(
    "build", (Join-Path $PSScriptRoot "OfficeToMarkdown.Installer.wixproj"),
    "-c", "Release",
    "-p:ProductVersion=$Version", "-p:AddInConfiguration=$Configuration", "-p:ListenerDir=$listenerOut",
    "-o", $artifacts)

Write-Host ""
Write-Host "완료: $(Join-Path $artifacts "OfficeToMarkdown-$Version.msi")"
