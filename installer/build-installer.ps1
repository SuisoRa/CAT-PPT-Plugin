# Builds CAT Release, stages files, produces CAT.msi and CAT-Setup.exe (WiX Burn).
# Prerequisites: Visual Studio (VSTO build), WiX Toolset v3.14+, offline prereq EXEs (see prereqs/README.md).

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$wixDir = Join-Path $PSScriptRoot "wix"
$stage = Join-Path $PSScriptRoot "staging\CAT"
$output = Join-Path $PSScriptRoot "output"
$prereqs = Join-Path $PSScriptRoot "prereqs"
$release = Join-Path $root "src\CAT\bin\Release"

function Find-Tool($names, $searchPaths) {
    foreach ($n in $names) {
        $c = Get-Command $n -ErrorAction SilentlyContinue
        if ($c) { return $c.Source }
    }
    foreach ($p in $searchPaths) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

$msbuild = Find-Tool @("msbuild") @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
)
if (-not $msbuild) { throw "MSBuild not found. Install Visual Studio with Office/VSTO workload." }

$wixBin = "${env:ProgramFiles(x86)}\WiX Toolset v3.14\bin"
$candle = Find-Tool @("candle") @(Join-Path $wixBin "candle.exe")
$light = Find-Tool @("light") @(Join-Path $wixBin "light.exe")
$heat = Find-Tool @("heat") @(Join-Path $wixBin "heat.exe")
if (-not $candle -or -not $light -or -not $heat) {
    throw "WiX Toolset v3.14 not found. Install from https://wixtoolset.org/ (candle/light/heat in PATH or Program Files (x86)\WiX Toolset v3.14\bin)."
}

$ndp = Join-Path $prereqs "ndp48-x86-x64-allos-enu.exe"
$vsto = Join-Path $prereqs "vstor_redist.exe"
if (-not (Test-Path $ndp)) { throw "Missing $ndp - see installer/prereqs/README.md" }
if (-not (Test-Path $vsto)) { throw "Missing $vsto - see installer/prereqs/README.md" }

Write-Host "Building CAT (Release, signed manifest if cert present)..."
& $msbuild (Join-Path $root "src\CAT\CAT.csproj") /t:Rebuild /p:Configuration=Release /v:minimal
if ($LASTEXITCODE -ne 0) { throw "CAT Release build failed." }
if (-not (Test-Path (Join-Path $release "CAT.vsto"))) { throw "CAT.vsto not found in $release - VSTO publish output missing." }

Write-Host "Staging install payload..."
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null
Copy-Item -Path (Join-Path $release "*") -Destination $stage -Recurse -Force
Get-ChildItem $stage -Filter "*.pdb" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force
Copy-Item -Path (Join-Path $root "assets\icon-library") -Destination (Join-Path $stage "icon-library") -Recurse -Force

$stagedWx = Join-Path $wixDir "StagedFiles.wxs"
Write-Host "Harvesting staged files (heat)..."
& $heat dir $stage `
    -cg CatStagedFiles `
    -dr INSTALLFOLDER `
    -gg -g1 -sfrag -srd -scom -sreg `
    -var var.StagedSource `
    -out $stagedWx
if ($LASTEXITCODE -ne 0) { throw "heat failed." }

New-Item -ItemType Directory -Path $output -Force | Out-Null
$wxsOut = Join-Path $wixDir "_build"
if (Test-Path $wxsOut) { Remove-Item $wxsOut -Recurse -Force }
New-Item -ItemType Directory -Path $wxsOut -Force | Out-Null
$candleOut = Join-Path $wxsOut ""

$ext = @("-ext", "WixUtilExtension")
Write-Host "Compiling MSI..."
& $candle @ext `
    -dStagedSource="$stage" `
    -dInstallerDir="$PSScriptRoot" `
    -out $candleOut `
    (Join-Path $wixDir "Product.wxs") `
    $stagedWx
if ($LASTEXITCODE -ne 0) { throw "candle (MSI) failed." }

$msi = Join-Path $output "CAT.msi"
& $light @ext -sice:ICE38 -sice:ICE64 -out $msi (Join-Path $wxsOut "Product.wixobj") (Join-Path $wxsOut "StagedFiles.wixobj")
if ($LASTEXITCODE -ne 0) { throw "light (MSI) failed." }

$bundleExt = @("-ext", "WixUtilExtension", "-ext", "WixBalExtension")
Write-Host "Compiling bootstrapper..."
& $candle @bundleExt -dPrereqsDir="$prereqs" -dOutputDir="$output" -out $candleOut (Join-Path $wixDir "Bundle.wxs")
if ($LASTEXITCODE -ne 0) { throw "candle (Bundle) failed." }

$setup = Join-Path $output "CAT-Setup.exe"
& $light @bundleExt -out $setup (Join-Path $wxsOut "Bundle.wixobj")
if ($LASTEXITCODE -ne 0) { throw "light (Bundle) failed." }

Write-Host ""
Write-Host "Done."
Write-Host "  MSI:    $msi"
Write-Host "  Setup:  $setup"
Write-Host ""
Write-Host "Install on a new PC: run CAT-Setup.exe (per-user, %LocalAppData%\Programs\CAT)."
Write-Host "Update icon library later: replace SVGs in that folder\icon-library\ (no reinstall)."
