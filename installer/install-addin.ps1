# Registers the CAT VSTO add-in for the current Windows user.
# Run after building Release (or pass -VstoPath to your CAT.vsto file).

param(
    [string]$VstoPath = (Join-Path $PSScriptRoot "..\src\CAT\bin\Release\CAT.vsto")
)

$VstoPath = (Resolve-Path -LiteralPath $VstoPath -ErrorAction Stop).Path
if ($VstoPath -notlike "*.vsto") { throw "Expected a .vsto file: $VstoPath" }

$manifest = "file:///" + ($VstoPath -replace '\\', '/') + "|vstolocal"
$key = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\CAT"

New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name "Description" -Value "CAT PowerPoint shortcuts and Auto format"
Set-ItemProperty -Path $key -Name "FriendlyName" -Value "CAT"
Set-ItemProperty -Path $key -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $key -Name "Manifest" -Value $manifest

Write-Host "Registered CAT add-in."
Write-Host "  Manifest: $manifest"
Write-Host "Tool icons: place PNGs in assets\icons (named cat*.png), build Release - they copy to bin\Release\icons."
Write-Host "Restart PowerPoint. If prompted, enable the add-in under File -> Options -> Add-ins -> COM Add-ins."
