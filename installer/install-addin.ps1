# Run after building Release (or pass -VstoPath to your CAT.vsto file).

param(
    [string]$VstoPath = (Join-Path $PSScriptRoot "..\src\CAT\bin\Release\CAT.vsto")
)

& (Join-Path $PSScriptRoot "Register-CatAddin.ps1") -VstoPath $VstoPath

Write-Host "Registered CAT add-in."
Write-Host "  VSTO: $((Resolve-Path -LiteralPath $VstoPath).Path)"
Write-Host "Restart PowerPoint. If prompted, enable the add-in under File -> Options -> Add-ins -> COM Add-ins."
