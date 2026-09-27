param(
    [Parameter(Mandatory = $true)]
    [string]$VstoPath
)

$ErrorActionPreference = "Stop"
if (-not [System.IO.Path]::IsPathRooted($VstoPath)) {
    $VstoPath = Join-Path (Get-Location).Path $VstoPath
}
$VstoPath = (Resolve-Path -LiteralPath $VstoPath -ErrorAction Stop).Path
if ($VstoPath -notlike "*.vsto") { throw "Expected a .vsto file: $VstoPath" }

$manifest = "file:///" + ($VstoPath -replace '\\', '/') + "|vstolocal"
$key = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\CAT"

New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name "Description" -Value "CAT PowerPoint shortcuts and Auto format"
Set-ItemProperty -Path $key -Name "FriendlyName" -Value "CAT"
Set-ItemProperty -Path $key -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $key -Name "Manifest" -Value $manifest
