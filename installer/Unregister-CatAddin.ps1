$key = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\CAT"
if (Test-Path -LiteralPath $key) {
    Remove-Item -LiteralPath $key -Recurse -Force -ErrorAction SilentlyContinue
}
