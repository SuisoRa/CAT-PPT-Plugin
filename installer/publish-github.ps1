# Creates github.com/<your-account>/CAT and pushes the current project.
# Prerequisites: Git + GitHub CLI installed; run `gh auth login` once.

param(
    [string]$RepoName = "CAT",
    [ValidateSet("private", "public")]
    [string]$Visibility = "private"
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

$gh = "${env:ProgramFiles}\GitHub CLI\gh.exe"
$git = "${env:ProgramFiles}\Git\bin\git.exe"
if (-not (Test-Path $gh)) { throw "GitHub CLI not found. Install from https://cli.github.com/" }
if (-not (Test-Path $git)) { throw "Git not found. Install from https://git-scm.com/" }

& $gh auth status | Out-Null
$login = & $gh api user -q .login
$email = "$login@users.noreply.github.com"

if (-not (Test-Path ".git")) {
    & $git init -b main
}

$hasHead = $false
try { & $git rev-parse HEAD 2>$null | Out-Null; $hasHead = $true } catch { }

if (-not $hasHead) {
    & $git add -A
    & $git -c "user.name=$login" -c "user.email=$email" commit -m "Initial commit: CAT PowerPoint add-in"
}

if (& $git remote 2>$null | Select-String -Pattern "^origin$" -Quiet) {
    Write-Host "Remote 'origin' already exists. Push with: git push -u origin main"
    exit 0
}

& $gh repo create $RepoName --$Visibility --source=. --remote=origin --push --description "CAT PowerPoint VSTO add-in (shortcuts, Auto format, Table Creator)"
Write-Host "Done. Repository: https://github.com/$login/$RepoName"
