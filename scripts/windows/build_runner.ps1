param(
  [string]$Name = "clearml_dataset_excel_runner",
  [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\\..")
if ($OutDir -eq "") {
  $OutDir = (Join-Path $RepoRoot "dist")
}

Write-Host "Repo: $RepoRoot"
Write-Host "Out : $OutDir"

py -3 -m pip install -U pip pyinstaller
py -3 -m pip install -r (Join-Path $RepoRoot "requirements.txt")
py -3 -m pip install -e $RepoRoot

Push-Location $RepoRoot
try {
  py -3 -m PyInstaller -F -n $Name -m clearml_dataset_excel.runner --distpath $OutDir --workpath (Join-Path $RepoRoot "build")
} finally {
  Pop-Location
}

$Exe = Join-Path $OutDir "$Name.exe"
Write-Output $Exe

