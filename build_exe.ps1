# Builds dist\VJ-TimesheetTracker.exe - a single file that runs without Python installed.
#   powershell -ExecutionPolicy Bypass -File build_exe.ps1
$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Here

$version = (Select-String -Path tracker_core.py -Pattern '^__version__ = "(.+)"').Matches[0].Groups[1].Value
Write-Host "Building VJ-TimesheetTracker $version"

# isolated build environment (does not touch your global Python)
if (-not (Test-Path .build-venv)) { python -m venv .build-venv }
$py = Join-Path $Here ".build-venv\Scripts\python.exe"
& $py -m pip install --quiet --disable-pip-version-check --upgrade pyinstaller pillow openpyxl
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }

& $py tools\make_icon.py
if ($LASTEXITCODE -ne 0) { throw "icon generation failed" }

# Windows version resource (shown in file Properties > Details)
$v = ($version.Split(".") + @("0", "0", "0"))[0..3] -join ", "
New-Item -ItemType Directory -Force build | Out-Null
@"
VSVersionInfo(
  ffi=FixedFileInfo(filevers=($v), prodvers=($v)),
  kids=[
    StringFileInfo([StringTable('040904B0', [
      StringStruct('CompanyName', 'codewithvj'),
      StringStruct('FileDescription', 'VJ-TimesheetTracker - timesheet from your git history'),
      StringStruct('FileVersion', '$version'),
      StringStruct('InternalName', 'VJ-TimesheetTracker'),
      StringStruct('LegalCopyright', 'MIT License'),
      StringStruct('OriginalFilename', 'VJ-TimesheetTracker.exe'),
      StringStruct('ProductName', 'VJ-TimesheetTracker'),
      StringStruct('ProductVersion', '$version')])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"@ | Set-Content -Encoding utf8 build\version.txt

& $py -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name VJ-TimesheetTracker `
    --icon "$Here\assets\icon.ico" `
    --version-file "$Here\build\version.txt" `
    --distpath "$Here\dist" --workpath "$Here\build\work" --specpath "$Here\build" `
    "$Here\timesheet_app.pyw"
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$exe = Join-Path $Here "dist\VJ-TimesheetTracker.exe"
$hash = (Get-FileHash $exe -Algorithm SHA256).Hash
"$hash  VJ-TimesheetTracker.exe" | Set-Content -Encoding ascii dist\SHA256SUMS.txt
Write-Host ""
Write-Host "Built: $exe  ($([math]::Round((Get-Item $exe).Length / 1MB, 1)) MB)"
Write-Host "SHA256: $hash"
