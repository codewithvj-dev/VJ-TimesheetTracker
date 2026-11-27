# VJ-TimesheetTracker installer (Windows, per-user, no admin needed)
#   powershell -ExecutionPolicy Bypass -File install.ps1 [-Time 19:00]
param([string]$Time = "19:00")
$ErrorActionPreference = "Stop"
$App  = "VJ-TimesheetTracker"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path

# 1. Python 3.9+
$py = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $py -or $py -like "*WindowsApps*") {
    $launcher = (Get-Command py -ErrorAction SilentlyContinue).Source
    if ($launcher) { $py = (& $launcher -3 -c "import sys; print(sys.executable)").Trim() }
}
if (-not $py) { throw "Python 3.9+ not found. Install it from https://www.python.org/downloads/ (tick 'Add to PATH')." }
$ver = & $py -c "import sys; print('%d.%d' % sys.version_info[:2])"
if ([version]$ver -lt [version]"3.9") { throw "Python $ver found; 3.9 or newer is required." }
& $py -c "import tkinter" 2>$null
if ($LASTEXITCODE -ne 0) { throw "This Python has no tkinter. Reinstall Python with the 'tcl/tk and IDLE' option." }
& git --version *> $null
if ($LASTEXITCODE -ne 0) { throw "git not found on PATH. Install Git for Windows first." }
$pyw = Join-Path (Split-Path $py) "pythonw.exe"
if (-not (Test-Path $pyw)) { $pyw = $py }
Write-Host "Python: $py ($ver)"

# 2. Dependencies
& $py -m pip install --user --quiet --disable-pip-version-check -r (Join-Path $Here "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "pip install failed." }

# 3. Daily update task
& $py (Join-Path $Here "tracker_core.py") schedule $Time
if ($LASTEXITCODE -ne 0) { throw "Could not create the scheduled task." }

# 4. Shortcuts (Desktop + Start Menu)
$shell = New-Object -ComObject WScript.Shell
$targets = @([Environment]::GetFolderPath("Desktop"), (Join-Path ([Environment]::GetFolderPath("Programs")) ""))
foreach ($dir in $targets) {
    $lnk = $shell.CreateShortcut((Join-Path $dir "$App.lnk"))
    $lnk.TargetPath = $pyw
    $lnk.Arguments = "`"$(Join-Path $Here 'timesheet_app.pyw')`""
    $lnk.WorkingDirectory = $Here
    $lnk.IconLocation = "$env:SystemRoot\System32\imageres.dll,186"
    $lnk.Description = "Timesheet from your git commits"
    $lnk.Save()
}

Write-Host ""
Write-Host "$App installed."
Write-Host "  Open it from the Desktop or Start Menu shortcut, then click 'Add project...'."
Write-Host "  Daily update: $Time   Data: $env:APPDATA\$App"
