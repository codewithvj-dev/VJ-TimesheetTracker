# VJ-TimesheetTracker uninstaller
#   powershell -ExecutionPolicy Bypass -File uninstall.ps1 [-RemoveData]
param([switch]$RemoveData)
$App = "VJ-TimesheetTracker"

Unregister-ScheduledTask -TaskName "$App Daily Update" -Confirm:$false -ErrorAction SilentlyContinue
foreach ($dir in @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs"))) {
    Remove-Item (Join-Path $dir "$App.lnk") -ErrorAction SilentlyContinue
}
if ($RemoveData) {
    Remove-Item (Join-Path $env:APPDATA $App) -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Removed task, shortcuts and saved data."
} else {
    Write-Host "Removed task and shortcuts. Your data is kept in $env:APPDATA\$App (use -RemoveData to delete it)."
}
