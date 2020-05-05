# Based on @nerdio01's version in https://github.com/microsoft/terminal/issues/1060

$icon = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\wt.ico"
if (Test-Path $icon) {
  Remove-Item $icon
}

Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Recurse -ErrorAction Ignore | Out-Null

Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell' -Recurse -ErrorAction Ignore | Out-Null

Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Recurse -ErrorAction Ignore | Out-Null

Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell' -Recurse -ErrorAction Ignore | Out-Null

Write-Host "Windows Terminal uninstalled from Windows Explorer context menu."
