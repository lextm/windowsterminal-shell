#Requires -RunAsAdministrator
#Requires -Version 6

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('Default', 'Flat', 'Mini')]
    [string] $Layout = 'Default',
    [Parameter()]
    [switch] $PreRelease
)

# Based on @nerdio01's version in https://github.com/microsoft/terminal/issues/1060

if ((Get-Process -Id $pid).Path -like "*WindowsApps*") {
    Write-Error "PowerShell installed via Microsoft Store is not supported. Learn other ways to install it from https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-windows?view=powershell-7 . Exit.";
    exit 1
}

$localCache = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\Cache"
if (Test-Path $localCache) {
    Remove-Item $localCache -Recurse
}

Write-Host "Use" $layout "layout."

if ($layout -eq "Default") {
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\shell\MenuTerminal' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\shell\MenuTerminalAdmin' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell' -Recurse -ErrorAction Ignore | Out-Null
} elseif ($layout -eq "Flat") {
    $rootKey = 'HKEY_CLASSES_ROOT\Directory\shell'
    foreach ($key in Get-ChildItem -Path "Registry::$rootKey") {
       if (($key.Name -like "$rootKey\MenuTerminal_*") -or ($key.Name -like "$rootKey\MenuTerminalAdmin_*")) {
          Remove-Item "Registry::$key" -Recurse -ErrorAction Ignore | Out-Null
       }
    }

    $rootKey = 'HKEY_CLASSES_ROOT\Directory\Background\shell'
    foreach ($key in Get-ChildItem -Path "Registry::$rootKey") {
       if (($key.Name -like "$rootKey\MenuTerminal_*") -or ($key.Name -like "$rootKey\MenuTerminalAdmin_*")) {
          Remove-Item "Registry::$key" -Recurse -ErrorAction Ignore | Out-Null
       }
    }
} elseif ($layout -eq "Mini") {
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\shell\MenuTerminalMini' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\shell\MenuTerminalAdminMini' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalMini' -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdminMini' -Recurse -ErrorAction Ignore | Out-Null
}

Write-Host "Windows Terminal uninstalled from Windows Explorer context menu."
