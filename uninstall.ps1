#Requires -Version 6

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('Default', 'Flat', 'Mini')]
    [string] $Layout = 'Default',
    [Parameter()]
    [switch] $PreRelease
)

#Specify target user if different from account running the script, else comment this out
$targetUser = "example.username"  # Optional - specify user. Otherwise, uncomment the below
# $targetUser = $Env:UserName

$targetUserObject = New-Object System.Security.Principal.NTAccount($targetUser)
$targetUserSID = $targetUserObject.Translate([System.Security.Principal.SecurityIdentifier]).value

$Env:LocalAppData = "$Env:HOMEDRIVE\Users\$targetUser\AppData\Local"
$Env:LOCALAPPDATA = $Env:LocalAppData

# Based on @nerdio01's version in https://github.com/microsoft/terminal/issues/1060

if ((Get-Process -Id $pid).Path -like "*WindowsApps*") {
    Write-Error "PowerShell installed via Microsoft Store is not supported. Learn other ways to install it from https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-windows?view=powershell-7 . Exit.";
    exit 1
}

if ((Test-Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\MenuTerminal") -and
    -not (Test-Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\shell\MenuTerminal")) {
    Write-Error "Please execute uninstall.old.ps1 to remove previous installation."
    exit 1
}

$localCache = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\Cache"
if (Test-Path $localCache) {
    Remove-Item $localCache -Recurse
}

Write-Host "Use" $layout "layout."

if ($layout -eq "Default") {
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\shell\MenuTerminal" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\Background\shell\MenuTerminal" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\ContextMenus\MenuTerminal\shell" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\shell\MenuTerminalAdmin" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdmin" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\ContextMenus\MenuTerminalAdmin\shell" -Recurse -ErrorAction Ignore | Out-Null
} elseif ($layout -eq "Flat") {
    $rootKey = 'HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell'
    foreach ($key in Get-ChildItem -Path "Registry::$rootKey") {
       if (($key.Name -like "$rootKey\MenuTerminal_*") -or ($key.Name -like "$rootKey\MenuTerminalAdmin_*")) {
          Remove-Item "Registry::$key" -Recurse -ErrorAction Ignore | Out-Null
       }
    }

    $rootKey = 'HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell'
    foreach ($key in Get-ChildItem -Path "Registry::$rootKey") {
       if (($key.Name -like "$rootKey\MenuTerminal_*") -or ($key.Name -like "$rootKey\MenuTerminalAdmin_*")) {
          Remove-Item "Registry::$key" -Recurse -ErrorAction Ignore | Out-Null
       }
    }
} elseif ($layout -eq "Mini") {
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\shell\MenuTerminalMini" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\shell\MenuTerminalAdminMini" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalMini" -Recurse -ErrorAction Ignore | Out-Null
    Remove-Item -Path "Registry::HKEY_USERS\$targetUserSID\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdminMini" -Recurse -ErrorAction Ignore | Out-Null
}

Write-Host "Windows Terminal uninstalled from Windows Explorer context menu."
