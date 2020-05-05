# Based on @nerdio01's version in https://github.com/microsoft/terminal/issues/1060

$executable = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
if (!(Test-Path $executable)) {
    Write-Host "Windows Terminal not detected. Learn hwo to install it from https://github.com/microsoft/terminal"
    exit 1
}

$command = "$executable -d ""%V."""
$elevated = "PowerShell -WindowStyle Hidden -Command ""Start-Process PowerShell.exe -WindowStyle Hidden -Verb RunAs -ArgumentList \""-Command ""$executable"" -d ""%V.""\"" """
$icon = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\wt.icon"

Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/microsoft/terminal/master/res/terminal.ico" -OutFile $icon  # Going to update my own to just grab icon from the appx package

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminal' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell' -Force | Out-Null
New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open' -Name 'MUIVerb' -PropertyType String -Value 'PowerShell' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open' -Name 'Icon' -PropertyType String -Value PowerShell.exe | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open\command' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open\command' -Name '(Default)' -PropertyType String -Value $command | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal (Admin)' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminalAdmin' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Name 'MUIVerb' -PropertyType String -Value 'PowerShell' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Name 'Icon' -PropertyType String -Value PowerShell.exe | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open\command' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open\command'-Name '(Default)' -PropertyType String -Value $elevated | Out-Null

Write-Host "Windows Terminal installed to Windows Explorer context menu."
