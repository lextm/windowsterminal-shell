# Based on @nerdio01's version in https://github.com/microsoft/terminal/issues/1060

$executable = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
if (!(Test-Path $executable)) {
    Write-Host "Windows Terminal not detected. Learn how to install it from https://github.com/microsoft/terminal"
    exit 1
}

$powershell = "$executable -p ""Windows PowerShell"" -d ""%V."""
$powershell2 = "PowerShell -WindowStyle Hidden -Command ""Start-Process PowerShell.exe -WindowStyle Hidden -Verb RunAs -ArgumentList \""-Command ""$executable"" -d ""%V."" -p ""Windows PowerShell""\"" """

$cmd = "$executable -p ""cmd"" -d ""%V."""
$cmd2 = "PowerShell -WindowStyle Hidden -Command ""Start-Process cmd.exe -WindowStyle Hidden -Verb RunAs -ArgumentList \""/c ""$executable -p ""cmd"" -d ""%V.""\"" """

$folder = (Get-ChildItem "$Env:ProgramFiles\WindowsApps" | Where-Object { $_.Name.StartsWith("Microsoft.WindowsTerminal_") } | Select-Object -First 1)
$actual = $folder.FullName + "\WindowsTerminal.exe"
if (Test-Path $actual) {
    # use app icon directly.
    Write-Host "found actual executable" $actual
    $icon = $actual
} else {
    # download from GitHub
    Write-Host "didn't find actual executable" $actual
    $icon = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\wt.icon"
    Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/microsoft/terminal/master/res/terminal.ico" -OutFile $icon
}

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminal' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell' -Force | Out-Null
New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open' -Name 'MUIVerb' -PropertyType String -Value 'Windows PowerShell' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open' -Name 'Icon' -PropertyType String -Value PowerShell.exe | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open\command' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\open\command' -Name '(Default)' -PropertyType String -Value $powershell | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\cmd' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\cmd' -Name 'MUIVerb' -PropertyType String -Value 'Command Prompt' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\cmd' -Name 'Icon' -PropertyType String -Value cmd.exe | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\cmd\command' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\cmd\command' -Name '(Default)' -PropertyType String -Value $cmd | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal (Admin)' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminalAdmin' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Name 'MUIVerb' -PropertyType String -Value 'Windows PowerShell' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Name 'Icon' -PropertyType String -Value PowerShell.exe | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open' -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open\command' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\open\command'-Name '(Default)' -PropertyType String -Value $powershell2 | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\cmd' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\cmd' -Name 'MUIVerb' -PropertyType String -Value 'Command Prompt' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\cmd' -Name 'Icon' -PropertyType String -Value cmd.exe | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\cmd' -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\cmd\command' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\cmd\command'-Name '(Default)' -PropertyType String -Value $cmd2 | Out-Null


Write-Host "Windows Terminal installed to Windows Explorer context menu."
