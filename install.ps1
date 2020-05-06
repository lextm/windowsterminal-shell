# Based on @nerdio01's version in https://github.com/microsoft/terminal/issues/1060

if ($PSVersionTable.PSVersion.Major -lt 6) {
    Write-Host "Must be executed in PowerShell 6 and above. Exit."
    exit 1
}

$executable = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
if (!(Test-Path $executable)) {
    Write-Host "Windows Terminal not detected. Learn how to install it from https://github.com/microsoft/terminal . Exit."
    exit 1
}

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
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal here' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminal' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminal' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell' -Force | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Force | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal (Admin) here' | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
New-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\MenuTerminalAdmin' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminalAdmin' | Out-Null

New-Item -Path 'Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell' -Force | Out-Null

$settings = Get-Content "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json" | Out-String | ConvertFrom-Json
$profiles = $settings.profiles | Where-Object { !$_.hidden }

foreach ($profile in $profiles) {
    $guid = $profile.guid
    $name = $profile.name
    if ($profile.commandline -match '(?<commandline>.+\.exe)(\s+.*)?') {
        $commandline = $Matches.commandline
    } else {
        $commandline = $null
    }

    $command = "$executable -p ""$name"" -d ""%V."""
    $elevated1 = "PowerShell -WindowStyle Hidden -Command ""Start-Process PowerShell.exe -WindowStyle Hidden -Verb RunAs -ArgumentList \""-Command ""$executable"" -d ""%V."" -p ""$name""\"" """
    $elevated2 = "PowerShell -WindowStyle Hidden -Command ""Start-Process cmd.exe -WindowStyle Hidden -Verb RunAs -ArgumentList \""/c ""$executable -p ""$name"" -d ""%V.""\"" """
    if ($commandline -eq "cmd.exe") {
        $elevated = $elevated2
    } else {
        $elevated = $elevated1
    }

    if ($null -ne $profile.icon) {
        $profileIcon = $profile.icon
    } elseif ($null -ne $commandline) {
        $profileIcon = $commandline
    } else {
        $profileIcon = $icon
    }

    if (($null -eq $profile.source) -or !($settings.disabledProfileSources -contains $profile.source)) {
        New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\$guid" -Force | Out-Null
        New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\$guid" -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
        New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\$guid" -Name 'Icon' -PropertyType String -Value $profileIcon | Out-Null
        
        New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\$guid\command" -Force | Out-Null
        New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminal\shell\$guid\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null

        New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\$guid" -Force | Out-Null
        New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\$guid" -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
        New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\$guid" -Name 'Icon' -PropertyType String -Value $profileIcon | Out-Null
        New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\$guid" -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
        
        New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\$guid\command" -Force | Out-Null
        New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuTerminalAdmin\shell\$guid\command" -Name '(Default)' -PropertyType String -Value $elevated | Out-Null
    }
}

Write-Host "Windows Terminal installed to Windows Explorer context menu."
