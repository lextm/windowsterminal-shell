# https://gist.github.com/darkfall/1656050
function ConvertTo-Icon
{
    <#
    .Synopsis
        Converts image to icons
    .Description
        Converts an image to an icon
    .Example
        ConvertTo-Icon -File .\Logo.png -OutputFile .\Favicon.ico
    #>
    [CmdletBinding()]
    param(
    # The file
    [Parameter(Mandatory=$true, Position=0,ValueFromPipelineByPropertyName=$true)]
    [Alias('Fullname')]
    [string]$File,
   
    # If provided, will output the icon to a location
    [Parameter(Position=1, ValueFromPipelineByPropertyName=$true)]
    [string]$OutputFile
    )
    
    begin {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing   
    }
    
    process {
        #region Load Icon
        $resolvedFile = $ExecutionContext.SessionState.Path.GetResolvedPSPathFromPSPath($file)
        if (-not $resolvedFile) { return }
        $inputBitmap = [Drawing.Image]::FromFile($resolvedFile)
        $width = $inputBitmap.Width
        $height = $inputBitmap.Height
        $size = New-Object Drawing.Size $width, $height
        $newBitmap = New-Object Drawing.Bitmap $inputBitmap, $size
        #endregion Load Icon

        #region Save Icon                     
        $memoryStream = New-Object System.IO.MemoryStream
        $newBitmap.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Png)

        $resolvedOutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outputFile)
        $output = [IO.File]::Create("$resolvedOutputFile")
        
        $iconWriter = New-Object System.IO.BinaryWriter($output)
        # 0-1 reserved, 0
        $iconWriter.Write([byte]0)
        $iconWriter.Write([byte]0)

        # 2-3 image type, 1 = icon, 2 = cursor
        $iconWriter.Write([short]1);

        # 4-5 number of images
        $iconWriter.Write([short]1);

        # image entry 1
        # 0 image width
        $iconWriter.Write([byte]$width);
        # 1 image height
        $iconWriter.Write([byte]$height);

        # 2 number of colors
        $iconWriter.Write([byte]0);

        # 3 reserved
        $iconWriter.Write([byte]0);

        # 4-5 color planes
        $iconWriter.Write([short]0);

        # 6-7 bits per pixel
        $iconWriter.Write([short]32);

        # 8-11 size of image data
        $iconWriter.Write([int]$memoryStream.Length);

        # 12-15 offset of image data
        $iconWriter.Write([int](6 + 16));

        # write image data
        # png data must contain the whole png data file
        $iconWriter.Write($memoryStream.ToArray());

        $iconWriter.Flush();
        $output.Close()               
        #endregion Save Icon

        #region Cleanup
        $memoryStream.Dispose()
        $newBitmap.Dispose()
        $inputBitmap.Dispose()
        #endregion Cleanup
    }
}

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
$localCache = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\Cache"
$actual = $folder.FullName + "\WindowsTerminal.exe"
if (Test-Path $actual) {
    # use app icon directly.
    Write-Host "found actual executable" $actual
    $icon = $actual
} else {
    # download from GitHub
    Write-Host "didn't find actual executable" $actual
    $icon = "$localCache\wt.icon"
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

    $profileIcon = $null;
    if ($null -ne $profile.icon) {
        if (Test-Path $profile.icon) {
            $profileIcon = $profile.icon  
        } elseif ($pr.ofile.icon.StartsWith("ms-appdata:///")) {
            Write-Host "Not implemented. Please report an issue at https://github.com/lextm/windowsterminal-shell/issues"   
        } else {
            Write-Host "Invalid profile icon found" $profile.icon ". Please report an issue at https://github.com/lextm/windowsterminal-shell/issues"
        }
    } else {
        $cachePng = "$folder\ProfileIcons\$guid.scale-200.png"
        if (Test-Path $cachePng) {
            if (!(Test-Path $localCache)) {
                New-Item $localCache -ItemType Directory | Out-Null
            }

            $cacheIcon = "$localCache\$guid.ico"
            ConvertTo-Icon -File $cachePng -OutputFile $cacheIcon
            $profileIcon = $cacheIcon
        }
    }

    if ($null -eq $profileIcon) {
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
