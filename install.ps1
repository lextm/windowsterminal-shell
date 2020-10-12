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

function Generate-HelperScript(
        # The cache folder
        [Parameter(Mandatory=$true)]
        [string]$cache)
{
    $content = 
    "Set shell = WScript.CreateObject(`"Shell.Application`")
     executable = WSCript.Arguments(0)
     folder = WScript.Arguments(1)
     If Wscript.Arguments.Count > 2 Then
         profile = WScript.Arguments(2)
         ' 0 at the end means to run this command silently
         shell.ShellExecute `"powershell`", `"Start-Process \`"`"`" & executable & `"\`"`" -ArgumentList \`"`"-p \`"`"\`"`"`" & profile & `"\`"`"\`"`" -d \`"`"\`"`"`" & folder & `"\`"`"\`"`" \`"`" `", `"`", `"runas`", 0
     Else
         ' 0 at the end means to run this command silently
         shell.ShellExecute `"powershell`", `"Start-Process \`"`"`" & executable & `"\`"`" -ArgumentList \`"`"-d \`"`"\`"`"`" & folder & `"\`"`"\`"`" \`"`" `", `"`", `"runas`", 0
     End If
    "
    Set-Content -Path "$cache/helper.vbs" -Value $content
}

# https://github.com/Duffney/PowerShell/blob/master/FileSystems/Get-Icon.ps1

Function Get-Icon {

    [CmdletBinding()]
    
    Param ( 
        [Parameter(Mandatory=$True, Position=1, HelpMessage="Enter the location of the .EXE file")]
        [string]$File,

        # If provided, will output the icon to a location
        [Parameter(Position=1, ValueFromPipelineByPropertyName=$true)]
        [string]$OutputFile
    )
    
    [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')  | Out-Null
    
    [System.Drawing.Icon]::ExtractAssociatedIcon($File).ToBitmap().Save($OutputFile)
}

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
    [Parameter(Mandatory=$true, Position=0, ValueFromPipelineByPropertyName=$true)]
    [Alias('Fullname')]
    [string]$File,
   
    # If provided, will output the icon to a location
    [Parameter(Position=1, ValueFromPipelineByPropertyName=$true)]
    [string]$OutputFile
    )
    
    begin {
        Add-Type -AssemblyName System.Drawing   
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

        #region Icon Size bound check
        if ($width -gt 255 -or $height -gt 255) {
            $ratio = ($height, $width | Measure-Object -Maximum).Maximum / 255
            $width /= $ratio
            $height /= $ratio
        }
        #endregion Icon Size bound check

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

function GetProgramFilesFolder(
    [Parameter(Mandatory=$true)]
    [bool]$includePreview)
{
    if (Test-Path "$env:USERPROFILE\scoop\apps\windows-terminal") {
        $result = "$env:USERPROFILE\scoop\apps\windows-terminal\current"
        return $result
    } else {
        $root = "$Env:ProgramFiles\WindowsApps"
    }
    $versionFolders = (Get-ChildItem $root | Where-Object {
            if ($includePreview) {
                $_.Name -like "Microsoft.WindowsTerminal_*__*" -or
                $_.Name -like "Microsoft.WindowsTerminalPreview_*__*"
            } else {
                $_.Name -like "Microsoft.WindowsTerminal_*__*"
            }
        })
    $foundVersion = $null
    $result = $null
    foreach ($versionFolder in $versionFolders) {
        if ($versionFolder.Name -match "[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+") {
            $version = [version]$Matches.0
            Write-Host "Found Windows Terminal version $version."
            if ($null -eq $foundVersion -or $version -gt $foundVersion) {
                $foundVersion = $version
                $result = $versionFolder.FullName
            }
        } else {
            Write-Warning "Found Windows Terminal unsupported version in $versionFolder."
        }
    }

    if ($null -eq $result) {
        Write-Error "Failed to find Windows Terminal actual folder under $root. To install menu items for Windows Terminal Preview, run with ""-Prerelease"" switch Exit."
        exit 1
    }

    if ($foundVersion -lt [version]"0.11") {
        Write-Warning "The latest version found is less than 0.11, which is not tested. The install script might fail in certain way."
    }

    return $result
}

function GetWindowsTerminalIcon(
    [Parameter(Mandatory=$true)]
    [string]$folder,
    [Parameter(Mandatory=$true)]
    [string]$localCache)
{
    $icon = "$localCache\wt.ico"
    $actual = $folder + "\WindowsTerminal.exe"
    if (Test-Path $actual) {
        # use app icon directly.
        Write-Host "Found actual executable $actual."
        $temp = "$localCache\wt.png"
        Get-Icon -File $actual -OutputFile $temp
        ConvertTo-Icon -File $temp -OutputFile $icon
    } else {
        # download from GitHub
        Write-Warning "Didn't find actual executable $actual so download icon from GitHub."
        Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/microsoft/terminal/master/res/terminal.ico" -OutFile $icon
    }

    return $icon
}

function GetActiveProfiles(
    [Parameter(Mandatory=$true)]
    [bool]$isPreview,
    [bool]$isScoop)
{
    if ($isScoop) {
        $file = "$env:LocalAppData\Microsoft\Windows Terminal\settings.json"
    }
    elseif ($isPreview) {
        $file = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json"
    } else {
        $file = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    }
    if (-not (Test-Path $file)) {
        Write-Error "Couldn't find profiles. Please run Windows Terminal at least once after installing it. Exit."
        exit 1
    }

    $settings = Get-Content $file | Out-String | ConvertFrom-Json
    if ($settings.profiles.PSObject.Properties.name -match "list") {
        $list = $settings.profiles.list
    } else {
        $list = $settings.profiles 
    }

    return $list | Where-Object { -not $_.hidden} | Where-Object { ($null -eq $_.source) -or -not ($settings.disabledProfileSources -contains $_.source) }
}

function GetProfileIcon (
    [Parameter(Mandatory=$true)]
    $profile,
    [Parameter(Mandatory=$true)]
    [string]$folder,
    [Parameter(Mandatory=$true)]
    [string]$localCache,
    [Parameter(Mandatory=$true)]
    [string]$defaultIcon,
    [Parameter(Mandatory=$true)]
    [bool]$isPreview,
    [bool]$isScoop)
{
    $guid = $profile.guid
    $name = $profile.name
    $result = $null
    $profilePng = $null
    $icon = $profile.icon
    if ($null -ne $icon) {
        if (Test-Path $icon) {
            # use user setting
            $profilePng = $icon  
        } elseif ($profile.icon -like "ms-appdata:///Roaming/*") {
            #resolve roaming cache
            if ($isPreview) {
                $profilePng = $icon -replace "ms-appdata:///Roaming", "$Env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\RoamingState" -replace "/", "\"
            } else {
                $profilePng = $icon -replace "ms-appdata:///Roaming", "$Env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\RoamingState" -replace "/", "\"
            }
        } elseif ($profile.icon -like "ms-appdata:///Local/*") {
            #resolve local cache
            if ($isPreview) {
                $profilePng = $icon -replace "ms-appdata:///Local", "$Env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState" -replace "/", "\"
            } else {
                $profilePng = $icon -replace "ms-appdata:///Local", "$Env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState" -replace "/", "\"
            }
        } elseif ($profile.icon -like "ms-appx:///*") {
            # resolve app cache
            $profilePng = $icon -replace "ms-appx://", $folder -replace "/", "\"
        } elseif ($profile.icon -like "*%*") {
            $profilePng = [System.Environment]::ExpandEnvironmentVariables($icon)
        } else {
            Write-Host "Invalid profile icon found $icon. Please report an issue at https://github.com/lextm/windowsterminal-shell/issues ."
        }
    }

    if (($null -eq $profilePng) -or -not (Test-Path $profilePng)) {
        # fallback to profile PNG
        $profilePng = "$folder\ProfileIcons\$guid.scale-200.png"
        if (-not (Test-Path($profilePng))) {
            if ($profile.source -eq "Windows.Terminal.Wsl") {
                $profilePng = "$folder\ProfileIcons\{9acb9455-ca41-5af7-950f-6bca1bc9722f}.scale-200.png"
            }
        }
    }

    if (Test-Path $profilePng) {        
        if ($profilePng -like "*.png") {
            # found PNG, convert to ICO
            $result = "$localCache\$guid.ico"
            ConvertTo-Icon -File $profilePng -OutputFile $result
        } elseif ($profilePng -like "*.ico") {
            $result = $profilePng
        } else {
            Write-Warning "Icon format is not supported by this script $profilePng. Please use PNG or ICO format."
        }
    } else {
        Write-Warning "Didn't find icon for profile $name."
    }

    if ($null -eq $result) {
        # final fallback
        $result = $defaultIcon
    }

    return $result
}

function CreateMenuItem(
    [Parameter(Mandatory=$true)]
    [string]$rootKey,
    [Parameter(Mandatory=$true)]
    [string]$name,
    [Parameter(Mandatory=$true)]
    [string]$icon,
    [Parameter(Mandatory=$true)]
    [string]$command,
    [Parameter(Mandatory=$true)]
    [bool]$elevated
)
{
    New-Item -Path $rootKey -Force | Out-Null
    New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $rootKey -Name 'Icon' -PropertyType String -Value $icon | Out-Null
    if ($elevated) {
        New-ItemProperty -Path $rootKey -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
    }

    New-Item -Path "$rootKey\command" -Force | Out-Null
    New-ItemProperty -Path "$rootKey\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null
}

function CreateProfileMenuItems(
    [Parameter(Mandatory=$true)]
    $profile,
    [Parameter(Mandatory=$true)]
    [string]$executable,
    [Parameter(Mandatory=$true)]
    [string]$folder,
    [Parameter(Mandatory=$true)]
    [string]$localCache,
    [Parameter(Mandatory=$true)]
    [string]$icon,
    [Parameter(Mandatory=$true)]
    [string]$layout,
    [Parameter(Mandatory=$true)]
    [bool]$isPreview,
    [bool]$isScoop)
{
    $guid = $profile.guid
    $name = $profile.name
    $command = """$executable"" -p ""$name"" -d ""%V."""
    $elevated = "wscript.exe ""$localCache/helper.vbs"" ""$executable"" ""%V."" ""$name"""
    $profileIcon = GetProfileIcon $profile $folder $localCache $icon $isPreview

    if ($layout -eq "Default") {
        $rootKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\ContextMenus\MenuTerminal\shell\$guid"
        $rootKeyElevated = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\ContextMenus\MenuTerminalAdmin\shell\$guid"
        CreateMenuItem $rootKey $name $profileIcon $command $false
        CreateMenuItem $rootKeyElevated $name $profileIcon $elevated $true
    } elseif ($layout -eq "Flat") {
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminal_$guid" "$name here" $profileIcon $command $false
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminalAdmin_$guid" "$name here as administrator" $profileIcon $elevated $true   
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminal_$guid" "$name here" $profileIcon $command $false
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdmin_$guid" "$name here as administrator" $profileIcon $elevated $true   
    }
}

function CreateMenuItems(
    [Parameter(Mandatory=$true)]
    [string]$executable,
    [Parameter(Mandatory=$true)]
    [string]$layout,
    [Parameter(Mandatory=$true)]
    [bool]$includePreview)
{
    $folder = GetProgramFilesFolder $includePreview
    $localCache = "$Env:LOCALAPPDATA\Microsoft\WindowsApps\Cache"

    if (-not (Test-Path $localCache)) {
        New-Item $localCache -ItemType Directory | Out-Null
    }

    Generate-HelperScript $localCache
    $icon = GetWindowsTerminalIcon $folder $localCache

    if ($layout -eq "Default") {
        # defaut layout creates two menus
        New-Item -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminal' -Force | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminal' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal here' | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminal' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminal' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminal' | Out-Null

        New-Item -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminal' -Force | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminal' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal here' | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminal' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminal' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminal' | Out-Null

        New-Item -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\ContextMenus\MenuTerminal\shell' -Force | Out-Null

        New-Item -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminalAdmin' -Force | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminalAdmin' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal here as administrator' | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminalAdmin' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminalAdmin' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminalAdmin' | Out-Null

        New-Item -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdmin' -Force | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdmin' -Name 'MUIVerb' -PropertyType String -Value 'Windows Terminal here as administrator' | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdmin' -Name 'Icon' -PropertyType String -Value $icon | Out-Null
        New-ItemProperty -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdmin' -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'Directory\\ContextMenus\\MenuTerminalAdmin' | Out-Null

        New-Item -Path 'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\ContextMenus\MenuTerminalAdmin\shell' -Force | Out-Null
    } elseif ($layout -eq "Mini") {
        $command = """$executable"" -d ""%V."""
        $elevated = "wscript.exe ""$localCache/helper.vbs"" ""$executable"" ""%V."""
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminalMini" "Windows Terminal here" $icon $command $false
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminalAdminMini" "Windows Terminal here as administrator" $icon $elevated $true   
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalMini" "Windows Terminal here" $icon $command $false
        CreateMenuItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\MenuTerminalAdminMini" "Windows Terminal here as administrator" $icon $elevated $true   
        return
    }

    $isPreview = $folder -like "*WindowsTerminalPreview*"
    $isScoop = $folder -like "*scoop\apps\windows-terminal*"
    $profiles = GetActiveProfiles $isPreview $isScoop
    foreach ($profile in $profiles) {
        CreateProfileMenuItems $profile $executable $folder $localCache $icon $layout $isPreview $isScoop
    }
}

# Based on @nerdio01's version in https://github.com/microsoft/terminal/issues/1060

if ((Test-Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\MenuTerminal") -and
    -not (Test-Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\MenuTerminal")) {
    Write-Error "Please execute uninstall.old.ps1 to remove previous installation."
    exit 1
}

if ($PSVersionTable.PSVersion.Major -lt 6) {
    Write-Error "Must be executed in PowerShell 6 and above. Learn how to install it from https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-windows?view=powershell-7 . Exit."
    exit 1
}

if (Test-Path "$env:USERPROFILE\scoop\apps\windows-terminal") {
        $executable = "$env:USERPROFILE\scoop\apps\windows-terminal\current\wt.exe"
} else {
        $executable = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
}

if (-not (Test-Path $executable)) {
    Write-Error "Windows Terminal not detected at $executable. Learn how to install it from https://github.com/microsoft/terminal (via Microsoft Store is recommended). Exit."
    exit 1
}

Write-Host "Use $Layout layout."

CreateMenuItems $executable $Layout $PreRelease

Write-Host "Windows Terminal installed to Windows Explorer context menu."
