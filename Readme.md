# PowerShell Scripts to Install/Uninstall Context Menu Items for Windows Terminal

![Menu items](menu.png)

## Install

1. [Install Windows Terminal](https://github.com/microsoft/terminal).
1. [Install PowerShell 7](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-windows?view=powershell-7).
1. Launch PowerShell 7 console as administrator, and run `install.ps1` to install context menu items to Windows Explorer.

## Uninstall
1. Run `uninstall.ps1` to uninstall context menu items from Windows Explorer.

## Notes
The current release only supports Windows 10 machines (Windows Terminal restriction).

The scripts must be run as administrator.

`install.ps1` and `uninstall.ps1` only manipulate Windows Explorer settings for the context menu items, and do not write to Windows Terminal settings.

Downloading Windows Terminal icon from GitHub (in `install.ps1`) requires internet connection, but in general is just an optional step that won't be executed in most cases.
