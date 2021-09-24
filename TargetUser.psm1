
function Get-TargetUser(
    [Parameter(Mandatory=$false)]
    [string]$defaultTargetUserName)
{

    if (!$defaultTargetUserName){ $defaultTargetUserName = $Env:UserName }

    $newTargetUserName = $defaultTargetUserName
    $newTargetUserSID = ""
    $validUserSelected = $false

    DO
    {

        $userInput = Read-Host "Please enter a username for the installation. Hit <Enter> to install for $defaultTargetUserName "

        if (!$userInput) {
            $newTargetUserName = $defaultTargetUserName
        } else {
            $newTargetUserName = $userInput
        }

        $newTargetUserObject = New-Object System.Security.Principal.NTAccount($newTargetUserName)
        try {
            $newTargetUserSID = $newTargetUserObject.Translate([System.Security.Principal.SecurityIdentifier]).value
            $validUserSelected = $true
        } catch {
            Write-Warning "Specified user $userInput not found"
        }
       
    } until ($validUserSelected)

    if ($newTargetUserName -eq $Env:UserName){
        $newTargetUserProfileDirectory = $env:USERPROFILE
        $newTargetUserLocalAppData = $env:LOCALAPPDATA
    } else {
        $newTargetUserProfileDirectory = Get-ItemPropertyValue "Registry::HKEY_USERS\$newTargetUserSID\Volatile Environment" -Name USERPROFILE
        $newTargetUserLocalAppData = Get-ItemPropertyValue "Registry::HKEY_USERS\$newTargetUserSID\Volatile Environment" -Name LOCALAPPDATA
    }

    return [PSCustomObject]@{
        userName = $newTargetUserName
        SID = $newTargetUserSID
        userProfile = $newTargetUserProfileDirectory
        localAppData = $newTargetUserLocalAppData
    }

}
