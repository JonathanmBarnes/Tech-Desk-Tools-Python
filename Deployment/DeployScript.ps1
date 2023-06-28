clear

# Check if the script is running with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Prompt the user to run the script as an administrator
    $arguments = "& '" + $MyInvocation.MyCommand.Definition + "'"
    Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList $arguments
    exit
}

###### Global Configuration - (Each path much be conatined within "quotes")

$TechToolsNew = Join-Path $PSScriptRoot "Tech Desk Tools.exe"

$AssetList = Join-Path $PSScriptRoot "Assets.txt"

## For Testing
#$AssetList = Join-Path $PSScriptRoot "AssetTest.txt"

$TechToolsName = "Tech Desk Tools.exe"

$ShortcutLocation = "c$\Users\Public\Desktop"

$ExeLocation = "c$\Program Files\"

$OldVersion = "Tech Desk Tools.exe" 

######

# Create a list of assets from the input text file
[string[]]$Assets = Get-Content -Path $AssetList

# Resolves name of files
$Script = Split-Path -Path $TechToolsNew -Leaf -Resolve
Echo "Application to be deployed: $Script"

# Analytics
$NumberOfAssets = $Assets.Length
$TechDeskToolsSuccess = 0
$TechDeskToolsPercent = 0

Echo ''
Echo 'List of assets:' $Assets `r
Echo 'Total number of assets:' $NumberOfAssets `r

$selection = Read-Host 'Do you still want to proceed? (y/n)'
if (!($selection.toLower() -eq "y")) {
   exit
}

$RemoveShortcut = 0 # 1 is yes 0 is no. You should not need to adjust this

Echo `r '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~' `r

ForEach($Computer in $Assets) 
{
    # Analytics
    Echo $Computer `r

    # Create new destination per device.
    $Destination = "\\$Computer\$ExeLocation"
    $PubDesktop = "\\$Computer\$ShortcutLocation"
    $ProcessName = $Script.TrimEnd(".exe")
    $OldProcess = "TechDeskTools3.exe"
    $CDrive = "\\$Computer\c$"

    # The Exe cannot be replaced if it's currently running so this command closes the program.
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            param($Script, $ProcessName, $OldProcess, $Computer)
            $ExeRun = Get-Process -Name $Script -ErrorAction Ignore
            $ProcessRun = Get-Process -Name $ProcessName -ErrorAction Ignore
	        $OldProcessRun = Get-Process -Name $OldProcess- -ErrorAction Ignore


        if ($ExeRun -or $ProcessRun -or $OldProcessRun) {
            Write-Host "Executable running on $env:COMPUTERNAME."
            $ExeRun | Stop-Process -Force -ErrorAction Ignore
            $ProcessRun | Stop-Process -Force -ErrorAction Ignore
            $OldProcessRun | Stop-Process -Force -ErrorAction Ignore

            Write-Host "Executable terminated on $env:COMPUTERNAME."
        } else {
            Write-Host "Executable is not running on $env:COMPUTERNAME."
        }
        } -ArgumentList $Script, $ProcessName, $OldProcess, $Computer


    # If there is a script with a duplicate name, delete it, and check if it was deleted successfully.
    $Dup = Test-Path -Path "$Destination\$Script" -PathType Leaf
    $Old = Test-Path -Path "$PubDesktop\$OldVersion" -PathType Leaf
    $Legacy = Test-Path -Path "$CDrive\Tech Desk Tools" -PathType Container

    if ($Dup -or $Old -or $Legacy)
    {
        
        Echo "Old Version Exists"
        Remove-Item "$Destination\$Script" -ErrorAction Ignore
        Remove-Item "$PubDesktop\$OldVersion" -ErrorAction Ignore
        Remove-Item "$CDrive\Tech Desk Tools" -Recurse -Force -ErrorAction Ignore

        $DupUp = Test-Path -Path "$Destination\$Script" -PathType Leaf
        $OldUp = Test-Path -Path "$PubDesktop\$OldVersion" -PathType Leaf
        $LegacyUp = Test-Path -Path "$CDrive\Tech Desk Tools" -PathType Container
        if ($DupUp -or $OldUp -or $LegacyUp)
        {

            Echo 'Could not remove old version'`r
            $OldNotRemovedComputers += $Computer + "`n"

        }
        else
        {

            Echo 'Successfully removed old version'`r

        }
        }
    else
    {
        Echo 'No Old Version Detected' `r
    }


    # Move the copy of the script to the new device, and check if it was successful.
    Echo "Copying Exe to Program Files"
    Copy-Item -Path "$TechToolsNew" -Destination $Destination
    if(Test-Path "$Destination\$Script" -PathType Leaf)
    {
        $TechDeskToolsSuccess++
        Unblock-File -Path "$Destination\$TechToolsName"
        Echo 'Successfully placed Tech Desk Tools at requested location.'
    }
    else
    {
        Echo 'Failed To Install Tech Desk Tools.'
        $FailInstallComputers += $Computer + "`n"
    }


    # Create Shortcut if it does not exist
    If(Test-Path "$PubDesktop\Tech Desk Tools.lnk" -PathType Leaf)
    {
    If($RemoveShortcut){
    Echo "Removing Existing Shortcut"
    Remove-Item "$PubDesktop\Tech Desk Tools.lnk" -ErrorAction Ignore
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$PubDesktop\Tech Desk Tools.lnk")
    $Shortcut.TargetPath = "$Destination\$TechToolsName"
    $Shortcut.Save()
    If(Test-Path "$PubDesktop\Tech Desk Tools.lnk" -PathType Leaf)
    {
    Echo "Shortcut created on Public Desktop" `r
    }}
    else {
     Echo "Shortcut Already Exists"
    }}
    else
    {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$PubDesktop\Tech Desk Tools.lnk")
    $Shortcut.TargetPath = "$Destination\$TechToolsName"
    $Shortcut.Save()
    If(Test-Path "$PubDesktop\Tech Desk Tools.lnk" -PathType Leaf)
    {
    Echo "Shortcut created on Public Desktop" `r
    }
    }
Echo '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~' `r
    }

    $TechDeskToolsPercent = ($TechDeskToolsSuccess / $NumberOfAssets) * 100


Echo '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~' `r

# Analytics
Echo 'Tech Desk Tools Success:' $TechDeskToolsSuccess `r
Echo 'Tech Desk Tools Percentage:' $TechDeskToolsPercent% `r

if ($FailInstallComputers -ne 0){
write-host "Install Failed on the following Computers"
write-host $FailInstallComputers
}

if ($OldNotRemovedComputers -ne 0){
write-host "Failed to remove an older version of Tech Desk Tools"
write-host $OldNotRemovedComputers
}

Read-Host 'Press any key to continue...';
exit
