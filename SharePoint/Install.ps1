#-------------------------------------------
#
#  Application: SharePoint Updates Script
#
#  Author: Zach Graf
#  Original Release: 03/01/2017
#
#  1.0 ZG 03/07/17 - Original Release
#  1.1 ZG 06/13/17 - Added support for CU installation
#  1.2 ZG 06/14/17 - June Security Updates/Hotfixes, AppVer set to 2
#  1.3 ZG 07/07/17 - Moved Reboot, corrected issues processing required patches based on $AppVer
#  1.4 ZG 07/12/17 - July Security Updates/Hotfixes, AppVer set to 3
#  1.5 ZG 08/09/17 - August Security Updates/Hotfixes, AppVer set to 4
#  1.6 ZG 09/13/17 - September Security Updates/Hotfixes, AppVer set to 5
#  1.7 ZG 10/10/17 - October Security Updates/Hotfixes, AppVer set to 6 
#  1.8 ZG 10/21/17 - AppVer set to 7 to re-run at the request of Middleware
#  1.9 ZG 10/23/17 - SharePoint 2016 updates/CUs added at the request of Middleware
#  2.0 ZG 10/25/17 - Added reboot if pending reboot is detected
#  2.1 ZG 11/01/17 - Modified to only update AppVer if updates are 100% successful to allow multiple runs
#  2.2 ZG 11/20/17 - November Security Updates/Hotfixes, AppVer set to 9
#  2.3 ZG 12/13/17 - December Security Updates/Hotfixes, AppVer set to 10
#  2.4 ZG 01/10/18 - January Security Updates/Hotfixes, AppVer set to 11
#  2.4 ZG 02/14/18 - February Security Updates/Hotfixes, AppVer set to 12
#  2.5 ZG 03/14/18 - March Security Updates/Hotfixes, AppVer set to 13
#  2.6 ZG 04/12/18 - April Security Updates/Hotfixes, AppVer set to 14
#  2.7 ZG 05/08/18 - May Security Updates/Hotfixes, AppVer set to 15
#  2.8 ZG 06/05/18 - Cumulative Updates added, AppVer set to 16
#  2.9 ZG 06/15/18 - June Security Updates/Hotfixes, AppVer set to 17
#  3.0 ZG 07/12/18 - July Security Updates/Hotfixes, AppVer set to 18
#  3.1 ZG 08/14/18 - August Security Updates/Hotfixes, AppVer set to 19
#  3.2 ZG 09/11/18 - September Security Updates/Hotfixes, AppVer set to 20
#  3.3 NA 10/09/18 - October Security Updates/Hotfixes, AppVer set to 21
#  3.4 NA 11/15/18 - November Security Updates/Hotfixes, AppVer set to 22
#  3.5 ZG 01/09/19 - January Security Updates/Hotfixes, AppVer set to 24
#  3.6 NA 02/14/19 - February Security Updates/Hotfixes, AppVer set to 25
#  3.7 NA 03/15/19 - March Security Updates/Hotfixes, AppVer set to 26
#  3.8 NA 04/11/19 - April Security Updates/Hotfixes, AppVer set to 27
#  3.9 NA 05/16/19 - May Security Updates/Hotfixes, AppVer set to 28
#  4.0 NA 06/12/19 - June Security Updates/Hotfixes and Cumulative Updates added, AppVer set to 29
#  4.1 PM 2019-JUN-27 - Updated the path from C:\windows\temp to E:\Share\Temp
#  4.2 PM 2019-JUL-09 - Added logic to create E:\Share\Temp if not exists and clean it after execution of Script
#  4.3 NA 07/11/19 - July Security Updates/Hotfixes, AppVer set to 30
#  4.4 NA 08/15/19 - August Security Updates/Hotfixes, AppVer set to 31
#  4.5 NA 09/11/19 - September Security Updates/Hotfixes, AppVer set to 32
#-------------------------------------------

Import-Module KSTools
$requiredVersion = ((Get-KSToolsVersion) -ge [version]"1.45.2")

if (!($requiredVersion)) { write-hoststatus "SharePoint Update Installer: installation error - KSTools version not met" "FAIL" ; exit 3 }

#-------------------------------------------
#  Script Constants
#-------------------------------------------

$AppFName = "Microsoft SharePoint Updates"
$AppVer = "32"
$AppKey = "Microsoft\SharePoint\Patch"
$WorkingDirPath = "E:\Share\Temp"
$RepoServer = $env:REPOSERVER
$ScriptRoot = $PSScriptRoot
if ($ScriptRoot -eq $null)
{
    $ScriptRoot = Split-Path $script:MyInvocation.MyCommand.Path
}

# Check if Source directory exists
if (Test-Path $WorkingDirPath)
{
	Write-HostStatus "E:\Share\Temp Exists" "INFO"
}
else{
	New-Item -ItemType Directory -Force -Path $WorkingDirPath
}
$errStatus = Get-ReturnCode -Value 'NoChanges'
$ConfigFile = "$ScriptRoot\files\Updates.ini"
$SP2010 = Test-Path "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0\WSS\InstalledProducts\9014*"
$SP2013 = Test-Path "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\15.0\WSS\InstalledProducts\9015*"
$SP2016 = Test-Path "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\16.0\WSS\InstalledProducts\9016*"
    if ($SP2010 -eq "True")
    {
        $SPVer = "2010"
        $CUList = Get-Ini "$ConfigFile" -Section "$SPVer" -Raw -Flatten
    }
    if ($SP2013 -eq "True")
    {
        $SPVer = "2013"
        $CUList = Get-Ini "$ConfigFile" -Section "$SPVer" -Raw -Flatten
    }
    if ($SP2016 -eq "True")
    {
        $SPVer = "2016"
        $CUList = Get-Ini "$ConfigFile" -Section "$SPVer" -Raw -Flatten
    }

#-------------------------------------------
#  Script Functions
#-------------------------------------------

function Get-ExitDesc($Value)
{
    switch ($Value)
    {
        "0"
        {
            return "ERROR_SUCCESS"
            break
        }
        "1"
        {
            return "Program Failure"
            break
        }
        "13"
        {
            return "ERROR_INVALID_DATA"
            break
        }
        "87"
        {
            return "ERROR_INVALID_PARAMETER"
            break
        }
        "1601"
        {
            return "ERROR_INSTALL_SERVICE_FAILURE"
            break
        }
        "1602"
        {
            return "ERROR_INSTALL_USEREXIT"
            break
        }
        "1603"
        {
            return "ERROR_INSTALL_FAILURE"
            break
        }
        "1604"
        {
            return "ERROR_INSTALL_SUSPEND"
            break
        }
        "1605"
        {
            return "ERROR_UNKNOWN_PRODUCT"
            break
        }
        "1606"
        {
            return "ERROR_UNKNOWN_FEATURE"
            break
        }
        "1607"
        {
            return "ERROR_UNKNOWN_COMPONENT"
            break
        }
        "1608"
        {
            return "ERROR_UNKNOWN_PROPERTY"
            break
        }
        "1609"
        {
            return "ERROR_INVALID_HANDLE_STATE"
            break
        }
        "1610"
        {
            return "ERROR_BAD_CONFIGURATION"
            break
        }
        "1611"
        {
            return "ERROR_INDEX_ABSENT"
            break
        }
        "1612"
        {
            return "ERROR_INSTALL_SOURCE_ABSENT"
            break
        }
        "1613"
        {
            return "ERROR_INSTALL_PACKAGE_VERSION"
            break
        }
        "1614"
        {
            return "ERROR_PRODUCT_UNINSTALLED"
            break
        }
        "1615"
        {
            return "ERROR_BAD_QUERY_SYNTAX"
            break
        }
        "1616"
        {
            return "ERROR_INVALID_FIELD"
            break
        }
        "1618"
        {
            return "ERROR_INSTALL_ALREADY_RUNNING"
            break
        }
        "1619"
        {
            return "ERROR_INSTALL_PACKAGE_OPEN_FAILED"
            break
        }
        "1620"
        {
            return "ERROR_INSTALL_PACKAGE_INVALID"
            break
        }
        "1621"
        {
            return "ERROR_INSTALL_UI_FAILURE"
            break
        }
        "1622"
        {
            return "ERROR_INSTALL_LOG_FAILURE"
            break
        }
        "1623"
        {
            return "ERROR_INSTALL_LANGUAGE_UNSUPPORTED"
            break
        }
        "1624"
        {
            return "ERROR_INSTALL_TRANSFORM_FAILURE"
            break
        }
        "1625"
        {
            return "ERROR_INSTALL_PACKAGE_REJECTED"
            break
        }
        "1626"
        {
            return "ERROR_FUNCTION_NOT_CALLED"
            break
        }
        "1627"
        {
            return "ERROR_FUNCTION_FAILED"
            break
        }
        "1628"
        {
            return "ERROR_INVALID_TABLE"
            break
        }
        "1629"
        {
            return "ERROR_DATATYPE_MISMATCH"
            break
        }
        "1630"
        {
            return "ERROR_UNSUPPORTED_TYPE"
            break
        }
        "1631"
        {
            return "ERROR_CREATE_FAILED"
            break
        }
        "1632"
        {
            return "ERROR_INSTALL_TEMP_UNWRITABLE"
            break
        }
        "1633"
        {
            return "ERROR_INSTALL_PLATFORM_UNSUPPORTED"
            break
        }
        "1634"
        {
            return "ERROR_INSTALL_NOTUSED"
            break
        }
        "1635"
        {
            return "ERROR_PATCH_PACKAGE_OPEN_FAILED"
            break
        }
        "1636"
        {
            return "ERROR_PATCH_PACKAGE_INVALID"
            break
        }
        "1637"
        {
            return "ERROR_PATCH_PACKAGE_UNSUPPORTED"
            break
        }
        "1638"
        {
            return "ERROR_PRODUCT_VERSION"
            break
        }
        "1639"
        {
            return "ERROR_INVALID_COMMAND_LINE"
            break
        }
        "1640"
        {
            return "ERROR_INSTALL_REMOTE_DISALLOWED"
            break
        }
        "1641"
        {
            return "ERROR_SUCCESS_REBOOT_INITIATED"
            break
        }
        "1642"
        {
            return "ERROR_PATCH_TARGET_NOT_FOUND"
            break
        }
        "3010"
        {
            return "ERROR_SUCCESS_REBOOT_REQUIRED"
            break
        }
        "17021"
        {
            return "Error: Creating temp folder"
            break
        }
        "17022"
        {
            return "Success: Reboot flag set"
            break
        }
        "17023"
        {
            return "Error: User cancelled installation"
            break
        }
        "17024"
        {
            return "Error: Creating folder failed"
            break
        }
        "17025"
        {
            return "Patch already installed"
            break
        }
        "17026"
        {
            return "Patch already installed to admin installation"
            break
        }
        "17027"
        {
            return "Installation source requires full file update"
            break
        }
        "17028"
        {
            return "No product installed for contained patch"
            break
        }
        "17029"
        {
            return "Patch failed to install"
            break
        }
        "17030"
        {
            return "Detection: Invalid CIF format"
            break
        }
        "17031"
        {
            return "Detection: Invalid baseline"
            break
        }
        "17034"
        {
            return "Error: Required patch does not apply to the machine"
            break
        }
        "17038"
        {
            return "You do not have sufficient privileges to complete this installation for all users of the machine. Log on as administrator and then retry this installation."
            break
        }
        "17044"
        {
            return "Installer was unable to run detection for this package."
            break
        }
        "17048"
        {
            return "This installation requires Windows Installer 3.1 or greater."
            break
        }
        "17301"
        {
            return "Error: General Detection error"
            break
        }
        "17302"
        {
            return "Error: Applying patch"
            break
        }
        "17303"
        {
            return "Error: Extracting file"
            break
        }
        "2359302"
        {
            return "ERROR_INSTALL_ALREADY_INSTALLED"
            break
        }
        "-2145124329"
        {
            return "ERROR_PATCH_NOT_APPLICABLE"
            break
        }
        default
        {
            return "NO_DESCRIPTION_AVAILABLE"
        }
    }
}
function Write-ExitCodeToEventLog($ExitCode, $Patch)
{
    $ExitDesc = Get-ExitDesc $ExitCode
    switch ($ExitDesc)
    {
        "ERROR_SUCCESS"
        {
            $LogType = "Information"
            $LogEventId = 100
            $LogMessage = "OK : Successfully installed $Patch"
            $SuccessIncrement = 1
            break
        }
        "ERROR_SUCCESS_REBOOT_REQUIRED"
        {
            $LogType = "Information"
            $LogEventId = 101
            $LogMessage = "OK : Successfully installed $Patch. Reboot required."
            $SuccessIncrement = 1
            break
        }
        "ERROR_SUCCESS_REBOOT_INITIATED"
        {
            $LogType = "Information"
            $LogEventId = 102
            $LogMessage = "OK : Successfully installed $Patch. Reboot initiated."
            $SuccessIncrement = 1
            break
        }
        "ERROR_INSTALL_ALREADY_INSTALLED"
        {
            $LogType = "Information"
            $LogEventId = 103
            $LogMessage = "OK : $Patch was already installed."
            $SuccessIncrement = 1
            break
        }
        "ERROR_PATCH_NOT_APPLICABLE"
        {
            $LogType = "Information"
            $LogEventId = 104
            $LogMessage = "OK : $Patch is not applicable to this server."
            $SuccessIncrement = 1
            break
        }
        "ERROR_PATCH_TARGET_NOT_FOUND"
        {
            $LogType = "Error"
            $LogEventId = 105
            $LogMessage = "FAIL : $Patch - Patch target not found."
            $SuccessIncrement = 1
            break
        }
        "Success: Reboot flag set"
        {
            $LogType = "Information"
            $LogEventId = 106
            $LogMessage = "OK : Successfully installed $Patch. Reboot required."
            $SuccessIncrement = 1
            break
        }
        "Patch already installed"
        {
            $LogType = "Information"
            $LogEventId = 107
            $LogMessage = "OK : $Patch - Patch already installed."
            $SuccessIncrement = 1
            break
        }
        "Patch already installed to admin installation"
        {
            $LogType = "Information"
            $LogEventId = 108
            $LogMessage = "OK : $Patch - Patch already installed to admin installation"
            $SuccessIncrement = 1
            break
        }
        "No product installed for contained patch"
        {
            $LogType = "Information"
            $LogEventId = 109
            $LogMessage = "OK : $Patch - No product installed for contained patch."
            $SuccessIncrement = 1
            break
        }
        "Error: Required patch does not apply to the machine"
        {
            $LogType = "Information"
            $LogEventId = 110
            $LogMessage = "OK : $Patch - Error: Required patch does not apply to the machine."
            $SuccessIncrement = 1
            break
        }
        default
        {
            $LogType = "Error"
            $LogEventId = 999
            $LogMessage = "FAIL : Unable to install $Patch. Result: $ExitCode - $ExitDesc"
            $SuccessIncrement = 0
            break
        }
    }
    New-EventLog -LogName KSTools -Source KPatch_SP -ErrorAction SilentlyContinue
    Write-EventLog -LogName KSTools -Source KPatch_SP -EventId $LogEventId -EntryType $LogType -Message $LogMessage
    return $SuccessIncrement
}

# abort installation if a reboot is pending 
if (Get-PendingReboot) 
{ 
    write-hoststatus "SharePoint Update Installer: installation error - pending reboot detected" "FAIL" ; exit 3010 
    Restart-Computer -Force
}

New-EventLog -LogName KSTools -Source KPatch_SP -ErrorAction SilentlyContinue
Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : Starting SharePoint patching script:" -Source KPatch_SP
$InstalledUpdates = Get-WmiObject -Query "Select HotfixID From Win32_QuickFixEngineering" | Select-Object -ExpandProperty HotfixID
Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : Found $($InstalledUpdates.Count) installed hotfixes" -Source KPatch_SP
$ErrorActionPreference = "Continue"

$stamp = (get-date -format "yyyy-MM-dd HH:mm")
write-hoststatus "SharePoint Update Installer: starting installation at $($stamp)" "INFO"

# ----------------------------------------------------------------------------
#  Subsystem Installations
# ----------------------------------------------------------------------------

# Save the current version so we can run multiple if this is more than one behind
$CurrentVersion = Get-RegistryValue "HKLM:\SOFTWARE\KPMG_US\APPVER\$AppKey" "Install"
if ($CurrentVersion -eq $null)
{
    $CurrentVersion = "0, 0"
}

$CurrentVersion = $CurrentVersion.Substring(0,$CurrentVersion.IndexOf(","))

if (Add-Subsystem $AppKey "Install" "SYSTEM" $AppVer) 
{
    Write-HostStatus "Installing SharePoint Updates" "INFO"

    for($x = [int]$CurrentVersion; $x -le [int]$AppVer; $x++)
    {
    if($x -gt $CurrentVersion)
    {
            $IniSection = Get-Ini "$ConfigFile" -Section "$x" -Raw -Flatten
    
            foreach ($Line in $IniSection)
            {
                $Updates += "$Line"
            }
 
        $UpdateList = $Updates.Split("|")
     }
     }

        $RebootRequired = $false
        $SuccessfulPatchCount = 0
        $TotalPatchCount = 0

        #----------------------------
        # Updates and HotFixes
        #----------------------------

        foreach ($Update in $UpdateList)
        {
            $DiskFreeSpace = Get-WmiObject -Query "Select DeviceID, FreeSpace FROM Win32_LogicalDisk Where DeviceID='E:'" | Select-Object -ExpandProperty FreeSpace
            if ($Update)
            {
                $TotalPatchCount += 1
                if ($Update -match "KB\d*")
                {
                    $KB = $Matches[0]
                }
                if ($InstalledUpdates -contains $KB)
                {
                    $SuccessfulPatchCount += 1
                    Add-Event -Type "INFORMATION" -EventID 1 -Message "OK : $KB is already installed (QFE Check)" -Source KPatch_SP
                }
                else
                {
                    if (Test-Path "$ScriptRoot\files\$Update.msu")
                    {
                        $UpdateFile = "$Update.msu"
                        $LogFile = "$Update.evt"
                    }
                    if (Test-Path "$ScriptRoot\files\$Update.exe")
                    {
                        $UpdateFile = "$Update.exe"
                        $LogFile = "$Update.log"
                    }
                    if ($UpdateFile)
                    {
                        $FileSize = ((Get-Item "$ScriptRoot\files\$UpdateFile").Length) * 2
                        if ($FileSize -ge $DiskFreeSpace)
                        {
                            $LogMessage = "FAIL : $ScriptRoot\files\$UpdateFile is larger than the available disk space, skipping patch"
                            Add-Event -Type "ERROR" -EventID 999 -Message $LogMessage -Source KPatch_SP
                        }
                        else
                        {
                            Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : Copying $UpdateFile to E:\Share\Temp" -Source KPatch_SP
                            Copy-File "$UpdateFile" "$ScriptRoot\files" "E:\Share\Temp"
                            if (Test-Path "E:\Share\Temp\$UpdateFile")
                            {
                                Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : Successfully copied $UpdateFile to E:\Share\Temp" -Source KPatch_SP
                            
                                $PatchCommand = "E:\Share\Temp\$UpdateFile /quiet /norestart /log:E:\Share\Temp\$LogFile"
                                Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : Installing $UpdateFile" -Source KPatch_SP
                                Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : Command: $PatchCommand" -Source KPatch_SP
                                $patchStatus = Run $PatchCommand True True
                            
                                Remove-File "E:\Share\Temp\$Update.*"
                                $SuccessfulPatchCount += Write-ExitCodeToEventLog -ExitCode $patchStatus -Patch $Update
                                If ((Get-ExitDesc -Value $patchStatus) -eq "ERROR_SUCCESS_REBOOT_REQUIRED")
                                {
                                $RebootRequired = $true
                                }
                            }
                            else
                            {
                                Add-Event -Type "ERROR" -EventID 10 -Message "FAIL : Failed to copy $UpdateFile to E:\Share\Temp" -Source KPatch_SP
                            }
                        }
                    }
                    else
                    {
                        Add-Event -Type "ERROR" -EventID 999 -Message "FAIL : Unable to find file for $Update on $ScriptRoot\files" -Source KPatch_SP
                    }
                }
            }
        }
    If($SuccessfulPatchCount -eq $TotalPatchCount)
    {
        Remove-Item "E:\Share\Temp\*.cab"
        Remove-Item "E:\Share\Temp\*_MSPLOG*.log"
		Remove-Item "E:\Share\Temp\*.exe"
        Close-Subsystem $AppKey "Install" SYSTEM $AppVer
    }
    else
    {
        Add-Event -Type "ERROR" -EventID 999 -Message "Not all updates were installed succssfully, please re-run." -Source KPatch_SP
    }
}

else
{
    Add-Event -Type "INFORMATION" -EventID 1 -Message "OK : Patch Revision is greater or equal to the repo patch revision, no patching necessary." -Source KPatch_SP
}

#----------------------------
# Cumulative Updates
#----------------------------

$AppFName = "Microsoft SharePoint Cumulative Updates"
$CUAppKey = "Microsoft\SharePoint\CU"
if (Add-Subsystem $CUAppKey "Install" "SYSTEM" $AppVer) 
{
    $CUSect = Get-Ini "$ConfigFile" -Section "$SPVer" -Raw -Flatten
        foreach ($entry in $CUSect)
        {
            $Updates+= "$entry"
        }
    $CUList = $entry.Split("[ | ]")
    $TempDir = Test-Path "E:\Share\Temp"
    if(!($TempDir))
    {
        New-Item E:\Share\Temp -ItemType Directory | Out-Null
    }

    Write-Host $CUSect
    Write-Host $CUList

    foreach ($CU in $CUList)
    {
        $DiskFreeSpace = Get-WmiObject -Query "Select DeviceID, FreeSpace FROM Win32_LogicalDisk Where DeviceID='E:'" | Select-Object -ExpandProperty FreeSpace
        
        Write-Host $KB
        Write-Host $SPVer

            if ($CU -match "KB\d*")
            {
                $KB = $Matches[0]
            }
            if ($InstalledUpdates -contains $KB)
            {
                $SuccessfulPatchCount += 1
                Add-Event -Type "INFORMATION" -EventID 1 -Message "OK : CU: $KB is already installed (QFE Check)" -Source KPatch_SP
            }
            else
            {
                Write-Host "Copying $CUList to E:\Share\Temp"
                Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : CU: Copying $CUList to E:\Share\Temp" -Source KPatch_SP
                Copy-File "*.*" "$ScriptRoot\files\CU\$SPVer" "E:\Share\Temp"
            }
            
            if (Test-Path "$ScriptRoot\files\CU\$SPVer\$CU.msu")
            {
                $CUFile = "$CU.msu"
                $CULog = "$CU.evt"
            }
            if (Test-Path "$ScriptRoot\files\CU\$SPVer\$CU.exe")
            {
                $CUFile = "$CU.exe"
                $CULog = "$CU.log"
            }

        if (Test-Path "$ScriptRoot\files\CU\$SPVer\$CUFile")
        {
                if (Test-Path "E:\Share\Temp\$CUFile")
                {
                    Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : CU: Successfully copied $CUFile to E:\Share\Temp" -Source KPatch_SP
                    $PatchCommand = "E:\Share\Temp\$CUFile /quiet /norestart /log:E:\Share\Temp\$CULog"
                    Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : CU: Installing $CUFile" -Source KPatch_SP
                    Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : CU: Command: $PatchCommand" -Source KPatch_SP
                    If($CUFile)
                    {
                        $TotalPatchCount += 1
                        $patchStatus = Run $PatchCommand True True
                        $SuccessfulPatchCount += Write-ExitCodeToEventLog -ExitCode $patchStatus -Patch $CUFile
                        If ((Get-ExitDesc -Value $patchStatus) -eq "ERROR_SUCCESS_REBOOT_REQUIRED")
                        {
                            $RebootRequired = $true
                        }
                        Start-Sleep -Seconds 10
                        Remove-Item "E:\Share\Temp\$CU.*"
                    }
                }
                else
                {
                    Add-Event -Type "ERROR" -EventID 10 -Message "FAIL : CU: Failed to copy $CUFile to E:\Share\Temp" -Source KPatch_SP
                }         
        }
        else
        {
            Add-Event -Type "ERROR" -EventID 999 -Message "FAIL : CU: Unable to find file for $CUFile on $ScriptRoot\files\CU\$SPVer" -Source KPatch_SP
        }
    }

    If($SuccessfulPatchCount -eq $TotalPatchCount)
    {
        Remove-Item "E:\Share\Temp\*.cab"
        Remove-Item "E:\Share\Temp\*_MSPLOG*.log"
		Remove-Item "E:\Share\Temp\*.exe"
        Close-Subsystem $CUAppKey "Install" SYSTEM $AppVer
    }
    else
    {
        Add-Event -Type "ERROR" -EventID 999 -Message "Not all updates were installed succssfully, please re-run." -Source KPatch_SP
    }
}
else
{
Add-Event -Type "INFORMATION" -EventID 1 -Message "OK : CU Revision is greater or equal to the repo patch revision, no patching necessary." -Source KPatch_SP
}

if($TotalPatchCount -gt 0)
    {
        $LogMessage = "$SuccessfulPatchCount of $TotalPatchCount patches didn't encounter an error (they were installed successfully, already installed, or not applicable)."
        Add-Event -Type "INFORMATION" -EventID 1 -Message $LogMessage -Source KPatch_SP 
        $LogMessage = "INFO : {0:P2} successful" -f $($SuccessfulPatchCount/$TotalPatchCount)
        Add-Event -Type "INFORMATION" -EventID 1 -Message ($LogMessage -replace ' %','%') -Source KPatch_SP
            
        Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : Patching completed." -Source KPatch_SP

        $stamp = (get-date -format "yyyy-MM-dd HH:mm")
        write-hoststatus "SharePoint Update Installer: installation completed at $($stamp)" "INFO"

        $RebootRequired = Get-PendingReboot

        If ($RebootRequired)
        {
            Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO : One or more patches requires a reboot to complete." -Source KPatch_SP
            $errStatus = 3010
            write-hoststatus "SharePoint Update Installer: Restarting Server" "INFO"
            Restart-Computer -Force
        }

        return $errStatus
    }
    else 
    {
        Add-Event -Type "INFORMATION" -EventID 1 -Message "INFO: There are no patches to apply at this time" -Source KPatch_SP
    }

return $errStatus
exit $errStatus

# SIG # Begin signature block
# MIIXkQYJKoZIhvcNAQcCoIIXgjCCF34CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUj7lE9xV7bytPRDSYUAKIOI6A
# M8qgghKxMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggTJMIIDsaADAgECAhAiXW95cB8gTnvQ6hqi4X7jMA0GCSqGSIb3DQEBCwUAMIGE
# MQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xHzAd
# BgNVBAsTFlN5bWFudGVjIFRydXN0IE5ldHdvcmsxNTAzBgNVBAMTLFN5bWFudGVj
# IENsYXNzIDMgU0hBMjU2IENvZGUgU2lnbmluZyBDQSAtIEcyMB4XDTE3MDYyNzAw
# MDAwMFoXDTIwMDkyNDIzNTk1OVowejELMAkGA1UEBhMCVVMxEzARBgNVBAgMCk5l
# dyBKZXJzZXkxETAPBgNVBAcMCE1vbnR2YWxlMREwDwYDVQQKDAhLUE1HIExMUDEd
# MBsGA1UECwwUTW9udHZhbGUgRGF0YSBDZW50ZXIxETAPBgNVBAMMCEtQTUcgTExQ
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAxjb4AQUf8OvAcZdFZUNb
# R/9wlX1mHJ7lrY5P07x7/b8tvk4FrcAsxRzCaqEQQtllgxFgc+ZtChH1oTPO7G28
# vv+PSd5mYiZ6AxprgnLW2oT5lss8OdKxc2Hs1EJGyACPdgNXKnRb4DjN9gaYUey8
# bbX8UgYBjergjwYO8euDOkjXNTGk47BOhs4kVufkt7s2iC94lRQ09xJhXVtFGqL9
# obRt1Z1PBCVxPyupScooTjIK4wTKy2KPEuKlbClyNquOX1ceYZp45498OyIhyD57
# tYedtYqtU2WyAO69tjVdysRuyENS7DEVw+zHqgD+fA/7OzzuPAzgNIaNem3zseTf
# CQIDAQABo4IBPjCCATowCQYDVR0TBAIwADAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwYQYDVR0gBFowWDBWBgZngQwBBAEwTDAjBggrBgEFBQcC
# ARYXaHR0cHM6Ly9kLnN5bWNiLmNvbS9jcHMwJQYIKwYBBQUHAgIwGQwXaHR0cHM6
# Ly9kLnN5bWNiLmNvbS9ycGEwHwYDVR0jBBgwFoAU1MAGIknrOUvdk+JcobhHdgly
# A1gwKwYDVR0fBCQwIjAgoB6gHIYaaHR0cDovL3JiLnN5bWNiLmNvbS9yYi5jcmww
# VwYIKwYBBQUHAQEESzBJMB8GCCsGAQUFBzABhhNodHRwOi8vcmIuc3ltY2QuY29t
# MCYGCCsGAQUFBzAChhpodHRwOi8vcmIuc3ltY2IuY29tL3JiLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAQEAscetsKMcvM1RcqxvfAitGIRf+6SWC9CPCLnaOCFXxxZBnPRl
# T/8t3JXNzUki4YatHuVV9g6+jfBF68lntWSGHt79yxauauIfMDzaA2Z3VcIFj+MS
# DVXRTVL4X3LFDpleKobCzr/L2clzf6UvS2w+GAsMaUTq+7LNyxmk+YZ04S/V/Q+u
# 336TIdsQeT7Q9hO7Z80Q9Uo1zIRZUoWl6eaJlDFae7xbSwFJVfDZGFcnNvBiy8AV
# 56pYblwWxu2t6AQYtIOwuAjSrpXleMz46mbbO7ES7LmtWIL8MrlOTBG9or9TMeGl
# AGz9Jzx7/aeBj79a6fdqEcMsfo4FnmKuPodZbDCCBUcwggQvoAMCAQICEHwbNTVK
# 59t050FfEWnKa6gwDQYJKoZIhvcNAQELBQAwgb0xCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5WZXJpU2lnbiwgSW5jLjEfMB0GA1UECxMWVmVyaVNpZ24gVHJ1c3QgTmV0
# d29yazE6MDgGA1UECxMxKGMpIDIwMDggVmVyaVNpZ24sIEluYy4gLSBGb3IgYXV0
# aG9yaXplZCB1c2Ugb25seTE4MDYGA1UEAxMvVmVyaVNpZ24gVW5pdmVyc2FsIFJv
# b3QgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkwHhcNMTQwNzIyMDAwMDAwWhcNMjQw
# NzIxMjM1OTU5WjCBhDELMAkGA1UEBhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENv
# cnBvcmF0aW9uMR8wHQYDVQQLExZTeW1hbnRlYyBUcnVzdCBOZXR3b3JrMTUwMwYD
# VQQDEyxTeW1hbnRlYyBDbGFzcyAzIFNIQTI1NiBDb2RlIFNpZ25pbmcgQ0EgLSBH
# MjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANeVQ9Tc32euOftSpLYm
# MQRw6beOWyq6N2k1lY+7wDDnhthzu9/r0XY/ilaO6y1L8FcYTrGNpTPTC3Uj1Wp5
# J92j0/cOh2W13q0c8fU1tCJRryKhwV1LkH/AWU6rnXmpAtceSbE7TYf+wnirv+9S
# rpyvCNk55ZpRPmlfMBBOcWNsWOHwIDMbD3S+W8sS4duMxICUcrv2RZqewSUL+6Mc
# ntimCXBx7MBHTI99w94Zzj7uBHKOF9P/8LIFMhlM07Acn/6leCBCcEGwJoxvAMg6
# ABFBekGwp4qRBKCZePR3tPNgKuZsUAS3FGD/DVH0qIuE/iHaXF599Sl5T7BEdG9t
# cv8CAwEAAaOCAXgwggF0MC4GCCsGAQUFBwEBBCIwIDAeBggrBgEFBQcwAYYSaHR0
# cDovL3Muc3ltY2QuY29tMBIGA1UdEwEB/wQIMAYBAf8CAQAwZgYDVR0gBF8wXTBb
# BgtghkgBhvhFAQcXAzBMMCMGCCsGAQUFBwIBFhdodHRwczovL2Quc3ltY2IuY29t
# L2NwczAlBggrBgEFBQcCAjAZGhdodHRwczovL2Quc3ltY2IuY29tL3JwYTA2BgNV
# HR8ELzAtMCugKaAnhiVodHRwOi8vcy5zeW1jYi5jb20vdW5pdmVyc2FsLXJvb3Qu
# Y3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIBBjApBgNVHREE
# IjAgpB4wHDEaMBgGA1UEAxMRU3ltYW50ZWNQS0ktMS03MjQwHQYDVR0OBBYEFNTA
# BiJJ6zlL3ZPiXKG4R3YJcgNYMB8GA1UdIwQYMBaAFLZ3+mlIR59TEtXC6gcydgfR
# lwcZMA0GCSqGSIb3DQEBCwUAA4IBAQB/68qn6ot2Qus+jiBUMOO3udz6SD4Wxw9F
# lRDNJ4ajZvMC7XH4qsJVl5Fwg/lSflJpPMnx4JRGgBi7odSkVqbzHQCR1YbzSIfg
# y8Q0aCBetMv5Be2cr3BTJ7noPn5RoGlxi9xR7YA6JTKfRK9uQyjTIXW7l9iLi4z+
# qQRGBIX3FZxLEY3ELBf+1W5/muJWkvGWs60t+fTf2omZzrI4RMD3R3vKJbn6Kmgz
# m1By3qif1M0sCzS9izB4QOCNjicbkG8avggVgV3rL+JR51EeyXgp5x5lvzjvAUoB
# CSQOFsQUecFBNzTQPZFSlJ3haO8I8OJpnGdukAsak3HUJgLDwFojMYIESjCCBEYC
# AQEwgZkwgYQxCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEfMB0GA1UECxMWU3ltYW50ZWMgVHJ1c3QgTmV0d29yazE1MDMGA1UEAxMs
# U3ltYW50ZWMgQ2xhc3MgMyBTSEEyNTYgQ29kZSBTaWduaW5nIENBIC0gRzICECJd
# b3lwHyBOe9DqGqLhfuMwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKA
# AKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFJtw9eCCFWHBDhnmVJlvpEkV
# xp82MA0GCSqGSIb3DQEBAQUABIIBAGErgtCtet236/2ra7R2d6jyo4i6Zolb0pJ7
# UAiYXK2CFaN72N391xSnvuUczNDg50sZmt9E0oyvVtPePa5uvE+4hLqiSMWoLG0W
# Ghk/NwXDlJ8pdJCQEQ07o1R6IR2J+XvO2XGunp0y0b0V77yY4swnRu/WeBWKEyqB
# UM6kiip2sbCeks64AzlVCbwzafL/W/9TszoFcEX+o+sMOGaci5m1a5Dm46onnMrX
# uxLhGDbFb6oNYgrQjbhgQFC9LV2kChoygY3tbaKPfXH8Ktu8WcVMTx9dxgog8tSY
# QibPe9e5rr/geSHEYnivomSM7ExmOjdQqyY/CYVjAKXr45bSr8OhggILMIICBwYJ
# KoZIhvcNAQkGMYIB+DCCAfQCAQEwcjBeMQswCQYDVQQGEwJVUzEdMBsGA1UEChMU
# U3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUgU3Rh
# bXBpbmcgU2VydmljZXMgQ0EgLSBHMgIQDs/0OMj+vzVuBNhqmBsaUDAJBgUrDgMC
# GgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MTkwOTEyMTUwMzIxWjAjBgkqhkiG9w0BCQQxFgQUTDQt2OHPbYoyFqm/p+1HAL19
# jK0wDQYJKoZIhvcNAQEBBQAEggEALFk0L0M12yD7Iz8SAQswP7G+vTMDud96Zs7q
# TZCpC9jYU/pPXBeftUZOH/jT+hIeRYG91rQJDuGkOIW6dxLaxspwU9pu1pRB10O/
# DSMfz9yzsjsBdcZLndS6mrNHomFJYL6R4fF1uXhyT986tCh0Eo4KI+qtB/gC1CsC
# f/t6Vk/QHmMTQAd3MgtUolSlp0VFdNk3sAXHW/5CuAZ9vnWm7xJM+oD77tbXY7pP
# Q+OeAia58eJEjVDh181AH2DdF4g+K/3HSTaL7arD/Pn2t4VgSdjw7L7K1EKtRWKh
# rph+QB1lbJCPFq1SiL23hZpgRim/DHH8aLS48vlgVEzM1Fqvfg==
# SIG # End signature block
