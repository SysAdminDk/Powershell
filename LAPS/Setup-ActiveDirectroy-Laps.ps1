<#

    DISCLAIMER:

    THE SCRIPT IS PROVIDED AS-IS, WITHOUT WARRANTY OF ANY KIND. USE AT YOUR OWN RISK.

    By running this script, you acknowledge that you have read and understood the disclaimer, and you agree to assume
    all responsibility for any failures, damages, or issues that may arise as a result of executing this script.

#>
<#
    .SYNOPSIS
     Extend Active Directory to support both Legacy LAPS and Windows LAPS
    
    .DESCRIPTION
     To extend Active Directory the user running the script needs to be a member of both Enterprise and Schema Admins, please remember to remove the membership again.

     Requires internet access to Download the Legacy LAPS files, unless they are provided in the directory provided on the LAPSFiles parameter.
     
     There will be created 4 GPO's and 2 WMI filters to support both Windows and Legacy LAPS.

         1. GPO "Manage LAPS Version"
            Will handle the LAPS installation on ALL clients where the GPO is linked, and will make sure that the best LAPS option is selected.

         2. GPO Legacy LAPS Settings"
            Will configure clients to save Local Administrator password in the Legacy LAPS properties.

         3. GPO "Windows LAPS Settings (Active Directory)"
            Will configure clients to save Local Administrator password in the Active Directory Windows LAPS properties.

         4. GPO "Windows LAPS Settings (Azure Active Directory)"
            Will configure clients to save Local Administrator password in the Azure Active Directory Windows LAPS properties.


    .NOTES
     Needs Enterprise and Schema Admins, please remember to remove the membership again   

    .LINK
     https://learn.microsoft.com/en-us/windows-server/identity/laps/laps-overview
     https://www.microsoft.com/en-us/download/details.aspx?id=46899


    .PARAMETER GPOPrefix 
     Defines the prefix of the GPO's and WMI filters that will be created.
     Defaults to Domain


    .PARAMETER LAPSFiles
     If the server where the script is executed, the Legacy LAPS file(s) need to be provided at the path specified.
     Can be downloaded at the following URI : https://www.microsoft.com/en-us/download/details.aspx?id=46899

     The x86 file is optional, and only required if there is member machines that requires it.


    .EXAMPLE
     Setup-ActiveDirectroy-Laps.ps1 -GPOPrefix "MyDomain"
     Setup-ActiveDirectroy-Laps.ps1 -GPOPrefix "MyDomain" -LAPSFiles "C:\_Install\LAPS"

#>
#requires -RunAsAdministrator
#requires -Modules ActiveDirectory, GroupPolicy

[CmdletBinding()]
Param(
  [Parameter(ValueFromPipelineByPropertyName=$true,Position=0)][string]$GPOPrefix = "LAPS",
  [Parameter(ValueFromPipelineByPropertyName=$true)][string]$LAPSFiles = $PSScriptRoot
)


##################################################################################
# DISCLAIMER [ Start ]
##################################################################################

Clear-Host
$CurrentColor = $host.UI.RawUI.ForegroundColor
$host.UI.RawUI.ForegroundColor = "red"
Write-Output "*******************************************************************************************************************"
Write-Output ""
Write-Output "DISCLAIMER: "
Write-Output ""
Write-Output ""
Write-Output "THE FOLLOWING POWERSHELL SCRIPT IS PROVIDED AS-IS, WITHOUT WARRANTY OF ANY KIND. USE AT YOUR OWN RISK."
Write-Output ""
Write-Output ""
Write-Output "By running this script, you acknowledge that you have read and understood the disclaimer, and you agree to assume"
Write-Output "all responsibility for any failures, damages, or issues that may arise as a result of executing this script."
Write-Output ""
Write-Output ""
Write-Output "Please note the following"
Write-Output ""
Write-Output "1. The script makes changes to the Active Directory Schema, which can have significant impacts on your environment."
Write-Output "2. It is strongly recommended that you run this script in a test or lab environment before executing it in production."
Write-Output "3. Performing changes in production without proper testing can lead to data loss, service disruptions,"
Write-Output "   or other unintended consequences."
Write-Output ""
Write-Output "Take appropriate precautions and ensure you have a backup of your Active Directory before running this script."
Write-Output ""
Write-Output ""
Write-Output "*******************************************************************************************************************"
$host.UI.RawUI.ForegroundColor = $CurrentColor

$title = "Update Active Directory"
$message = "Do you want to run this script and prepare you domain for LAPS ?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes Install."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Just quit"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

if (($host.ui.PromptForChoice($title, $message, $options, 0)) -eq 1) {
    Write-Output "Quit"
    break
} else {
    Clear-Host
    Write-Output "Verifying Prerequisites"
}

##################################################################################
# DISCLAIMER [ End ]
##################################################################################


##################################################################################
# Prerequisites check [ Start ]
##################################################################################

# --------------------------------------------------------------------------------
# Check the GPO prefix
# --------------------------------------------------------------------------------
Write-Output "Prerequisites : Verify selected GPO prifix is valid"
if ($GPOPrefix -Notmatch "^[a-zA-Z]+$") { # Only allow letters
    throw "The GPO name prefix contains invalid characters, unable to continue"
}
else {
    Write-Output "Prerequisites : GPO prefix is valid, will continue"
}


# --------------------------------------------------------------------------------
# Verify required files are avalible
# --------------------------------------------------------------------------------
Write-Output "Prerequisites : Verify required GPO files are avalible"
if ( (!(Test-Path -Path "$PSScriptRoot\GPO")) -AND (Test-Path -Path "$PSScriptRoot\GPO.zip") ) {
    Expand-Archive -Path "$PSScriptRoot\GPO.zip" -DestinationPath $PSScriptRoot
}
if ( (!(Test-Path -Path "$PSScriptRoot\GPO")) -AND (((Get-ChildItem -Path $PSScriptRoot -Recurse).FullName).Count -ne 59) ) {
    Throw "Required files and folders missing, unable to continue"
}
if (!(Test-Path -Path "$PSScriptRoot\GPO\Policy Dependencies\Manage-Laps-Version.ps1")) {
    Throw "Manage LAPS version script not found, unable to continue"
}


# --------------------------------------------------------------------------------
# Read Registry for Installed Windows version
# --------------------------------------------------------------------------------
Write-Output "Prerequisites : Get Windows version and Architecture"
$WindowsVersion = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Select-Object ProductName, CurrentBuild, UBR


# --------------------------------------------------------------------------------
# Check for Windows version
# --------------------------------------------------------------------------------
Write-Output "Prerequisites : Detect which version of LAPS current server can support"
if ( ($WindowsVersion.CurrentMajorVersionNumber -ne "10") -and ($WindowsVersion.ProductName -notlike "*Server*") ) {
    throw "Not running on a supported Windows Server version, unable to continue"
} else {
    # Ensure April patch is installed (KB5025230).
    Switch ($WindowsVersion) {
        {($_.CurrentBuild -eq 17763) -and ($_.UBR -ge 4252)} { $Laps = "Windows" }
        {($_.CurrentBuild -eq 20348) -and ($_.UBR -ge 1668)} { $Laps = "Windows" }
        {($_.CurrentBuild -gt 20348)                       } { $Laps = "Windows" }

        default {
            # --
            # Quit if April patch not installed (KB5025230)
            # --
            Throw "Missing KB5025230 and Windows LAPS unsupported, unable to continue"
            break
        }
    }
}


# --------------------------------------------------------------------------------
#  Quit if running on x86.
# --------------------------------------------------------------------------------
if ($($env:PROCESSOR_ARCHITECTURE) -ne "AMD64") {
    Throw "Missing KB5025230 and Windows LAPS supported, unable to continue"
}


# --------------------------------------------------------------------------------
# Connect to Domain and Make sure all AD commands use the PDC
# --------------------------------------------------------------------------------
Write-Output "Prerequisites : Connect to PDC, and setting default server for *AD* commands"
$CurrentDomain = Get-ADDomain
if ($null -eq $($CurrentDomain.PDCEmulator)) {
    throw "Failed to connect to Active Directory, unable to continue"
} else {
    $PSDefaultParameterValues = @{
        "*AD*:Server" = $CurrentDomain.PDCEmulator
    }
}


# --------------------------------------------------------------------------------
# Verify Current user is Domain User and member of Schema and Enterprise Admins.
# --------------------------------------------------------------------------------
Write-Output "Prerequisites : Verify Current user is a member of Enterprise and Schema Admins"
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
try {
    $ADUser = Get-AdUser -Identity ($CurrentUser -split("\\"))[-1]
}
Catch {
    throw "Script is running in local user context, unable to continue"
}

if (!(Get-ADGroupMember -Identity "Enterprise Admins" | Where-Object {$_.distinguishedName -eq $ADUser.DistinguishedName})) { 
    throw "User is NOT a member of Enterprise Admins, unable to continue"
}

if (!(Get-ADGroupMember -Identity "Schema Admins" | Where-Object {$_.distinguishedName -eq $ADUser.DistinguishedName})) { 
    throw "User is NOT a member of Schema Admins, unable to continue"
}


# --------------------------------------------------------------------------------
# Verify Domain Functional Level
# --------------------------------------------------------------------------------
If ($CurrentDomain.DomainMode -ne "Windows2016Domain") {
    Write-Warning "To fully support Windows LAPS, the Domain functional level needs to be 2016, please upgrade prior to configuring Windows LAPS Password encryption"
}


##################################################################################
# Prerequisites check [ End ]
##################################################################################


##################################################################################
# Main script [ Start ]
##################################################################################


# --------------------------------------------------------------------------------
# Download Legacy LAPS installation files.
# --------------------------------------------------------------------------------
Write-Output "Main : Download LAPS install files"
if (!(Test-Path -Path "$LapsFiles\LAPS.x64.msi")) {
    try {
        Invoke-WebRequest -Uri "https://download.microsoft.com/download/C/7/A/C7AAD914-A8A6-4904-88A1-29E657445D03/LAPS.x64.msi" -OutFile "$LapsFiles\LAPS.x64.msi"
    } catch {
        Throw "Failed to download the LAPS.x64.msi, unable to continue"
    }
}
if (!(Test-Path -Path "$LapsFiles\LAPS.x86.msi")) {
    try {
        Invoke-WebRequest -Uri "https://download.microsoft.com/download/C/7/A/C7AAD914-A8A6-4904-88A1-29E657445D03/LAPS.x86.msi" -OutFile "$LapsFiles\LAPS.x86.msi"
    } catch {
        Write-Warning "Main : Failed to download the LAPS.x86.msi, please copy it to sysvol if there is x86 machines in the Company where you need to support Legacy LAPS"
    }
}


# --------------------------------------------------------------------------------
# Check if Legacy LAPS is installed
# --------------------------------------------------------------------------------
$LegacyLAPS = Test-Path -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{97E2CA7B-B657-4FF7-A6DB-30ECC73E1E28}"
if (!($LegacyLAPS)) {
    Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/i `"$LapsFiles\LAPS.x64.msi`" ADDLOCAL=Management.PS,Management.ADMX ALLUSERS=1 /qn /quiet" -Wait

    # Make sure we wait for installation
    Start-Sleep -Seconds 5
}


# --------------------------------------------------------------------------------
# Load Legacy LAPS powershell module.
# --------------------------------------------------------------------------------
If ( ($LegacyLAPS) -AND (Test-Path -Path "$(($env:PSModulePath -split(";"))[-1])\admpwd.ps\AdmPwd.PS.dll") ) {
    Import-Module "$(($env:PSModulePath -split(";"))[-1])\admpwd.ps\AdmPwd.PS.dll"
} else {
    Write-Verbose "Main : Legacy Laps is installed, missing the Powershell Module (Update Install)"
    Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/i {97E2CA7B-B657-4FF7-A6DB-30ECC73E1E28} ADDLOCAL=Management.PS /quiet" -Wait

    if (Test-Path -Path "$(($env:PSModulePath -split(";"))[-1])\admpwd.ps\AdmPwd.PS.dll") {
        Import-Module "$(($env:PSModulePath -split(";"))[-1])\admpwd.ps\AdmPwd.PS.dll" -Force
    } else {
        Throw "Failed to install Legacy LAPS powershell module, unable to continue"
    }
}


# --------------------------------------------------------------------------------
# Verify Legacy LAPS Policy Definitions is installed.
# --------------------------------------------------------------------------------
$PolicyDefinitions = @("AdmPwd.admx","en-US\AdmPwd.adml")
If (!(Test-Path -Path "C:\Windows\PolicyDefinitions\$($PolicyDefinitions[0])")) {
    Write-Verbose "Main : Legacy Laps is installed, missing the Policy Definitions (Update Install)"
    Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/i {97E2CA7B-B657-4FF7-A6DB-30ECC73E1E28} ADDLOCAL=Management.ADMX /quiet" -Wait
}


# --------------------------------------------------------------------------------
# Add Windows LAPS Policy Definitions.
# --------------------------------------------------------------------------------
if ($Laps -eq "Windows") {
    $PolicyDefinitions += @("LAPS.admx","en-US\LAPS.adml")
}


# --------------------------------------------------------------------------------
# Copy selected Policy Definitions
# --------------------------------------------------------------------------------
Write-Verbose "Main : Copy Local Policy Definitions SYSVOL"
if (!(Test-Path "\\$($CurrentDomain.DNSRoot)\SYSVOL\$($CurrentDomain.DNSRoot)\Policies\PolicyDefinitions")) {
    Write-Warning "Main : Central Store for Group Policy Administrative Templates is missing, please create the central store and copy the Templates"
} else {
    Foreach ($File in $PolicyDefinitions) {
        $SourceFilePath = "C:\Windows\PolicyDefinitions\$File"
        $TargetFilePath = "\\$($CurrentDomain.DNSRoot)\SYSVOL\$($CurrentDomain.DNSRoot)\Policies\PolicyDefinitions\$File"
    
        If ( (Test-Path -Path $SourceFilePath) -AND (!(Test-Path -Path $TargetFilePath)) ) {
            Copy-Item -Path $SourceFilePath -Destination $TargetFilePath
        }
        if (!($TargetFilePath)) {
            Write-Warning "Main : Unable to copy $File to SYSVOL"
        }
    }
}


# --------------------------------------------------------------------------------
# Update AD Schema to hold Legacy LAPS properties
# --------------------------------------------------------------------------------
Try {
    Write-Verbose "Main : Check Legacy LAPS schema properties"
    $Null = Get-AdObject -Identity "CN=ms-mcs-admpwd,CN=Schema,$($CurrentDomain.SubordinateReferences | Where-Object {$_ -like '*Config*'})"
} Catch {
    Write-Verbose "Main : Updating Schema to suport Legacy LAPS"
    Update-AdmPwdADSchema
}


# --------------------------------------------------------------------------------
# Update AD Schema to hold Windows LAPS properties
# --------------------------------------------------------------------------------
Try {
    Write-Verbose "Main : Check Windows LAPS schema properties"
    $Null = Get-AdObject -Identity "CN=ms-LAPS-Password,CN=Schema,$($CurrentDomain.SubordinateReferences | Where-Object {$_ -like '*Config*'})" 
} Catch {
    Write-Verbose "Main : Updating Schema to suport Windows LAPS"
    Update-LapsADSchema -Confirm:$false
}


# --------------------------------------------------------------------------------
# Create WMI filters
# --------------------------------------------------------------------------------
$WMIFilters = @()
$WMIFilters += "Detect Legacy LAPS; Select * From CIM_Datafile Where Name = `"C:\\Program Files\\LAPS\\CSE\\AdmPwd.dll`""
$WMIFilters += "Detect Windows LAPS; Select * From CIM_Datafile Where Name = `"C:\\Windows\\System32\\lapscsp.dll`""

# Build the date field in required format
$now = (Get-Date).ToUniversalTime()
$msWMICreationDate = ($now.Year).ToString("0000") + ($now.Month).ToString("00") + ($now.Day).ToString("00") + ($now.Hour).ToString("00") + ($now.Minute).ToString("00") + ($now.Second).ToString("00") + "." + ($now.Millisecond * 1000).ToString("000000") + "-000"

# Create WMI filters
foreach ($Line in $WMIFilters) {
    $Name = "$GPOPrefix - $($($Line -Split("; "))[0])"
    $Query = $($Line -Split("; "))[1]

    # Create a new GUID
    $NewWMIGUID = [string]"{" + ([System.Guid]::NewGuid()) + "}"

    $Attr = @{
        "msWMI-Name" = $Name;
        "msWMI-Parm1" = "Created by LAPS configuration script";
        "msWMI-Parm2" = "1;3;10;$($Query.Length.ToString());WQL;root\CIMv2;$Query;"
        "msWMI-Author" = "Administrator@" + $($CurrentDomain.DNSRoot);
        "msWMI-ID" = $NewWMIGUID;
        "instanceType" = 4;
        "showInAdvancedViewOnly" = "TRUE";
        "distinguishedname" = "CN=" + $NewWMIGUID + ",CN=SOM,CN=WMIPolicy,CN=System," + $($CurrentDomain.DistinguishedName);
        "msWMI-ChangeDate" = $msWMICreationDate;
        "msWMI-CreationDate" = $msWMICreationDate
        }

    if (!(Get-ADObject -Filter { msWMI-Name -eq $Name })) {
        Write-Output "Main : Create WMI filter $Name"
        New-ADObject -name $NewWMIGUID -type "msWMI-Som" -Path "CN=SOM,CN=WMIPolicy,CN=System,$($CurrentDomain.DistinguishedName)" -OtherAttributes $Attr | Out-Null
    }
}


# --------------------------------------------------------------------------------
# Create / Import GroupPolicy (Update scheduled task, and copy required LAPS files)
# --------------------------------------------------------------------------------
$GPOImport = Get-ChildItem -Path "$PSScriptRoot\GPO" -Recurse -Depth 1 | Where-Object {$_.FullName -Like "*{*}*"}
Foreach ($GPO in $GPOimport) {
    $GPOProperty = New-Object -Type PSObject -Property @{
        'Guid'  = $($GPO.Name)
        'Name' = $(([XML](Get-Content -Path "$($GPO.FullName)\backup.xml")).GroupPolicyBackupScheme.GroupPolicyObject.GroupPolicyCoreSettings.DisplayName.InnerText) -replace("Domain",$GPOPrefix)
    }
    Write-Output "Main : Create GPO $($GPOProperty.Name)"

    # Change scheduled task prior to import.
    If (Test-Path -Path "$($GPO.FullName)\DomainSysvol\GPO\Machine\Preferences\ScheduledTasks") {
        
        # Read Scheduled task from GPO
        [XML]$ScheduleXML = Get-Content -Path "$($GPO.FullName)\DomainSysvol\GPO\Machine\Preferences\ScheduledTasks\ScheduledTasks.xml"
        $CurrentArguemnts = $ScheduleXML.ScheduledTasks.TaskV2.Properties.Task.Actions.Exec.Arguments -split(" ")
        $NewArguemnts = ($CurrentArguemnts | Select-Object -SkipLast 1) -Join(" ")
        $ScriptName = Split-Path ($CurrentArguemnts[-1]) -Leaf

        # --
        # Create empty GPO (need the ID for the Path)
        # --
        $NewGPO = Get-GPO -Name $($GPOProperty.Name) -ErrorAction SilentlyContinue
        if ($null -eq $NewGPO) {
             $NewGPO = New-GPO -Name $($GPOProperty.Name)
        }
        $NewGPOPath = "\\$($CurrentDomain.DNSRoot)\SYSVOL\$($CurrentDomain.DNSRoot)\Policies\{$($NewGPO.ID)}\Machine\Scripts\Startup"
        $ScheduleXML.ScheduledTasks.TaskV2.Properties.Task.Actions.Exec.Arguments = $($NewArguemnts + " `"$NewGPOPath\$ScriptName")
        $ScheduleXML.save("$($GPO.FullName)\DomainSysvol\GPO\Machine\Preferences\ScheduledTasks\ScheduledTasks.xml")
        
        Import-GPO -BackupId $GPOProperty.Guid -Path $(Split-Path $($GPO.FullName) -Parent) -TargetName $($GPOProperty.Name) -CreateIfNeeded | Out-Null

        # Copy Legacy Laps installation files
        if (!(Test-Path $NewGPOPath)) {
            New-Item -Path $NewGPOPath -ItemType Directory | Out-Null
        }
        if (Test-Path $NewGPOPath) {
            Copy-Item -Path "$PSScriptRoot\GPO\Policy Dependencies\Manage-Laps-Version.ps1" -Destination $NewGPOPath
            Copy-Item -Path "$LapsFiles\LAPS.x64.msi" -Destination $NewGPOPath
            Copy-Item -Path "$LapsFiles\LAPS.x86.msi" -Destination $NewGPOPath
        }

    } else {
        Import-GPO -BackupId $GPOProperty.Guid -Path $(Split-Path $($GPO.FullName) -Parent) -TargetName $($GPOProperty.Name) -CreateIfNeeded | Out-Null

        $WMIFilter = $(New-Object Microsoft.GroupPolicy.GPDomain).SearchWmiFilters($(New-Object Microsoft.GroupPolicy.GPSearchCriteria)) | Where-Object {$_.Name -like "$GPOPrefix - *$(($($GPOProperty.Name) -split(" "))[2])*"}

        if ($null -ne ($WMIFilter).Name) {
            $FilterGPO = Get-GPO -Name $($GPOProperty.Name) -ErrorAction SilentlyContinue
            $FilterGPO.WmiFilter = $WMIFilter
        }
    }
}


# --------------------------------------------------------------------------------
# Script Done
# --------------------------------------------------------------------------------
Write-Output "Main : Domain is now prepared to Support Legacy and Windows LAPS"


# --------------------------------------------------------------------------------
# Prompt for cleanup
# --------------------------------------------------------------------------------
$title = "Delete Files"
$message = "Do you want to cleanup the downloaded files ?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Do cleanup."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Quit"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)


# --------------------------------------------------------------------------------
# Cleanup and Remove Legacy LAPS
# --------------------------------------------------------------------------------
if (($host.ui.PromptForChoice($title, $message, $options, 0)) -eq 0) {
    Write-Output "Cleanup"
    Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/x {97E2CA7B-B657-4FF7-A6DB-30ECC73E1E28} /qn /promptrestart" -Wait

    if (Test-Path -Path "$PSScriptRoot\GPO.zip") {
        Remove-Item -Path "$PSScriptRoot\GPO.zip"
    }
    if (Test-Path -Path "$PSScriptRoot\GPO") {
        Remove-Item -Path "$PSScriptRoot\GPO" -Recurse -Force
    }
    if (Test-Path -Path "$LapsFiles\LAPS.x64.msi") {
        Remove-Item -Path "$LapsFiles\LAPS.x64.msi" -Force
    }
    if (Test-Path -Path "$LapsFiles\LAPS.x86.msi") {
        Remove-Item -Path "$LapsFiles\LAPS.x86.msi" -Force
    }
    Write-Output "Cleanup done";
}

# Exit script with a pause to allow the user to read the output before closing the window
Write-Output 'Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
