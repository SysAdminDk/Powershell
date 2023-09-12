<#
.DESCRIPTION
    This script update default permissions on selected Schema Classes
    
    Copy to a Domain Controller or other Server with write permissions to Schema, and run with permissions to modify the Schema.

    There will be created a backup file of current permissions on the selected objects, prior to applying the updated ACL's, that backup can be used to restore if needed
#>

##################################################################################
# DISCLAIMER [ Start ]
##################################################################################

clear
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

$title = "Update Active Directory"
$message = "Do you want to run this script and update default schema ACL's ?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes Update."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Just quit"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

if (($host.ui.PromptForChoice($title, $message, $options, 0)) -eq 1) {
    Write-Output ""
} else {
    Clear
    Write-Output "Verifying Prerequisites"
}

##################################################################################
# DISCLAIMER [ End ]
##################################################################################



# --
# Load Required Modules
# --
Import-Module ActiveDirectory


# --
# Default Variables
# --
$DefaultPath = $PSScriptRoot
$ADForest = Get-ADForest
$ADDomain = Get-ADDomain


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

if (!(Get-ADGroupMember -Identity "Enterprise Admins" | Where {$_.distinguishedName -eq $ADUser.DistinguishedName})) { 
    throw "User is NOT a member of Enterprise Admins, unable to continue"
}

if (!(Get-ADGroupMember -Identity "Schema Admins" | Where {$_.distinguishedName -eq $ADUser.DistinguishedName})) { 
    throw "User is NOT a member of Schema Admins, unable to continue"
}


# --
# Selected Schema Classes.
# --
$SchemaClasses = @(
    "Computer",
    "Foreign-Security-Principal",
    "Group","groupOfUniqueNames",
    "inetOrgPerson",
    "ms-DS-Group-Managed-Service-Account",
    "ms-DS-Managed-Service-Account",
    "Organizational-Unit",
    "Print-Queue",
    "User",
    "Group-Policy-Container"
    )


# --
# Save the Default Security Descriptor from selected Classes
# --
$DefaultSecurityDescriptors = @()
ForEach ($Class in $SchemaClasses) {
    Write-Verbose "Quering SecurityDescriptor in Schema for class $Class"
    $Identity = $("CN=$Class,CN=Schema,CN=Configuration," + $ADDomain.DistinguishedName)
    $DefaultSecurityDescriptors += Get-ADObject -Server $ADForest.SchemaMaster -Identity $Identity -Properties defaultSecurityDescriptor | Select-Object Name,DistinguishedName,defaultSecurityDescriptor
}
$DefaultSecurityDescriptors | Export-Csv -Path "$DefaultPath\BackupSecurityDescriptors.csv" -NoTypeInformation -Encoding UTF8 -Delimiter "|"


# --
# Update the Default Security Descriptor on selected Classes
# --

<#
# To restore to Default Schema Permissions, use this csv file
$DefaultSecurityDescriptors = Import-Csv -Path "$DefaultPath\Refrence\DefaultSecurityDescriptors.csv" -Encoding UTF8 -Delimiter "|"
#>

<#
# To restore the backup if needed, use this csv file
$DefaultSecurityDescriptors = Import-Csv -Path "$DefaultPath\BackupSecurityDescriptors.csv" -Encoding UTF8 -Delimiter "|"
#>

$DefaultSecurityDescriptors = Import-Csv -Path "$DefaultPath\UpdateSecurityDescriptors.csv" -Encoding UTF8 -Delimiter "|"
Foreach ($Class in $DefaultSecurityDescriptors) {
    Write-Output "Update SecurityDescriptor on Schema class $($Class.Name)"
    $Identity = $($Class.DistinguishedName -replace("DC=DOMAIN,DC=REF",(Get-ADDomain).DistinguishedName))
    Set-ADObject -Identity $Identity -Replace @{defaultSecurityDescriptor=$($Class.defaultSecurityDescriptor)} -Server $ADDomain.PDCEmulator
}

