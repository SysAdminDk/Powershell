<#
.DESCRIPTION
    This script update default permissions on selected Schema Classes
    
    Copy to a Domain Controller or other Server with write permissions to Schema, and run with permissions to modify the Schema.

    There will be created a backup file of current permissions on the selected objects, prior to applying the updated ACL's, that backup can be used to restore if needed
#>

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


# --
# Verify curent user have Schema Admin rights.
# --
Write-Verbose "Verify entered credentials have permissions to access Schema"
if (((Get-AdUser -Identity $($Credentials.UserName -split("\\"))[-1] -Properties MemberOf) -Like "*\Schema Admins").Count -eq 0) {
    Write-Output "The current user is NOT a member of the Schema Admins group."
    break
} else {
    Write-Verbose "Get Forrest and Domain information"
    $ADForest = Get-ADForest -Credential $Credentials
    $ADDomain = Get-ADDomain -Credential $Credentials
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
    #$DefaultSecurityDescriptors += Get-ADObject -Server $ADForest.SchemaMaster -Identity $Identity -Properties defaultSecurityDescriptor -Credential $Credentials | Select-Object Name,DistinguishedName,defaultSecurityDescriptor
    $DefaultSecurityDescriptors += Get-ADObject -Server $ADForest.SchemaMaster -Identity $Identity -Properties defaultSecurityDescriptor | Select-Object Name,DistinguishedName,defaultSecurityDescriptor
}
$DefaultSecurityDescriptors | Export-Csv -Path "$DefaultPath\BackupSecurityDescriptors.csv" -NoTypeInformation -Encoding UTF8 -Delimiter "|"


# --
# Update the Default Security Descriptor on selected Classes
# --
#$DefaultSecurityDescriptors = Import-Csv -Path "$DefaultPath\DefaultSecurityDescriptors.csv" -Encoding UTF8 -Delimiter "|"
#$DefaultSecurityDescriptors = Import-Csv -Path "$DefaultPath\BackupSecurityDescriptors.csv" -Encoding UTF8 -Delimiter "|"
$DefaultSecurityDescriptors = Import-Csv -Path "$DefaultPath\UpdateSecurityDescriptors.csv" -Encoding UTF8 -Delimiter "|"
Foreach ($Class in $DefaultSecurityDescriptors) {
    Write-Output "Update SecurityDescriptor on Schema class $($Class.Name)"
    $Identity = $($Class.DistinguishedName -replace("DC=DOMAIN,DC=REF",(Get-ADDomain).DistinguishedName))
    #Set-ADObject -Identity $Identity -Replace @{defaultSecurityDescriptor=$($Class.defaultSecurityDescriptor)} -Credential $Credentials
    Set-ADObject -Identity $Identity -Replace @{defaultSecurityDescriptor=$($Class.defaultSecurityDescriptor)} -Server $ADDomain.PDCEmulator
}

