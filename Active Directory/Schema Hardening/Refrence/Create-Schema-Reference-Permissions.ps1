<#
.DESCRIPTION
    This script update default permissions on selected Schema Classes
    
    Copy to a Domain Controller or other Server with write permissions to Schema, and run with permissions to modify the Schema.

    There will be created a backup file of current permissions on the selected objects, prior to applying the updated ACL's, that backup can be used to restore if needed
#>


# --
# Default path for CSV files
# --
#$DefaultPath = $PSScriptRoot
$DefaultPath = "Z:\Active Directory\Hardening Schema"


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
$ADForest = Get-ADForest
$ADDomain = Get-ADDomain

$DefaultSecurityDescriptors = @()
ForEach ($Class in $SchemaClasses) {
    $DefaultSecurityDescriptors += Get-ADObject -Server $ADForest.SchemaMaster -Identity "CN=$Class,CN=Schema,CN=Configuration,$($ADDomain.DistinguishedName)" -Properties defaultSecurityDescriptor | `
    Select-Object -Property Name,@{n="DistinguishedName"; e={$_.DistinguishedName -replace($ADDomain.DistinguishedName,"DC=DOMAIN,DC=REF")}},defaultSecurityDescriptor
}
$DefaultSecurityDescriptors | Export-Csv -Path "$DefaultPath\DefaultSecurityDescriptors.csv" -NoTypeInformation -Encoding UTF8 -Delimiter "|"
