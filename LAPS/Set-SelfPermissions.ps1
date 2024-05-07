<#

    DISCLAIMER:

    THE SCRIPT IS PROVIDED AS-IS, WITHOUT WARRANTY OF ANY KIND. USE AT YOUR OWN RISK.

    By running this script, you acknowledge that you have read and understood the disclaimer, and you agree to assume
    all responsibility for any failures, damages, or issues that may arise as a result of executing this script.

#>

<#

 Search Current Domain for all Computers, extract OUs names and select unique OUs
 Select the OUs where computers needs permission to write the Windows AND Legacy LAPS password

#>


# --
# Get all Computers from domain.
# - Exclude the Domain Controllers OU and the Computers container.
# --
$AllComputers = (Get-ADComputer -Filter * | Where {$_.DistinguishedName -notlike "*Computers*" -AND $_.DistinguishedName -notlike "*Domain Controllers*"})


# --
# Remove Computer Name from DistinguishedName and select unique OUs
# - Output to GridView for selection of OUs where the permissions need to be applied.
# --
$OUsUnique = $AllComputers | Select-Object @{l='DN';e={$_.DistinguishedName -replace "^CN=.+?(?<!\\),"}} -Unique
$OUsUnique | Out-GridView -Title "Select OUs" -OutputMode Multiple | % {
    Write-Output "Setting LAPS Self permissions on `"$($_.DN)`""
    Set-LapsADComputerSelfPermission -Identity $_.DN | Out-Null
    Set-AdmPwdComputerSelfPermission -OrgUnit $_.DN | Out-Null
}


# --
# Alternative
# --


# --
# Find all OU's where the LAPS policy is linked, and set required SELF permissions
# --
Get-GPO -All | Where {$_.DisplayName -like '*LAPS*'} | Foreach {
    $GPOReport = [XML](Get-GPOReport -Guid $_.Id -ReportType Xml)
    $GPOReport.GPO.LinksTo.SOMPath | Foreach {
        $Filter = $_
        $OuName = (Get-ADObject -Filter * -Properties CanonicalName | Where {$_.CanonicalName -eq $Filter}).DistinguishedName
        Set-LapsADComputerSelfPermission -Identity $OuName
        Set-AdmPwdComputerSelfPermission -OrgUnit $OuName
    }
}




# --
# Allow LAPS managers password lookup
# --
