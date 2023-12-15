<#

 Search Current Domain for all Computers, extract OUs names and selct the unique ones
 Select the OUs where computers needs permission to write the Windows AND Legacy LAPS password

#>

# --
# Get all computer objects from domain.
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
