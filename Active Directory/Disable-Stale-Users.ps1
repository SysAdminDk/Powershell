<#
    .DESCRIPTION
     Find and disable Inactive Users

#>

$LastLogon = (Get-Date).Adddays(-180)
$CurrentDate = Get-Date -Format "dd-MM-yyyy"

$Users = Get-ADUser -Filter { LastLogonTimeStamp -lt $LastLogon -and Enabled -eq 'True' } -Properties Description
Foreach ($User in $Users) {
    Write-Verbose "Disable User : $($User.DistinguishedName)"

    Disable-ADAccount -Identity $User.DistinguishedName
    Set-AdUser -Identity $User.DistinguishedName -Description $("[Disabled, $CurrentDate] $($User.Description)")
}
