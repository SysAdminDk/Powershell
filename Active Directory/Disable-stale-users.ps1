<#
    .DESCRIPTION
     Find and disable users that havnt been logged in the last 6 month

#>

$CurrentDate = Get-Date -Format "dd-MM-yyyy"
$LastLogon = (Get-Date).Adddays(-180)
$Users = Get-ADUser -Filter { LastLogonTimeStamp -lt $LastLogon -and Enabled -eq 'True' } -Properties Description
Foreach ($User in $Users) {
    Disable-ADAccount -Identity $User.DistinguishedName
    Set-AdUser -Identity $User.DistinguishedName -Description $("[Disabled, $CurrentDate] $($User.Description)")
}

