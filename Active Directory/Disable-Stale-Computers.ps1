<#
    .DESCRIPTION
     Find and disable Inactive Computers

#>

$LastLogon = (Get-Date).Adddays(-90)
$CurrentDate = Get-Date -Format "dd-MM-yyyy"

$Computers = Get-ADComputer -Filter { LastLogonTimeStamp -lt $LastLogon -and Enabled -eq 'True' }  -Properties Description
Foreach ($Computer in $Computers) {
    Write-Verbose "Disable Computer : $($Computer.DistinguishedName)"

    Disable-ADAccount -Identity $Computer.DistinguishedName
    Set-ADComputer -Identity $Computer.DistinguishedName -Description $("[Disabled, $CurrentDate] $($Computer.Description)")
}
