<#

    Inventory all Local Users and GroupsMembers

#>


# List all Windows Servers in the domain, excluding Domain Controllers
# ----------------------------------------------------------------------------------------------------
$AllServers = Get-ADComputer -Filter "operatingSystem -like 'Windows Server*' -and Enabled -eq 'True'"
$AllServers = $AllServers | Where {$_.DistinguishedName -Notlike "*Domain Controllers*"}


# Run Inventory
# ----------------------------------------------------------------------------------------------------
$Inventory = @()
Foreach ($Server in $AllServers) {

    # Test connection with WinRM (needed for remote PowerShell)
    # ----------------------------------------------------------------------------------------------------
    $NetWinRM = Test-NetConnection -ComputerName $Server.DNSHostName -CommonTCPPort WINRM
    If (!($NetWinRM.TcpTestSucceeded)) {
        Write-Warning "Unable to connect to $($Server.DNSHostName)"
    } else {

        # Connect and get all Local Users and Group Members
        # ----------------------------------------------------------------------------------------------------
        $Inventory += Invoke-Command -ComputerName $Server.DNSHostName -Authentication NegotiateWithImplicitCredential -ScriptBlock {

            $LocalUserandGroupData = @()
            Get-LocalUser | Select-Object -Property Name, Enabled, PasswordExpires, PasswordLastSet, PasswordRequired | Foreach {

                $Users = New-Object -TypeName psobject
                $Users | Add-Member -MemberType NoteProperty -Name "UserName" -Value $_.Name
                $Users | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $_.Enabled
                $Users | Add-Member -MemberType NoteProperty -Name "PasswordExpires" -Value $_.PasswordExpires
                $Users | Add-Member -MemberType NoteProperty -Name "PasswordLastSet" -Value $_.PasswordLastSet
                $Users | Add-Member -MemberType NoteProperty -Name "PasswordRequired" -Value $_.PasswordRequired
                $Users | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $Null
                $Users | Add-Member -MemberType NoteProperty -Name "MemberName" -Value $Null
                $Users | Add-Member -MemberType NoteProperty -Name "Source" -Value $Null
                $Users | Add-Member -MemberType NoteProperty -Name "Class" -Value $Null

                $LocalUserandGroupData += $Users

            }

            Get-LocalGroup | Foreach {
                $GroupName = $_.Name
                Get-LocalGroupMember $GroupName | Foreach {

                    $Groups = New-Object -TypeName psobject
                    $Groups | Add-Member -MemberType NoteProperty -Name "UserName" -Value $Null
                    $Groups | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $Null
                    $Groups | Add-Member -MemberType NoteProperty -Name "PasswordExpires" -Value $Null
                    $Groups | Add-Member -MemberType NoteProperty -Name "PasswordLastSet" -Value $Null
                    $Groups | Add-Member -MemberType NoteProperty -Name "PasswordRequired" -Value $Null
                    $Groups | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $GroupName
                    $Groups | Add-Member -MemberType NoteProperty -Name "MemberName" -Value $_.Name
                    $Groups | Add-Member -MemberType NoteProperty -Name "Source" -Value $_.PrincipalSource
                    $Groups | Add-Member -MemberType NoteProperty -Name "Class" -Value $_.ObjectClass

                    $LocalUserandGroupData += $Groups

                }
            }

            Return $LocalUserandGroupData
        }
    }
}


# Export to CSV
# ----------------------------------------------------------------------------------------------------
$Inventory | Select-Object PSComputerName, UserName, Enabled, PasswordExpires, PasswordLastSet, PasswordRequired, GroupName, MemberName, Source, Class | `
    Export-Csv -Path "C:\TS-Data\ServerInventory.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"
