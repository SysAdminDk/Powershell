<#
    Script to Inventory Scheduled Tasks on all servers

    Excludes tasks thats running as "System", "Network Service" and "local Service"
    - If all tasks are to be listed change $ExcludeAccounts variable at line 19

    Outputs to screen, no CSV or log.

#>

# List all Windows Servers in the domain, excluding Domain Controllers (There shouldn't be any)
# ----------------------------------------------------------------------------------------------------
$AllServers = Get-ADComputer -Filter "operatingSystem -like 'Windows Server*'"
$AllServers = $AllServers | Where {$_.DistinguishedName -notlike "*Domain*"}


# Exclude from inventory
# ----------------------------------------------------------------------------------------------------
$ExcludeAccounts = @($null,"System","Network Service", "local Service")


Foreach ($Server in $AllServers) {

    # Test connection with WinRM (needed for remote PowerShell)
    # ----------------------------------------------------------------------------------------------------
    $NetWinRM = Test-NetConnection -ComputerName $Server.DNSHostName -CommonTCPPort WINRM
    If (!($NetWinRM.TcpTestSucceeded)) {
        Write-Warning "Unable to connect to $($Server.DNSHostName)"
    } else {

        Write-Output "$($Server.DNSHostName) : WinRM ok"
        Write-Output "Connect WinRM"

        # Connect and get all shcduled tasks
        # ----------------------------------------------------------------------------------------------------
        Invoke-Command -ComputerName $Server.DNSHostName -Authentication NegotiateWithImplicitCredential -ScriptBlock {

            $Command = Get-Command -Name "Get-ScheduledTask" -ErrorAction SilentlyContinue
            if ($Command -eq $null) {

                Write-Warning "$($ENV:Computername) : The PowerShell command `"Get-ScheduledTask`" is not recognized"

            } else {
                $Tasks = Get-ScheduledTask | Select-Object TaskName, State, @{Name="RunAs";Expression={ $_.principal.userid }} | Where {$_.RunAs -notin $Using:ExcludeAccounts}
                if ($Tasks) {
                    Write-Output $Tasks
                }
            }
        }
    }
}
