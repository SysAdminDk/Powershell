<#
    .DISCLAIMER

    THE SCRIPT IS PROVIDED AS-IS, WITHOUT WARRANTY OF ANY KIND. USE AT YOUR OWN RISK.

    By running this script, you acknowledge that you have read and understood the disclaimer, and you agree to assume
    all responsibility for any failures, damages, or issues that may arise as a result of executing this script.

    .DESCRIPTION
    Restore the DHCP server with leases, to the latest backup file found on the supplied Path

    .PARAMETER Path 
    Specifies where the backup file will be retrieved.
    This can be a local path or a UNC path.
    
    .EXAMPLE
    .\Restore-Dhcp-Server.ps1 -Path "\\FILE01\Backup\DHCP"
#>
[CmdletBinding()]
Param(
  [Parameter(ValueFromPipelineByPropertyName=$true,Position=0)][string]$Path
)

# --------------------------------------------------------------------------------------------------
# Restore the DHCP server
# --------------------------------------------------------------------------------------------------
$LatestBackup = Get-ChildItem -Path $Path | Sort-Object CreationTime -Descending | Select-Object -First 1
Try {
    Import-DhcpServer -Leases -File "$($LatestBackup.fullname)" -BackupPath "C:\Windows\Temp" -force
} catch {
    Write-Output "Unable to Restore the DHCP server"
    Write-Output $_
}
