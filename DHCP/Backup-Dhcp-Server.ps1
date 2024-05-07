<#
    .DISCLAIMER

    THE SCRIPT IS PROVIDED AS-IS, WITHOUT WARRANTY OF ANY KIND. USE AT YOUR OWN RISK.

    By running this script, you acknowledge that you have read and understood the disclaimer, and you agree to assume
    all responsibility for any failures, damages, or issues that may arise as a result of executing this script.
    
    .DESCRIPTION
    Create a backup of DHCP scope and leases.

    .PARAMETER Path 
    Specifies where the backup file will be created.
    This can be a local path or a UNC path, but the user running the script must have write permissions to the path.
    
    .EXAMPLE
    .\Backup-Dhcp-Server.ps1 -Path "\\FILE01\Backup\DHCP"

#>
[CmdletBinding()]
Param(
  [Parameter(ValueFromPipelineByPropertyName=$true,Position=0)][string]$Path
)

#
# Create the folder if missing.
#
if ((!Test-Path -Path $Path)) {
    New-Item -Path $Path -ItemType Directory -Force | Out-Null
}

#
# Backup DHCP server
#
Try {
    $FileName = (get-date).ToString('dd-MM-yyyy')
    Export-DhcpServer -Leases -File "$Path\DHCP-$FileName.xml"
} Catch {
    Write-Output "Unable to backup DHCP server"
    Write-Output $_
}
