<#
    .DESCRIPTION
    Download and import selected security baselines from Microsoft Security Compliance Toolkit

    .PARAMETER Path 
    Specifies where the backup file will be created.
    This can be local path or a UNC path.
    
    .EXAMPLE
    .\Backup-Dhcp-Server.ps1 -Path "\\FILE01\Backup\DHCP"

#>

#
# Create folder if missing.
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
