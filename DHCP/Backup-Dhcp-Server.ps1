<#
    .DESCRIPTION
    Download and import selected security baselines from Microsoft Security Compliance Toolkit

    .PARAMETER DownloadID
    
    .EXAMPLE
    .\Backup-Dhcp-Server.ps1 -Path "\\FILE01\Backup\DHCP"

#>

# --------------------------------------------------------------------------------------------------
# Backup DHCP server
# --------------------------------------------------------------------------------------------------
$FileName = (get-date).ToString('dd-MM-yyyy')
Export-DhcpServer -Leases -File "$UNCPath\DHCP-Backup-$FileName.xml"
