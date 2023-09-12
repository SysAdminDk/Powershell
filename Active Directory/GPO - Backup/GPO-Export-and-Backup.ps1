<#
    .SYNOPSIS
    Backup all GPO's that have changed since last time this script have run (Using the same BackupFlder)


    .DESCRIPTION
    Script exports and documents the GPO's in Active Directory, writes an CSV file where each GPO have been linked, and saves a HTML GPO report.
    - If there are any files in the SCRIPTS folder in the GPO, they will be copied to the backup folder.


    .PARAMETER BackupFolder
    Set the folder where the export and reports will be stored.


    .EXAMPLE
    .\GPO-Export-and-Backup.ps1 -BackupFolder "Path to where GPO export is stored" -Verbose

#>

param (
    [parameter(ValueFromPipeline)][string]$BackupFolder = $PSScriptRoot
)

# --
# Get Date
# --
$FileDate = Get-Date -Format "dd-MM-yyyy"
$GpoFilePath = $($BackupFolder + "\" + $FileDate)
If (!(Test-Path -Path $GpoFilePath)) {
    New-Item -Path $GpoFilePath -ItemType Directory | Out-Null
}


# --
# Get latest export date (Only export policy that have changed or added since)
# --
#$LatestExportTime = Get-Date -Date "06-09-2023"
Write-Verbose "Find latest GPO backup in $BackupFolder"
$LatestExportTime = $(Get-Date -Date ((Get-ChildItem -Path $GpoFilePath | Sort-Object CreationTime -Descending | Select-Object -First 1).LastWriteTime).ToShortDateString())

# --
# Import modules
# --
Write-Verbose "Import Required modules"
Import-Module ActiveDirectory


# --
# Get Domain Info
# --
Write-Verbose "Get Domain info and find/makeup SysVol Path"
$Domain = Get-ADDomain
$SysVolFolder = "\\" + $($Domain.DNSRoot) + "\sysvol\" + $($Domain.DNSRoot) + "\Policies\"


# --
# Backup changed Group Policies
# --
Write-Verbose "Get GPO's changed since $LatestExportTime"
$GPOs = Get-GPO -All | Where { $_.ModificationTime -gt $LatestExportTime }


$OutReport = @()
Foreach ($GPO in $GPOs) {
    Write-Verbose "Export $($GPO.DisplayName)"
    if (!(Test-Path -Path "$GpoFilePath\$($GPO.DisplayName)")) {
        New-Item -Path "$GpoFilePath\$($GPO.DisplayName)" -ItemType Directory | Out-Null
    }
    Backup-GPO -Guid $GPO.ID -Path "$GpoFilePath\$($GPO.DisplayName)" | Out-Null

    $UserPolicyFiles = Get-ChildItem -Path $($SysVolFolder + "{" + $($GPO.ID) + "}\User\Scripts") -File -Recurse
    if ($UserPolicyFiles.Count -ne 0) {
        Write-Verbose "Copy scripts from $($GPO.DisplayName) User"
        New-Item -Path "$GpoFilePath\$($GPO.DisplayName)\UserFiles" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path $($SysVolFolder + "{" + $($GPO.ID) + "}\User\Scripts") -Destination $($GpoFilePath + "\" + $($GPO.DisplayName) + "\UserFiles") -Recurse | Out-Null
    }

    $ComputerPolicyFiles = Get-ChildItem -Path $($SysVolFolder + "{" + $($GPO.ID) + "}\Machine\Scripts") -File -Recurse
    if ($ComputerPolicyFiles.Count -ne 0) {
        Write-Verbose "Copy scripts from $($GPO.DisplayName) Machine"
        New-Item -Path "$GpoFilePath\$($GPO.DisplayName)\MachineFiles" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path $($SysVolFolder + "{" + $($GPO.ID) + "}\Machine\Scripts") -Destination $($GpoFilePath + "\" + $($GPO.DisplayName) + "\MachineFiles") -Recurse | Out-Null
    }

    Write-Verbose "Create GPO HTML report"
    [XML]$GPReport = Get-GPOReport -ReportType Xml -Guid $GPO.ID
    Get-GPOReport -ReportType Html -Guid $GPO.ID -Path $($GpoFilePath + "\" + $($GPO.DisplayName) + "\" + $($GPO.DisplayName) + ".html")

    Write-Verbose "Document the OU's where the Policy is linked"
    if (($GPReport.GPO.LinksTo).Count -eq 0) {
        $OutReport += [PSCustomObject]@{
        "Name" = $GPReport.GPO.Name
        "Link" = ""
        "Link Enabled" = ""
        "ComputerEnabled" = $GPReport.GPO.Computer.Enabled
        "UserEnabled" = $GPReport.GPO.User.Enabled
        "WmiFilter" = $GPO.WmiFilter
        "GpoApply" = (Get-GPPermissions -Guid $GPO.ID -All | Where {$_.Permission -eq "GpoApply"}).Trustee.Name
        "SDDL" = $($GPReport.GPO.SecurityDescriptor.SDDL.'#text')
        }
    } else {

        foreach ($i in $GPReport.GPO.LinksTo) {
            $OutReport += [PSCustomObject]@{
            "Name" = $GPReport.GPO.Name
            "Link" = $i.SOMPath
            "Link Enabled" = $i.Enabled
            "ComputerEnabled" = $GPReport.GPO.Computer.Enabled
            "UserEnabled" = $GPReport.GPO.User.Enabled
            "WmiFilter" = $GPO.WmiFilter
            "GpoApply" = (Get-GPPermissions -Guid $GPO.ID -All | Where {$_.Permission -eq "GpoApply"}).Trustee.Name
            "SDDL" = $($GPReport.GPO.SecurityDescriptor.SDDL.'#text')
            }
        }

    }

}
$OutReport | Export-Csv -Path "$(Split-Path -Path $GpoFilePath -Parent)\$FileDate-GPO-Link-Report.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8
