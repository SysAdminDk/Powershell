<#
    .DESCRIPTION
     Backup ALL Group Policy and the files located in gpo path.
     CSV report Group Policy links.

#>


# --
# Where to store the GPO backup files.
# --
$BasePath = "C:\Temp\GPO-Backup"

# --
# Create output folder.
# --
$FileDate = Get-Date -Format "dd-MM-yyyy"
$FilePath = "$BasePath\$FileDate"
If (!(Test-Path -Path $FilePath)) {
    New-Item -Path $FilePath -ItemType Directory | Out-Null
}


# --
# Import modules
# --
Import-Module ActiveDirectory


# --
# Get Domain Info
# --
$Domain = Get-ADDomain

# --
# Backup ALL Group Policies
# --
$GPOs = Get-GPO -All

$OutReport = @()
Foreach ($GPO in $GPOs) {
    Write-Output "Export $($GPO.DisplayName)"
    if (!(Test-Path -Path "$FilePath\$($GPO.DisplayName)")) {
        New-Item -Path "$FilePath\$($GPO.DisplayName)" -ItemType Directory | Out-Null
    }
    Backup-GPO -Guid $GPO.ID -Path "$FilePath\$($GPO.DisplayName)" | Out-Null
    $UserPolicyFiles = Get-ChildItem -Path "\\$($Domain.DNSRoot)\sysvol\$($Domain.DNSRoot)\Policies\{$($GPO.ID)}\User\Scripts" -File -Recurse
    if ($UserPolicyFiles.Count -ne 0) {
        # If there are ANY files located in the User Policy Path, make a backup
        New-Item -Path "$FilePath\$($GPO.DisplayName)\UserFiles" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path "\\$($Domain.DNSRoot)\sysvol\$($Domain.DNSRoot)\Policies\{$($GPO.ID)}\User\Scripts" -Destination "$FilePath\$($GPO.DisplayName)\UserFiles" -Recurse | Out-Null
    }
    $ComputerPolicyFiles = Get-ChildItem -Path "\\$($Domain.DNSRoot)\sysvol\$($Domain.DNSRoot)\Policies\{$($GPO.ID)}\Machine\Scripts" -File -Recurse
    if ($ComputerPolicyFiles.Count -ne 0) {
        # If there are ANY files located in the Machine Policy Path, make a backup
        New-Item -Path "$FilePath\$($GPO.DisplayName)\MachineFiles" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path "\\$($Domain.DNSRoot)\sysvol\$($Domain.DNSRoot)\Policies\{$($GPO.ID)}\Machine\Scripts" -Destination "$FilePath\$($GPO.DisplayName)\MachineFiles" -Recurse | Out-Null
    }

    # Save HTML report.
    Get-GPOReport -ReportType Html -Guid $GPO.ID -Path "$FilePath\$($GPO.DisplayName)\$($GPO.DisplayName).html"

    # Need to create CSV to show where the policy are linked
    [XML]$GPReport = Get-GPOReport -ReportType Xml -Guid $GPO.ID
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
$OutReport | Export-Csv -Path "$(Split-Path -Path $FilePath -Parent)\$FileDate-GPO-Link-Report.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8
