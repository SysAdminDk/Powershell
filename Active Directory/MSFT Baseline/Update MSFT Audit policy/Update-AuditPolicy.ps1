# --
# Find the MSFT*Domain Controller* GPO
# --
$GPO = Get-GPO -All | Where {$_.DisplayName -like 'MSFT*Windows Server*- Domain Controller'}

$GPO | Foreach {

    # --
    # Backup the GPO
    # --
    If (!(Test-Path -Path "$($env:TEMP)\MSFT")) {
        New-Item -Path "$($env:TEMP)\MSFT" -ItemType Directory | Out-Null
    }
    $Backup = $_ | Backup-GPO -Path "$($env:TEMP)\MSFT"


    # --
    # Import Audit.CSV for change.
    # --
    #$GPOBackupFolder = Get-ChildItem -Path "$($env:TEMP)\MSFT"
    $GPOBackupFolder = Get-ChildItem -Path $Backup.BackupDirectory -Filter "{$($Backup.Id.guid)}"

    if (Test-Path -Path "$($GPOBackupFolder.FullName)\DomainSysvol\GPO\Machine\microsoft\windows nt\Audit\Audit.csv") {
        $AuditConfig = Import-Csv "$($GPOBackupFolder.FullName)\DomainSysvol\GPO\Machine\microsoft\windows nt\Audit\Audit.csv"
    }


    # --
    # Update Audit Kerberos Service Ticket Operations (Pingcacle)
    # --
    $AuditConfig | Where {$_.Subcategory -eq "Audit Kerberos Service Ticket Operations"} | ForEach-Object {
        $_.'Setting Value' = 3;
        $_.'Inclusion Setting' = "Success and Failure"
    }

    # --
    # Create Audit DPAPI Activity (Pingcacle)
    # --
    $AuditConfig += ([PSCustomObject]@{
        "Machine Name" = ""
        "Policy Target" = "System"
        "Subcategory" = "Audit DPAPI Activity"
        "Subcategory GUID" = "{0cce922d-69ae-11d9-bed3-505054503030}"
        "Inclusion Setting" = "Success"
        "Exclusion Setting" = ""
        "Setting Value" = "1"
    })

    # --
    # Create audit Logoff (Pingcacle)
    # --
    $AuditConfig += ([PSCustomObject]@{
        "Machine Name" = ""
        "Policy Target" = "System"
        "Subcategory" = "Audit Logoff"
        "Subcategory GUID" = "{0cce9216-69ae-11d9-bed3-505054503030}"
        "Inclusion Setting" = "Success"
        "Exclusion Setting" = ""
        "Setting Value" = "1"
    })

    # --
    # Export audit scv in right format
    # --
    $AuditConfig | ConvertTo-Csv -NoTypeInformation | ForEach-Object { $_ -replace('"',"")} |`
        Out-File "$($GPOBackupFolder.FullName)\DomainSysvol\GPO\Machine\microsoft\windows nt\Audit\audit.csv"

    # --
    # Import GPO, overwrite orginal
    # --
    Import-GPO -BackupId $GPOBackupFolder.Name -Path "$($env:TEMP)\MSFT" -TargetName $_.DisplayName -CreateIfNeeded | Out-Null


    # --
    # Cleanup GPO backup
    # --
    Remove-Item -Path $GPOBackupFolder.FullName -Confirm:$false -Recurse -Force

}