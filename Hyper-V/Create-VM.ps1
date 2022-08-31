# --
# Import module
# --
Import-Module "$PSScriptRoot\VM-Functions.psm1"
<#
Import-Module "C:\Users\Administrator\Documents\VM-Functions.psm1" -Verbose -Force
#>


# --
# Set default Path & Switch Name
# --
$DefaultPath = "D:\Virtual Machines"
$Defaultwitch = "Default Virtual Switch"


Do {
    # --
    # Request required parameters
    # --
    do {
        $ServerName=[string]$(Read-Host -Prompt "Servername ([A-Z]{2,4}-[A-Z]{2,15}-\d\d)")
        if ((Get-VM -Name "$ServerName" -ErrorAction SilentlyContinue) -ne $null) {
               Write-Output "$ServerName alreay exists, please use another name"
            $exists = $true
        } else {
            $exists = $false
        }
        Write-Verbose "Servername = $ServerName"
    } until ($ServerName -cmatch '^[a-zA-Z]{2,4}-[a-zA-Z0-9-]{2,15}-[T0-9]{2}$' -and $exists -eq $false)

    do {
        $CpuCount=[int]$(Read-Host -Prompt "Amount off required processors (2|4|6|8)")
        Write-Verbose "Cpu Count = $CpuCount"
    } until ($CpuCount -match '^(2|4|6|8)$')

    do {
        $MemoryMax=[int]$(Read-Host -Prompt "Amount off maximum memory (2|4|6|8|16)")
        Write-Verbose "Maximum Dynamic Memory = $MemoryMax"
    } until ($MemoryMax -match '^(2|4|6|8|16)$')

    do {
        $DiskSize=[int]$(Read-Host -Prompt "Amount off diskspace (50>100)")
        Write-Verbose "OS Disk Size = $DiskSize"
    } until ($DiskSize -ge 50 -and $DiskSize -le 100)


    Write-Verbose "Creating VM $ServerName"
    New-CreateVM -ServerName $ServerName -MemoryMax $MemoryMax -CpuCount $CpuCount -DefaultPath $DefaultPath -Defaultwitch $Defaultwitch

    Write-Verbose "Creating OS disk for $ServerName"
    New-CreateDrive -ServerName $ServerName -DiskSize $DiskSize

    $MultiDisk = (Read-Host -Prompt "Create any data drives ?").ToLower()
    if ($MultiDisk -eq "y") {
        do {
            $DiskSize=[int]$(Read-Host -Prompt "Amount off diskspace (>50 <1024)")
            New-CreateDrive -ServerName $ServerName -DiskSize $DiskSize
            $exit = (Read-Host -Prompt "More drives ?").ToLower()
        } until ($exit -eq "n")
    }

    $noMore = (Read-Host -Prompt "Create new VM ?").ToLower()
} until ($noMore -eq "n")
