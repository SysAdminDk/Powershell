Import-Module "$PSScriptRoot\VM-Functions.psm1"
Import-Module "C:\Users\Administrator\Documents\VM-Functions.psm1"


# --
# Remove and cleanup the selected VM.
# --
do {

    $ToBeRemoved = select-VirtualMachine -Hostname localhost
    do {
        $DeleteConfirm = Read-Host -Prompt "Are you sure the $($ToBeRemoved.Name) is to be deleted ?"
    } until ($DeleteConfirm -match "y|Y|n|N")

    if ($DeleteConfirm -match "y|Y") {
        $DiskPath = $(Get-VMHardDiskDrive -VMName $ToBeRemoved.Name).Path
        $DiskPathParent = $(Split-Path -Path $DiskPath -Parent).Split("\\")
        $VMInfo = Get-VM -Name $ToBeRemoved.Name
        $VmPath = $VMInfo.Path.Split("\\")
    
        $PathTest = Compare-Object -ReferenceObject $DiskPathParent -DifferenceObject $VmPath -PassThru
        If (!($PathTest -eq "Virtual Hard Disks")) {
            # Remove VHD
            Write-Output "Remove disk"
        }
        If ($VMInfo.State -ne "Off") {
            Stop-VM -Name $ToBeRemoved.Name -Force
            Start-Sleep -Seconds 5
        }
        Remove-VM -Name $ToBeRemoved.Name -Force -Confirm:$false
        Remove-Item -Path $(Resolve-Path $($VmPath -join("\"))).Path -Recurse -Force -Confirm:$false
    }

    $exit = (Read-Host -Prompt "Remove more VMs ?").ToLower()
}
until($exit -eq "n")
