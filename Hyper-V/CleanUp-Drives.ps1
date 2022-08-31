# --
# Get Hyper-V Directory configuration
# --
$DefaultHardDiskPath = (Get-VMHost).VirtualHardDiskPath
$DefaultMachinePath = (Get-VMHost).VirtualMachinePath


# --
# Move VM's if folder Dont match name
# --
$VirtualMachines = Get-VM
foreach ($VirtualMachine in $VirtualMachines) {
    $VmConfigPath = "$DefaultMachinePath\$($VirtualMachine.Name)"
    $VmDiskPath = "$DefaultHardDiskPath\$($VirtualMachine.Name)\Virtual Hard Disks"
    
    if ($($VirtualMachine.ConfigurationLocation) -ne $VmConfigPath) {
        Write-Host "Move Config > $VmConfigPath"
        Move-VMStorage -Name $VirtualMachine.Name -DestinationStoragePath "$VmConfigPath"
    }

    $VMDisks = Get-VMHardDiskDrive -VMName $VirtualMachine.Name
    foreach ($VmDisk in $VMDisks) {
        if ($($VMDisk.path) -Notlike "$DefaultHardDiskPath\$($VirtualMachine.Name)\*") {
            $VHDFileName = Split-Path $($VmDisk.Path) -Leaf

            Write-Host "Move Disk $($VMDisk.path) > $VmDiskPath\$VHDFileName"

            Move-VMStorage -Name $VirtualMachine.Name -Vhds @{"SourceFilePath"=$($VmDisk.Path);"DestinationFilePath"="$VmDiskPath\$VHDFileName"}
        }
    }
}


# --
# Get all drives in server where there are Virtual Machine folders
# --
#$VirtualMachinesPath = @()
#foreach ($Drive in (Get-Volume).DriveLetter) {
#    if (Test-Path "$Drive`:\Virtual Machines") {
#        $VirtualMachinesPath += Get-ChildItem -Path "$Drive`:\Virtual Machines" -Depth 0
#    }
#}

# --
# If No VM exists, remove the folder
# --
#foreach ($Directory in $VirtualMachinesPath) {
#    if (!(Get-VM -Name $Directory.name -ErrorAction SilentlyContinue)) {
#        Write-Host "Remove Folder $($Directory.FullName)"
#        #Remove-Item -Path $($Directory.FullName) -Recurse -Force
#    }
#}


#foreach ($Drive in (Get-Volume).DriveLetter) {
#    Get-ChildItem -Directory -Path $Drive -Recurse | Where { (Get-ChildItem -Path $_.FullName).Count -ne 0 } | select -expandproperty FullName
#
#
#
#}


#(Get-ChildItem D:\ -Recurse -Directory | ? { -Not ($_.EnumerateFiles('*',1) | Select-Object -First 1) }).FullName #| Remove-Item -Recurse
