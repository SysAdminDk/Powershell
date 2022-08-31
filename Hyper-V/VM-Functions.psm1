# --
# List VM's
# --
function select-VirtualMachine {
    Param (
          [Parameter(Mandatory=$false)][String]$Hostname = "localhost"
          )

    $VirtualMachines = Get-VM
    Clear-Host
    Write-Output "----- Virtual Machines on $Hostname -----"
    for ($i=1; $i -lt ($VirtualMachines.Length+1); $i++) { 
        Write-Host "[$i] : $($VirtualMachines[$i-1].Name)"
    }

    Write-Output "-----------------------------------------"
    $selection = Read-Host "Select Virtual Machine"

    Return ($VirtualMachines[$selection-1])
}


# --
# Create VM - Funcction
# --
function New-CreateVM {
    Param (
          [Parameter(Mandatory=$false)][String]$ServerName,
          [Parameter(Mandatory=$false)][String]$MemoryMax,
          [Parameter(Mandatory=$false)][String]$CpuCount,
          [Parameter(Mandatory=$false)][String]$DefaultPath,
          [Parameter(Mandatory=$false)][String]$Defaultwitch
          )

    if ((Get-VM -Name "$ServerName" -ErrorAction SilentlyContinue) -ne $null) {
        Throw "$ServerName alreay exists, please use another name"
    }

    Try {
        New-VM –Name $ServerName -Generation 2 –MemoryStartupBytes 1024MB -Path "$DefaultPath" -switchname $Defaultwitch -NoVHD | Out-Null
        Set-VMMemory –VMName $ServerName -DynamicMemoryEnabled $true -MaximumBytes (1Gb*$MemoryMax) -MinimumBytes (1Gb*1) -StartupBytes (1Gb*1) | Out-Null
        Set-VMProcessor –VMName $ServerName –count $CpuCount | Out-Null
        Set-VMNetworkAdapterVlan -VMName $ServerName -VlanId 8 -Access
    } catch {
        Write-Error "Unable to create VM"
    }
}


# --
# Create Virtual Disk - Funcction
# --
function New-CreateDrive {
    Param (
          [Parameter(Mandatory=$false)][String]$ServerName,
          [Parameter(Mandatory=$false)][String]$DiskSize
          )

    $ServerPath = (Get-VM -Name "$ServerName" -ErrorAction SilentlyContinue).Path
    if ($ServerPath -eq $Null) {
        Throw "Unable to find path of virtual machine : $ServerName"
    }

    # Get Next drive number
    $ExistingDrive = ((Get-ChildItem -Path "$ServerPath\Virtual Hard Disks" -Filter *.vhdx -ErrorAction SilentlyContinue).Name)
    if ($ExistingDrive -eq $Null)  {
        $NextAvail = 0
    } else {
        $Numbers = ($ExistingDrive -replace $ServerName) -Replace '\D+'
        $NextAvail = Compare $(0..63) $Numbers -PassThru | Select-Object -first 1
    }


    $VHDxFile = "$ServerPath\Virtual Hard Disks\$ServerName`_Disk_$NextAvail.vhdx"
    $VHDxFileParrent = Split-Path -Path $VHDxFile -Parent

    if (!(Test-Path $VHDxFileParrent)) {
        New-Item -ItemType directory -Path $VHDxFileParrent | Out-Null
    }

    if ( (Test-Path $VHDxFileParrent) -and (!(Test-Path $VHDxFile)) ) {
        New-VHD -Path $VHDxFile -SizeBytes (1GB*$DiskSize) | Out-Null
    }
    if (Test-Path $VHDxFile) {
        # --
        # Find next avalible controler location
        # --
        $UsedSaSId = (Get-VMHardDiskDrive -VMName $ServerName).ControllerLocation
        if ($UsedSaSId -eq $Null)  {
            $NextSaSId = 0
        } else {
            $NextSaSId = Compare $(0..63) $Numbers -PassThru | Select-Object -first 1
        }
        
        Add-VMHardDiskDrive -VMName $ServerName -ControllerNumber 0 -ControllerLocation $NextSaSId –Path $VHDxFile | Out-Null
    }
#    Return $True
}

Export-ModuleMember Select-VirtualMachine, New-CreateVM, New-CreateDrive



# --
# Test GUI for the script
# --
