## --------------------------------------------------------------------------------------------------
# Set Variables
## --------------------------------------------------------------------------------------------------
$DiskImage	= "E:\Shares\ISO Images\en_windows_10_business_editions_version_1809_updated_sep_2020_x64_dvd_f7317bec.iso"
$VHDXFile	= "D:\Virtual Machines\_Reference\Windows_OS_Disk.vhdx"

## --------------------------------------------------------------------------------------------------
# Create Refrence VIM
## --------------------------------------------------------------------------------------------------
if (Test-Path $DiskImage) {
	Mount-DiskImage -ImagePath $DiskImage
	$MountDrive = $((Get-DiskImage -ImagePath $DiskImage | get-volume).DriveLetter) + ":"
}

Get-WindowsImage -ImagePath "$MountDrive\Sources\install.wim" | ft ImageIndex, ImageName

if (!(Test-Path $VHDXFile)) {
	# Create New VHDx
	New-VHD -Path $VHDXFile -Dynamic -SizeBytes 50Gb | Out-Null
	Mount-DiskImage -ImagePath $VHDXFile
	$VHDXDisk = Get-DiskImage -ImagePath $VHDXFile | Get-Disk
	$VHDXDiskNumber = [string]$VHDXDisk.Number

	# Create Partitions
	Initialize-Disk -Number $VHDXDiskNumber –PartitionStyle GPT -Verbose
	$VHDXDrive1 = New-Partition -DiskNumber $VHDXDiskNumber -GptType '{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}' -Size 499MB
	$VHDXDrive1 | Format-Volume -FileSystem FAT32 -NewFileSystemLabel System -Confirm:$false | Out-Null
	$VHDXDrive2 = New-Partition -DiskNumber $VHDXDiskNumber -GptType '{e3c9e316-0b5c-4db8-817d-f92df00215ae}' -Size 128MB
	$VHDXDrive3 = New-Partition -DiskNumber $VHDXDiskNumber -GptType '{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}' -UseMaximumSize
	$VHDXDrive3 | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$false | Out-Null
	Add-PartitionAccessPath -DiskNumber $VHDXDiskNumber -PartitionNumber $VHDXDrive1.PartitionNumber -AssignDriveLetter
	$VHDXDrive1 = Get-Partition -DiskNumber $VHDXDiskNumber -PartitionNumber $VHDXDrive1.PartitionNumber
	Add-PartitionAccessPath -DiskNumber $VHDXDiskNumber -PartitionNumber $VHDXDrive3.PartitionNumber -AssignDriveLetter
	$VHDXDrive3 = Get-Partition -DiskNumber $VHDXDiskNumber -PartitionNumber $VHDXDrive3.PartitionNumber
	$VHDXVolume1 = [string]$VHDXDrive1.DriveLetter+":"
	$VHDXVolume3 = [string]$VHDXDrive3.DriveLetter+":"

	#Extract Server 2016 Standard, and apply to VHDx
	Expand-WindowsImage -ImagePath "$MountDrive\Sources\install.wim" -Index 3 -ApplyPath $VHDXVolume3\ -ErrorAction Stop -LogPath Out-Null

	# Apply BootFiles
	cmd /c "$VHDXVolume3\Windows\system32\bcdboot $VHDXVolume3\Windows /s $VHDXVolume1 /f ALL"

	# Change ID on FAT32 Partition
	$DiskPartTextFile = New-Item "diskpart.txt" -type File -force
	Set-Content $DiskPartTextFile "select disk $VHDXDiskNumber"
	Add-Content $DiskPartTextFile "Select Partition 2"
	Add-Content $DiskPartTextFile "Set ID=c12a7328-f81f-11d2-ba4b-00a0c93ec93b OVERRIDE"
	Add-Content $DiskPartTextFile "GPT Attributes=0x8000000000000000"
	cmd /c "diskpart.exe /s .\diskpart.txt"

	Dismount-DiskImage -ImagePath $DiskImage
	Dismount-DiskImage -ImagePath $VHDXFile
}


#Copy-Item -Path "F:\*" -Destination "E:\Win 10 - Enterprise" -Recurse -Exclude @("*install.wim*")
#Get-WindowsImage -ImagePath "$MountDrive\Sources\install.wim" | ft ImageIndex, ImageName
#Export-WindowsImage -SourceImagePath "$MountDrive\sources\install.wim" -SourceIndex 3 -DestinationImagePath "E:\Win 10 - Enterprise\sources\Install.wim" -DestinationName "Windows 10 - Enterprise"
