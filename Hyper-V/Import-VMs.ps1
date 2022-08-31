$VMinport = (Get-ChildItem "D:\Virtual Machines").Name

foreach ($VMName in $VMinport) {
    if (!(Get-VM -Name $VMName -ErrorAction SilentlyContinue)) {
        Write-Host "Create $VMName"
        Import-VM -Path $(Get-ChildItem -Path "D:\Virtual Machines\$VMName\Virtual Machines\*vmcx").FullName -Register -Verbose
    }
}
