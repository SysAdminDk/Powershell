<#
    .DESCRIPTION
    Export all computer objects to CSV file, for inventory and review.

#>


# --
# Where to store the csv files.
# --
$BasePath = "C:\Temp\CSV-Files"

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
# List All Servers
# --
Get-ADComputer -Filter "operatingSystem -like 'Windows Server*'" -Properties operatingSystem,OperatingSystemVersion,PasswordLastSet,LastLogonDate,whenCreated | `
Select-Object Name, OperatingSystem, OperatingSystemVersion, Enabled, PasswordLastSet, LastLogonDate, whenCreated, DistinguishedName | `
Export-Csv -Path "$FilePath\All-Servers.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8


# --
# List All Computers
# --
Get-ADComputer -Filter "operatingSystem -NotLike 'Windows Server*'" -Properties operatingSystem,OperatingSystemVersion,PasswordLastSet,LastLogonDate,whenCreated
Select-Object Name, OperatingSystem, OperatingSystemVersion, Enabled, PasswordLastSet, LastLogonDate, whenCreated, DistinguishedName | `
Export-Csv -Path "$FilePath\All-Computers.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8
