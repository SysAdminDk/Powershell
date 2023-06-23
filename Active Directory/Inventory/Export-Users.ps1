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
# List All Users with selected properties
# --
Get-AdUser -Filter "Enabled -eq 'True'" -Properties DisplayName, DistinguishedName, lastLogonTimestamp, PasswordNeverExpires, pwdLastSet, badPasswordTime, LockedOut, adminCount, whenChanged, whenCreated, Enabled |`
Select-Object -Property "Name", "DisplayName", "UserPrincipalName", "Enabled", "PasswordNeverExpires", `
@{N="lastLogonTime"; E={if ($_.LastLogontimestamp -eq $Null) { } else { [DateTime]::FromFileTimeUtc($_.LastLogontimestamp) } }}, `
@{N="PasswordLastSet"; E={ if ($_.pwdLastSet -eq $Null) { } else { [DateTime]::FromFileTimeUtc($_.pwdLastSet).ToString() } }}, `
@{N="badPasswordTime"; E={ if ($_.badPasswordTime -eq $Null) { } else { [DateTime]::FromFileTimeUtc($_.badPasswordTime).ToString() } }}, `
"LockedOut", "adminCount", "whenChanged", "whenCreated", "DistinguishedName" `
| Export-Csv -Path "$FilePath\All-Users.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8

# --
# List users with SPN's
# --
Get-AdUser -Filter "ServicePrincipalName -like '*'" -Properties Name, PasswordNeverExpires, pwdLastSet, Enabled, adminCount, ServicePrincipalNames, memberof |`
Select-Object -Property Name, PasswordNeverExpires, Enabled, `
@{N="PasswordLastSet"; E={ if ($_.pwdLastSet -eq $Null) { } else { [DateTime]::FromFileTimeUtc($_.pwdLastSet).ToString() } }}, `
@{N="SPNs"; E={$_ | Select-Object -ExpandProperty ServicePrincipalNames}}, `
@{N="MemberOff"; E={$_ | Select-Object -ExpandProperty memberof}} `
| Export-Csv -Path "$FilePath\SPN_Users.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8

# --
# List users with Password NOT reguired
# --
get-adobject -ldapfilter "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=32))" -properties whenCreated, samaccountname, useraccountcontrol, pwdLastSet, Enabled |`
Select-Object Name, Samaccountname, DistinguishedName, useraccountcontrol, `
@{N="PasswordLastSet"; E={ if ($_.pwdLastSet -eq $Null) { } else { [DateTime]::FromFileTimeUtc($_.pwdLastSet).ToString() } }}, whenCreated `
| Export-Csv -Path "$FilePath\UserAccountControl.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8
