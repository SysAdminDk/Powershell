# --
# Set default variables
# --
$DefaultWMIPath = "CN=SOM,CN=WMIPolicy,CN=System,$((Get-ADRootDSE).defaultNamingContext)"
$msWMIAuthor = "Administrator@" + [System.DirectoryServices.ActiveDirectory.Domain]::getcurrentdomain().name
$NewWMIOwner = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name


# --
# Build the date field in required format
# --
$now = (Get-Date).ToUniversalTime()
$msWMICreationDate = ($now.Year).ToString("0000") + ($now.Month).ToString("00") + ($now.Day).ToString("00") + ($now.Hour).ToString("00") + ($now.Minute).ToString("00") + ($now.Second).ToString("00") + "." + ($now.Millisecond * 1000).ToString("000000") + "-000"


# --
# Get WMI filters from file.
# --
#$Data = Get-Content -Path "Q:\Active Directory\WMI-Filters\WMI Query.txt"
$Data = Get-Content -Path "$PSScriptRoot\WMI Query.txt"

# --
# Create WMI filters and Set permissions
# --
foreach ($Line in $Data) {
    $Name = $($Line -Split("; "))[0]
    $Query = $($Line -Split("; "))[1]

    $NewWMIGUID = [string]"{" + ([System.Guid]::NewGuid()) + "}"
    $NewWMIDN = "CN=$NewWMIGUID,$DefaultWMIPath"
    $NewWMICN = $NewWMIGUID
    $NewWMIdistinguishedname = $NewWMIDN
    $NewWMIID = $NewWMIGUID

    $Attr = @{
        "msWMI-Name" = $Name;
        "msWMI-Parm1" = "Created by script";
        "msWMI-Parm2" = $msWMIParm2 = "1;3;10;" + $Query.Length.ToString() + ";WQL;root\CIMv2;" + $Query + ";"
        "msWMI-Author" = $msWMIAuthor;
        "msWMI-ID" = $NewWMIID;
        "instanceType" = 4;
        "showInAdvancedViewOnly" = "TRUE";
        "distinguishedname" = $NewWMIdistinguishedname;
        "msWMI-ChangeDate" = $msWMICreationDate;
        "msWMI-CreationDate" = $msWMICreationDate
        }

    if (!(Get-ADObject -Filter { msWMI-Name -eq $Name })) {
        Write-Output "Create WMI filter - $Name"

        New-ADObject -name $NewWMICN -type "msWMI-Som" -Path $DefaultWMIPath -OtherAttributes $Attr | Out-Null
    }
}
