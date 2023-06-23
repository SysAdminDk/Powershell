<#
    .DESCRIPTION
     Save ALL Acl's from selected objects found in Active Directory

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


#--
# Import module(s)
# --
Import-Module ActiveDirectory


#-- 
# Get Domain information
# --
$ADRootDSE = Get-ADRootDSE
$DomainInfo = Get-ADDomain


# --
# List ALL Objects in Active Directory
# --
$ObjectList = @()
$ObjectList += "$($DomainInfo.DistinguishedName)"
$ObjectList += Get-ADObject -Filter "ObjectClass -eq 'organizationalUnit' -or ObjectClass -eq 'group' -or ObjectClass -eq 'user' -or ObjectClass -eq 'computer'" -SearchBase "$($DomainInfo.DistinguishedName)"


# --
# Get permissions of selected objects in the Domain
# --
$Accesslist = @()
foreach ($Object in $ObjectList) {

    # --
    # Verify that we can find the object. (The script have extended runtime, there can be made changes while the script is running)
    # --
    try {
        $VerifyObject = Get-ADObject -Identity $($Object.DistinguishedName)
        $objectPath = "Microsoft.ActiveDirectory.Management.dll\ActiveDirectory:://RootDSE/$($VerifyObject.DistinguishedName)"
    } catch {
        #Write-Output $_
        continue
    }


    # --
    # Get Object ACL
    # --
    $ACLList = Get-Acl -Path $objectPath


    # --
    # Prepare the Export
    # --
    foreach ($ACL in $($ACLList.Access)) {

        if ($ACL.IsInherited -eq $False) {

            # --
            # Prepare next row in export csv
            # --
            $ExportAcl = New-Object -TypeName psobject

            # --
            # Add the ACL properties to row in export csv
            # --
            $ExportAcl | Add-Member -MemberType NoteProperty -Name "ADObject" -Value $Object.DistinguishedName
            $ExportAcl | Add-Member -MemberType NoteProperty -Name "ADObjectType" -Value $Object.ObjectClass
            $ExportAcl | Add-Member -MemberType NoteProperty -Name "ActiveDirectoryRights" -Value $ACL.ActiveDirectoryRights
            $ExportAcl | Add-Member -MemberType NoteProperty -Name "InheritanceType" -Value $ACL.InheritanceType

            $ExportAcl | Add-Member -MemberType NoteProperty -Name "ObjectTypeGuid" -Value $ACL.ObjectType.Guid

            $ExportAcl | Add-Member -MemberType NoteProperty -Name "InheritedObjectTypeGuid" -Value $ACL.InheritedObjectType.Guid

            $ExportAcl | Add-Member -MemberType NoteProperty -Name "ObjectFlags" -Value $ACL.ObjectFlags
            $ExportAcl | Add-Member -MemberType NoteProperty -Name "AccessControlType" -Value $ACL.AccessControlType

            # --
            # Convert well-known SIDs to account name
            # --
            if ($($ACL.IdentityReference.Value) -like "S-1-5-32-*") {
                $ExportAcl | Add-Member -MemberType NoteProperty -Name "IdentityReference" -Value (Get-ADObject -Filter "objectSid -eq '$($ACL.IdentityReference.Value)'").Name
                $ExportAcl | Add-Member -MemberType NoteProperty -Name "IdentityReferenceSid" -Value $ACL.IdentityReference.Value
            } else {
                $ExportAcl | Add-Member -MemberType NoteProperty -Name IdentityReference -Value $ACL.IdentityReference.Value
                try {
                    # --
                    # Convert Account Name to SID
                    # --
                    $ObjectSID = (Get-ADObject -Filter "Name -eq '$((($ACL.IdentityReference.Value) -Split("\\"))[1])'" -Properties objectSid).objectSid.Value
                    if (!($ObjectSID)) {
                        $ExportAcl | Add-Member -MemberType NoteProperty -Name IdentityReferenceSid -Value "NULL"
                    } else {
                        $ExportAcl | Add-Member -MemberType NoteProperty -Name IdentityReferenceSid -Value $ObjectSID
                    }
                } catch {
                }
            }

            $ExportAcl | Add-Member -MemberType NoteProperty -Name IsInherited -Value $ACL.IsInherited
            $ExportAcl | Add-Member -MemberType NoteProperty -Name "InheritanceFlags" -Value $ACL.InheritanceFlags
            $ExportAcl | Add-Member -MemberType NoteProperty -Name "PropagationFlags" -Value $ACL.PropagationFlags

            # --
            # Add row to export csv
            # --
            $Accesslist += $ExportAcl
        }
    }
}


# --
# Save CSV
# --
$Accesslist | Export-Csv "$FilePath\ADObject_ACL_list.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation
