
<#
    .DESCRIPTION
     Delegate computer join rights to specified Group  

    .Requirements
     The group have to exist.

    .PARAMETER Name
     Define the group to delegate control to

    .PARAMETER Path
    Define the OU where the ACL is apllied to

#>

Param(
        [Parameter(Mandatory)]$Name = "Delegate Control - Join Computers",
        [Parameter(Mandatory)]$Path = "OU=OSDeploy,OU=Endpoints,OU=Company Endpoints,DC=Prod,DC=SysAdmins,DC=Dk"
    )


# Get need domain information
$DomainInfo = Get-ADDomain
$DomainName = $DomainInfo.NetBIOSName


try {
    # Verify the group exists
    Write-Verbose "Find required Group [$Name]"

    $Group = Get-ADGroup -Identity $Name
    $ACLAccount = New-Object System.Security.Principal.NTAccount "$DomainName","$($Group.Name)"
}
catch {
    Write-Output "The required group does not exist"
    break
}


try {
    # Verify the OU exists, and current ACL can be retrived
    Write-Verbose "Get current ACL of provided OU [$Path]"

    $ACL = Get-Acl -Path "AD:\$Path"
}
catch {
    Write-Output "Unable to get current ACL on provided OU"
    break
}

Write-Verbose "Update ACL with new rules"
$ACL.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $ACLAccount,'ReadProperty,WriteProperty','Allow',$([GUID]::Parse('4c164200-20c0-11d0-a768-00aa006e0529')),'Descendents',$([GUID]::Parse('bf967a86-0de6-11d0-a285-00aa003049e2'))))
$ACL.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $ACLAccount,'ReadProperty,WriteProperty','Allow',$([GUID]::Parse('bf967a68-0de6-11d0-a285-00aa003049e2')),'Descendents',$([GUID]::Parse('bf967a86-0de6-11d0-a285-00aa003049e2'))))
$ACL.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $ACLAccount,'Self','Allow',$([GUID]::Parse('72e39547-7b18-11d1-adef-00c04fd8d5cd')),'Descendents',$([GUID]::Parse('bf967a86-0de6-11d0-a285-00aa003049e2'))))
$ACL.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $ACLAccount,'Self','Allow',$([GUID]::Parse('f3a64788-5306-11d1-a9c5-0000f80367c1')),'Descendents',$([GUID]::Parse('bf967a86-0de6-11d0-a285-00aa003049e2'))))
$ACL.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $ACLAccount,'ExtendedRight','Allow',$([GUID]::Parse('00299570-246d-11d0-a768-00aa006e0529')),'Descendents',$([GUID]::Parse('bf967a86-0de6-11d0-a285-00aa003049e2'))))
$ACL.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $ACLAccount,'CreateChild, DeleteChild','Allow',$([GUID]::Parse('bf967a86-0de6-11d0-a285-00aa003049e2')),'All'))


try {
    Write-Verbose "Save updated ACL to provided OU"

    Set-ACl -Path "AD:\$Path" -AclObject $ACL
}
catch {
    Write-Output "Unable to save the new ACL"
}
