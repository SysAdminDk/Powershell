<#
    .DESCRIPTION
    Export all groups with listed members, to CSV file, for inventory and review.

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
# List All groups
# --
$Groups = Get-ADGroup -Filter * -Properties member

$GroupList = @()
Foreach ($Group in $Groups) {

    Foreach ($Member in $($Group.member)) {
        $ExportGroups = New-Object -TypeName psobject
        $ExportGroups | Add-Member -MemberType NoteProperty -Name "Name" -Value $Group.Name
        $ExportGroups | Add-Member -MemberType NoteProperty -Name "Category" -Value $Group.GroupCategory
        $ExportGroups | Add-Member -MemberType NoteProperty -Name "Scope" -Value $Group.GroupScope
        $ExportGroups | Add-Member -MemberType NoteProperty -Name "Member" -Value $Member
        $ExportGroups | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $Group.DistinguishedName
        
        $GroupList += $ExportGroups
    }
}

$GroupList | Export-Csv -Path "$FilePath\Groups_With_Members.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8
