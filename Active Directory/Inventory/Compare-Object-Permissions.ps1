<#
    .DESCRIPTION
     Compares Object Permissions exported today with latest Export file, and saves the DIFF to CSV for change log

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
# Get csv files for compare
# --
$Files = Get-ChildItem -Path $FilePath -Include "*ADObject_ACL_list.csv*" -Recurse | Sort-Object -Property CreationTimeUtc
$TodayFile = $Files | Where {$_.FullName -Like "*$FileDate*"}
$YesterdayFile = $Files | Where { $_.FullName -NotLike "*$FileDate*" } | Select-Object -Last 1

Write-Output "Today File : $($TodayFile.FullName)"
Write-Output "Previus File : $($YesterdayFile.FullName)"


# --
# Make sure that we can find both files
# --
If (!(Test-Path -Path $TodayFile.FullName)) {
    Throw "Cant find csv file from today ($TodayFile)"
}

If (!(Test-Path -Path $YesterdayFile.FullName)) {
    Throw "Cant find csv file from today ($LastWeekFile)"
}


# --
# Import the Files for Compare
# --
$Objects = @{
  ReferenceObject = (Get-Content -Path $TodayFile)
  DifferenceObject = (Get-Content -Path $YesterdayFile)
}


# --
# Compare the arrays
# --
$Diff = Compare-Object @Objects


# --
# Prepare the Output in CSV format
# --
$OutCSVDiff = @()
Foreach ($rule in $Diff) {
    $DiffObjects = ( ($rule.InputObject).Split(";") ).Replace('"','')

    $DiffAcl = New-Object -TypeName psobject
    
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "ADObject" -Value $DiffObjects[0]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "ADObjectType" -Value $DiffObjects[1]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "ActiveDirectoryRights" -Value $DiffObjects[2]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "InheritanceType" -Value $DiffObjects[3]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "ObjectTypeGuid" -Value $DiffObjects[4]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "InheritedObjectTypeGuid" -Value $DiffObjects[5]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "ObjectFlags" -Value $DiffObjects[6]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "AccessControlType" -Value $DiffObjects[7]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "IdentityReference" -Value $DiffObjects[8]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "IdentityReferenceSid" -Value $DiffObjects[9]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "IsInherited" -Value $DiffObjects[10]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "InheritanceFlags" -Value $DiffObjects[11]
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "PropagationFlags" -Value $DiffObjects[12]

    if ($rule.SideIndicator -eq "<=") {
        $DiffAcl | Add-Member -MemberType NoteProperty -Name "Action" -Value "Added"
    } Else {
        $DiffAcl | Add-Member -MemberType NoteProperty -Name "Action" -Value "Removed"
    }
    $DiffAcl | Add-Member -MemberType NoteProperty -Name "SideIndicator" -Value $rule.SideIndicator

    $OutCSVDiff += $DiffAcl
}


# --
# Save the Diff file
# --
$OutCSVDiff | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Out-File "$FilePath\$FileDate\Diff-From-Yesterday.csv"
