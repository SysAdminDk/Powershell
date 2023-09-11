<#
    .NOTES
        Name    : Import-CIS-Baselines.ps1
        Author  : Jan Kristensen (Truesec)

        Version : 1.0
        Date    : 28-03-2023

    .DESCRIPTION
        Import selected security baselines from CIS Microsoft Windows Server Build Kits

        Files are downloaded from here, please note that you must have a subscription to access the files !

        2012   : https://workbench.cisecurity.org/files/3903/download
        2012R2 : https://workbench.cisecurity.org/files/3901/download
        2016   : https://workbench.cisecurity.org/files/3864/download
        2019   : https://workbench.cisecurity.org/files/3793/download
        2022   : https://workbench.cisecurity.org/files/4290/download


    .PARAMETER Path 
        Specifies where the files will be extracted and imported from

    .PARAMETER Cleanup 
        Specifies whether to remove the extracted files after the script have run.

    .EXAMPLE
        .\Import-MSFT-Baselines.ps1 -Path "C:\Windows\temp" -Cleanup Yes

    .EXAMPLE
        .\Import-MSFT-Baselines.ps1 -Path "C:\Windows\temp" -Cleanup Yes

    .EXAMPLE
        .\Import-MSFT-Baselines.ps1 -Path "C:\Windows\temp" -Cleanup Yes

#>
# 
# Request required script options.
# 
param (
    [Parameter(ValueFromPipeline)]
    $Path=$PSScriptRoot,

    [Parameter(ValueFromPipeline)]
    [ValidateSet("Yes","No")]
    $Cleanup="Yes"

)


#
# Create Output folders, if not exists
#
if (!(Test-Path $Path)) {
    Write-Verbose "Create download directory"
    New-Item -Path $Path -ItemType Directory -Force | Out-Null
}


#
# Extract GPO baselines
#
Write-Verbose "Extract the Group Policy files"
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

# Get ZIP files to extract
$Files = (Get-ChildItem -Path "$Path\*.zip")
foreach ($File in $Files) {
    $DestinationFolder = $($File.Name -replace(".zip"))
    $ZipFile = [IO.Compression.ZipFile]::OpenRead($File.FullName)

    $ZipFile.Entries | ? { $_.FullName -like "*{*}*" } | ForEach-Object {
        $OutFile = Join-Path $Path $(Join-Path $DestinationFolder "{$(($_.FullName -split("{"))[1])")
        if (!(Test-Path -LiteralPath $(Split-Path $OutFile -Parent))) {
            New-Item -Path $(Split-Path $OutFile -Parent) -ItemType Directory -Force | Out-Null
        }

        if ($_ -notlike "*/") {
            [IO.Compression.ZipFileExtensions]::ExtractToFile($_, $OutFile, $true)
        }
    }
    $ZipFile.Dispose()
}


#
# Select Policy to import
#
Write-Verbose "List all avaliable policy files"
$GPOList = Get-ChildItem -Path $Path -Recurse -Directory -Filter "{*}"
if ($GPOList.Length -eq 0) {
    Write-Error "Unable to find Policy to import"
    break
}
$GPOMap = @()
Foreach ($GPO in $GPOList) {
    $GPOMap += New-Object -Type PSObject -Property @{
        'Guid'  = $($GPO.Name)
        'Package' = $($GPO.FullName).Replace("$Path\","").Split("\\")[0]
        'Name' = $(([XML](Get-Content -Path "$($GPO.FullName)\backup.xml")).GroupPolicyBackupScheme.GroupPolicyObject.GroupPolicyCoreSettings.DisplayName.InnerText) -replace("SCM ","MSFT ")
    }

}
Write-Verbose "Show GPO list, please select which GPOs to import"
$Selected = $($GPOMap | Select-Object -Property "Name","Guid","Package" | Sort-Object -Descending -Property "Package" | Out-GridView -OutputMode Multiple -Title "Select Group Policy(s) to import")
if ($Selected.Length -eq 0) {
    Write-Error "Please select which GPOs to import"
    break
}


#
# Import selected GPOs 
#
Foreach ($GPO in $Selected) {
    $GpoPath = (Get-ChildItem -Path $Path -Recurse | Where {$_.Name -eq $($GPO.Guid)}).Parent
    Write-Verbose "Import GPO : `"$($GPO.Name)`""
    try {
        Import-GPO -BackupId $GPO.Guid -Path $GpoPath.FullName -TargetName "$($GPO.Name)" -CreateIfNeeded | Out-Null
    } catch {
        Write-Output "Unable to import GPO, please verify that you user have the required permissions"
        Write-Output $_
    }
}


#
# Cleanup
#
if ($Cleanup -eq "Yes") {
    Write-Verbose "Cleanup policy folders"
    Get-ChildItem -Path $Path -Directory -Exclude "*.zip","*.ps1" | Remove-Item -Recurse -Force
}