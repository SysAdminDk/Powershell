<#
    .DESCRIPTION
    Download and import selected security baselines from Microsoft Security Compliance Toolkit

    .PARAMETER DownloadID
    Specifies the ID from Microsoft Download
    - If Download ID have changed please find latest by searching for "Microsoft Security Compliance Toolkit"
    - Curent URL = https://www.microsoft.com/en-us/download/details.aspx?id=55319
    - Curent ID = 55319

    .PARAMETER Path 
    Specifies where the dowloaded files will be saved, and extracted

    .PARAMETER Action
    Specifies which of the actions to preform.
    Download - Only download and extract til GPO files, requires internet access.
    Install - Only import the GPO files, requires the GPO folders to be avalible in Root of the Path, requires write access to Active Directory.
    DownloadAndInstall - Does both of the above actions, requires internet access and write access to Active Directory.

    .PARAMETER Cleanup 
    Specifies whether we remove the files when the script have run.

    .EXAMPLE
    .\Import-MSFT-Baselines.ps1 -DownloadID 55319 -Path "C:\Windows\temp" -Action Download -Cleanup Yes

    .EXAMPLE
    .\Import-MSFT-Baselines.ps1 -DownloadID 55319 -Path "C:\Windows\temp" -Action Install -Cleanup Yes

    .EXAMPLE
    .\Import-MSFT-Baselines.ps1 -DownloadID 55319 -Path "C:\Windows\temp" -Action DownloadAndInstall -Cleanup Yes

#>
# 
# Request required script options.
# 
param (
    [Parameter(ValueFromPipeline)]
    [string[]]$DownloadID=55319,

    [Parameter(ValueFromPipeline)]
    $Path=$PSScriptRoot,

    [Parameter(ValueFromPipeline)]
    [ValidateSet("Download","Install","DownloadAndInstall")]
    $Action="DownloadAndInstall",

    [Parameter(ValueFromPipeline)]
    [ValidateSet("Yes","No")]
    $Cleanup="Yes"

)


#
# Create Output folders, if not exists
#
if ( (!(Test-Path $Path)) -and ($Action -ne "Install") ) {
    Write-Verbose "Create download directory"
    New-Item -Path $Path -ItemType Directory -Force | Out-Null
}

if ($Action -ne "Install") {
    if (!(Test-Path "$Path\ZIP")) {
        Write-Verbose "Create temp ZIP directory"
        New-Item -Path "$Path\ZIP" -ItemType Directory -Force | Out-Null
    }


    #
    # Download MS Security Baselines
    #
    Write-Verbose "Download MSFT security baselines"
    $HTML = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$DownloadID"
    $HTML.Content -match 'downloadData={(?<data>.*)}}' | Out-Null

    Foreach ($URI in $($Matches[0] -split("url:") -split(",id") -match("https")) -replace("`"")) {
        $FileName = $URI.Split("/")[-1]
        Write-Verbose "Download $FileName"
        Invoke-WebRequest -Uri "$URI" -OutFile "$Path\ZIP\$FileName"
    }


    #
    # Extract GPO baselines
    #
    Write-Verbose "Extract the Group Policy files"
    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

    # Get ZIP files to extract
    $Files = (Get-ChildItem -Path "$Path\ZIP\*")
    foreach ($file in $Files) {
        $DestinationFolder = $($file.Name -replace(".zip"))
        $ZipFile = [IO.Compression.ZipFile]::OpenRead($file)

        $ZipFile.Entries | ? { $_.FullName -like "*/GPOs/{*"; } | `
        ForEach-Object {
            # Remove leading Directory from Path, if not GPO
            $ZipBasePath = ($_.FullName -split("/"))[0]
            if ($ZipBasePath -notlike "GPO*") {
                $DestinationFile = $_.FullName -replace("$ZipBasePath/","")
            } else {
                $DestinationFile = $_.FullName
            }
            $OutFile = Join-Path $Path $(Join-Path $DestinationFolder $($DestinationFile -replace("/","\")))

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
    # Remove ZIP files
    #
    if ($Cleanup -eq "Yes") {
        Remove-Item -Path "$Path\ZIP" -Recurse -Force
    }
}

#
# First part done, notify the we are done.
#
if ($Action -eq "Download") {
    Write-Output "Copy `"$Path`" content to server with write access to Active Directory for import of the baseline GPOs"
}


if ($Action -ne "Download") {
    #
    # Select Policy to import
    #
    Write-Verbose "List all avaliable policy files"
    $GPOList = Get-ChildItem -Path $Path -Recurse | Where {$_.name -like "{*}"}
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
    Write-Verbose "Show GPO list, please selct what to import"
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

}


#
# Cleanup
#
if ( ($Action -ne "Download") -and ($Cleanup -eq "Yes") ) {
    Write-Verbose "Cleanup policy folders"
    Get-ChildItem -Path $Path -Directory | Remove-Item -Recurse -Force
}
