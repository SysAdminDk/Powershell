<#
    .SYNOPSIS
    Extract Free / Busy data from local Outlook  profile.

    .DESCRIPTION
    Quick (fairly) code to extract free busy data from local outlook, and present in a excel sheet to make the bosses happy
    Perhaps an update will come later to use the olbusystatus to indicate Driving / Working @Customer/ Working @ Home

    .LINK
    https://learn.microsoft.com/en-us/office/vba/api/outlook.recipient.freebusy
    https://learn.microsoft.com/en-us/office/vba/api/outlook.olbusystatus


    .PARAMETER MaxWeeks
    Defines the number of weeks into the future the scripts extracts freebusy data.

    .PARAMETER Outfile
    Defines path to and filename of the Excel file gennerated by this script.
    - Defaults to C:\Users\Username\OneDrive**\Desktop\FreeBusy-date*.xlsx
    
    .PARAMETER WorkHours
    Defines Working Hours
    - Defaults to 0800 to 1600

    .PARAMETER NightHours
    Defines Hours of the day where I normaly sleep.
    - Defaults to 2200 - 0600

    .PARAMETER Cleanup
    Remove the temp CSV file
    The temp CSV file is saved as $ENV:Temp\FreeBusy-date*.csv


     * Date of the fisrt monday after the script have been run.
     ** If onedrive is installed, and desktop is synced.



    .EXAMPLE
    Just run the script, verbose can be used!
    .\Get-OutlookFreeBusy.ps1 -MaxWeeks 10 -Verbose
    .\Get-OutlookFreeBusy.ps1  -MaxWeeks 10 -Outfile "Path\To\file.xlsx" -WorkHours 08:00-16:00 -Cleanup Yes

#>
param (
    [Parameter(ValueFromPipeline)][Int]$MaxWeeks = 10,
    [Parameter(ValueFromPipeline)][String]$WorkHours = "08:00-16:30",
    [Parameter(ValueFromPipeline)][String]$NightHours = "22:00-06:00",
    [parameter(ValueFromPipeline)][string]$Outfile = "$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop))\FreeBusy-$((Get-date).ToShortDateString()).xlsx",
    [Parameter(ValueFromPipeline)][ValidateSet("Yes","No")]$Cleanup="No"
)

# Set next Monday as start date for FreeBusy Qyery
Write-Verbose "Find next Monday as start of week"
$Today = Get-date
While ($Today.DayOfWeek -ne "Monday") { $Today = $Today.AddDays(1) }
Write-Verbose "Query start date $($Today.ToShortDateString())"
Write-Verbose "Query end date $($Today.AddDays($MaxWeeks*7))"
Write-Verbose "Workdays retrived from Outlook $($MaxWeeks*5)"

# Connect to local outlook, to extract free/busy from local account.
Write-Verbose "Connecting to local Outlok to get Free/bussy data"
$Outlook = New-Object -ComObject Outlook.Application -Verbose:$False
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$Namespace = $Outlook.GetNamespace('MAPI')

Write-Verbose "Extract username of local Outlook Profile"
$Recipient = ($Namespace.Accounts | Select-Object -Property DisplayName)[0]
$objRecipient = $Namespace.CreateRecipient($Recipient.DisplayName)

# Get Free/Busy data for selected weeks
Write-Verbose "Get Free/bussy data for $MaxWeeks Weeks"
$FreeBusyWeeks = @()
for ($w=0; $w -lt $MaxWeeks; $w++) {
    $AddDays = $($w * 7)
    $QyeryDate = $Today.AddDays($(($AddDays)))
    $Week = $objRecipient.FreeBusy("#$(Get-Date -Date $($QyeryDate) -UFormat "%m/%d/%Y")#", 30, $False)
    $oWeek = New-Object PSObject
    $oWeek | Add-Member -type NoteProperty -Name "Date" -Value $($QyeryDate.ToShortDateString())
    $oWeek | Add-Member -type NoteProperty -Name "WeekData" -Value $($Week.Substring(0,240))
    $FreeBusyWeeks += $oWeek
}

# Prepare CSV Array
$OutCSVArray = @()

# Loop thru the Weeks and find free busy times.
foreach ($FreeBusyWeek in $FreeBusyWeeks) {
    Write-Verbose "Get week start date, and split Free/busy time to 30 min parts"
    $WeekStartDate = $(Get-Date -Date $FreeBusyWeek.Date)
    $FreeBusyData = $FreeBusyWeek.WeekData -split "(.{48})" -ne ""

    foreach ($FreeBusyarray in $FreeBusyData) {

        Write-Verbose "Create Workday Free/Busy object - $(($WeekStartDate.Date).ToShortDateString())"
        $OutCSVItem = New-Object PSObject
        $OutCSVItem | Add-Member -type NoteProperty -Name "Date" -Value $($WeekStartDate.Date)

        $utilization = $(((100 / 8) * ( (($FreeBusyArray) | % { $_ -split "(.)" -ne "" | measure-object -sum }).sum / 2) ) /100 ).ToString("P")
        $OutCSVItem | Add-Member -type NoteProperty -Name "Utilization" -Value $utilization

        for ($i=0; $i -lt ($FreeBusyarray).Length; $i++) {

            $TimeOfDay = $(Get-Date -Date (Get-Date -Date "00:00:00").AddMinutes($i*30) -Format "HH:mm")

            if (($FreeBusyArray[$i]) -eq "1") {
                $FreeBusy = "Busy"
            } else {
                $FreeBusy = "Free"
            }

            if ( ($TimeOfDay -le "$(($WorkHours -Split("-"))[0])") -OR ($TimeOfDay -ge "$(($WorkHours -Split("-"))[1])") )  {
                if ($FreeBusy -eq "free") {
                    $OutCSVItem | Add-Member -type NoteProperty -Name $TimeOfDay -Value ""
                } else {
                    $OutCSVItem | Add-Member -type NoteProperty -Name $TimeOfDay -Value "*$FreeBusy"
                }
            } else {
                $OutCSVItem | Add-Member -type NoteProperty -Name $TimeOfDay -Value $FreeBusy
            }

        }
        $OutCSVArray += $OutCSVItem

        $WeekStartDate = $WeekStartDate.AddDays(1)
    }
}
Write-Verbose "Save temp CSV file for import to Excel"
$OutCSVArray | Export-Csv -Path "$env:TEMP\FreeBusy-$(($Today).ToShortDateString()).csv" -NoTypeInformation -Delimiter ";" -Encoding Unicode -Force


# Create the Excel sheet.
Write-Verbose "Create new Excel workbook"
$Excel = New-Object -ComObject excel.application -Verbose:$False
#if ($VerbosePreference -eq "Continue") {
    $Excel.Visible = $True
#} else {
#    $Excel.Visible = $False
#}
$Excel.DisplayAlerts = $False

$Workbook = $Excel.Workbooks.Add()
$Workbook.Title = "Free/Busy Times"

$WS = $Workbook.Worksheets
$WS = $WS.Item(1)
$WS.Name = "Free|Busy Times"

# Select Worksheet
$WS = $Workbook.Worksheets  | Where {$_.Name -eq "Free|Busy Times"}
$WS.Select()

# Open the CSV file, Import data from CSV and remove the connection
Write-Verbose "Connect to the temp CSV and import data"
$TxtConnector = ("TEXT;" + "$env:TEMP\FreeBusy-$(($Today).ToShortDateString()).csv")
$CellRef = $WS.Range("A1")
$Connector = $WS.QueryTables.add($TxtConnector,$CellRef)
$WS.QueryTables.item($Connector.name).TextFileCommaDelimiter = $false
$WS.QueryTables.item($Connector.name).TextFileSemicolonDelimiter = $True
$WS.QueryTables.item($Connector.name).TextFileParseType  = 1

#$WS.QueryTables.item($Connector.name).TextFileDecimalSeparator = "."
#$WS.QueryTables.item($Connector.name).TextFileThousandsSeparator = "," 

$WS.QueryTables.item($Connector.name).Refresh() | Out-Null
$WS.QueryTables.item($Connector.name).delete()

# Autofit the columns, freeze the top row
Write-Verbose "Change size of the Colums"
$WS.UsedRange.EntireColumn.AutoFit() | Out-Null
$WS.Application.ActiveWindow.SplitRow = 1
$WS.Application.ActiveWindow.SplitColumn = 1
$WS.Application.ActiveWindow.FreezePanes = $true

# Set formating colur to the cels based on content.
Write-Verbose "Add coloring"
$Selection = $WS.Range("C2:$($WS.UsedRange.SpecialCells(11).Address($false,$false))")
$Selection.FormatConditions.Add(2, 0, "=LEFT(C2;LEN(`"Free`"))=`"Free`"") | Out-Null
$Selection.FormatConditions.Item($Selection.FormatConditions.Count).Interior.Color = 65280
$Selection.FormatConditions.Add(2, 0, "=LEFT(C2;LEN(`"Busy`"))=`"Busy`"") | Out-Null
$Selection.FormatConditions.Item($Selection.FormatConditions.Count).Interior.Color = 255
$Selection.FormatConditions.Add(2, 0, "=LEFT(C2;LEN(`"*Busy`"))=`"*Busy`"") | Out-Null
$Selection.FormatConditions.Item($Selection.FormatConditions.Count).Interior.Color = 255
$Selection.FormatConditions.Item(3).Font.FontStyle = "Bold Italic"

$Selection = $WS.Columns.item('B')
$Selection.FormatConditions.Add(1, 5, "=0,9") | Out-Null
$Selection.FormatConditions.Item($Selection.FormatConditions.Count).Font.FontStyle = "Bold Italic"
$Selection.FormatConditions.Item($Selection.FormatConditions.Count).Interior.Color = 9359529
$Selection.FormatConditions.Add(1, 5, "=0,7") | Out-Null
$Selection.FormatConditions.Item($Selection.FormatConditions.Count).Interior.Color = 9359529


# Hide selected Night hours
#Write-Verbose "Hide selected NightHours"
for ($i=3; $i -le $($WS.UsedRange.rows.count); $i++) {
    if ( ($($WS.cells.Item(1, $i).text) -ge $($NightHours -split("-"))[0]) -or ($($WS.cells.Item(1, $i).text) -lt $($NightHours -split("-"))[1]) ) {
        $ws.Columns.Item($i).hidden=$true
    }
}

# Set date & percentage format
Write-Verbose "Change format of date and util colums"
$WS.Columns.item('B').NumberFormat = "0,00%"
$WS.Columns.item('A').NumberFormatlocal = "dddd - d. mmmm åååå"

# Save the Workbook
if ($Cleanup -eq "Yes") {
    Write-Verbose "Save and close excel"
    $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
    $Workbook.SaveAs($OutFile, $xlFixedFormat)
    $Workbook.Close()
    $Excel.Quit()
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

# Cleanup
if ($Cleanup -eq "Yes") {
    Write-Verbose "Cleanup temp CSV file"
    Remove-Item "$env:TEMP\FreeBusy-$(($Today).ToShortDateString()).csv"
}
