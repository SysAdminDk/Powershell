
# --------------------------------------------------
# Encode - Create-Report.ps1
# --------------------------------------------------
$CreateReportPs1 = "# --
# Base path
# --
`$Root = `"C:\inetpub\PingCastle`"

#--
# Create Report folder
# --
`$FolderDate = Get-Date -Format `"dd-MM-yyyy - HH-mm`"
if (!(Test-Path -Path `"`$Root\`$FolderDate`")) {
    New-Item -Path `"`$Root\`$FolderDate`" -ItemType Directory | Out-null
}

#--
# Run PingCastle
#--
if (Test-Path -Path `"`$(`$ENV:ProgramFiles)\PingCastle\PingCastle.exe`") {
    Start-Process -FilePath `"`$(`$ENV:ProgramFiles)\PingCastle\PingCastle.exe`" -ArgumentList `"--healthcheck`" -NoNewWindow -WorkingDirectory `"`$Root\`$FolderDate`" -Wait
}"

$Bytes = [System.Text.Encoding]::UTF8.GetBytes($CreateReportPs1)
$EncodedText =[Convert]::ToBase64String($Bytes)
$EncodedText


# --------------------------------------------------
# Encode - Default.asp
# --------------------------------------------------
$DefaultASP = $test = "<%
Response.Write(`"<!DOCTYPE html>`")
Response.Write(`"<html lang='en'>`")
Response.Write(`"<head>`")
Response.Write(`"<meta charset='UTF-8'>`")
Response.Write(`"<meta name='viewport' content='width=device-width, initial-scale=1.0'>`")
Response.Write(`"<meta http-equiv='X-UA-Compatible' content='ie=edge'>`")

function addLeadingZero(value)
  addLeadingZero = value
  if value < 10 then
    addLeadingZero = `"0`" & value
  end if
end function

Dim today, sYear, sMonth, sDay, folderdate, fs, folder, subfolder, file, item, url

today = now()
sYear = Year(today)
sMonth = addLeadingZero(Month(today))
sDay = addLeadingZero(Day(today))

set fs = CreateObject(`"Scripting.FileSystemObject`")
set folder = fs.GetFolder(Server.MapPath(`"./`"))

for each item in folder.SubFolders
  folderdate = Split(item.name,`" - `")
  if folderdate(0) = sDay & `"-`" & sMonth & `"-`" & sYear Then

    set subfolder = fs.GetFolder(Server.MapPath(`"./`" & item.name & `"/`"))
    for each subitem in subfolder.Files
      if InStr(subitem.name,`".html`") Then
        
        destinationurl = `"./`" & item.name & `"/`" & subitem.name
        Response.Write(`"    <meta HTTP-EQUIV='refresh' content='5;url=`" & destinationurl & `"'>`")

      End If
    next
  End If
next

Response.Write(`"<title>PingCastle Reports.</title>`")
Response.Write(`"</head>`")
Response.Write(`"<body>`")

if destinationurl <> `"`" Then

  Response.Write(`"<H3>You wil be redirected to the latest PingCastle report in 5 sec.</H3>`")
  Response.Write(`"Click <a href='`" & destinationurl & `"'>here</a> to skip the wait.<br>`")

  if fs.FileExists(Server.MapPath(`".`") & `"\ad_hc_rules_list.html`") Then
    Response.Write(`"<br><br>`")
    Response.Write(`"<a href='./ad_hc_rules_list.html'>PingCastle Healthcheck rules</a>`")
  End If

Else

  Response.Write(`"<H3>No PingCastle reports avalible</H3>`")

End If

if folder.SubFolders.count > 0 Then
  Response.Write(`"<br><br>`")
  Response.Write(`"If you need to se a older reports, use the link below to goto the list of older reports.<br>`")
  Response.Write(`"<br>`")
  Response.Write(`"<a href='./list.asp'>List all report dates</a><br>`")
  Response.Write(`"</body>`")
  Response.Write(`"</html>`")
End If
%>"

$Bytes = [System.Text.Encoding]::UTF8.GetBytes($DefaultASP)
$EncodedText =[Convert]::ToBase64String($Bytes)
$EncodedText




# --------------------------------------------------
# Encode - List.asp
# --------------------------------------------------
$ListASP = "<%
Response.Write(`"<!DOCTYPE html>`")
Response.Write(`"<html lang='en'>`")
Response.Write(`"<head>`")
Response.Write(`"<meta charset='UTF-8'>`")
Response.Write(`"<meta name='viewport' content='width=device-width, initial-scale=1.0'>`")
Response.Write(`"<meta http-equiv='X-UA-Compatible' content='ie=edge'>`")
Response.Write(`"<title>PingCastle Reports.</title>`")
Response.Write(`"</head>`")
Response.Write(`"<body>`")

dim fs, folder, subfolder, file, item, subitem, url

set fs = CreateObject(`"Scripting.FileSystemObject`")
set folder = fs.GetFolder(Server.MapPath(`".`"))

for each item in folder.SubFolders
  set subfolder = fs.GetFolder(Server.MapPath(item.name))
  for each subitem in subfolder.Files
    if InStr(subitem.name,`".html`") Then
      Response.Write(`"<a href='./`" & item.name & `"/`" & subitem.name & `"'>`" & item.name & `"</a><br>`" & vbCrLf)
    End if
  next
next

Response.Write(`"</body>`")
Response.Write(`"</html>`")
%>"

$Bytes = [System.Text.Encoding]::UTF8.GetBytes($ListASP)
$EncodedText =[Convert]::ToBase64String($Bytes)
$EncodedText
