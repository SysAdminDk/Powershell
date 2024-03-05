<%
Response.Write("<!DOCTYPE html>")
Response.Write("<html lang='en'>")
Response.Write("<head>")
Response.Write("<meta charset='UTF-8'>")
Response.Write("<meta name='viewport' content='width=device-width, initial-scale=1.0'>")
Response.Write("<meta http-equiv='X-UA-Compatible' content='ie=edge'>")
Response.Write("<title>PingCastle Reports.</title>")
Response.Write("</head>")
Response.Write("<body>")

dim fs, folder, subfolder, file, item, subitem, url

set fs = CreateObject("Scripting.FileSystemObject")
set folder = fs.GetFolder(Server.MapPath("."))

for each item in folder.SubFolders
  set subfolder = fs.GetFolder(Server.MapPath(item.name))
  for each subitem in subfolder.Files
    if InStr(subitem.name,".html") Then
      Response.Write("<a href='./" & item.name & "/" & subitem.name & "'>" & item.name & "</a><br>" & vbCrLf)
    End if
  next
next

Response.Write("</body>")
Response.Write("</html>")
%>
