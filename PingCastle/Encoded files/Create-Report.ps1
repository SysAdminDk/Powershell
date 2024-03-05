# --
# Base path
# --
$Root = "C:\inetpub\PingCastle"

#--
# Create Report folder
# --
$FolderDate = Get-Date -Format "dd-MM-yyyy - HH-mm"
if (!(Test-Path -Path "$Root\$FolderDate")) {
    New-Item -Path "$Root\$FolderDate" -ItemType Directory | Out-null
}

#--
# Run PingCastle
#--
if (Test-Path -Path "$($ENV:ProgramFiles)\PingCastle\PingCastle.exe") {
    Start-Process -FilePath "$($ENV:ProgramFiles)\PingCastle\PingCastle.exe" -ArgumentList "--healthcheck" -NoNewWindow -WorkingDirectory "$Root\$FolderDate" -Wait
}
