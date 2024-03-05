#RunAs Administrator
<#
    This script Installs IIS with the required features to support Windows authentication and authorization

    Create a Scheduled task that run the PingCastleAutoUpdate every Friday at 05:00
    Create a Scheduled task that run the PingCastle every day at 06:00

    Create two .asp files in the created IIS Application directory to redirect to latest report and link to listing of All avalible reports.

#>


# --------------------------------------------------
# Active Directory Group - Allowed access to the web site
# - The Group have to be created in Active Directory
# --------------------------------------------------
param (
    [cmdletbinding()]
    [Parameter(ValueFromPipeline)]
    [string[]]$ADGroupName = "PingCastle Report Readers"
)


# --------------------------------------------------
# Install Required Festures
# --------------------------------------------------
$ToolsToInstall = @(
    "Web-Server",
    "Web-WebServer",
    "Web-Common-Http",
    "Web-Default-Doc",
    "Web-Dir-Browsing",
    "Web-Http-Errors",
    "Web-Static-Content",
    "Web-Http-Redirect",
    "Web-Health",
    "Web-Http-Logging",
    "Web-Performance",
    "Web-Stat-Compression",
    "Web-Security",
    "Web-Filtering",
    "Web-Url-Auth",
    "Web-Windows-Auth",
    "Web-App-Dev",
    "Web-ASP",
    "Web-ISAPI-Ext",
    "Web-Mgmt-Tools",
    "Web-Mgmt-Console"
	)
Write-Verbose "Installing required Windows Features"
Install-WindowsFeature -Name $ToolsToInstall -IncludeManagementTools


# --------------------------------------------------
# Create Folder structure
# --------------------------------------------------
Write-Verbose "Create Folders"
$BasePath = "$($ENV:SystemDrive)\InetPub\PingCastle"
If (!(Test-Path -Path $BasePath)) {
    Write-Verbose "Create Folders, $BasePath"
    New-Item -Path $BasePath -Itemtype Directory
}


# --------------------------------------------------
# Download PingCastle
# --------------------------------------------------
Write-Verbose "Download PingCastle"
$DownloadPath = "$($ENV:USERPROFILE)\Downloads"

# Get latest version download link and name
$LatestRelease = (Invoke-WebRequest "https://api.github.com/repos/vletoux/pingcastle/releases" | ConvertFrom-Json)[0]
$Uri = $LatestRelease.assets.browser_download_url
$OutFile = $LatestRelease.assets.name

if (!(Test-Path -Path "$DownloadPath\$OutFile")) {
    Write-Verbose "Resolved latest stable version, $($LatestRelease.Name)"
    Invoke-WebRequest -Uri $Uri -OutFile "$DownloadPath\$OutFile"
}


# --------------------------------------------------
# Extract PingCastle.
# --------------------------------------------------
$AppPath = "$($ENV:ProgramFiles)\PingCastle"
if (!(Test-Path -Path $AppPath)) {
    Write-Verbose "Create Folders, $AppPath"
    New-Item -Path $AppPath -ItemType Directory
}
if (!(Test-Path -Path "$AppPath\PingCastle.exe")) {
    Write-Verbose "Extract PingCastle to $AppPath"

    if (Test-Path -Path "$DownloadPath\$OutFile") {
        Expand-Archive -Path "$DownloadPath\$OutFile" -DestinationPath "$AppPath" -Force
    }

    # Dump all PingCaste rules, for refrence.
    if (Test-Path -Path "$AppPath\PingCastle.exe") {
        Write-Verbose "Create refrence HTML with all PingCastle rules"
        Start-Process -FilePath "$($ENV:ProgramFiles)\PingCastle\PingCastle.exe" -ArgumentList "--rules" -NoNewWindow -WorkingDirectory "$BasePath" -Wait
    }
}


# --------------------------------------------------
# Create Scheduled Task - PingCastle Auto Update
# --------------------------------------------------
Write-Verbose "Create PingCastle Auto Update Task"
$Scheduletrigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Friday -At "05:00"
$ScheduleSettings = New-ScheduledTaskSettingsSet
$ScheduleAction = New-ScheduledTaskAction -Execute "$($ENV:ProgramFiles)\PingCastle\PingCastleAutoUpdater.exe"
$SchedulePrincipal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Limited
$ScheduledTask = New-ScheduledTask -Action $ScheduleAction -Trigger $Scheduletrigger -Settings $ScheduleSettings -Principal $SchedulePrincipal
Register-ScheduledTask -TaskName "Run PingCastle Auto Update" -InputObject $ScheduledTask


# --------------------------------------------------
# Write "Create-Report.ps1"
# - The content of the string is created with "Encode files.ps1"
# --------------------------------------------------
Write-Verbose "Create Scheduler powershell script"
$CodeCreateReportPs1 = "IyAtLQ0KIyBCYXNlIHBhdGgNCiMgLS0NCiRSb290ID0gIkM6XGluZXRwdWJcUGluZ0Nhc3RsZSINCg0KIy0tDQojIENyZWF0ZSBSZXBvcnQgZm9sZGVyDQojIC0tDQokRm9sZGVyRGF0ZSA9IEdldC1EYXRlIC1Gb3JtYXQgImRkLU1NLXl5eXkgLSBISC1tbSINCmlmICghKFRlc3QtUGF0aCAtUGF0aCAiJFJvb3RcJEZvbGRlckRhdGUiKSkgew0KICAgIE5ldy1JdGVtIC1QYXRoICIkUm9vdFwkRm9sZGVyRGF0ZSIgLUl0ZW1UeXBlIERpcmVjdG9yeSB8IE91dC1udWxsDQp9DQoNCiMtLQ0KIyBSdW4gUGluZ0Nhc3RsZQ0KIy0tDQppZiAoVGVzdC1QYXRoIC1QYXRoICIkKCRFTlY6UHJvZ3JhbUZpbGVzKVxQaW5nQ2FzdGxlXFBpbmdDYXN0bGUuZXhlIikgew0KICAgIFN0YXJ0LVByb2Nlc3MgLUZpbGVQYXRoICIkKCRFTlY6UHJvZ3JhbUZpbGVzKVxQaW5nQ2FzdGxlXFBpbmdDYXN0bGUuZXhlIiAtQXJndW1lbnRMaXN0ICItLWhlYWx0aGNoZWNrIiAtTm9OZXdXaW5kb3cgLVdvcmtpbmdEaXJlY3RvcnkgIiRSb290XCRGb2xkZXJEYXRlIiAtV2FpdA0KfQ=="
[System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($CodeCreateReportPs1)) | Out-File "$($ENV:ProgramFiles)\PingCastle\Create-Report.ps1" -Encoding utf8 -Force


# --------------------------------------------------
# Create Scheduled Task
# --------------------------------------------------
Write-Verbose "Create PingCastle Scheduled Task"
$PowershellPath = "$(($ENV:PATH) -split(";") | Where {$_ -like '*Powershell*'})Powershell.exe"

$Scheduletrigger = New-ScheduledTaskTrigger -Daily -At "06:00"
$ScheduleSettings = New-ScheduledTaskSettingsSet
$ScheduleAction = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$($ENV:ProgramFiles)\PingCastle\Create-Report.ps1"
$SchedulePrincipal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Limited
$ScheduledTask = New-ScheduledTask -Action $ScheduleAction -Trigger $Scheduletrigger -Settings $ScheduleSettings -Principal $SchedulePrincipal
Register-ScheduledTask -TaskName "Run PingCastle - Daily" -InputObject $ScheduledTask


# --------------------------------------------------
# Create Web Application
# --------------------------------------------------
Write-Verbose "Create PingCastle IIS Application"
$DefaultSite = Get-IISSite
$DefaultAppPool = Get-IISAppPool
$Roles = "$($ENV:USERDOMAIN)\$ADGroupName"

$PingCastleWeb = New-WebApplication -Name "PingCastle" -Site $DefaultSite.name -PhysicalPath "$BasePath" -ApplicationPool $DefaultAppPool.Name

# Diasble Anonymous and Enabme Windows Authentication
Set-WebConfigurationProperty -Filter "/system.webServer/security/authentication/anonymousAuthentication" -Name "enabled" -Value "false" -PSPath "IIS:\" -Location "$($DefaultSite.name)/$($PingCastleWeb.Name)"
Set-WebConfigurationProperty -Filter "/system.webServer/security/authentication/windowsAuthentication" -Name "enabled" -Value "true" -PSPath "IIS:\" -Location "$($DefaultSite.name)/$($PingCastleWeb.Name)"

# Remove Default Authorization and Add Group Authorization rule
Remove-WebConfigurationProperty -PSPath "MACHINE/WEBROOT/APPHOST/Default Web Site/PingCastle" -Filter 'system.webServer/security/authorization' -name '.' -AtElement @{users='*';roles='';verbs=''}
Add-WebConfigurationProperty -PSPath "MACHINE/WEBROOT/APPHOST/Default Web Site/PingCastle" -Filter 'system.webServer/security/authorization' -name '.' -value @{accessType='Allow';roles=$Roles}


# --------------------------------------------------
# Create Default.asp
# - The content of the string is created with "Encode files.ps1"
# --------------------------------------------------
Write-Verbose "Create ASP files"
$CodeDefaultAsp = "PCUNClJlc3BvbnNlLldyaXRlKCI8IURPQ1RZUEUgaHRtbD4iKQ0KUmVzcG9uc2UuV3JpdGUoIjxodG1sIGxhbmc9J2VuJz4iKQ0KUmVzcG9uc2UuV3JpdGUoIjxoZWFkPiIpDQpSZXNwb25zZS5Xcml0ZSgiPG1ldGEgY2hhcnNldD0nVVRGLTgnPiIpDQpSZXNwb25zZS5Xcml0ZSgiPG1ldGEgbmFtZT0ndmlld3BvcnQnIGNvbnRlbnQ9J3dpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xLjAnPiIpDQpSZXNwb25zZS5Xcml0ZSgiPG1ldGEgaHR0cC1lcXVpdj0nWC1VQS1Db21wYXRpYmxlJyBjb250ZW50PSdpZT1lZGdlJz4iKQ0KDQpmdW5jdGlvbiBhZGRMZWFkaW5nWmVybyh2YWx1ZSkNCiAgYWRkTGVhZGluZ1plcm8gPSB2YWx1ZQ0KICBpZiB2YWx1ZSA8IDEwIHRoZW4NCiAgICBhZGRMZWFkaW5nWmVybyA9ICIwIiAmIHZhbHVlDQogIGVuZCBpZg0KZW5kIGZ1bmN0aW9uDQoNCkRpbSB0b2RheSwgc1llYXIsIHNNb250aCwgc0RheSwgZm9sZGVyZGF0ZSwgZnMsIGZvbGRlciwgc3ViZm9sZGVyLCBmaWxlLCBpdGVtLCB1cmwNCg0KdG9kYXkgPSBub3coKQ0Kc1llYXIgPSBZZWFyKHRvZGF5KQ0Kc01vbnRoID0gYWRkTGVhZGluZ1plcm8oTW9udGgodG9kYXkpKQ0Kc0RheSA9IGFkZExlYWRpbmdaZXJvKERheSh0b2RheSkpDQoNCnNldCBmcyA9IENyZWF0ZU9iamVjdCgiU2NyaXB0aW5nLkZpbGVTeXN0ZW1PYmplY3QiKQ0Kc2V0IGZvbGRlciA9IGZzLkdldEZvbGRlcihTZXJ2ZXIuTWFwUGF0aCgiLi8iKSkNCg0KZm9yIGVhY2ggaXRlbSBpbiBmb2xkZXIuU3ViRm9sZGVycw0KICBmb2xkZXJkYXRlID0gU3BsaXQoaXRlbS5uYW1lLCIgLSAiKQ0KICBpZiBmb2xkZXJkYXRlKDApID0gc0RheSAmICItIiAmIHNNb250aCAmICItIiAmIHNZZWFyIFRoZW4NCg0KICAgIHNldCBzdWJmb2xkZXIgPSBmcy5HZXRGb2xkZXIoU2VydmVyLk1hcFBhdGgoIi4vIiAmIGl0ZW0ubmFtZSAmICIvIikpDQogICAgZm9yIGVhY2ggc3ViaXRlbSBpbiBzdWJmb2xkZXIuRmlsZXMNCiAgICAgIGlmIEluU3RyKHN1Yml0ZW0ubmFtZSwiLmh0bWwiKSBUaGVuDQogICAgICAgIA0KICAgICAgICBkZXN0aW5hdGlvbnVybCA9ICIuLyIgJiBpdGVtLm5hbWUgJiAiLyIgJiBzdWJpdGVtLm5hbWUNCiAgICAgICAgUmVzcG9uc2UuV3JpdGUoIiAgICA8bWV0YSBIVFRQLUVRVUlWPSdyZWZyZXNoJyBjb250ZW50PSc1O3VybD0iICYgZGVzdGluYXRpb251cmwgJiAiJz4iKQ0KDQogICAgICBFbmQgSWYNCiAgICBuZXh0DQogIEVuZCBJZg0KbmV4dA0KDQpSZXNwb25zZS5Xcml0ZSgiPHRpdGxlPlBpbmdDYXN0bGUgUmVwb3J0cy48L3RpdGxlPiIpDQpSZXNwb25zZS5Xcml0ZSgiPC9oZWFkPiIpDQpSZXNwb25zZS5Xcml0ZSgiPGJvZHk+IikNCg0KaWYgZGVzdGluYXRpb251cmwgPD4gIiIgVGhlbg0KDQogIFJlc3BvbnNlLldyaXRlKCI8SDM+WW91IHdpbCBiZSByZWRpcmVjdGVkIHRvIHRoZSBsYXRlc3QgUGluZ0Nhc3RsZSByZXBvcnQgaW4gNSBzZWMuPC9IMz4iKQ0KICBSZXNwb25zZS5Xcml0ZSgiQ2xpY2sgPGEgaHJlZj0nIiAmIGRlc3RpbmF0aW9udXJsICYgIic+aGVyZTwvYT4gdG8gc2tpcCB0aGUgd2FpdC48YnI+IikNCg0KICBpZiBmcy5GaWxlRXhpc3RzKFNlcnZlci5NYXBQYXRoKCIuIikgJiAiXGFkX2hjX3J1bGVzX2xpc3QuaHRtbCIpIFRoZW4NCiAgICBSZXNwb25zZS5Xcml0ZSgiPGJyPjxicj4iKQ0KICAgIFJlc3BvbnNlLldyaXRlKCI8YSBocmVmPScuL2FkX2hjX3J1bGVzX2xpc3QuaHRtbCc+UGluZ0Nhc3RsZSBIZWFsdGhjaGVjayBydWxlczwvYT4iKQ0KICBFbmQgSWYNCg0KRWxzZQ0KDQogIFJlc3BvbnNlLldyaXRlKCI8SDM+Tm8gUGluZ0Nhc3RsZSByZXBvcnRzIGF2YWxpYmxlPC9IMz4iKQ0KDQpFbmQgSWYNCg0KaWYgZm9sZGVyLlN1YkZvbGRlcnMuY291bnQgPiAwIFRoZW4NCiAgUmVzcG9uc2UuV3JpdGUoIjxicj48YnI+IikNCiAgUmVzcG9uc2UuV3JpdGUoIklmIHlvdSBuZWVkIHRvIHNlIGEgb2xkZXIgcmVwb3J0cywgdXNlIHRoZSBsaW5rIGJlbG93IHRvIGdvdG8gdGhlIGxpc3Qgb2Ygb2xkZXIgcmVwb3J0cy48YnI+IikNCiAgUmVzcG9uc2UuV3JpdGUoIjxicj4iKQ0KICBSZXNwb25zZS5Xcml0ZSgiPGEgaHJlZj0nLi9saXN0LmFzcCc+TGlzdCBhbGwgcmVwb3J0IGRhdGVzPC9hPjxicj4iKQ0KICBSZXNwb25zZS5Xcml0ZSgiPC9ib2R5PiIpDQogIFJlc3BvbnNlLldyaXRlKCI8L2h0bWw+IikNCkVuZCBJZg0KJT4="
[System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($CodeDefaultAsp)) | Out-File "$BasePath\Default.asp" -Encoding utf8 -Force


# --------------------------------------------------
# Create List.asp
# - The content of the string is created with "Encode files.ps1"
# --------------------------------------------------
$CodeListAsp = "PCUNClJlc3BvbnNlLldyaXRlKCI8IURPQ1RZUEUgaHRtbD4iKQ0KUmVzcG9uc2UuV3JpdGUoIjxodG1sIGxhbmc9J2VuJz4iKQ0KUmVzcG9uc2UuV3JpdGUoIjxoZWFkPiIpDQpSZXNwb25zZS5Xcml0ZSgiPG1ldGEgY2hhcnNldD0nVVRGLTgnPiIpDQpSZXNwb25zZS5Xcml0ZSgiPG1ldGEgbmFtZT0ndmlld3BvcnQnIGNvbnRlbnQ9J3dpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xLjAnPiIpDQpSZXNwb25zZS5Xcml0ZSgiPG1ldGEgaHR0cC1lcXVpdj0nWC1VQS1Db21wYXRpYmxlJyBjb250ZW50PSdpZT1lZGdlJz4iKQ0KUmVzcG9uc2UuV3JpdGUoIjx0aXRsZT5QaW5nQ2FzdGxlIFJlcG9ydHMuPC90aXRsZT4iKQ0KUmVzcG9uc2UuV3JpdGUoIjwvaGVhZD4iKQ0KUmVzcG9uc2UuV3JpdGUoIjxib2R5PiIpDQoNCmRpbSBmcywgZm9sZGVyLCBzdWJmb2xkZXIsIGZpbGUsIGl0ZW0sIHN1Yml0ZW0sIHVybA0KDQpzZXQgZnMgPSBDcmVhdGVPYmplY3QoIlNjcmlwdGluZy5GaWxlU3lzdGVtT2JqZWN0IikNCnNldCBmb2xkZXIgPSBmcy5HZXRGb2xkZXIoU2VydmVyLk1hcFBhdGgoIi4iKSkNCg0KZm9yIGVhY2ggaXRlbSBpbiBmb2xkZXIuU3ViRm9sZGVycw0KICBzZXQgc3ViZm9sZGVyID0gZnMuR2V0Rm9sZGVyKFNlcnZlci5NYXBQYXRoKGl0ZW0ubmFtZSkpDQogIGZvciBlYWNoIHN1Yml0ZW0gaW4gc3ViZm9sZGVyLkZpbGVzDQogICAgaWYgSW5TdHIoc3ViaXRlbS5uYW1lLCIuaHRtbCIpIFRoZW4NCiAgICAgIFJlc3BvbnNlLldyaXRlKCI8YSBocmVmPScuLyIgJiBpdGVtLm5hbWUgJiAiLyIgJiBzdWJpdGVtLm5hbWUgJiAiJz4iICYgaXRlbS5uYW1lICYgIjwvYT48YnI+IiAmIHZiQ3JMZikNCiAgICBFbmQgaWYNCiAgbmV4dA0KbmV4dA0KDQpSZXNwb25zZS5Xcml0ZSgiPC9ib2R5PiIpDQpSZXNwb25zZS5Xcml0ZSgiPC9odG1sPiIpDQolPg=="
[System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($CodeListAsp)) | Out-File "$BasePath\List.asp" -Encoding utf8 -Force
