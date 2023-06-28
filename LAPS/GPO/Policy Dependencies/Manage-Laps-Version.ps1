<#
    .SYNOPSIS
     Install or Uninstall Legacy LAPS based on Windows version and Patch level
    
    .DESCRIPTION
    Windows LAPS supported versions

    Windows 11 22H2 - April 11 2023 Update
    - OS Build 10.0.22621.1555

    Windows 11 21H2 - April 11 2023 Update
    - OS Build 10.0.22000.1817

    Windows 10 - April 11 2023 Update
    - OS Build 10.0.19042.2846
    - OS Build 10.0.19044.2846
    - OS Build 10.0.19045.2846

    Windows Server 2022 - April 11 2023 Update
    - OS Build 10.0.20348.1668

    Windows Server 2019 - April 11 2023 Update
    - OS Build 10.0.17763.4252

#>


# --------------------------------------------------------------------------------
# Read registry keys for Installed Windows version
# --------------------------------------------------------------------------------
$Registry_Windows_Version = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion"


# --------------------------------------------------------------------------------
# Windows LAPS Supported Versions, rest is Legacy LAPS
# --------------------------------------------------------------------------------
Switch ($($Registry_Windows_Version.CurrentBuild)) {
    "17763" { if ($($Registry_Windows_Version.UBR) -gt "4252") { $Laps = "Windows" } }
    "19042" { if ($($Registry_Windows_Version.UBR) -gt "2846") { $Laps = "Windows" } }
    "19044" { if ($($Registry_Windows_Version.UBR) -gt "2846") { $Laps = "Windows" } }
    "19045" { if ($($Registry_Windows_Version.UBR) -gt "2846") { $Laps = "Windows" } }
    "20348" { if ($($Registry_Windows_Version.UBR) -gt "1668") { $Laps = "Windows" } }
    "22000" { if ($($Registry_Windows_Version.UBR) -gt "1817") { $Laps = "Windows" } }
    "22621" { if ($($Registry_Windows_Version.UBR) -gt "1555") { $Laps = "Windows" } }
    default { $Laps = "Legacy" }
}


# --------------------------------------------------------------------------------
# Cleck if Legacy LAPS is installed.
# --------------------------------------------------------------------------------
$LegacyLAPS = Test-Path -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{97E2CA7B-B657-4FF7-A6DB-30ECC73E1E28}"


# --------------------------------------------------------------------------------
# Install or uninstall Legacy Laps
# --------------------------------------------------------------------------------
if ( ($Laps -eq "Legacy") -AND (!($LegacyLAPS)) ) {

    $Path = "$(Split-Path -Parent $MyInvocation.MyCommand.Definition)"
    $Path.TrimEnd('\')

    if ( ($($env:PROCESSOR_ARCHITECTURE) -eq "x86") -and (Test-Path -Path "$Path\LAPS.x86.msi") ) {
        Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/i `"$Path\LAPS.x86.msi`" ADDLOCAL=CSE /qn" -Wait
    }
    if ( ($($env:PROCESSOR_ARCHITECTURE) -eq "AMD64") -and (Test-Path -Path "$Path\LAPS.x64.msi") ) {
        Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/i `"$Path\LAPS.x64.msi`" ADDLOCAL=CSE /qn" -Wait
    }

} elseif (($Laps -eq "Windows") -AND ($LegacyLAPS)) {

    Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/x {97E2CA7B-B657-4FF7-A6DB-30ECC73E1E28} /quiet"

}
