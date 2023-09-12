<#
    .SYNOPSIS
     Deletes ALL unlinked and "AllSettingsDisabled" GPOs from current Domain
    
    .NOTES

#>

##################################################################################
# DISCLAIMER [ Start ]
##################################################################################

clear
Write-Output "*******************************************************************************************************************"
Write-Output ""
Write-Output "DISCLAIMER: "
Write-Output ""
Write-Output "THE FOLLOWING POWERSHELL SCRIPT IS PROVIDED AS-IS, WITHOUT WARRANTY OF ANY KIND. USE AT YOUR OWN RISK."
Write-Output ""
Write-Output "By running this script, you acknowledge that you have read and understood the disclaimer, and you agree to assume"
Write-Output "all responsibility for any failures, damages, or issues that may arise as a result of executing this script."
Write-Output ""
Write-Output "Take appropriate precautions and ensure you have a backup of your Active Directory before running this script."
Write-Output ""
Write-Output "*******************************************************************************************************************"

$title = "Cleanup GPOs"
$message = "Do you want to run this script and remove unused GPOs ?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes cleanup."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Just quit"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

if (($host.ui.PromptForChoice($title, $message, $options, 0)) -eq 1) {
    Write-Output ""
    break
} else {
    Write-Output "Verifying Prerequisites"
}

##################################################################################
# DISCLAIMER [ End ]
##################################################################################





# --
# Get ALL GPOs from current domain
# --
$GPOs = Get-GPO -All


# --
# Remve all inactive GPO's
# --
$GPOs | Where {$_.GpoStatus -like '*AllSettingsDisabled*'} | Remove-GPO


# --
# Remve all Unlinked GPO's
# --
Foreach ($GPO in $GPOs) {
    [XML]$GPReport = Get-GPOReport -ReportType Xml -Guid $GPO.ID
    if (($GPReport.GPO.LinksTo).Count -eq 0) {
        $GPO | Remove-GPO
    }
}
