Sample scripts to install and provide support for Legacy LAPS and the new Windows LAPS.


On any Windows 10/11 or Server 2016/2019/2022.
This has to be run prior to the other scripts, the WMI filters will be linked to the created GPO's
- This script will create 2 WMI filters, to support Legacy LAPS on 2016 and older Server OS, and New LAPS on 2019 and newer.
- Run : .\Create WMI Filters\Create-WMIfilters.ps1


On Windows 10/11 or Server 2016/2019/2022, the April update is NOT installed.
- This script will create GPO to install LAPS "AdmPwd GPO Extension", configure Legacy LAPS settings and link the required WMI filter.
- Using credentials that have permissions to create GPO's and install GPO extentions
- Run : .\Install Legacy LAPS\Prepare Domain.ps1


On Windows 10/11 or Server 2016/2019/2022, the April update is installed.
- This script will create GPO to remove Legacy LAPS on servers where the April 2023 patch can be installed, configure Windows LAPS settings and link the required WMI filter.
- Using credentials that have permissions to create GPO's
- Run .\Install Windows LAPS\Prepare Domain.ps1

