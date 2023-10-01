# Start-Transcript $ENV:TEMP\OfficeUtil.log -Append
Start-Transcript $ENV:TEMP\OfficeUtil.log

##################################################
#                 SET VARIABLES                  #
##################################################

$ScriptUrl = "https://raw.githubusercontent.com/technoluc/scripts/main/officeutil/OfficeUtil.ps1"
$OfficeUtilPath = "C:\OfficeUtil"

$odtInstaller = "C:\OfficeUtil\odtInstaller.exe"
$odtPath = "C:\Program Files\OfficeDeploymentTool"
$setupExe = "C:\Program Files\OfficeDeploymentTool\setup.exe"
$configuration21XML = "C:\Program Files\OfficeDeploymentTool\config21.xml"
$configuration365XML = "C:\Program Files\OfficeDeploymentTool\config365.xml"

# OfficeScrubber
$ArchiveUrl = "https://github.com/abbodi1406/WHD/raw/master/scripts/OfficeScrubber_11.7z"
$ScrubberPath = "C:\OfficeUtil\OfficeScrubber"
$ScrubberArchive = "OfficeScrubber_11.7z"
$ScrubberCmd = "OfficeScrubber.cmd"
$ScrubberFullPath = Join-Path -Path $ScrubberPath -ChildPath $ScrubberCmd


$OfficeRemovalToolUrl = "https://raw.githubusercontent.com/technoluc/msoffice-removal-tool/main/msoffice-removal-tool.ps1"
$OfficeRemovalTool = "msoffice-removal-tool.ps1"
$OfficeRemovalToolPath = Join-Path -Path $OfficeUtilPath -ChildPath $OfficeRemovalTool

# Unattended Arguments for Office Installation
$UnattendedArgs21 = "/configure `"$configuration21XML`""
$UnattendedArgs365 = "/configure `"$configuration365XML`""
$odtInstallerArgs = "/extract:`"c:\Program Files\OfficeDeploymentTool`" /quiet"


# Check if script was run as Administrator, relaunch if not
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Output "OfficeUtil needs to be run as Administrator. Attempting to relaunch."
  Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -UseBasicParsing `"$ScriptUrl`" | Invoke-Expression" 
  break
}
