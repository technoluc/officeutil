# Start-Transcript $ENV:TEMP\OfficeUtil.log -Append
# Start-Transcript $ENV:TEMP\OfficeUtil.log

##################################################
#                 SET VARIABLES                  #
##################################################

$ScriptUrl = "https://raw.githubusercontent.com/technoluc/officeutil/update/OfficeUtil.ps1"
$OfficeUtilPath = "C:\OfficeUtil"

$odtInstallerPath = Join-Path -Path $OfficeUtilPath -ChildPath "odtInstaller.exe"
$odtPath = "C:\Program Files\OfficeDeploymentTool"
$setupExePath = Join-Path -Path $odtPath -ChildPath "setup.exe"
$configuration21XMLPath = Join-Path -Path $odtPath -ChildPath "config21.xml"
$configuration365XMLPath = Join-Path -Path $odtPath -ChildPath "config365.xml"

# OfficeScrubber
$ScrubberPath = Join-Path -Path $OfficeUtilPath -ChildPath "OfficeScrubber"
$ScrubberBaseUrl = "https://github.com/abbodi1406/WHD/raw/master/scripts/OfficeScrubber_11.7z"
$ScrubberArchiveName = "OfficeScrubber_11.7z"
$ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

$ScrubberCmdName = "OfficeScrubber.cmd"
$ScrubberCmdPath = Join-Path -Path $ScrubberPath -ChildPath $ScrubberCmdName

# Office Removal Tool
$OfficeRemovalToolUrl = "https://raw.githubusercontent.com/technoluc/msoffice-removal-tool/main/msoffice-removal-tool.ps1"
$OfficeRemovalToolName = "msoffice-removal-tool.ps1"
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
