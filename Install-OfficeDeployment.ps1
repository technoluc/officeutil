##################################################
#                 SET VARIABLES                  #
##################################################
$odtPath = "C:\Program Files\OfficeDeploymentTool"
$configuration21XML = "C:\Program Files\OfficeDeploymentTool\config21.xml"
$configuration365XML = "C:\Program Files\OfficeDeploymentTool\config365.xml"

# Step 2: Check if required files are present
$requiredFiles = @(
  @{
    Name       = "config365.xml";
    PrettyName = "Office 365 Business Configuration File";
    Url        = "https://github.com/technoluc/winutil/raw/main-custom/office/config365.xml"
  },
  @{
    Name       = "config21.xml";
    PrettyName = "Office 21 Pro Plus Configuration File";
    Url        = "https://github.com/technoluc/winutil/raw/main-custom/office/config21.xml"
  }
)

# Step 3: Get the full path of the required files
foreach ($fileInfo in $requiredFiles) {
  $filePath = Join-Path -Path $odtPath -ChildPath $fileInfo.Name

  # Step 4: Test if the required file exists
  if (-not (Test-Path -Path $filePath -PathType Leaf)) {
      Write-Host ("Downloading $($fileInfo.PrettyName)...") -ForegroundColor Cyan
      $downloadUrl = $fileInfo.Url
      Invoke-WebRequest -Uri $downloadUrl -OutFile $filePath
    }
}


