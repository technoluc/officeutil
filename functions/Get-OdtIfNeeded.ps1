function Get-OdtIfNeeded {
  $OdtInstalled = Test-Path "$setupExePath"

  
  if (-not $OdtInstalled) {
    $IntallerPresent = Test-Path "$odtInstallerPath"

    if (-not $IntallerPresent) {

      if (-not (Test-Path "$OfficeUtilPath")) {
        New-Item -Path $OfficeUtilPath -ItemType Directory -Force | Out-Null
        Write-Host "$OfficeUtilPath created"
      }
      if (-not (Test-Path "$odtPath")) {
        New-Item -Path $odtPath -ItemType Directory -Force | Out-Null
        Write-Host "$odtPath created"
      }
      $URL = $(Get-ODTUri)
      # $URL = "https://officecdn.microsoft.com/pr/wsus/setup.exe" # Backup URL
      Invoke-WebRequest -Uri $URL -OutFile $odtInstallerPath
    }
    Start-Process -Wait $odtInstallerPath -ArgumentList $odtInstallerArgs
    
    # Check for successful installation
    $OdtInstalled = Test-Path "$setupExePath" 
    if ($OdtInstalled) {
      Write-Host "Office Deployment Tool has been successfully installed." -ForegroundColor Green
    }
    else {
      Write-Host "Error: Office Deployment Tool installation failed." -ForegroundColor Red
    }
  
    # # Remove the temporary installation file
    # Remove-Item -Path $InstallerPath -Force
  } 
  else {
    Write-Host "Office Deployment Tool is already installed." -ForegroundColor Green
  }
}