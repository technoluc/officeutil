function Install-OdtIfNeeded {
  $OdtInstalled = Test-Path "$setupExePath"

  if (-not $OdtInstalled) {
    if (-not (Test-Path "$odtInstallerPath")) {
      
      if (-not (Test-Path "$OfficeUtilPath")) {
        New-Item -Path $OfficeUtilPath -ItemType Directory -Force | Out-Null
        Write-Host "$OfficeUtilPath created"
      }
      if (-not (Test-Path "$odtPath")) {
        New-Item -Path $odtPath -ItemType Directory -Force | Out-Null
        Write-Host "$odtPath created"
      }
      $URL = $(Get-ODTUri)
      Download-File -url $URL -outputPath $odtInstallerPath
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
  }
  else {
    Write-Host ""
    Write-Host "Office Deployment Tool is already installed." -ForegroundColor Green
  }
}
