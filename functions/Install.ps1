function Install-7ZipIfNeeded {
  $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"

  if (-not $7ZipInstalled) {
      Write-Host "7-Zip is not installed. Installing..."
      $InstallerUrl = "https://www.7-zip.org/a/7z2301-x64.exe"
      $InstallerPath = Join-Path -Path $env:TEMP -ChildPath "7zInstaller.exe"
      
      # Download and install 7-Zip with /S for silent installation
      Download-File -url $InstallerUrl -outputPath $InstallerPath 
      Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait
      
      # Check for successful installation
      $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"
      if ($7ZipInstalled) {
          Write-Host "7-Zip has been successfully installed."
      }
      else {
          Write-Host "Error: 7-Zip installation failed."
      }

      # Remove the temporary installation file
      Remove-Item -Path $InstallerPath -Force
  }
  else {
      Write-Host "7-Zip is already installed."
  }
}

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

function Install-Office {
  param (
      [string]$product
  )

  Install-OdtIfNeeded

  if ($product -eq "Office 365 Business") {
      $configXMLPath = $configuration365XMLPath
      $downloadURL = $configuration365XMLUrl
      $unattendedArgs = $UnattendedArgs365
  }
  elseif ($product -eq "Office 2021 Pro Plus") {
      $configXMLPath = $configuration21XMLPath
      $downloadURL = $configuration21XMLUrl
      $unattendedArgs = $UnattendedArgs21
  }
  else {
      Write-Host "Invalid product selection."
      return
  }

  if (-not (Test-Path -Path $configXMLPath -PathType Leaf)) {
      Write-Host "Downloading $product Configuration File..." -ForegroundColor Cyan
      Download-File -url $downloadURL -outputPath $configXMLPath
  }

  Write-Host -NoNewline "Install Microsoft '$product'? ( Y / N ): "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""

  switch ($choice) {
      'y' {
          Write-Host "Installation started. Don't close this window" -ForegroundColor Green
          Start-Process -Wait $setupExePath -ArgumentList "$unattendedArgs"
          Write-Host "Installation of Microsoft $product completed." -ForegroundColor Green
          Show-OfficeMainMenu
      }
      'n' {
          Write-Host "Exiting..."
          Show-OfficeInstallMenu
      }
      default {
          Write-Host -NoNewLine "Invalid option. Press any key to try again... "
          $x = [System.Console]::ReadKey().KeyChar
          Invoke-Logo
          Show-OfficeInstallMenu
      }
  }
}
