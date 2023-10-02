# Function to install Office 365 Business
function Install-Office365 {
  Install-OdtIfNeeded
  if (-not (Test-Path -Path $configuration365XMLPath -PathType Leaf)) {
    Write-Host "Downloading Office 365 Business Configuration File..." -ForegroundColor Cyan
    $downloadUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config365.xml"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $configuration365XMLPath
  }
  Write-Host -NoNewline "Install Microsoft Office 365 Business? ( Y / N ): "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  switch ($choice) {
    'y' {
      Write-Host "Installation started. Don't close this window" -ForegroundColor Green
      Start-Process -Wait $setupExePath -ArgumentList "$UnattendedArgs365"
      Write-Host "Installation completed." -ForegroundColor Green
      Show-OfficeMainMenu
    }
    'n' {
      Write-Host "Exiting..."
      Show-OfficeInstallMenu
      # exit
    }
    default {
      Write-Host -NoNewLine "Invalid option. Press any key to try again... "
      $x = [System.Console]::ReadKey().KeyChar
      Invoke-Logo
      Install-Office365
    }
  }

}

# Function to install Office 2021 Pro Plus
function Install-Office21 {
  Install-OdtIfNeeded
  if (-not (Test-Path -Path $configuration21XMLPath -PathType Leaf)) {
    Write-Host "Downloading Office 2021 Pro Plus Configuration File..." -ForegroundColor Cyan
    $downloadUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config21.xml"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $configuration21XMLPath
  }
  Write-Host -NoNewline "Install Microsoft Office 2021 Pro Plus? ( Y / N ): "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  switch ($choice) {
    'y' {
      Write-Host "Installation started. Don't close this window" -ForegroundColor Green
      Start-Process -Wait $setupExePath -ArgumentList "$UnattendedArgs21"
      Write-Host "Installation completed." -ForegroundColor Green
      Show-OfficeMainMenu
    }
    'n' {
      Write-Host "Exiting..."
      Show-OfficeInstallMenu
      # exit
    }
    default {
      Write-Host -NoNewLine "Invalid option. Press any key to go back... "
      $x = [System.Console]::ReadKey().KeyChar
      Invoke-Logo
      Install-Office21
    }
  }
}

