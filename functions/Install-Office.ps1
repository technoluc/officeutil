# Function to install Office 365 Business
function Install-Office365 {
  Get-OdtIfNeeded
  if (-not (Test-Path -Path $configuration365XMLPath -PathType Leaf)) {
    Write-Host "Downloading Office 365 Business Configuration File..." -ForegroundColor Cyan
    $downloadUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config365.xml"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $configuration365XMLPath
  }
  Write-Host -NoNewline "Install Microsoft Office 365 Business? ( Y / N ) "
  $choice = [System.Console]::ReadKey().KeyChar
  switch ($choice) {
    'y' {
      Write-Host "Installation started. Don't close this window" -ForegroundColor Green
      Start-Process -Wait $setupExePath -ArgumentList "$UnattendedArgs365"
      Write-Host "Installation completed." -ForegroundColor Green
    }
    'n' {
      Write-Host "Exiting..."
      # exit
    }
    default {
      Write-Host -NoNewLine "Invalid option. Press any key to go back... "
      $x = [System.Console]::ReadKey().KeyChar
      Install-Office365
    }
  }

}

# Function to install Office 2021 Pro Plus
function Install-Office21 {
  Get-OdtIfNeeded
  if (-not (Test-Path -Path $configuration21XMLPath -PathType Leaf)) {
    Write-Host "Downloading Office 2021 Pro Plus Configuration File..." -ForegroundColor Cyan
    $downloadUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config21.xml"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $configuration21XMLPath
  }
  Write-Host -NoNewline "Install Microsoft Office 2021 Pro Plus? ( Y / N ) "
  $choice = [System.Console]::ReadKey().KeyChar
  switch ($choice) {
    'y' {
      Write-Host "Installation started. Don't close this window" -ForegroundColor Green
      Start-Process -Wait $setupExePath -ArgumentList "$UnattendedArgs21"
      Write-Host "Installation completed." -ForegroundColor Green
    }
    'n' {
      Write-Host "Exiting..."
      # exit
    }
    default {
      Write-Host -NoNewLine "Invalid option. Press any key to go back... "
      $x = [System.Console]::ReadKey().KeyChar
      Install-Office21
    }
  }
}

