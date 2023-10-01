function Test-OfficeInstalled {
  if (Test-Path "C:\Program Files\Microsoft Office") {
    Write-Host "Microsoft Office is already installed." -ForegroundColor Yellow
    Write-Host "Run OfficeRemoverTool and OfficeScrubber to remove the previous installation first."
    Write-Host "Or run Massgrave.dev Microsoft Activation Scripts to activate Office / Windows."

    return $true
  }
  else {
    return $false
  }
}
