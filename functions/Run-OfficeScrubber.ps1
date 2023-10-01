function Run-OfficeScrubber {

  Write-Host "Select [R] Remove all Licenses option in OfficeScrubber." -ForegroundColor Yellow
  try {
    Get-OfficeScrubber -ScrubberBaseUrl "$ScrubberBaseUrl" -OfficeUtilPath "$OfficeUtilPath" -ScrubberArchiveName "$ScrubberArchiveName"
  }
  catch {
    Write-Host "Fout opgetreden: $_"
  }

  Start-Process -Verb runas -FilePath "cmd.exe" -ArgumentList "/C $ScrubberCmdPath "
}

