function Invoke-OfficeScrubber {

  try {
    Get-OfficeScrubber
  }
  catch {
    Write-Host "Fout opgetreden: $_"
  }
  finally {
    Write-Host "Select [R] Remove all Licenses option in OfficeScrubber." -ForegroundColor Yellow

  }

  Start-Process -Verb runas -FilePath "cmd.exe" -ArgumentList "/C $ScrubberCmdPath "
}

