function Run-OfficeScrubber {
  Write-Host "Select [R] Remove all Licenses option in OfficeScrubber." -ForegroundColor Yellow
  Get-OfficeScrubber -ArchiveUrl $ArchiveUrl -ScrubberPath $ScrubberPath -ScrubberArchive $ScrubberArchive
  Start-Process -Verb runas -FilePath "cmd.exe" -ArgumentList "/C $ScrubberFullPath "
  Read-Host "Press Enter to continue..."
  Remove-Item -Path $ScrubberFullPath -Force
}
