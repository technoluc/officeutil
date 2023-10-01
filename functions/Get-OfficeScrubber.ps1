function Get-OfficeScrubber {
  param (
    [string]$7zPath = "C:\Program Files\7-Zip\7z.exe"
  )

  # Combine the path to the archive
  $ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

  # Create the directory if it doesn't exist yet
  if (-not (Test-Path -Path $ScrubberPath -PathType Container)) {
    New-Item -Path $ScrubberPath -ItemType Directory -Force | Out-Null
  }

  try {
    # Download and extract the archive using 7-Zip
    Invoke-WebRequest -Uri $ScrubberBaseUrl -OutFile $ScrubberArchivePath -UseBasicParsing
    & $7zPath x $ScrubberArchivePath -o"$ScrubberPath" -y > $null

    Write-Host "The archive has been successfully downloaded and extracted to: $ScrubberPath" -ForegroundColor Green
  }
  catch {
    Write-Host "An error occurred while downloading and extracting the archive: $_" -ForegroundColor Red
  }
  finally {
    # Clean up: Remove the downloaded archive
    if (Test-Path -Path $ScrubberArchivePath -PathType Leaf) {
      Remove-Item -Path $ScrubberArchivePath -Force
    }
  }
}
