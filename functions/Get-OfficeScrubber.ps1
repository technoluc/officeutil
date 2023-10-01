function Get-OfficeScrubber {
  param (
    [string]$ScrubberBaseUrl,
    [string]$OfficeUtilPath,
    [string]$ScrubberArchiveName,
    [string]$7zPath = "C:\Program Files\7-Zip\7z.exe"
  )

  # Combine the path to the archive
  $ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

  # Create the directory if it doesn't exist yet
  if (-not (Test-Path -Path $OfficeUtilPath -PathType Container)) {
    New-Item -Path $OfficeUtilPath -ItemType Directory ;
  }

  try {
    # Download the archive
    Invoke-WebRequest -Uri $ScrubberBaseUrl -OutFile $ScrubberArchivePath

    # Extract the archive using the full path to 7z
    & $7zPath x $ScrubberArchivePath -o"$OfficeUtilPath" -y | Out-Null

    Write-Host "The archive has been successfully downloaded and extracted to: $OfficeUtilPath"
  }
  catch {
    Write-Host "An error occurred while downloading and extracting the archive: $_"
  }
  finally {
    # Clean up: Remove the downloaded archive
    Remove-Item -Path $ScrubberArchivePath -Force
  }
}
