#===========================================================================
# Shows the form
#===========================================================================

# Invoke-WPFFormVariables

# Toon het hoofdmenu
Show-MainMenu

Invoke-Logo
Write-Host -NoNewLine "Press F to delete $OfficeUtilPath or any other key to quit: "
$choice = [System.Console]::ReadKey().KeyChar
Write-Host ""
switch ($choice) {
  'f' {
    # Clean up: Remove the downloaded archive
    Stop-Script
  }
  'default' {
    Invoke-Logo
  }
}

# Stop-Script

# Stop-Transcript
