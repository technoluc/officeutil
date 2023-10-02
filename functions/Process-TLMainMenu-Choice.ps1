function Process-TLMainMenu-Choice {
    param (
        [string]$choice
    )
  
    switch ($choice) {
        'q' {
            Write-Host "Exiting..."
        }
        '0' {
            Write-Host "Exiting..."
        }
        '1' {
            Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-RestMethod win.technoluc.nl | Invoke-Expression" -Wait
            Show-TLMainMenu
        }
        '2' {
            # Check if script was run as Administrator, relaunch if not
            if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Clear-Host
                Start-Process -FilePath powershell.exe -ArgumentList "Invoke-RestMethod https://github.com/technoluc/recycle-bin-themes/raw/test-exit/RecycleBinThemes.ps1 | Invoke-Expression" -Wait -NoNewWindow
                Show-TLMainMenu
                }
            else {
                Read-Host "BinUtil can't be run as Administrator..."
                break
            }

        }
        '3' {
            # Check if script was run as Administrator, relaunch if not
            if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Write-Output "OfficeUtil needs to be run as Administrator. Attempting to relaunch."
                Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -UseBasicParsing `"$ScriptUrl`" | Invoke-Expression" 
                break
            }
            Show-OfficeMainMenu
        }
        default {
            # Read-Host "Press Enter to continue..."
            # Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-TLMainMenu
        }
    }
}
  