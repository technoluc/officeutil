function Install-Office {
    param (
        [string]$product
    )

    Install-OdtIfNeeded

    if ($product -eq "Office 365 Business") {
        $configXMLPath = $configuration365XMLPath
        $downloadURL = $configuration365XMLUrl
        $unattendedArgs = $UnattendedArgs365
    }
    elseif ($product -eq "Office 2021 Pro Plus") {
        $configXMLPath = $configuration21XMLPath
        $downloadURL = $configuration21XMLUrl
        $unattendedArgs = $UnattendedArgs21
    }
    else {
        Write-Host "Invalid product selection."
        return
    }

    if (-not (Test-Path -Path $configXMLPath -PathType Leaf)) {
        Write-Host "Downloading $product Configuration File..." -ForegroundColor Cyan
        Download-File -url $downloadURL -outputPath $configXMLPath
    }

    Write-Host -NoNewline "Install Microsoft $product? ( Y / N ): "
    $choice = [System.Console]::ReadKey().KeyChar
    Write-Host ""

    switch ($choice) {
        'y' {
            Write-Host "Installation started. Don't close this window" -ForegroundColor Green
            Start-Process -Wait $setupExePath -ArgumentList "$unattendedArgs"
            Write-Host "Installation of Microsoft $product completed." -ForegroundColor Green
            Show-OfficeMainMenu
        }
        'n' {
            Write-Host "Exiting..."
            Show-OfficeInstallMenu
        }
        default {
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Invoke-Logo
            Show-OfficeInstallMenu
        }
    }
}
