# OfficeUtil.ps1

## Introduction
OfficeUtil.ps1 is a PowerShell script designed to simplify the installation, removal, and management of Microsoft Office on Windows systems. It automates the process of deploying the Office Deployment Tool (ODT), installing Office 365 Business and Office 2021 Pro Plus, as well as providing options for uninstalling Office, activating Office/Windows, and running various maintenance tools.

**Disclaimer:** Always review scripts before execution, and ensure that you trust the source. This script is designed for automated deployment scenarios and is not intended for manual modification.

## Prerequisites
- PowerShell
- 7-Zip (automatically installed if not present)
- Administrator privileges (required for some operations)

## Usage
To quickly run the script, use the following command:
```powershell
$url = "https://get.technoluc.nl/office"
irm $url | iex
```
Replace `$url` with the raw GitHub URL of the OfficeUtil.ps1 script. <br>
Use `$url = "https://raw.githubusercontent.com/technoluc/officeutil/main/OfficeUtil.ps1"` if the other link isn't working properly.


## Features
- Automated download and installation of Office Deployment Tool (ODT)
- Installation of Microsoft Office 365 Business and Office 2021 Pro Plus
- Uninstallation of Microsoft Office using various methods
- Activation of Microsoft Office/Windows using scripts from [massgrave.dev](https://massgrave.dev)
- Running OfficeScrubber for license removal

## Configuration
Several variables at the beginning of the script can be customized to match your environment. Please review and adjust these variables according to your needs.

## Warning
- **Do not modify this file directly:** It is automatically generated and will be overwritten.
- **Always check scripts before running them:** Ensure you understand the script's functionality and trust its source.

## Important Note
Before running the script, it is recommended to review the entire script to understand its behavior and ensure it aligns with your requirements.

## License
This script is provided under the [MIT License](LICENSE.md). Feel free to use, modify, and distribute it according to the terms of the license.

## Acknowledgments
- [TechnoLuc](https://github.com/technoluc) - For the initial creation of the OfficeUtil script.
- [massgrave.dev](https://massgrave.dev) - For providing activation scripts.

## Author
TechnoLuc

## Support
For any issues or questions, please open an [issue](https://github.com/yourusername/yourrepository/issues).

---
*Note: The above template assumes you have a license file named LICENSE.md in the same directory as the script. If you don't have one, you can remove the "License" section or replace it with the appropriate licensing information.*
