
# exo-directsend

An interactive, menu-driven PowerShell script for managing direct send operations with Exchange Online (Exo). Designed for administrators, this tool provides a guided interface to perform direct send tasks without needing to write or automate scripts.

## Features

- Menu-driven interface for direct send management
- Built with the Exchange Online Management PowerShell Module
- User-friendly and suitable for manual administration
- Customizable and extendable PowerShell script


## Prerequisites

- PowerShell 7 or higher (**required**)
- Exchange Online account with appropriate permissions
- Internet connectivity

## Getting Started

1. **Clone the repository:**

   ```powershell
   git clone https://github.com/gruberaaron/exo-directsend.git
   cd exo-directsend
   ```

2. **Review and update the script:**
   - Open `exo-directsend.ps1` in your preferred editor.
   - Update any configuration variables as needed.


3. **Run the script:**

   > **Note:** This script requires PowerShell 7 or higher. If you are using Windows PowerShell 5.1, you must [install PowerShell 7](https://github.com/PowerShell/PowerShell/releases/latest) and run the script using `pwsh`.

   ```powershell
   pwsh .\exo-directsend.ps1
   ```

## Usage

Customize the script parameters as needed for your environment. Refer to the comments in `exo-directsend.ps1` for available options and usage examples.

## Contributing

Contributions are welcome! Please open issues or submit pull requests for improvements or bug fixes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Disclaimer

This script is provided as-is without warranty. Use at your own risk. Always test in a non-production environment before deploying to production.
