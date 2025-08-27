# Exchange Online Direct Send Management PowerShell Scripts

This repository contains PowerShell scripts to help manage Exchange Online direct send settings and connectors. These tools are designed for Microsoft 365 administrators who need to configure, test, and manage direct send scenarios in their tenant.

## Features

- Connect to Exchange Online (with module check and install prompt)
- View and modify the 'rejectdirectsend' setting
- Disable direct send with confirmation
- Send a test message using direct send (requires valid tenant users)
- List all inbound connectors
- Create new inbound connectors (with prompts for all required parameters)
- Add pre-configured inbound connectors for KnowBe4 and Securence

## Usage

1. **Prerequisites**
   - PowerShell 5.1 or later
   - Global Admin credentials for the target Microsoft 365 tenant
   - Internet access to install the ExchangeOnlineManagement module if not already present

2. **Running the Script**
   - Open a PowerShell terminal.
   - Navigate to the directory containing the script files.
   - Run the script:

     ```powershell
     .\exo-directsend.ps1
     ```

   - Follow the on-screen menu prompts.

3. **Important Notes**
   - When connecting to Exchange Online, you must use a Global Admin account from the target tenant.
   - For the test message feature, both sender and recipient addresses must be valid users within your Microsoft 365 tenant.
   - The script will prompt to install the ExchangeOnlineManagement module if it is not already installed.

## Menu Options

1. Connect to Exchange Online
2. Show 'rejectdirectsend' setting
3. Disable direct send
4. Send test message using direct send
5. List inbound connectors
6. Create new inbound connector
7. Add inbound connector for KnowBe4
8. Add inbound connector for Securence
9. Exit

## License

This repository is licensed under the MIT License. See the `LICENSE` file for details.
