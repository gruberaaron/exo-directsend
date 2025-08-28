# --- Version check: compare local script to latest on GitHub ---
$githubRawUrl = 'https://raw.githubusercontent.com/gruberaaron/powershell/main/exo-directsend.ps1'
$localScriptPath = $MyInvocation.MyCommand.Path

try {
    $localContent = Get-Content -Path $localScriptPath -Raw -Encoding UTF8
    $remoteContent = Invoke-WebRequest -Uri $githubRawUrl -UseBasicParsing | Select-Object -ExpandProperty Content
    # Normalize line endings to LF and trim trailing whitespace
    $localContentNorm = ($localContent -replace '\r\n?', "`n") -replace '\s+$',''
    $remoteContentNorm = ($remoteContent -replace '\r\n?', "`n") -replace '\s+$',''
    $localHash = [System.BitConverter]::ToString((New-Object -TypeName System.Security.Cryptography.SHA256Managed).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($localContentNorm))).Replace('-','').ToLower()
    $remoteHash = [System.BitConverter]::ToString((New-Object -TypeName System.Security.Cryptography.SHA256Managed).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($remoteContentNorm))).Replace('-','').ToLower()
    if ($localHash -ne $remoteHash) {
        Write-Host "WARNING: This script is not the latest version from GitHub." -ForegroundColor Yellow
        $update = Read-Host "Would you like to download and run the latest version now? (y/n)"
        if ($update -eq 'y') {
            $tempPath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($localScriptPath)) -ChildPath 'exo-directsend-latest.ps1'
            try {
                Invoke-WebRequest -Uri $githubRawUrl -OutFile $tempPath -UseBasicParsing
                Write-Host "Launching the latest version..." -ForegroundColor Green
                Start-Process pwsh -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", $tempPath
                exit
            } catch {
                Write-Host "Failed to download or launch the latest version: $_" -ForegroundColor Red
                Read-Host "Press Enter to continue with the current version..."
            }
        } else {
            Read-Host "Press Enter to continue with the current version..."
        }
    }
} catch {
    Write-Host "Could not check for script updates: $_" -ForegroundColor DarkYellow
}
# --- End version check ---

# Global variable to track connected tenant name
$Global:TenantName = "Not connected to a tenant"
# exo-directsend.ps1


Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

function Show-Menu {
    Clear-Host
    Write-Host "Exchange Online Direct Send Management" -ForegroundColor Cyan
    Write-Host ""
    Write-Host $Global:TenantName -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1) Connect to Exchange Online"
    Write-Host "2) Show 'rejectdirectsend' setting"
    Write-Host "3) Disable direct send"
    Write-Host "4) Send test message using direct send"
    Write-Host "5) List inbound connectors"
    Write-Host "6) Create new inbound connector"
    Write-Host "7) Add inbound connector for KnowBe4"
    Write-Host "8) Add inbound connector for Securence"
    Write-Host "9) Disconnect and Exit"
}

function Connect-ExchangeOnlineSession {
    # Check if ExchangeOnlineManagement module is installed
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "The ExchangeOnlineManagement module is not installed." -ForegroundColor Yellow
        $install = Read-Host "Do you want to install it now? (y/n)"
        if ($install -eq 'y') {
            try {
                Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
                Write-Host "Module installed successfully."
            } catch {
                Write-Host "Failed to install module: $_" -ForegroundColor Red
                return
            }
        } else {
            Write-Host "Cannot connect without ExchangeOnlineManagement. Returning to menu..." -ForegroundColor Red
            return
        }
    }
    Write-Host "WARNING: You must use a Global Admin account from the target Microsoft 365 tenant to connect." -ForegroundColor Yellow
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline 6>$null
        $org = Get-OrganizationConfig
        $Global:TenantName = "Connected to tenant: $($org.Name)"
        Write-Host "Connected to Exchange Online."
    } catch {
        $Global:TenantName = "Not connected to a tenant"
        Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    }
}

function Show-RejectDirectSend {
    Get-OrganizationConfig | Format-List rejectdirectsend
    Pause
}

function Disable-DirectSend {
    $confirm = Read-Host "Are you sure you want to disable direct send? (y/n)"
    if ($confirm -eq 'y') {
        try {
            Set-OrganizationConfig -RejectDirectSend $true
            Write-Host "Direct send has been disabled."
        } catch {
            Write-Host "Failed to disable direct send: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Operation cancelled."
    }
    Pause
}

function Send-TestDirectSend {
    Write-Host "NOTE: Both the sender and recipient email addresses must be valid users within your Microsoft 365 tenant for this test to succeed." -ForegroundColor Yellow
    $mx = Read-Host "Enter recipient domain MX record"
    $from = Read-Host "Enter sender email address"
    $to = Read-Host "Enter recipient email address"
    $subject = "Direct Send Test"
    $body = "This is a test message sent using direct send."

    $smtp = New-Object Net.Mail.SmtpClient($mx,25)
    $smtp.EnableSsl = $false
    $mail = New-Object Net.Mail.MailMessage($from, $to, $subject, $body)
    try {
        $smtp.Send($mail)
        Write-Host "Test message sent successfully."
    } catch {
        Write-Host "Failed to send message: $_"
    }
    Pause
}

function Show-InboundConnectors {
    Get-InboundConnector | Format-list -Property Name,Enabled,ConnectorType,SenderIPAddresses,`
    TlsSenderCertificateName,RestrictDomainsToIPAddresses,RestrictDomainsToCertificate,RequireTls
    Pause
}

function New-Connector {
    $name = Read-Host "Enter connector name"
    $types = @("Partner", "OnPremises")
    Write-Host "Select connector type:"
    for ($i=0; $i -lt $types.Count; $i++) {
        Write-Host "$($i+1)) $($types[$i])"
    }
    $typeIndex = [int](Read-Host "Enter number (1 or 2)") - 1
    $type = $types[$typeIndex]
    $ips = Read-Host "Enter sender IP address(es)/subnet(s) (comma separated)"
    $tls = Read-Host "Require TLS? (y/n)"
    $enabled = Read-Host "Enable connector? (y/n)"

    try {
        New-InboundConnector -Name $name `
            -ConnectorType $type `
            -SenderIPAddresses ($ips -split ",") `
            -SenderDomains '*' `
            -RequireTls ($tls -eq 'y') `
            -Enabled ($enabled -eq 'y') | Out-Null
        Write-Host "Inbound connector created."
    } catch {
        Write-Host "Failed to create inbound connector: $_" -ForegroundColor Red
    }
    Pause
}

function Add-KnowBe4Connector {
    $name = "KnowBe4 Inbound"
    $subnet = "147.160.167.0/26"
    try {
        New-InboundConnector -Name $name `
            -ConnectorType Partner `
            -SenderIPAddresses $subnet `
            -SenderDomains '*' `
            -RequireTls $true `
            -Enabled $true | Out-Null
        Write-Host "KnowBe4 inbound connector created."
    } catch {
        Write-Host "Failed to create KnowBe4 inbound connector: $_" -ForegroundColor Red
    }
    Pause
}

function Add-SecurenceConnector {
    $name = "Securence Inbound"
    $subnet = "216.17.3.0/24"
    try {
        New-InboundConnector -Name $name `
            -ConnectorType Partner `
            -SenderIPAddresses $subnet `
            -SenderDomains '*' `
            -RequireTls $true `
            -Enabled $true | Out-Null
        Write-Host "Securence inbound connector created."
    } catch {
        Write-Host "Failed to create Securence inbound connector: $_" -ForegroundColor Red
    }
    Pause
}

function Exit-Script {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
    } catch {
        # Ignore errors if not connected
    }
    $Global:TenantName = "Not connected to a tenant"
    Write-Host "Exiting..."
    exit
}

do {
    Show-Menu
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Connect-ExchangeOnlineSession }
        '2' { Show-RejectDirectSend }
        '3' { Disable-DirectSend }
        '4' { Send-TestDirectSend }
        '5' { Show-InboundConnectors }
        '6' { New-Connector }
        '7' { Add-KnowBe4Connector }
        '8' { Add-SecurenceConnector }
        '9' { Exit-Script }
        default { Write-Host "Invalid selection. Try again."; Pause }
    }
} while ($true)
