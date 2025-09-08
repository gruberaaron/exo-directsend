
# exo-directsend.ps1
# <#
# .SYNOPSIS
#     Exchange Online Direct Send Management Script
#
# .DESCRIPTION
#     Provides a menu-driven interface for managing Exchange Online direct send settings and connectors.
#     Includes auto-update/version check, auto-connect, and best-practice error handling.
#
# .AUTHOR
#     Aaron Gruber
#
# .LICENSE
#     BSD 3-Clause License (see LICENSE file)
#
# .NOTES
#     - Requires PowerShell 5.1+ and ExchangeOnlineManagement module.
#     - No admin rights required for module install (uses -Scope CurrentUser).
#     - For feedback or contributions, visit: https://github.com/gruberaaron/exo-directsend
#
# .CREDITS
#     Developed by Aaron Gruber. Inspired by Microsoft documentation and community best practices.
#>

$ScriptVersion = '1.1.0'

# --- Version check: compare local script version to latest on GitHub ---
# Use GitHub API to get the latest release tag
$githubApiUrl = 'https://api.github.com/repos/gruberaaron/exo-directsend/releases/latest'
$localScriptPath = $MyInvocation.MyCommand.Path

try {
    $headers = @{ 'User-Agent' = 'PowerShell' }
    $response = Invoke-WebRequest -Uri $githubApiUrl -Headers $headers -UseBasicParsing | ConvertFrom-Json
    $latestTag = $response.tag_name
    # Remove leading 'v' if present (e.g., v1.2.3 -> 1.2.3)
    if ($latestTag -match '^v') { $latestTag = $latestTag.Substring(1) }
    function Compare-Version($v1, $v2) {
        $a = $v1 -split '\.' | ForEach-Object { [int]$_ }
        $b = $v2 -split '\.' | ForEach-Object { [int]$_ }
        for ($i=0; $i -lt 3; $i++) {
            if ($a[$i] -gt $b[$i]) { return 1 }
            elseif ($a[$i] -lt $b[$i]) { return -1 }
        }
        return 0
    }
    if ($latestTag -and (Compare-Version $latestTag $ScriptVersion) -gt 0) {
        Write-Host "WARNING: A newer version ($latestTag) is available on GitHub. You are running $ScriptVersion." -ForegroundColor Yellow
        $update = Read-Host "Would you like to download and run the latest version now? (y/n)"
        if ($update -eq 'y') {
            $downloadUrl = "https://raw.githubusercontent.com/gruberaaron/exo-directsend/main/exo-directsend.ps1"
            $tempPath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($localScriptPath)) -ChildPath 'exo-directsend-latest.ps1'
            try {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $tempPath -UseBasicParsing
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
# If any error occurs during version check, prompt user to continue
} catch {
    Write-Host "Could not check for script updates: $_" -ForegroundColor Yellow
    if ($_.Exception.Response) {
        Write-Host ("Status Code: {0}" -f $_.Exception.Response.StatusCode.value__)
    }
    Write-Host ("Exception Message: {0}" -f $_.Exception.Message)
    Read-Host "Press Enter to continue with the current version..."
}
# --- End version check ---

# Global variable to track connected tenant name
$Global:TenantName = "Not connected to a tenant"

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

# Automatically connect to Exchange Online at launch
Connect-ExchangeOnlineSession

function Show-Menu {
    Clear-Host
    Write-Host "Exchange Online Direct Send Management " -NoNewline -ForegroundColor Cyan
    Write-Host "(Version: $ScriptVersion)" -ForegroundColor DarkGray
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
    Write-Host "9) Add inbound connector for Zix"
    Write-Host "10) Disconnect and Exit"
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

function Add-ZixConnector {
    $name = "Zix Inbound"
    $subnets = @("63.71.13.0/24","63.71.14.0/24","63.71.15.0/24","199.30.236.0/24","91.209.6.0/24")
    try {
        New-InboundConnector -Name $name `
            -ConnectorType Partner `
            -SenderIPAddresses $subnets `
            -SenderDomains '*' `
            -RequireTls $true `
            -Enabled $true | Out-Null
        Write-Host "Zix inbound connector created."
    } catch {
        Write-Host "Failed to create Zix inbound connector: $_" -ForegroundColor Red
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
        '9' { Add-ZixConnector }
        '10' { Exit-Script }
        default { Write-Host "Invalid selection. Try again."; Pause }
    }
} while ($true)
