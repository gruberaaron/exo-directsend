# exo-directsend.ps1

Import-Module ExchangeOnlineManagement -ErrorAction Stop

function Show-Menu {
    Clear-Host
    Write-Host "Exchange Online Direct Send Management" -ForegroundColor Cyan
    Write-Host "1) Connect to Exchange Online"
    Write-Host "2) Show 'rejectdirectsend' setting"
    Write-Host "3) Disable direct send"
    Write-Host "4) Send test message using direct send"
    Write-Host "5) List inbound connectors"
    Write-Host "6) Create new inbound connector"
    Write-Host "7) Add inbound connector for KnowBe4"
    Write-Host "8) Add inbound connector for Securence"
    Write-Host "9) Exit"
}

function Connect-ExchangeOnlineSession {
    Connect-ExchangeOnline
}

function Show-RejectDirectSend {
    Get-OrganizationConfig | Format-List rejectdirectsend
    Pause
}

function Disable-DirectSend {
    $confirm = Read-Host "Are you sure you want to disable direct send? (y/n)"
    if ($confirm -eq 'y') {
        Set-OrganizationConfig -RejectDirectSend $true
        Write-Host "Direct send has been disabled."
    } else {
        Write-Host "Operation cancelled."
    }
    Pause
}

function Send-TestDirectSend {
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
    Get-InboundConnector | Format-Table -Property Name,Enabled,ConnectorType,SenderIPAddresses,RequireTls
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

    New-InboundConnector -Name $name `
        -ConnectorType $type `
        -SenderIPAddresses ($ips -split ",") `
        -RequireTls ($tls -eq 'y') `
        -Enabled ($enabled -eq 'y')

    Write-Host "Inbound connector created."
    Pause
}

function Add-KnowBe4Connector {
    $name = "KnowBe4 Inbound"
    $subnet = "147.160.167.0/26"
    New-InboundConnector -Name $name `
        -ConnectorType Partner `
        -SenderIPAddresses $subnet `
        -RequireTls $true `
        -Enabled $true
    Write-Host "KnowBe4 inbound connector created."
    Pause
}

function Add-SecurenceConnector {
    $name = "Securence Inbound"
    $subnet = "216.17.3.0/24"
    New-InboundConnector -Name $name `
        -ConnectorType Partner `
        -SenderIPAddresses $subnet `
        -RequireTls $true `
        -Enabled $true
    Write-Host "Securence inbound connector created."
    Pause
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
        '9' { Write-Host "Exiting..."; exit }
        default { Write-Host "Invalid selection. Try again."; Pause }
    }
} while ($true)