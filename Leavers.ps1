# convert multiple M365 user mailboxes to shared mailboxes, confirm, then remove direct licenses.
# authored by Sam Lynas

function Ensure-InstalledModuleVersion {
    param(
        [string]$Name,
        [string]$Version
    )

    $installed = Get-Module -ListAvailable $Name | Where-Object { $_.Version -eq [version]$Version }
    if (-not $installed) {
        Write-Host "Installing $Name $Version..." -ForegroundColor Yellow
        Install-Module $Name -RequiredVersion $Version -Scope CurrentUser -Force
    }
}

function Ensure-ImportedModuleVersion {
    param(
        [string]$Name,
        [string]$Version
    )

    $loaded = Get-Module $Name | Where-Object { $_.Version -eq [version]$Version }
    if (-not $loaded) {
        Write-Host "Importing $Name $Version..." -ForegroundColor Yellow
        Import-Module $Name -RequiredVersion $Version
    }
}

# collect users until blank
$userEmails = @()
do {
    $inputEmail = Read-Host "Enter a user's email address (leave blank to start processing)"
    $inputEmail = $inputEmail.Trim()
    if (-not [string]::IsNullOrWhiteSpace($inputEmail)) {
        $userEmails += $inputEmail
    }
} until ([string]::IsNullOrWhiteSpace($inputEmail))

if ($userEmails.Count -eq 0) {
    Write-Error "No email addresses were entered. Aborting."
    exit 1
}

# install only if missing
if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}

Ensure-InstalledModuleVersion -Name "Microsoft.Graph.Authentication" -Version "2.35.0"
Ensure-InstalledModuleVersion -Name "Microsoft.Graph.Users" -Version "2.35.0"
Ensure-InstalledModuleVersion -Name "Microsoft.Graph.Users.Actions" -Version "2.35.0"

# import only if not already loaded
Ensure-ImportedModuleVersion -Name "Microsoft.Graph.Authentication" -Version "2.35.0"
Ensure-ImportedModuleVersion -Name "Microsoft.Graph.Users" -Version "2.35.0"
Ensure-ImportedModuleVersion -Name "Microsoft.Graph.Users.Actions" -Version "2.35.0"

if (-not (Get-Module ExchangeOnlineManagement)) {
    Write-Host "Importing ExchangeOnlineManagement..." -ForegroundColor Yellow
    Import-Module ExchangeOnlineManagement
}

# connect GRAPH FIRST, then EXO
$mgContext = Get-MgContext -ErrorAction SilentlyContinue
if (-not $mgContext) {
    Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All"
}

Connect-ExchangeOnline

# track results
$successfulShared = @()
$failedShared = @()

# pass 1: convert all mailboxes first
foreach ($userEmail in $userEmails) {
    Write-Host ""
    Write-Host "Starting conversion for $userEmail" -ForegroundColor Yellow

    try {
        Set-Mailbox -Identity $userEmail -Type Shared -ErrorAction Stop

        $maxTries = 10
        $tryCount = 0

        do {
            Start-Sleep -Seconds 30
            $mailbox = Get-Mailbox -Identity $userEmail -ErrorAction Stop
            $tryCount++
            Write-Host "Check $tryCount/$maxTries - Mailbox type for ${userEmail}: $($mailbox.RecipientTypeDetails)" -ForegroundColor Cyan
        } until ($mailbox.RecipientTypeDetails -eq "SharedMailbox" -or $tryCount -ge $maxTries)

        if ($mailbox.RecipientTypeDetails -eq "SharedMailbox") {
            Write-Host "$userEmail confirmed as SharedMailbox." -ForegroundColor Green
            $successfulShared += $userEmail
        }
        else {
            Write-Host "$userEmail failed to convert to SharedMailbox after $maxTries checks." -ForegroundColor Red
            $failedShared += $userEmail
        }
    }
    catch {
        Write-Host "Conversion failed for ${userEmail}: $($_.Exception.Message)" -ForegroundColor Red
        $failedShared += $userEmail
    }
}

# pass 2: remove licenses only for successfully converted mailboxes
foreach ($userEmail in $successfulShared) {
    Write-Host ""
    Write-Host "Starting license removal for $userEmail" -ForegroundColor Yellow

    try {
        $user = Get-MgUser -UserId $userEmail -ErrorAction Stop
        $licenses = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop

        if ($licenses) {
            $skuIds = @($licenses | Select-Object -ExpandProperty SkuId)
            Write-Host "Removing $($skuIds.Count) direct license(s) from $userEmail..." -ForegroundColor Yellow
            Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses $skuIds -ErrorAction Stop
            Start-Sleep -Seconds 15
        }
        else {
            Write-Host "No direct licenses found on $userEmail." -ForegroundColor Yellow
        }

        $remainingLicenses = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop

        if (($remainingLicenses | Measure-Object).Count -eq 0) {
            Write-Host "$userEmail is a SharedMailbox and has no direct licenses assigned." -ForegroundColor Green
        }
        else {
            Write-Host "$userEmail is a SharedMailbox, but some licenses still remain." -ForegroundColor Yellow
            Write-Host "These may be group-assigned and cannot be removed directly with Set-MgUserLicense." -ForegroundColor Yellow
            $remainingLicenses | Select-Object SkuPartNumber, SkuId | Format-Table -AutoSize
        }
    }
    catch {
        Write-Host "License removal failed for ${userEmail}: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Processing complete." -ForegroundColor Cyan
Write-Host "Converted successfully: $($successfulShared.Count)" -ForegroundColor Green
Write-Host "Failed conversion: $($failedShared.Count)" -ForegroundColor Red

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
