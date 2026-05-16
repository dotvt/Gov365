
param(

    [ValidateRange(1,3650)]
    [int]$ExpectedValidity = 365,

    [ValidateRange(1,365)]
    [int]$ExpirationThreshold = 30,

    [ValidateSet("Yes","No")]
    [string]$NotifyExpiration = "No",

    [ValidateSet("Yes","No")]
    [string]$NotifyNonCompliant = "No",

    [string]$SenderMailbox = "entra-alerts@contoso.com",

    [string]$TemplatePath = "./templates/credential-alert-fr.html"
)

# =========================
# FLAGS
# =========================

$NotifyExpirationBool = $NotifyExpiration -eq "Yes"
$NotifyNonCompliantBool = $NotifyNonCompliant -eq "Yes"

# =========================
# GRAPH CONNECTION
# =========================

Connect-MgGraph -Scopes `
"Application.Read.All",
"Directory.Read.All",
"User.Read.All",
"Mail.Send"

Import-Module Microsoft.Graph.Applications

# =========================
# LOAD TEMPLATE
# =========================

$template = Get-Content $TemplatePath -Raw

# =========================
# STORAGE
# =========================

$ownerAlerts = @{}
$csvLog = @()

function Add-Alert {
    param($Owner, $Item)

    if (-not $ownerAlerts.ContainsKey($Owner)) {
        $ownerAlerts[$Owner] = @()
    }

    $ownerAlerts[$Owner] += $Item
}

function Add-CsvLog {
    param(
        $App,
        $Credential,
        $Type,
        $Reason,
        $OwnerStatus,
        $MailSent,
        $Owners,
        $SecretDurationDays,
        $ExpirationDate,
        $RemainingDays,
        $NonCompliant
    )

    $script:csvLog += [PSCustomObject]@{
        Application              = $App
        Credential              = $Credential
        Type                    = $Type
        OwnerStatus             = $OwnerStatus
        AppOwners               = ($Owners -join "; ")
        SecretDurationDays      = $SecretDurationDays
        ExpirationDate          = $ExpirationDate
        RemainingDays           = $RemainingDays
        BaselineDays            = $ExpectedValidity
        ExpirationThresholdDays = $ExpirationThreshold
        ComplianceStatus        = if ($NonCompliant) { "NON_COMPLIANT" } else { "OK" }
        Reason                  = $Reason
        MailSent                = $MailSent
    }
}

# =========================
# GET APPLICATIONS
# =========================

$apps = Get-MgApplication -All

foreach ($app in $apps) {

    Write-Host "Scanning: $($app.DisplayName)" -ForegroundColor DarkGray

    $owners = Get-MgApplicationOwner -ApplicationId $app.Id -All

    $ownerEmails = @(
        $owners |
        Where-Object { $_.AdditionalProperties.mail } |
        ForEach-Object { $_.AdditionalProperties.mail }
    )

    $hasOwners = $ownerEmails.Count -gt 0

    # =========================
    # SECRETS
    # =========================

    foreach ($secret in @($app.PasswordCredentials)) {

        $durationDays = [math]::Round((New-TimeSpan $secret.StartDateTime $secret.EndDateTime).TotalDays)
        $remainingDays = [math]::Round((New-TimeSpan (Get-Date) $secret.EndDateTime).TotalDays)

        $nonCompliant = $durationDays -gt $ExpectedValidity

        $notify =
            ($NotifyNonCompliantBool -and $nonCompliant) -or
            ($NotifyExpirationBool -and $remainingDays -le $ExpirationThreshold)

        if ($notify) {

            $reason = if ($nonCompliant) {
                "NON-CONFORME : durée du secret ($durationDays jours) > baseline ($ExpectedValidity jours)"
            } else {
                "Expiration proche ($remainingDays jours)"
            }

            if (-not $hasOwners) {

                Add-CsvLog $app.DisplayName $secret.DisplayName "Secret" $reason "NO_OWNER" "NO" @() $durationDays $secret.EndDateTime $remainingDays $nonCompliant

            } else {

                foreach ($mail in $ownerEmails) {

                    Add-Alert $mail ([PSCustomObject]@{
                        AppName       = $app.DisplayName
                        Credential    = $secret.DisplayName
                        Type          = "Secret"
                        Expiration    = $secret.EndDateTime
                        RemainingDays = $remainingDays
                        NonCompliant  = $nonCompliant
                    })
                }

                Add-CsvLog $app.DisplayName $secret.DisplayName "Secret" $reason "OWNER_FOUND" "YES" $ownerEmails $durationDays $secret.EndDateTime $remainingDays $nonCompliant
            }
        }
    }

    # =========================
    # CERTIFICATES
    # =========================

    foreach ($cert in @($app.KeyCredentials)) {

        $durationDays = [math]::Round((New-TimeSpan $cert.StartDateTime $cert.EndDateTime).TotalDays)
        $remainingDays = [math]::Round((New-TimeSpan (Get-Date) $cert.EndDateTime).TotalDays)

        $nonCompliant = $durationDays -gt $ExpectedValidity

        $notify =
            ($NotifyNonCompliantBool -and $nonCompliant) -or
            ($NotifyExpirationBool -and $remainingDays -le $ExpirationThreshold)

        if ($notify) {

            $reason = if ($nonCompliant) {
                "NON-CONFORME : durée du certificat ($durationDays jours) > baseline ($ExpectedValidity jours)"
            } else {
                "Expiration proche ($remainingDays jours)"
            }

            if (-not $hasOwners) {

                Add-CsvLog $app.DisplayName $cert.DisplayName "Certificate" $reason "NO_OWNER" "NO" @() $durationDays $cert.EndDateTime $remainingDays $nonCompliant

            } else {

                foreach ($mail in $ownerEmails) {

                    Add-Alert $mail ([PSCustomObject]@{
                        AppName       = $app.DisplayName
                        Credential    = $cert.DisplayName
                        Type          = "Certificate"
                        Expiration    = $cert.EndDateTime
                        RemainingDays = $remainingDays
                        NonCompliant  = $nonCompliant
                    })
                }

                Add-CsvLog $app.DisplayName $cert.DisplayName "Certificate" $reason "OWNER_FOUND" "YES" $ownerEmails $durationDays $cert.EndDateTime $remainingDays $nonCompliant
            }
        }
    }
}

# =========================
# SEND EMAILS
# =========================

foreach ($owner in $ownerAlerts.Keys) {

    $rows = ""

    foreach ($item in $ownerAlerts[$owner]) {

        $appName = $item.AppName
        $credential = $item.Credential
        $type = $item.Type
        $expiration = $item.Expiration
        $remainingDays = $item.RemainingDays
        $nonCompliant = $item.NonCompliant

        if ($type -eq "Secret") {
            $typeBg = "#dbeafe"
            $typeColor = "#1d4ed8"
        } else {
            $typeBg = "#e0e7ff"
            $typeColor = "#4338ca"
        }

        $daysColor = if ($remainingDays -le 7) { "#dc2626" }
        elseif ($remainingDays -le 30) { "#f59e0b" }
        else { "#16a34a" }

        $complianceLabel = if ($nonCompliant) {
            "NON CONFORME (durée du secret)"
        } else {
            "OK"
        }

        $complianceColor = if ($nonCompliant) { "#dc2626" } else { "#16a34a" }

        $rows += @"
<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #e5e7eb;margin-bottom:14px;'>

<tr><td style='background:#f8fafc;padding:14px 16px;border-bottom:1px solid #e5e7eb;'>
<div style='font-size:15px;font-weight:900;color:#111827;'>$appName</div>
</td></tr>

<tr><td style='padding:14px 16px;'>

<div><span style='background:$typeBg;color:$typeColor;font-size:11px;font-weight:800;padding:4px 8px;'>$type</span></div>

<div style='margin-top:8px;font-size:13px;'><b>Credential :</b> $credential</div>

<div style='font-size:13px;'><b>Expiration :</b> $expiration</div>

<div style='margin-top:10px;'>
<b>Jours restants :</b>
<span style='font-weight:900;color:$daysColor;font-size:15px;'>$remainingDays</span>
</div>

<div>
<b>Conformité :</b>
<span style='font-weight:900;color:$complianceColor;'>$complianceLabel</span>
</div>

</td></tr>

</table>
"@
    }

    $body = $template -replace "{Rows}", $rows

    try {
        Send-MgUserMail -UserId $SenderMailbox -BodyParameter @{
            Message = @{
                Subject = "[Entra ID] Synthèse des identifiants applicatifs"
                Body = @{ ContentType = "HTML"; Content = $body }
                Importance = "High"
                ToRecipients = @(@{ EmailAddress = @{ Address = $owner } })
            }
        }

        Write-Host "Sent -> $owner" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed -> $owner" -ForegroundColor Red
    }
}

# =========================
# CSV EXPORT
# =========================

$csvPath = "./EntraID_Credential_Audit.csv"

$csvLog | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Invoke-Item $csvPath

Write-Host "CSV generated and opened: $csvPath" -ForegroundColor Cyan
Write-Host "Done." -ForegroundColor Green