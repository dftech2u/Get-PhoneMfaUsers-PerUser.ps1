<#
.SYNOPSIS
    Loops through every user and reports who has SMS and/or voice registered for MFA.
    Also checks whether the tenant has Entra ID P1/P2 (needed for the faster registration report).

.PARAMETER TenantId
    Optional. Tenant ID or domain (handy when working across client tenants).

.PARAMETER IncludeGuests
    Include guest (B2B) accounts. Excluded by default.

.PARAMETER IncludeDisabled
    Include disabled accounts. Excluded by default.

.EXAMPLE
    .\Get-PhoneMfaUsers-PerUser.ps1
    .\Get-PhoneMfaUsers-PerUser.ps1 -TenantId contoso.onmicrosoft.com -IncludeDisabled

.NOTES
    Requires: Microsoft.Graph PowerShell SDK (Install-Module Microsoft.Graph -Scope CurrentUser)
    Works in Windows PowerShell 5.1 and PowerShell 7.
    Permissions: User.Read.All, UserAuthenticationMethod.Read.All, Organization.Read.All, AuditLog.Read.All
    Role: Global Reader or Authentication Administrator (minimum) to read other users' methods.
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [switch]$IncludeGuests,
    [switch]$IncludeDisabled,
    [string]$OutputPath = ".\PhoneMfaUsers_$(Get-Date -Format yyyyMMdd_HHmm).csv"
)

# ---------------------------------------------------------------
# Connect
# ---------------------------------------------------------------
$connectParams = @{
    Scopes    = 'User.Read.All','UserAuthenticationMethod.Read.All','Organization.Read.All','AuditLog.Read.All'
    NoWelcome = $true
}
if ($TenantId) { $connectParams.TenantId = $TenantId }
Connect-MgGraph @connectParams
$ctx = Get-MgContext
Write-Host "Connected to tenant $($ctx.TenantId) as $($ctx.Account)`n" -ForegroundColor Cyan

# ---------------------------------------------------------------
# 1. Licensing check
# ---------------------------------------------------------------
Write-Host "=== Entra ID P1/P2 license check ===" -ForegroundColor Cyan

# 1a. Look for an active SKU that includes the P1 or P2 service plan
#     (covers standalone Entra ID P1/P2, M365 Business Premium, E3/E5, EMS, etc.)
$premiumPlans = 'AAD_PREMIUM','AAD_PREMIUM_P2'
$premiumSkus = Get-MgSubscribedSku -All | Where-Object {
    $_.CapabilityStatus -eq 'Enabled' -and
    ($_.ServicePlans | Where-Object { $_.ServicePlanName -in $premiumPlans -and $_.ProvisioningStatus -eq 'Success' })
}

if ($premiumSkus) {
    Write-Host "  SKU check:    P1/P2 found in $(($premiumSkus.SkuPartNumber) -join ', ')" -ForegroundColor Green
} else {
    Write-Host "  SKU check:    No active SKU includes Entra ID P1/P2" -ForegroundColor Yellow
}

# 1b. Live probe of the registration report endpoint - confirms whether it actually works,
#     and separates "no license" from "no permission" when it doesn't.
try {
    $null = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/reports/authenticationMethods/userRegistrationDetails?$top=1' -ErrorAction Stop
    Write-Host "  Report probe: Registration report API is available (the fast script will work here)" -ForegroundColor Green
}
catch {
    $msg = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
    $reason = if ($msg -match 'NonPremiumTenant|premium|SKU') {
        'Tenant lacks Entra ID P1/P2 licensing'
    } elseif ($msg -match 'Authorization_RequestDenied|Forbidden|insufficient') {
        'Permission/role issue (not licensing) - check AuditLog.Read.All consent and your admin role'
    } else {
        "Unknown error: $msg"
    }
    Write-Host "  Report probe: NOT available - $reason" -ForegroundColor Yellow
}
Write-Host ""

# ---------------------------------------------------------------
# 2. Get users
# ---------------------------------------------------------------
$userParams = @{
    All      = $true
    Property = 'Id','UserPrincipalName','DisplayName','AccountEnabled','UserType'
}
if (-not $IncludeDisabled) { $userParams.Filter = 'accountEnabled eq true' }

$users = Get-MgUser @userParams
if (-not $IncludeGuests) { $users = $users | Where-Object { $_.UserType -ne 'Guest' } }
$users = @($users)
$total = $users.Count
Write-Host "Checking $total users (this makes 2 Graph calls per user)...`n" -ForegroundColor Cyan

# ---------------------------------------------------------------
# 3. Loop through every user
# ---------------------------------------------------------------
$results  = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[object]]::new()
$voicePrefs = 'voiceMobile','voiceAlternateMobile','voiceOffice'
$i = 0

foreach ($u in $users) {
    $i++
    Write-Progress -Activity 'Checking MFA methods' -Status "$i of $total : $($u.UserPrincipalName)" -PercentComplete (($i / [math]::Max($total,1)) * 100)

    # All registered methods in one call
    try {
        $methods = @((Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$($u.Id)/authentication/methods" -ErrorAction Stop).value)
    }
    catch {
        $failures.Add([PSCustomObject]@{ UPN = $u.UserPrincipalName; Error = $_.Exception.Message })
        continue
    }

    $phones = @($methods | Where-Object { $_['@odata.type'] -eq '#microsoft.graph.phoneAuthenticationMethod' })
    if ($phones.Count -eq 0) { continue }

    # mobile = SMS + voice; alternateMobile / office = voice only
    $phoneTypes = @($phones | ForEach-Object { $_['phoneType'] })
    $canSms     = $phoneTypes -contains 'mobile'
    $capability = if ($canSms) { 'SMS + Voice' } else { 'Voice only' }

    # Non-phone methods (ignore password and email, which aren't MFA sign-in methods)
    $otherMethods = @($methods | ForEach-Object {
        $_['@odata.type'] -replace '^#microsoft\.graph\.', '' -replace 'AuthenticationMethod$', ''
    } | Where-Object { $_ -notin 'phone','password','email' })

    # Default MFA method (beta endpoint - this is what tells SMS apart from voice)
    $defaultMethod = $null; $systemPreferred = $null
    try {
        $pref = Invoke-MgGraphRequest -Method GET -Uri "beta/users/$($u.Id)/authentication/signInPreferences" -ErrorAction Stop
        $raw  = $pref['userPreferredMethodForSecondaryAuthentication']
        $defaultMethod = if ($raw -eq 'sms') { 'SMS' } elseif ($raw -in $voicePrefs) { "Voice ($raw)" } elseif ($raw) { $raw } else { 'Not set' }
        $systemPreferred = $pref['systemPreferredAuthenticationMethod']
    }
    catch {
        $defaultMethod = 'Unable to read'
    }

    $smsSignIn = ($phones | Where-Object { $_['phoneType'] -eq 'mobile' } | ForEach-Object { $_['smsSignInState'] }) -join ', '

    $results.Add([PSCustomObject]@{
        DisplayName      = $u.DisplayName
        UPN              = $u.UserPrincipalName
        Enabled          = $u.AccountEnabled
        UserType         = $u.UserType
        PhoneCapability  = $capability
        PhoneTypes       = $phoneTypes -join ', '
        PhoneNumbers     = ($phones | ForEach-Object { $_['phoneNumber'] }) -join ', '
        DefaultMfaMethod = $defaultMethod
        SystemPreferred  = $systemPreferred
        PhoneOnly        = ($otherMethods.Count -eq 0)
        OtherMethods     = $otherMethods -join ', '
        SmsSignInState   = $smsSignIn
    })
}
Write-Progress -Activity 'Checking MFA methods' -Completed

# ---------------------------------------------------------------
# 4. Output
# ---------------------------------------------------------------
$results | Sort-Object UPN |
    Format-Table DisplayName, UPN, PhoneCapability, DefaultMfaMethod, PhoneOnly -AutoSize

$results | Sort-Object UPN | Export-Csv $OutputPath -NoTypeInformation

Write-Host "=== Summary ===" -ForegroundColor Cyan
Write-Host "  Users checked:                $total"
Write-Host "  With a phone method:          $($results.Count)"
Write-Host "    SMS + Voice capable:        $(@($results | Where-Object PhoneCapability -eq 'SMS + Voice').Count)"
Write-Host "    Voice only:                 $(@($results | Where-Object PhoneCapability -eq 'Voice only').Count)"
Write-Host "    Default method = SMS:       $(@($results | Where-Object DefaultMfaMethod -eq 'SMS').Count)"
Write-Host "    Default method = Voice:     $(@($results | Where-Object DefaultMfaMethod -like 'Voice*').Count)"
Write-Host "    Phone is ONLY MFA method:   $(@($results | Where-Object PhoneOnly).Count)" -ForegroundColor Yellow
Write-Host "  Exported to $OutputPath" -ForegroundColor Green
Write-Host "  Closing connection to Tenant" -ForegroundColor Red
Disconnect-MgGraph

if ($failures.Count -gt 0) {
    $failPath = $OutputPath -replace '\.csv$', '_errors.csv'
    $failures | Export-Csv $failPath -NoTypeInformation
    Write-Host "  $($failures.Count) users could not be read - see $failPath" -ForegroundColor Red
}
