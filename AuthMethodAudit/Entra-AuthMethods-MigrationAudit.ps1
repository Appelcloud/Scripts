#Requires -Version 5.1
<#
.SYNOPSIS
    Microsoft Entra ID Authentication Methods Migration Audit Script
    
.DESCRIPTION
    This script audits your Microsoft Entra ID tenant for the MC678069 mandatory migration 
    from Legacy MFA and SSPR to the unified Authentication Methods policy.
    Deadline: September 30, 2025
    
.NOTES
    Version:        1.0
    Author:         Alexander Appelby
    Creation Date:  2025
    Purpose:        MC678069 - Legacy MFA/SSPR to Authentication Methods Migration
    
.EXAMPLE
    .\Entra-AuthMethods-MigrationAudit.ps1
    
.REQUIREMENTS
    - Microsoft Graph PowerShell SDK
    - Global Administrator or Authentication Policy Administrator role
    - Tenant must have Microsoft Entra ID P1 or P2
    - (For Excel export) ImportExcel PowerShell module
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = (Get-Location).Path,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipModuleCheck,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportHTML = $true,
    
    # Create a single Excel workbook with multiple sheets instead of separate CSVs
    [Parameter(Mandatory=$false)]
    [switch]$ExportExcel = $true,
    
    # Legacy CSV export (off by default)
    [Parameter(Mandatory=$false)]
    [switch]$ExportCSV = $false,
    
    
    # Include resource accounts (shared mailboxes/rooms). If not set, script
    # will attempt to discover and exclude resource accounts via Places API.
    [Parameter(Mandatory=$false)]
    [switch]$IncludeResources = $false,
    
    # Optional: read policy status (requires Policy.Read.All)
    [Parameter(Mandatory=$false)]
    [switch]$IncludePolicyStatus = $false
)

# Script configuration
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$Script:TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:ReportDate = Get-Date -Format "MMMM dd, yyyy HH:mm"

# Color codes for console output
$Script:Colors = @{
    Success = "Green"
    Warning = "Yellow"
    Error = "Red"
    Info = "Cyan"
    Header = "Magenta"
}

#region Helper Functions

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Get-SafeFileName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name
    )
    # Replace characters invalid or awkward in file names with underscore
    return ($Name -replace "[^A-Za-z0-9._-]+", "_")
}

function Test-Prerequisites {
    Write-ColorOutput "`n=== Checking Prerequisites ===" -Color $Script:Colors.Header
    
    if (-not $SkipModuleCheck) {
        try {
            # Check for Microsoft Graph module
            $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph* | Where-Object {$_.Name -eq "Microsoft.Graph.Authentication"}
            
            if (-not $graphModule) {
                Write-ColorOutput "Microsoft Graph PowerShell SDK not found. Installing..." -Color $Script:Colors.Warning
                Install-Module Microsoft.Graph -Scope CurrentUser -Force
                Write-ColorOutput "Microsoft Graph module installed successfully." -Color $Script:Colors.Success
            }
            else {
                Write-ColorOutput "Microsoft Graph module found: v$($graphModule.Version)" -Color $Script:Colors.Success
            }
            
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Import-Module Microsoft.Graph.Reports -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

            if ($ExportExcel -and -not $SkipModuleCheck) {
                # Ensure ImportExcel is available for Excel export
                $importExcel = Get-Module -ListAvailable -Name ImportExcel
                if (-not $importExcel) {
                    Write-ColorOutput "ImportExcel module not found. Installing..." -Color $Script:Colors.Warning
                    Install-Module ImportExcel -Scope CurrentUser -Force
                    Write-ColorOutput "ImportExcel module installed successfully." -Color $Script:Colors.Success
                }
                Import-Module ImportExcel -ErrorAction Stop
            }
            
        }
        catch {
            Write-ColorOutput "Failed to load required modules: $_" -Color $Script:Colors.Error
            exit 1
        }
    }
    
    # Test output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    Write-ColorOutput "Prerequisites check completed." -Color $Script:Colors.Success
}

function Connect-MicrosoftGraph {
    Write-ColorOutput "`n=== Connecting to Microsoft Graph ===" -Color $Script:Colors.Header
    
    try {
        # Least-privileged by default; add optional scopes based on selected features
        $requiredScopes = @(
            "User.Read.All",
            "UserAuthenticationMethod.Read.All",
            "Reports.Read.All",
            "Organization.Read.All"
        )
        if ($IncludePolicyStatus) { $requiredScopes += "Policy.Read.All" }
        if (-not $IncludeResources) { $requiredScopes += "Place.Read.All" }

        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        
        $context = Get-MgContext
        Write-ColorOutput "Connected to tenant: $($context.TenantId)" -Color $Script:Colors.Success
        Write-ColorOutput "Account: $($context.Account)" -Color $Script:Colors.Info

        # Fetch tenant display name for report naming
        $tenantDisplayName = $null
        try {
            $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization?`$select=displayName,tenantId,verifiedDomains"
            if ($org -and $org.value -and $org.value.Count -gt 0) {
                $tenantDisplayName = $org.value[0].displayName
                if (-not $tenantDisplayName -and $org.value[0].verifiedDomains) {
                    $defaultDomain = ($org.value[0].verifiedDomains | Where-Object { $_.isDefault -or $_.isInitial } | Select-Object -First 1).name
                    if ($defaultDomain) { $tenantDisplayName = $defaultDomain }
                }
            }
        } catch { }

        if (-not $tenantDisplayName) { $tenantDisplayName = $context.TenantId }
        $Script:TenantName = Get-SafeFileName -Name $tenantDisplayName

        # Return minimal context with tenant info
        return [PSCustomObject]@{
            TenantId    = $context.TenantId
            Account     = $context.Account
            TenantName  = $tenantDisplayName
        }
    }
    catch {
        Write-ColorOutput "Failed to connect to Microsoft Graph: $_" -Color $Script:Colors.Error
        exit 1
    }
}

#endregion

#region Data Collection Functions

function Test-EntraP1Requirement {
    Write-ColorOutput "`n=== Checking Tenant License (Entra ID P1+) ===" -Color $Script:Colors.Header
    try {
        $uri = "https://graph.microsoft.com/v1.0/subscribedSkus?`$select=skuPartNumber,servicePlans,capabilityStatus"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $skus = @()
        if ($resp -and $resp.value) { $skus = $resp.value }
        $hasP1 = $false

        foreach ($sku in $skus) {
            if ($sku.servicePlans) {
                foreach ($sp in $sku.servicePlans) {
                    $name = $sp.servicePlanName
                    $status = $sp.provisioningStatus
                    if (($name -eq 'AAD_PREMIUM' -or $name -eq 'AAD_PREMIUM_P2') -and $status -eq 'Success') {
                        $hasP1 = $true
                        break
                    }
                }
            }
            if ($hasP1) { break }
        }

        if (-not $hasP1) {
            Write-ColorOutput "License check failed: Entra ID P1/P2 not detected. This script requires at least P1." -Color $Script:Colors.Error
            exit 1
        }

        Write-ColorOutput "License check passed: Entra ID P1/P2 available." -Color $Script:Colors.Success
    }
    catch {
        Write-ColorOutput "Unable to verify tenant license: $_" -Color $Script:Colors.Error
        exit 1
    }
}

function Get-AllUsers {
    Write-ColorOutput "`n=== Retrieving User Information ===" -Color $Script:Colors.Header
    
    try {
        $users = @()
        $nextLink = $null
        $count = 0
        
        do {
            if ($nextLink) {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink
            }
            else {
                # Exclude resource accounts (e.g., rooms/shared) by default via accountEnabled filter
                $filter = if ($IncludeResources) {
                    "userType eq 'Member'"
                } else {
                    "userType eq 'Member' and accountEnabled eq true"
                }
                $select = "id,userPrincipalName,displayName,assignedLicenses,accountEnabled,mail"
                $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=$select&`$top=999"
            }
            
            $users += $response.value
            $nextLink = $response.'@odata.nextLink'
            $count += $response.value.Count
            
            Write-ColorOutput "Retrieved $count users..." -Color $Script:Colors.Info
            
        } while ($nextLink)
        
        Write-ColorOutput "Total users retrieved: $($users.Count)" -Color $Script:Colors.Success
        return $users
    }
    catch {
        Write-ColorOutput "Failed to retrieve users: $_" -Color $Script:Colors.Error
        return @()
    }
}

function Get-UserAuthenticationMethods {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Users
    )
    
    Write-ColorOutput "`n=== Analyzing Authentication Methods ===" -Color $Script:Colors.Header
    
    $results = @()
    $processedCount = 0
    $totalUsers = $Users.Count
    
    foreach ($user in $Users) {
        $processedCount++
        
        if ($processedCount % 50 -eq 0) {
            Write-ColorOutput "Processing user $processedCount of $totalUsers..." -Color $Script:Colors.Info
        }
        
        try {
            # Get authentication methods for user using beta endpoint via direct Graph request
            $authMethods = @()
            $authUri = "https://graph.microsoft.com/beta/users/$($user.id)/authentication/methods"
            do {
                $authResponse = Invoke-MgGraphRequest -Method GET -Uri $authUri -ErrorAction Stop
                if ($authResponse -and $authResponse.value) {
                    $authMethods += $authResponse.value
                }
                $authUri = $authResponse.'@odata.nextLink'
            } while ($authUri)
            
            # Initialize method flags
            $methodDetails = @{
                Email = $false
                Fido2 = $false
                MicrosoftAuthenticator = $false
                Phone = $false
                SoftwareOath = $false
                WindowsHello = $false
                TemporaryAccessPass = $false
                Certificate = $false
                Password = $false
            }
            
            $methodsList = @()
            
            foreach ($method in $authMethods) {
                # Support both typed SDK objects and raw Graph JSON from Invoke-MgGraphRequest
                $methodType = $null
                try {
                    if ($method -is [System.Collections.IDictionary]) {
                        $methodType = $method['@odata.type']
                    }
                    elseif ($method.PSObject -and ($method.PSObject.Properties.Name -contains '@odata.type')) {
                        $methodType = $method.'@odata.type'
                    }
                    elseif ($method.AdditionalProperties) {
                        $methodType = $method.AdditionalProperties['@odata.type']
                    }
                } catch { }
                
                switch ($methodType) {
                    "#microsoft.graph.emailAuthenticationMethod" { 
                        $methodDetails.Email = $true
                        $methodsList += "Email"
                    }
                    "#microsoft.graph.fido2AuthenticationMethod" { 
                        $methodDetails.Fido2 = $true
                        $methodsList += "FIDO2/Passkey"
                    }
                    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { 
                        $methodDetails.MicrosoftAuthenticator = $true
                        $methodsList += "Microsoft Authenticator"
                    }
                    "#microsoft.graph.phoneAuthenticationMethod" { 
                        $methodDetails.Phone = $true
                        $methodsList += "Phone"
                    }
                    "#microsoft.graph.softwareOathAuthenticationMethod" { 
                        $methodDetails.SoftwareOath = $true
                        $methodsList += "Software OATH"
                    }
                    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { 
                        $methodDetails.WindowsHello = $true
                        $methodsList += "Windows Hello"
                    }
                    "#microsoft.graph.temporaryAccessPassAuthenticationMethod" { 
                        $methodDetails.TemporaryAccessPass = $true
                        $methodsList += "TAP"
                    }
                    "#microsoft.graph.x509CertificateAuthenticationMethod" { 
                        $methodDetails.Certificate = $true
                        $methodsList += "Certificate"
                    }
                    "#microsoft.graph.passwordAuthenticationMethod" { 
                        $methodDetails.Password = $true
                    }
                }
            }
            
            # Determine authentication strength
            $hasStrongMethod = $methodDetails.Fido2 -or $methodDetails.MicrosoftAuthenticator -or 
                              $methodDetails.WindowsHello -or $methodDetails.Certificate
            
            $hasWeakMethod = $methodDetails.Phone -or $methodDetails.Email
            
            $authStrength = if ($hasStrongMethod) { "Strong" } 
                           elseif ($hasWeakMethod) { "Weak" } 
                           else { "None" }
            
            # MFA Status
            $mfaStatus = if ($authMethods.Count -gt 1) { "Enabled" } else { "Disabled" }
            
            $userResult = [PSCustomObject]@{
                UserPrincipalName = $user.userPrincipalName
                DisplayName = $user.displayName
                AccountEnabled = $user.accountEnabled
                MFAStatus = $mfaStatus
                AuthenticationStrength = $authStrength
                MethodCount = $authMethods.Count - 1  # Subtract password
                RegisteredMethods = ($methodsList -join ", ")
                HasEmail = $methodDetails.Email
                HasFido2 = $methodDetails.Fido2
                HasAuthenticatorApp = $methodDetails.MicrosoftAuthenticator
                HasPhone = $methodDetails.Phone
                HasSoftwareOath = $methodDetails.SoftwareOath
                HasWindowsHello = $methodDetails.WindowsHello
                HasTAP = $methodDetails.TemporaryAccessPass
                HasCertificate = $methodDetails.Certificate
                NeedsAction = ($authStrength -eq "None" -or $authStrength -eq "Weak")
                RequiredActions = ""
            }
            
            # Determine required actions
            $actions = @()
            if ($authStrength -eq "None") {
                $actions += "Register at least one authentication method"
            }
            elseif ($authStrength -eq "Weak") {
                $actions += "Register a strong authentication method (Authenticator App, FIDO2, or Windows Hello)"
            }
            if (-not $methodDetails.Email -and -not $methodDetails.Phone) {
                $actions += "Register email or phone for SSPR"
            }
            
            $userResult.RequiredActions = $actions -join "; "
            
            $results += $userResult
        }
        catch {
            Write-ColorOutput "Error processing user $($user.userPrincipalName): $_" -Color $Script:Colors.Warning
            
            $results += [PSCustomObject]@{
                UserPrincipalName = $user.userPrincipalName
                DisplayName = $user.displayName
                AccountEnabled = $user.accountEnabled
                MFAStatus = "Error"
                AuthenticationStrength = "Unknown"
                MethodCount = 0
                RegisteredMethods = "Error retrieving methods"
                HasEmail = $false
                HasFido2 = $false
                HasAuthenticatorApp = $false
                HasPhone = $false
                HasSoftwareOath = $false
                HasWindowsHello = $false
                HasTAP = $false
                HasCertificate = $false
                NeedsAction = $true
                RequiredActions = "Unable to determine - manual review required"
            }
        }
    }
    
    Write-ColorOutput "Authentication methods analysis completed." -Color $Script:Colors.Success
    return $results
}

function Get-AuthenticationMethodsRegistrationDetails {
    Write-ColorOutput "`n=== Getting Registration Details from Reports ===" -Color $Script:Colors.Header
    
    try {
        $registrationDetails = @()
        $uri = "https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails"
        
        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri
            $registrationDetails += $response.value
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        Write-ColorOutput "Retrieved registration details for $($registrationDetails.Count) users" -Color $Script:Colors.Success
        return $registrationDetails
    }
    catch {
        Write-ColorOutput "Failed to retrieve registration details: $_" -Color $Script:Colors.Error
        return @()
    }
}

function Get-CurrentPolicyStatus {
    Write-ColorOutput "`n=== Checking Current Policy Configuration ===" -Color $Script:Colors.Header
    
    try {
        $authPolicy = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/authenticationMethodsPolicy" -ErrorAction Stop
        
        $policyStatus = [PSCustomObject]@{
            MigrationState = if ($authPolicy.policyMigrationState) { $authPolicy.policyMigrationState } else { "Unknown" }
            LastModified = if ($authPolicy.lastModifiedDateTime) { $authPolicy.lastModifiedDateTime } else { "Unknown" }
            RegistrationEnforcement = $authPolicy.registrationEnforcement
            ReportingSuspicious = $authPolicy.reportSuspiciousActivitySettings
        }
        
        Write-ColorOutput "Current migration state: $($policyStatus.MigrationState)" -Color $Script:Colors.Info
        
        return $policyStatus
    }
    catch {
        Write-ColorOutput "Warning: Could not retrieve policy status (may require additional permissions)" -Color $Script:Colors.Warning
        Write-ColorOutput "Error details: $_" -Color $Script:Colors.Warning
        return [PSCustomObject]@{
            MigrationState = "Unable to determine"
            LastModified = "Unknown"
            RegistrationEnforcement = $null
            ReportingSuspicious = $null
        }
    }
}

#endregion

#region Analysis Functions

function Get-MigrationStatistics {
    param(
        [Parameter(Mandatory=$true)]
        [array]$AuthMethodsData,
        
        [Parameter(Mandatory=$false)]
        [array]$RegistrationDetails
    )
    
    Write-ColorOutput "`n=== Calculating Migration Statistics ===" -Color $Script:Colors.Header
    
    $stats = [PSCustomObject]@{
        TotalUsers = $AuthMethodsData.Count
        EnabledAccounts = ($AuthMethodsData | Where-Object {$_.AccountEnabled -eq $true}).Count
        DisabledAccounts = ($AuthMethodsData | Where-Object {$_.AccountEnabled -eq $false}).Count
        
        # MFA Statistics
        MFAEnabled = ($AuthMethodsData | Where-Object {$_.MFAStatus -eq "Enabled"}).Count
        MFADisabled = ($AuthMethodsData | Where-Object {$_.MFAStatus -eq "Disabled"}).Count
        MFAPercentage = 0
        
        # Authentication Strength
        StrongAuth = ($AuthMethodsData | Where-Object {$_.AuthenticationStrength -eq "Strong"}).Count
        WeakAuth = ($AuthMethodsData | Where-Object {$_.AuthenticationStrength -eq "Weak"}).Count
        NoAuth = ($AuthMethodsData | Where-Object {$_.AuthenticationStrength -eq "None"}).Count
        
        # Method Distribution
        UsersWithEmail = ($AuthMethodsData | Where-Object {$_.HasEmail}).Count
        UsersWithPhone = ($AuthMethodsData | Where-Object {$_.HasPhone}).Count
        UsersWithAuthenticator = ($AuthMethodsData | Where-Object {$_.HasAuthenticatorApp}).Count
        UsersWithFido2 = ($AuthMethodsData | Where-Object {$_.HasFido2}).Count
        UsersWithWindowsHello = ($AuthMethodsData | Where-Object {$_.HasWindowsHello}).Count
        UsersWithCertificate = ($AuthMethodsData | Where-Object {$_.HasCertificate}).Count
        
        # Action Required
        UsersNeedingAction = ($AuthMethodsData | Where-Object {$_.NeedsAction}).Count
        UsersCompliant = ($AuthMethodsData | Where-Object {-not $_.NeedsAction}).Count
        
        # SSPR Statistics (if registration details available)
        SSPRCapable = 0
        SSPRNotCapable = 0
        
        # Risk Categories
        HighRiskUsers = 0
        MediumRiskUsers = 0
        LowRiskUsers = 0
    }
    
    # Calculate percentages
    if ($stats.TotalUsers -gt 0) {
        $stats.MFAPercentage = [math]::Round(($stats.MFAEnabled / $stats.TotalUsers) * 100, 2)
    }
    
    # SSPR Statistics from registration details
    if ($RegistrationDetails) {
        $stats.SSPRCapable = ($RegistrationDetails | Where-Object {$_.isSSPRCapable}).Count
        $stats.SSPRNotCapable = ($RegistrationDetails | Where-Object {-not $_.isSSPRCapable}).Count
    }
    
    # Risk categorization
    foreach ($user in $AuthMethodsData) {
        if ($user.AuthenticationStrength -eq "None") {
            $stats.HighRiskUsers++
        }
        elseif ($user.AuthenticationStrength -eq "Weak") {
            $stats.MediumRiskUsers++
        }
        else {
            $stats.LowRiskUsers++
        }
    }
    
    return $stats
}

function Get-MigrationReadiness {
    param(
        [Parameter(Mandatory=$true)]
        $Statistics,
        
        [Parameter(Mandatory=$false)]
        $PolicyStatus
    )
    
    $readiness = [PSCustomObject]@{
        OverallStatus = "Not Ready"
        ReadinessScore = 0
        CriticalIssues = @()
        Warnings = @()
        Recommendations = @()
    }
    
    # Calculate readiness score (out of 100)
    $score = 0
    
    # MFA Coverage (40 points)
    if ($Statistics.MFAPercentage -ge 95) { $score += 40 }
    elseif ($Statistics.MFAPercentage -ge 80) { $score += 30 }
    elseif ($Statistics.MFAPercentage -ge 60) { $score += 20 }
    elseif ($Statistics.MFAPercentage -ge 40) { $score += 10 }
    
    # Strong Authentication (30 points)
    $strongAuthPercentage = if ($Statistics.TotalUsers -gt 0) { 
        ($Statistics.StrongAuth / $Statistics.TotalUsers) * 100 
    } else { 0 }
    
    if ($strongAuthPercentage -ge 80) { $score += 30 }
    elseif ($strongAuthPercentage -ge 60) { $score += 20 }
    elseif ($strongAuthPercentage -ge 40) { $score += 10 }
    
    # SSPR Coverage (20 points)
    if ($Statistics.SSPRCapable -gt 0 -or $Statistics.SSPRNotCapable -gt 0) {
        $ssprPercentage = ($Statistics.SSPRCapable / ($Statistics.SSPRCapable + $Statistics.SSPRNotCapable)) * 100
        if ($ssprPercentage -ge 80) { $score += 20 }
        elseif ($ssprPercentage -ge 60) { $score += 15 }
        elseif ($ssprPercentage -ge 40) { $score += 10 }
    }
    
    # No high-risk users (10 points)
    if ($Statistics.HighRiskUsers -eq 0) { $score += 10 }
    elseif ($Statistics.HighRiskUsers -le 5) { $score += 5 }
    
    $readiness.ReadinessScore = $score
    
    # Determine overall status
    if ($score -ge 85) {
        $readiness.OverallStatus = "Ready"
    }
    elseif ($score -ge 60) {
        $readiness.OverallStatus = "Partially Ready"
    }
    else {
        $readiness.OverallStatus = "Not Ready"
    }
    
    # Critical Issues
    if ($Statistics.MFAPercentage -lt 50) {
        $readiness.CriticalIssues += "Less than 50% of users have MFA enabled"
    }
    if ($Statistics.HighRiskUsers -gt 10) {
        $readiness.CriticalIssues += "$($Statistics.HighRiskUsers) users have no authentication methods registered"
    }
    if ($PolicyStatus -and $PolicyStatus.MigrationState -eq "preMigration") {
        $readiness.CriticalIssues += "Migration has not been started (still in preMigration state)"
    }
    
    # Warnings
    if ($Statistics.WeakAuth -gt ($Statistics.TotalUsers * 0.3)) {
        $readiness.Warnings += "More than 30% of users only have weak authentication methods"
    }
    if ($Statistics.UsersWithAuthenticator -lt ($Statistics.TotalUsers * 0.5)) {
        $readiness.Warnings += "Less than 50% of users have Microsoft Authenticator configured"
    }
    
    # Recommendations
    $readiness.Recommendations += "Enable MFA for all users (currently $($Statistics.MFAPercentage)%)"
    $readiness.Recommendations += "Migrate users from phone/SMS to Microsoft Authenticator or FIDO2"
    $readiness.Recommendations += "Ensure all users register at least 2 authentication methods"
    $readiness.Recommendations += "Complete migration before September 30, 2025 deadline"
    
    if ($PolicyStatus -and $PolicyStatus.MigrationState -eq "preMigration") {
        $readiness.Recommendations += "Start migration by setting policy state to 'migrationInProgress'"
    }
    
    return $readiness
}

#endregion

#region Report Generation Functions

function Get-ResourceEmailAddresses {
    Write-ColorOutput "`n=== Discovering Resource Mailboxes (Rooms/Sharedmailboxes) ===" -Color $Script:Colors.Header
    $resourceEmails = @()
    try {
        $uri = "https://graph.microsoft.com/v1.0/places/microsoft.graph.room?`$select=emailAddress&`$top=999"
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($resp.value) {
                $resourceEmails += ($resp.value | ForEach-Object { $_.emailAddress })
            }
            $uri = $resp.'@odata.nextLink'
        } while ($uri)
    } catch {
        Write-ColorOutput "Warning: Could not enumerate rooms via Places API: $_" -Color $Script:Colors.Warning
    }
    
    try {
        # Workspace emailAddress is only available in beta; avoid $select to prevent schema errors
        $uri = "https://graph.microsoft.com/beta/places/microsoft.graph.workspace?`$top=999"
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($resp.value) {
                $resourceEmails += ($resp.value | ForEach-Object { $_.emailAddress } | Where-Object { $_ })
            }
            $uri = $resp.'@odata.nextLink'
        } while ($uri)
    } catch {
        Write-ColorOutput "Warning: Could not enumerate workspaces via Places API: $_" -Color $Script:Colors.Warning
    }
    
    $clean = ($resourceEmails | Where-Object { $_ } | ForEach-Object { $_.ToLower() } | Select-Object -Unique)
    Write-ColorOutput "Discovered $($clean.Count) resource addresses (rooms/sharedmailboxes)" -Color $Script:Colors.Info
    return $clean
}

function Export-CSVReports {
    param(
        [Parameter(Mandatory=$true)]
        $AuthMethodsData,
        
        [Parameter(Mandatory=$true)]
        $Statistics,
        
        [Parameter(Mandatory=$true)]
        $Readiness
    )
    
    Write-ColorOutput "`n=== Generating CSV Reports ===" -Color $Script:Colors.Header
    
    try {
        # Main authentication methods report
        $mainReportPath = Join-Path $OutputPath ("AuthMethods_MigrationReport_{0}_Detailed.csv" -f $Script:TenantName)
        $AuthMethodsData | Export-Csv -Path $mainReportPath -NoTypeInformation
        Write-ColorOutput "Detailed report exported: $mainReportPath" -Color $Script:Colors.Success
        
        # Users needing action report
        $actionRequiredUsers = $AuthMethodsData | Where-Object {$_.NeedsAction}
        if ($actionRequiredUsers) {
            $actionReportPath = Join-Path $OutputPath ("AuthMethods_MigrationReport_{0}_ActionRequired.csv" -f $Script:TenantName)
            $actionRequiredUsers | Select-Object UserPrincipalName, DisplayName, AuthenticationStrength, RequiredActions | 
                Export-Csv -Path $actionReportPath -NoTypeInformation
            Write-ColorOutput "Action required report exported: $actionReportPath" -Color $Script:Colors.Success
        }
        
        # Summary statistics
        $summaryPath = Join-Path $OutputPath ("AuthMethods_MigrationReport_{0}_Summary.csv" -f $Script:TenantName)
        $Statistics | Export-Csv -Path $summaryPath -NoTypeInformation
        Write-ColorOutput "Summary statistics exported: $summaryPath" -Color $Script:Colors.Success
        
        return $true
    }
    catch {
        Write-ColorOutput "Failed to export CSV reports: $_" -Color $Script:Colors.Error
        return $false
    }
}

function Export-ExcelReport {
    param(
        [Parameter(Mandatory=$true)]
        $AuthMethodsData,
        
        [Parameter(Mandatory=$true)]
        $Statistics,
        
        [Parameter(Mandatory=$true)]
        $Readiness
    )
    
    Write-ColorOutput "`n=== Generating Excel Report ===" -Color $Script:Colors.Header
    
    try {
        $excelPath = Join-Path $OutputPath ("AuthMethods_MigrationReport_{0}.xlsx" -f $Script:TenantName)
        
        # Detailed sheet (blue styling)
        $AuthMethodsData | Export-Excel -Path $excelPath -WorksheetName "Detailed Report" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "Detailed" -TableStyle Medium2
        
        # Action Required sheet (include SSPR contact requirement-only users)
        $actionRequiredUsers = $AuthMethodsData | Where-Object { $_.NeedsAction -or ($_.RequiredActions -match 'Register email or phone for SSPR') }
        if ($actionRequiredUsers) {
            $actionRequiredUsers | Select-Object UserPrincipalName, DisplayName, AuthenticationStrength, RequiredActions |
                Export-Excel -Path $excelPath -WorksheetName "Action Required" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "ActionRequired" -TableStyle Medium2 -Append
        }
        
        # Summary sheet (single row)
        @($Statistics) | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "Summary" -TableStyle Medium2 -Append
        
        Write-ColorOutput "Excel report generated: $excelPath" -Color $Script:Colors.Success
        return $true
    }
    catch {
        Write-ColorOutput "Failed to generate Excel report: $_" -Color $Script:Colors.Error
        return $false
    }
}

function Export-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        $AuthMethodsData,
        
        [Parameter(Mandatory=$true)]
        $Statistics,
        
        [Parameter(Mandatory=$true)]
        $Readiness,
        
        [Parameter(Mandatory=$false)]
        $PolicyStatus,
        
        [Parameter(Mandatory=$false)]
        $TenantInfo
    )
    
    Write-ColorOutput "`n=== Generating HTML Report ===" -Color $Script:Colors.Header
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Authentication Methods Migration Report - MC678069</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: #f5f6f8;
            min-height: 100vh;
            padding: 20px;
        }
        .container { 
            max-width: 1400px; 
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: #4f46e5;
            color: white;
            padding: 40px;
            text-align: center;
        }
        .header h1 { 
            font-size: 2.5em; 
            margin-bottom: 10px;
            font-weight: 300;
        }
        .header p { 
            font-size: 1.2em; 
            opacity: 0.9;
        }
        .deadline-banner {
            background: #ff6b6b;
            color: white;
            padding: 15px;
            text-align: center;
            font-weight: bold;
            font-size: 1.1em;
        }
        .content { padding: 40px; }
        
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }
        .summary-card {
            background: #4f46e5;
            color: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        .summary-card h3 {
            font-size: 0.9em;
            opacity: 0.9;
            margin-bottom: 10px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .summary-card .value {
            font-size: 2.5em;
            font-weight: bold;
        }
        .summary-card .subtitle {
            font-size: 0.9em;
            opacity: 0.8;
            margin-top: 5px;
        }
        
        .readiness-section {
            background: #f8f9fa;
            padding: 30px;
            border-radius: 15px;
            margin-bottom: 40px;
        }
        .readiness-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
        }
        .readiness-status {
            font-size: 2em;
            font-weight: bold;
        }
        .status-ready { color: #28a745; }
        .status-partial { color: #ffc107; }
        .status-notready { color: #dc3545; }
        
        /* pie chart removed */
        
        .issues-container {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        .issue-box {
            padding: 20px;
            border-radius: 10px;
            border-left: 4px solid;
        }
        .critical-issues {
            background: #ffeaea;
            border-color: #dc3545;
        }
        .warnings {
            background: #fff3cd;
            border-color: #ffc107;
        }
        .recommendations {
            background: #d4edda;
            border-color: #28a745;
        }
        .issue-box h4 {
            margin-bottom: 10px;
            color: #333;
        }
        .issue-box ul {
            list-style-position: inside;
            color: #666;
        }
        .issue-box li {
            margin-bottom: 8px;
        }
        
        .chart-container {
            background: white;
            padding: 30px;
            border-radius: 15px;
            margin-bottom: 40px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        }
        .chart-container h2 {
            margin-bottom: 20px;
            color: #333;
        }
        .chart-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 30px;
        }
        .chart {
            text-align: center;
        }
        .bar-chart {
            display: flex;
            justify-content: space-around;
            align-items: flex-end;
            height: 200px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
        }
        .bar {
            width: 60px;
            background: #4f46e5;
            border-radius: 5px 5px 0 0;
            position: relative;
            transition: all 0.3s;
        }
        .bar:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 15px rgba(79, 70, 229, 0.4);
        }
        .bar-label {
            position: absolute;
            bottom: -25px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 0.8em;
            white-space: nowrap;
        }
        .bar-value {
            position: absolute;
            top: -25px;
            left: 50%;
            transform: translateX(-50%);
            font-weight: bold;
            color: #4f46e5;
        }
        
        .table-container {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            margin-bottom: 40px;
            overflow-x: auto;
        }
        .table-container h2 {
            margin-bottom: 20px;
            color: #333;
        }
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th {
            background: #4f46e5;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 500;
        }
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
        }
        tr:hover {
            background: #f8f9fa;
        }
        .needs-action {
            background: #ffeaea;
        }
        .compliant {
            background: #d4edda;
        }
        
        .method-badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 0.75em;
            margin: 2px;
            background: #e9ecef;
            color: #495057;
        }
        .method-badge.strong {
            background: #28a745;
            color: white;
        }
        .method-badge.weak {
            background: #ffc107;
            color: #333;
        }
        
        .footer {
            background: #f8f9fa;
            padding: 30px;
            text-align: center;
            color: #666;
        }
        
        @media print {
            body { background: white; }
            .container { box-shadow: none; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Authentication Methods Migration Report</h1>
            <p>Microsoft Message Center: MC678069</p>
            <p>Generated: $Script:ReportDate</p>
"@

    if ($TenantInfo) {
        $htmlContent += "<p>Tenant: $($TenantInfo.TenantId)</p>"
    }

    $htmlContent += @"
        </div>
        
        <div class="deadline-banner">
            ⚠️ DEADLINE: September 30, 2025 - Legacy MFA and SSPR policies will be retired.
        </div>
        
        <div class="content">
            <!-- Summary Statistics -->
            <div class="summary-grid">
                <div class="summary-card">
                    <h3>Total Users</h3>
                    <div class="value">$($Statistics.TotalUsers)</div>
                    <div class="subtitle">$($Statistics.EnabledAccounts) active</div>
                </div>
                <div class="summary-card">
                    <h3>MFA Coverage</h3>
                    <div class="value">$($Statistics.MFAPercentage)%</div>
                    <div class="subtitle">$($Statistics.MFAEnabled) users protected</div>
                </div>
                <div class="summary-card">
                    <h3>Strong Auth</h3>
                    <div class="value">$($Statistics.StrongAuth)</div>
                    <div class="subtitle">Users with modern methods</div>
                </div>
                <div class="summary-card">
                    <h3>Action Required</h3>
                    <div class="value">$($Statistics.UsersNeedingAction)</div>
                    <div class="subtitle">Users need configuration</div>
                </div>
            </div>
            
            <!-- Readiness Assessment -->
            <div class="readiness-section">
                <div class="readiness-header">
                    <div>
                        <h2>Migration Readiness Assessment</h2>
                        <div class="readiness-status status-$(($Readiness.OverallStatus -replace ' ','').ToLower())">
                            Status: $($Readiness.OverallStatus)
                        </div>
"@

    if ($PolicyStatus) {
        $htmlContent += "<p style='margin-top: 10px; color: #666;'>Current Migration State: <strong>$($PolicyStatus.MigrationState)</strong></p>"
    }

    $htmlContent += @"
                    </div>
                    <div style="font-size: 1.1em; color: #333; font-weight: 600;">
                        Readiness Score: $($Readiness.ReadinessScore)/100
                    </div>
                </div>
                
                <div class="issues-container">
"@

    # Critical Issues
    if ($Readiness.CriticalIssues.Count -gt 0) {
        $htmlContent += @"
                    <div class="issue-box critical-issues">
                        <h4>🚨 Critical Issues</h4>
                        <ul>
"@
        foreach ($issue in $Readiness.CriticalIssues) {
            $htmlContent += "<li>$issue</li>"
        }
        $htmlContent += @"
                        </ul>
                    </div>
"@
    }

    # Warnings
    if ($Readiness.Warnings.Count -gt 0) {
        $htmlContent += @"
                    <div class="issue-box warnings">
                        <h4>⚠️ Warnings</h4>
                        <ul>
"@
        foreach ($warning in $Readiness.Warnings) {
            $htmlContent += "<li>$warning</li>"
        }
        $htmlContent += @"
                        </ul>
                    </div>
"@
    }

    # Recommendations
    $htmlContent += @"
                    <div class="issue-box recommendations">
                        <h4>✅ Recommendations</h4>
                        <ul>
"@
    foreach ($rec in $Readiness.Recommendations) {
        $htmlContent += "<li>$rec</li>"
    }
    $htmlContent += @"
                        </ul>
                    </div>
                </div>
            </div>
            
            <!-- Authentication Methods Distribution -->
            <div class="chart-container">
                <h2>Authentication Methods Distribution</h2>
                <div class="bar-chart">
                    <div class="bar" style="height: $(($Statistics.UsersWithAuthenticator / $Statistics.TotalUsers * 100))%">
                        <span class="bar-value">$($Statistics.UsersWithAuthenticator)</span>
                        <span class="bar-label">Authenticator</span>
                    </div>
                    <div class="bar" style="height: $(($Statistics.UsersWithPhone / $Statistics.TotalUsers * 100))%">
                        <span class="bar-value">$($Statistics.UsersWithPhone)</span>
                        <span class="bar-label">Phone</span>
                    </div>
                    <div class="bar" style="height: $(($Statistics.UsersWithEmail / $Statistics.TotalUsers * 100))%">
                        <span class="bar-value">$($Statistics.UsersWithEmail)</span>
                        <span class="bar-label">Email</span>
                    </div>
                    <div class="bar" style="height: $(($Statistics.UsersWithFido2 / $Statistics.TotalUsers * 100))%">
                        <span class="bar-value">$($Statistics.UsersWithFido2)</span>
                        <span class="bar-label">FIDO2</span>
                    </div>
                    <div class="bar" style="height: $(($Statistics.UsersWithWindowsHello / $Statistics.TotalUsers * 100))%">
                        <span class="bar-value">$($Statistics.UsersWithWindowsHello)</span>
                        <span class="bar-label">Win Hello</span>
                    </div>
                </div>
                <div style="margin-top: 14px; background: #e8f4ff; color: #084298; border: 1px solid #b6daff; padding: 12px 14px; border-radius: 8px; font-size: 1em;">
                    <strong>Note:</strong> The full lists users are included in the Excel export (see the sheets).
                </div>
            </div>
            
            <!-- Users Requiring Action -->
"@

    $htmlContent += @"
            <div class="table-container">
                <h2>Users Requiring Action</h2>
                <table>
                    <thead>
                        <tr>
                            <th>User Principal Name</th>
                            <th>Display Name</th>
                            <th>Auth Strength</th>
                            <th>Current Methods</th>
                            <th>Required Actions</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    # Add top 50 users needing action, including SSPR contact requirement
    $allActionUsers = $AuthMethodsData | Where-Object { $_.NeedsAction -or ($_.RequiredActions -match 'Register email or phone for SSPR') }
    $actionUsers = $allActionUsers | Select-Object -First 50
    foreach ($user in $actionUsers) {
        $rowClass = if ($user.AuthenticationStrength -eq "None") { "needs-action" } else { "" }
        $strengthClass = if ($user.AuthenticationStrength -eq "Strong") { "strong" } 
                        elseif ($user.AuthenticationStrength -eq "Weak") { "weak" } 
                        else { "" }
        
        $htmlContent += @"
                        <tr class="$rowClass">
                            <td>$($user.UserPrincipalName)</td>
                            <td>$($user.DisplayName)</td>
                            <td><span class="method-badge $strengthClass">$($user.AuthenticationStrength)</span></td>
                            <td>$($user.RegisteredMethods)</td>
                            <td>$($user.RequiredActions)</td>
                        </tr>
"@
    }

    if ($actionUsers.Count -eq 50 -and $allActionUsers.Count -gt 50) {
        $htmlContent += @"
                        <tr>
                            <td colspan="5" style="text-align: center; font-style: italic;">
                                ... and $($allActionUsers.Count - 50) more users. See Excel/CSV export for complete list.
                            </td>
                        </tr>
"@
    }

    $htmlContent += @"
                    </tbody>
                </table>
            </div>
"@

    $htmlContent += @"
            
            <!-- Distribution Summary -->
            <div class="chart-container">
                <h2>Risk Categories</h2>
                <div class="chart-grid">
                    <div class="chart">
                        <h3>Risk Categories</h3>
                        <div style="display: flex; justify-content: space-around; margin-top: 20px;">
                            <div style="text-align: center;">
                                <div style="font-size: 2em; color: #dc3545; font-weight: bold;">$($Statistics.HighRiskUsers)</div>
                                <div style="color: #666;">High Risk</div>
                            </div>
                            <div style="text-align: center;">
                                <div style="font-size: 2em; color: #ffc107; font-weight: bold;">$($Statistics.MediumRiskUsers)</div>
                                <div style="color: #666;">Medium Risk</div>
                            </div>
                            <div style="text-align: center;">
                                <div style="font-size: 2em; color: #28a745; font-weight: bold;">$($Statistics.LowRiskUsers)</div>
                                <div style="color: #666;">Low Risk</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Official Migration Guide -->
            <div class="readiness-section">
                <h2>Official Migration Guide</h2>
                <p style="margin-top: 10px;">
                    See Microsoft's step-by-step guide for managing Authentication Methods and planning your migration:
                    <a href="https://learn.microsoft.com/en-us/entra/identity/authentication/how-to-authentication-methods-manage" target="_blank" rel="noopener">
                        Authentication methods management – Microsoft Entra
                    </a>
                </p>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>Microsoft MC678069 Compliance Report</strong></p>
            <p>Legacy MFA and SSPR must be migrated to Authentication Methods policy by September 30, 2025</p>
            <p>Report Generated: $Script:ReportDate | Script Version: 1.0</p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $htmlPath = Join-Path $OutputPath ("AuthMethods_MigrationReport_{0}.html" -f $Script:TenantName)
        $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
        Write-ColorOutput "HTML report generated: $htmlPath" -Color $Script:Colors.Success
        return $true
    }
    catch {
        Write-ColorOutput "Failed to generate HTML report: $_" -Color $Script:Colors.Error
        return $false
    }
}

#endregion

#region Main Execution

function Main {
    # Pretty banner
    Write-ColorOutput (@"

=========================================================================
    MICROSOFT ENTRA ID AUTHENTICATION METHODS MIGRATION AUDIT
    Message Center: MC678069
    Deadline: September 30, 2025
=========================================================================
"@
) -Color $Script:Colors.Header

    # Test prerequisites
    Test-Prerequisites
    
    # Connect to Microsoft Graph
    $tenantInfo = Connect-MicrosoftGraph
    
    # Require Entra ID P1 or P2
    Test-EntraP1Requirement
    
    # Get all users
    $users = Get-AllUsers
    if ($users.Count -eq 0) {
        Write-ColorOutput "No users found. Exiting." -Color $Script:Colors.Error
        return
    }
    
    # Exclude resources by default: remove rooms/workspaces and typically unlicensed mailboxes
    if (-not $IncludeResources) {
        $resourceEmails = Get-ResourceEmailAddresses
        $resourceSet = @{}
        foreach ($e in $resourceEmails) { $resourceSet[$e] = $true }
        $before = $users.Count
        $users = $users | Where-Object {
            $upn = ([string]$_.userPrincipalName).ToLower()
            $mail = ([string]$_.mail).ToLower()
            $notRoom = (-not $resourceSet.ContainsKey($upn)) -and (-not $resourceSet.ContainsKey($mail))
            $licenseCount = @($_.assignedLicenses).Count
            $notUnlicensed = ($licenseCount -gt 0)
            $notRoom -and $notUnlicensed
        }
        $after = $users.Count
        Write-ColorOutput "Filtered resources/unlicensed: $before -> $after users" -Color $Script:Colors.Info
    }

    # Get authentication methods for all users
    $authMethodsData = Get-UserAuthenticationMethods -Users $users
    
    # Get registration details
    $registrationDetails = Get-AuthenticationMethodsRegistrationDetails
    
    # Get current policy status (optional)
    $policyStatus = $null
    if ($IncludePolicyStatus) {
        $policyStatus = Get-CurrentPolicyStatus
    }
    
    # Calculate statistics
    $statistics = Get-MigrationStatistics -AuthMethodsData $authMethodsData -RegistrationDetails $registrationDetails
    
    # Assess migration readiness
    $readiness = Get-MigrationReadiness -Statistics $statistics -PolicyStatus $policyStatus
    
    # Display summary
    Write-ColorOutput "`n=== MIGRATION READINESS SUMMARY ===" -Color $Script:Colors.Header
    Write-ColorOutput "Overall Status: $($readiness.OverallStatus)" -Color $(
        if ($readiness.OverallStatus -eq "Ready") { $Script:Colors.Success }
        elseif ($readiness.OverallStatus -eq "Partially Ready") { $Script:Colors.Warning }
        else { $Script:Colors.Error }
    )
    Write-ColorOutput "Readiness Score: $($readiness.ReadinessScore)/100" -Color $Script:Colors.Info
    Write-ColorOutput "Total Users: $($statistics.TotalUsers)" -Color $Script:Colors.Info
    Write-ColorOutput "MFA Coverage: $($statistics.MFAPercentage)%" -Color $Script:Colors.Info
    Write-ColorOutput "Users Needing Action: $($statistics.UsersNeedingAction)" -Color $(
        if ($statistics.UsersNeedingAction -eq 0) { $Script:Colors.Success }
        else { $Script:Colors.Warning }
    )
    
    # Export reports
    if ($ExportExcel) {
        Export-ExcelReport -AuthMethodsData $authMethodsData -Statistics $statistics -Readiness $readiness | Out-Null
    }
    if ($ExportCSV) {
        Export-CSVReports -AuthMethodsData $authMethodsData -Statistics $statistics -Readiness $readiness | Out-Null
    }

    if ($ExportHTML) {
        Export-HTMLReport -AuthMethodsData $authMethodsData -Statistics $statistics -Readiness $readiness -PolicyStatus $policyStatus -TenantInfo $tenantInfo | Out-Null
    }
    
    Write-ColorOutput "`n=== AUDIT COMPLETE ===" -Color $Script:Colors.Success
    Write-ColorOutput "Reports have been saved to: $OutputPath" -Color $Script:Colors.Info
    
    # Disconnect from Microsoft Graph
    Disconnect-MgGraph | Out-Null
}

# Execute main function
Main

#endregion
