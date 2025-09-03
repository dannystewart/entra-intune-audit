<#
.SYNOPSIS
    User and Device Master Audit Script - Comprehensive analysis of active users and devices

.DESCRIPTION
    This script performs a comprehensive audit of active users and devices following organizational
    changes. It connects to both Entra ID and Intune to gather detailed information about:
    - Active (non-disabled) user accounts
    - Computers in Entra that have been active since the specified cutoff date (default: 30 days ago)
    - Devices in Intune that have been active since the same date
    - Device join status, enrollment details, and user associations

    Unless otherwise specified, the script will create a timestamped file in the current directory:
    MasterAudit_YYYY-MM-DD_HH-mm-ss.xlsx

.PARAMETER DeviceActivityCutoffDate
    The cutoff date for determining active devices. Defaults to 30 days, as that's usually all the
    activity data available. Dates are UTC and should be in format "2025-06-07 00:00:00". You can
    also use relative dates with "(Get-Date).AddDays(-7)" where 7 is the number of days ago. (Be
    sure to use a negative value to calculate a past date.)

.PARAMETER ActiveUsersGroup
    The display name of the user group to retrieve members from.

.PARAMETER OutputPath
    The filename (or directory path) where the Excel export will be saved.

.EXAMPLE
    .\Get-MasterAuditGeneric.ps1 -ActiveUsersGroup "All Org Staff"
    Runs the script, exporting the final file to the current directory: MasterAudit_YYYY-MM-DD_HH-mm-ss.xlsx

.EXAMPLE
    .\Get-MasterAudit.ps1 -ActiveUsersGroup "All Org Staff" -DeviceActivityCutoffDate (Get-Date).AddDays(-7)
    Uses 7 days ago as the cutoff date

.EXAMPLE
    .\Get-MasterAudit.ps1 -ActiveUsersGroup "All Org Staff" -DeviceActivityCutoffDate "2025-06-07 00:00:00"
    Uses a custom specified cutoff date

.EXAMPLE
    .\Get-MasterAudit.ps1 -ActiveUsersGroup "All Org Staff" -OutputPath "C:\Audit"
    Exports the final file to C:\Audit\MasterAudit_YYYY-MM-DD_HH-mm-ss.xlsx

.EXAMPLE
    .\Get-MasterAudit.ps1 -ActiveUsersGroup "All Org Staff" -OutputPath "C:\OneDrive\MasterAudit.xlsx"
    Creates or updates the specified file with the exact name provided

.NOTES
    Author: Danny Stewart
    Version: 1.1.0
    License: MIT
    Created: 2025-09-03
    Repository: https://github.com/dannystewart/entra-intune-audit
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ActiveUsersGroup,

    [Parameter(Mandatory = $false)]
    [DateTime]$DeviceActivityCutoffDate = (Get-Date).AddDays(-30),

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "."
)

# Script configuration
$ErrorActionPreference = "Continue"
$ProgressPreference = "Continue"

# Initialize timestamp first
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Determine if OutputPath is a directory or full file path
$isDirectory = if ($OutputPath -eq "." -or (Test-Path $OutputPath -PathType Container)) { $true } else { $false }
$outputDirectory = if ($isDirectory) { $OutputPath } else { Split-Path $OutputPath -Parent }
$outputFilename = if ($isDirectory) { "MasterAudit_$timestamp.xlsx" } else { Split-Path $OutputPath -Leaf }

# Initialize logging
$logPath = Join-Path $outputDirectory "MasterAudit.log"

# Add run separator to log file
$separator = "`n" + ("=" * 50) + "`nNEW RUN: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" + ("=" * 50)
Add-Content -Path $logPath -Value $separator -Encoding UTF8

function Write-AuditLog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $logEntry = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $logEntry -ForegroundColor $(
        switch ($Level) {
            "INFO" { "White" }
            "WARNING" { "Yellow" }
            "ERROR" { "Red" }
            "SUCCESS" { "Green" }
        }
    )

    Add-Content -Path $logPath -Value $logEntry -Encoding UTF8
}

function Test-GraphConnection {
    try {
        $context = Get-MgContext
        if (-not $context) {
            Write-AuditLog "No active Graph connection found. Attempting to connect..." -Level "WARNING"
            Connect-MgGraph -Scopes "User.Read.All", "Device.Read.All", "DeviceManagementManagedDevices.Read.All", "Group.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Directory.Read.All" -NoWelcome
            $context = Get-MgContext
        }

        if ($context) {
            # Test basic connectivity
            $testUser = Get-MgUser -Top 1 -ErrorAction SilentlyContinue
            if ($testUser) {
                Write-AuditLog "Connected to Graph API." -Level "SUCCESS"
                return $true
            }
        }

        Write-AuditLog "Failed to establish Graph connection." -Level "ERROR"
        return $false
    } catch {
        Write-AuditLog "Graph connection error: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Get-ActiveUsers {
    Write-AuditLog "Retrieving users from '$ActiveUsersGroup' group..."

    try {
        # Get the specified user group
        $staffGroup = Get-MgGroup -Filter "displayName eq '$ActiveUsersGroup'"
        if (-not $staffGroup) {
            Write-AuditLog "Could not find '$ActiveUsersGroup' group!" -Level "ERROR"
            throw "'$ActiveUsersGroup' group not found."
        }

        # Get group members
        $groupMembers = Get-MgGroupMember -GroupId $staffGroup.Id -All
        $userIds = $groupMembers | ForEach-Object { $_.Id }

        Write-AuditLog "Processing $($userIds.Count) user members from the group." -Level "INFO"

        # Test if SignInActivity is accessible before processing batches
        $canAccessSignInActivity = $false

        try {
            # Test with a single user to see if SignInActivity works
            $testUser = Get-MgUser -Top 1 -Property "Id,SignInActivity" -ErrorAction SilentlyContinue -ErrorVariable testError
            if ($testUser -and -not $testError) {
                $canAccessSignInActivity = $true
            } else {
                throw $testError[0]
            }
        } catch {
            $errorMessage = $_.Exception.Message
            $isRoleError = $errorMessage -like "*Authentication_RequestFromUnsupportedUserRole*" -or
            $errorMessage -like "*User is not in the allowed roles*" -or
            $errorMessage -like "*Forbidden*" -or
            $_.Exception.Response.StatusCode -eq 403

            if ($isRoleError) {
                Write-AuditLog "SignInActivity access denied. Proceeding without sign-in data." -Level "WARNING"
            } else {
                Write-AuditLog "Unexpected error testing SignInActivity access: $($_.Exception.Message)" -Level "WARNING"
            }
            $canAccessSignInActivity = $false
        }

        # Determine properties to request based on SignInActivity availability
        $userProperties = if ($canAccessSignInActivity) {
            "Id,DisplayName,UserPrincipalName,CreatedDateTime,SignInActivity,Department,JobTitle,OfficeLocation"
        } else {
            "Id,DisplayName,UserPrincipalName,CreatedDateTime,Department,JobTitle,OfficeLocation"
        }

        # Process users in batches
        $batchSize = 15  # Graph API filter limit for OR clauses
        $totalBatches = [Math]::Ceiling($userIds.Count / $batchSize)

        $users = @()
        for ($i = 0; $i -lt $userIds.Count; $i += $batchSize) {
            $currentBatch = [Math]::Floor($i / $batchSize) + 1
            Write-Progress -Activity "Retrieving User Details" -Status "Processing batch $currentBatch of $totalBatches" -PercentComplete (($currentBatch / $totalBatches) * 100)

            $batchIds = $userIds[$i..([Math]::Min($i + $batchSize - 1, $userIds.Count - 1))]
            $idFilter = ($batchIds | ForEach-Object { "id eq '$_'" }) -join ' or '

            try {
                # Use the appropriate properties based on the earlier test
                $batchUsers = Get-MgUser -Filter $idFilter -Property $userProperties -All -ErrorAction Stop
                $users += $batchUsers
            } catch {
                Write-AuditLog "Error retrieving batch ${currentBatch}: $($_.Exception.Message)" -Level "ERROR"
            }
        }

        Write-Progress -Activity "Retrieving User Details" -Completed
        Write-AuditLog "Found $($users.Count) users in '$ActiveUsersGroup' group." -Level "SUCCESS"
        Write-AuditLog "NOTE: Cutoff is not applied for users to ensure visibility on all active user accounts." -Level "INFO"

        $userResults = @()
        $counter = 0

        foreach ($user in $users) {
            $counter++
            if ($counter % 50 -eq 0) {
                Write-Progress -Activity "Processing Users" -Status "Processing user $counter of $($users.Count)" -PercentComplete (($counter / $users.Count) * 100)
            }

            $lastSignIn = $null
            if ($user.SignInActivity -and $user.SignInActivity.LastSignInDateTime) {
                $lastSignIn = $user.SignInActivity.LastSignInDateTime
            }

            $userResults += [PSCustomObject]@{
                UserId             = $user.Id
                DisplayName        = $user.DisplayName
                UserPrincipalName  = $user.UserPrincipalName
                CreatedDateTime    = $user.CreatedDateTime
                LastSignInDateTime = $lastSignIn
                Department         = $user.Department
                JobTitle           = $user.JobTitle
                OfficeLocation     = $user.OfficeLocation
                ExportDate         = Get-Date
            }
        }

        Write-Progress -Activity "Processing Users" -Completed
        return $userResults | Sort-Object UserPrincipalName
    } catch {
        Write-AuditLog "Error retrieving users: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

function Get-EntraDevices {
    Write-AuditLog "Retrieving Entra devices active since $($DeviceActivityCutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))..."

    try {
        # Get all devices with comprehensive properties
        $allDevices = Get-MgDevice -All -Property "Id,DisplayName,DeviceId,OperatingSystem,OperatingSystemVersion,TrustType,IsCompliant,IsManaged,ApproximateLastSignInDateTime,RegistrationDateTime,DeviceOwnership,Model,Manufacturer"

        Write-AuditLog "Retrieved $($allDevices.Count) total Entra devices." -Level "INFO"

        # Filter for devices that have been active since cutoff date
        $activeDevices = $allDevices | Where-Object {
            # Filter out Hybrid Azure AD Joined devices (servers/VMs)
            $_.TrustType -ne "ServerAd"
        } | Where-Object {
            if ($_.ApproximateLastSignInDateTime) {
                $_.ApproximateLastSignInDateTime -ge $DeviceActivityCutoffDate
            } else {
                # Include devices without sign-in data for manual review
                $true
            }
        }

        Write-AuditLog "Found $($activeDevices.Count) devices in Entra active since $($DeviceActivityCutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))." -Level "SUCCESS"

        $deviceResults = @()
        $counter = 0

        foreach ($device in $activeDevices) {
            $counter++
            if ($counter % 25 -eq 0) {
                Write-Progress -Activity "Processing Entra Devices" -Status "Processing device $counter of $($activeDevices.Count)" -PercentComplete (($counter / $activeDevices.Count) * 100)
            }

            # Get device owner from Entra using Get-MgDeviceRegisteredOwner
            $primaryUserUPN = $null
            $primaryUserName = $null
            try {
                $owners = Get-MgDeviceRegisteredOwner -DeviceId $device.Id -ErrorAction SilentlyContinue
                if ($owners) {
                    foreach ($owner in $owners) {
                        # Only process user owners
                        if ($owner.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                            try {
                                $user = Get-MgUser -UserId $owner.Id -Property "DisplayName,UserPrincipalName" -ErrorAction SilentlyContinue
                                if ($user) {
                                    $primaryUserUPN = $user.UserPrincipalName
                                    $primaryUserName = $user.DisplayName
                                    break # Use first user owner found
                                }
                            } catch {
                                # Continue to next owner if this one fails
                            }
                        }
                    }
                }
            } catch {
                # Device owner lookup failed, continue without user association
            }

            $deviceResults += [PSCustomObject]@{
                # Device IDs
                DeviceId                 = $device.Id
                DeviceObjectId           = $device.DeviceId
                # Device Name
                DeviceName               = $device.DisplayName
                # User Info
                PrimaryUserName          = $primaryUserName
                PrimaryUserUPN           = $primaryUserUPN
                # Operating System
                OperatingSystem          = $device.OperatingSystem
                OSVersion                = $device.OperatingSystemVersion
                # Hardware Info
                Manufacturer             = $device.Manufacturer
                Model                    = $device.Model
                # Device Status/Config
                TrustType                = $device.TrustType
                JoinStatus               = switch ($device.TrustType) {
                    "AzureAd" { "Azure AD Joined" }
                    "ServerAd" { "Hybrid Azure AD Joined" }
                    "Workplace" { "Azure AD Registered" }
                    default { if ($device.TrustType) { $device.TrustType } else { "UNKNOWN" } }
                }
                IsCompliant              = $device.IsCompliant
                IsManaged                = $device.IsManaged
                DeviceOwnership          = $device.DeviceOwnership
                # Activity Dates
                RegistrationDateTime     = $device.RegistrationDateTime
                LastSignInDateTime       = $device.ApproximateLastSignInDateTime
                # Report Metadata
                ActiveSinceCutoff        = if ($device.ApproximateLastSignInDateTime) {
                    $device.ApproximateLastSignInDateTime -ge $DeviceActivityCutoffDate
                } else {
                    "UNKNOWN"
                }
                ExportDate               = Get-Date
                DeviceActivityCutoffDate = $DeviceActivityCutoffDate
            }
        }

        Write-Progress -Activity "Processing Entra Devices" -Completed
        return $deviceResults | Sort-Object PrimaryUserUPN, DeviceName
    } catch {
        Write-AuditLog "Error retrieving Entra devices: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

function Get-IntuneDevices {
    Write-AuditLog "Retrieving Intune devices active since $($DeviceActivityCutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))..."

    try {
        # Get all managed devices
        $allDevices = Get-MgDeviceManagementManagedDevice -All

        Write-AuditLog "Retrieved $($allDevices.Count) total Intune managed devices." -Level "INFO"

        # Filter for devices that have been active since cutoff date
        $activeDevices = $allDevices | Where-Object {
            if ($_.LastSyncDateTime) {
                $_.LastSyncDateTime -ge $DeviceActivityCutoffDate
            } else {
                # Include devices without sync data for manual review
                $true
            }
        }

        Write-AuditLog "Found $($activeDevices.Count) devices in Intune active since $($DeviceActivityCutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))." -Level "SUCCESS"

        $deviceResults = @()
        $counter = 0

        foreach ($device in $activeDevices) {
            $counter++
            if ($counter % 25 -eq 0) {
                Write-Progress -Activity "Processing Intune Devices" -Status "Processing device $counter of $($activeDevices.Count)" -PercentComplete (($counter / $activeDevices.Count) * 100)
            }

            # Get associated user information
            $associatedUser = $null
            if ($device.UserPrincipalName) {
                try {
                    $associatedUser = Get-MgUser -Filter "userPrincipalName eq '$($device.UserPrincipalName)'" -Property "DisplayName,UserPrincipalName" -ErrorAction SilentlyContinue
                } catch {
                    Write-AuditLog "Warning: Could not retrieve user info for $($device.UserPrincipalName)" -Level "WARNING"
                }
            }

            $deviceResults += [PSCustomObject]@{
                # Device IDs
                DeviceId                 = $device.Id
                AzureADDeviceId          = $device.AzureADDeviceId
                # Device Name
                DeviceName               = $device.DeviceName
                # User Info
                UserDisplayName          = if ($associatedUser) { $associatedUser.DisplayName } else { $device.UserDisplayName }
                UserPrincipalName        = $device.UserPrincipalName
                # Operating System
                OperatingSystem          = $device.OperatingSystem
                OSVersion                = $device.OSVersion
                # Hardware Info
                Manufacturer             = $device.Manufacturer
                Model                    = $device.Model
                SerialNumber             = $device.SerialNumber
                # Device Status/Config
                EnrollmentType           = $device.EnrollmentType
                ComplianceState          = $device.ComplianceState
                DeviceOwnership          = $device.ManagedDeviceOwnerType
                IsSupervised             = $device.IsSupervised
                IsEncrypted              = $device.IsEncrypted
                # Activity Dates
                LastSyncDateTime         = $device.LastSyncDateTime
                EnrolledDateTime         = $device.EnrolledDateTime
                # Report Metadata
                ActiveSinceCutoff        = if ($device.LastSyncDateTime) {
                    $device.LastSyncDateTime -ge $DeviceActivityCutoffDate
                } else {
                    "Unknown"
                }
                ExportDate               = Get-Date
                DeviceActivityCutoffDate = $DeviceActivityCutoffDate
            }
        }

        Write-Progress -Activity "Processing Intune Devices" -Completed
        return $deviceResults | Sort-Object UserPrincipalName, DeviceName
    } catch {
        Write-AuditLog "Error retrieving Intune devices: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

function Compare-EntraIntuneDevices {
    param(
        [array]$EntraDevices,
        [array]$IntuneDevices
    )

    Write-AuditLog "Comparing Entra and Intune device inventories..." -Level "INFO"

    # Create lookup tables for efficient matching
    $entraLookup = @{}
    $intuneLookup = @{}

    # Build Entra lookup by DeviceObjectId (maps to AzureADDeviceId in Intune)
    foreach ($device in $EntraDevices) {
        if ($device.DeviceObjectId) {
            $entraLookup[$device.DeviceObjectId.ToLower()] = $device
        }
    }

    # Build Intune lookup by AzureADDeviceId
    foreach ($device in $IntuneDevices) {
        if ($device.AzureADDeviceId) {
            $intuneLookup[$device.AzureADDeviceId.ToLower()] = $device
        }
    }

    # Find matches and mismatches
    $matchedDevices = @()
    $entraOnlyDevices = @()
    $intuneOnlyDevices = @()

    # Check Entra devices against Intune
    foreach ($entraDevice in $EntraDevices) {
        if ($entraDevice.DeviceObjectId) {
            $deviceId = $entraDevice.DeviceObjectId.ToLower()
            if ($intuneLookup.ContainsKey($deviceId)) {
                # Device exists in both systems
                $intuneDevice = $intuneLookup[$deviceId]
                $matchedDevices += [PSCustomObject]@{
                    # Device IDs
                    DeviceObjectId   = $entraDevice.DeviceObjectId
                    # Device Name
                    DeviceName       = $entraDevice.DeviceName
                    # User Info (Intune preferred as requested)
                    IntuneUserName   = $intuneDevice.UserDisplayName
                    IntuneUser       = $intuneDevice.UserPrincipalName
                    EntraUserName    = $entraDevice.PrimaryUserName
                    EntraUser        = $entraDevice.PrimaryUserUPN
                    UserMismatch     = if ($entraDevice.PrimaryUserUPN -and $intuneDevice.UserPrincipalName) {
                        $entraDevice.PrimaryUserUPN.ToLower() -ne $intuneDevice.UserPrincipalName.ToLower()
                    } else { $false }
                    # Operating System
                    OperatingSystem  = $entraDevice.OperatingSystem
                    # Device Status/Config
                    EntraJoinStatus  = $entraDevice.JoinStatus
                    IntuneCompliance = $intuneDevice.ComplianceState
                    IntuneOwnership  = $intuneDevice.DeviceOwnership
                    # Activity Dates
                    EntraLastSignIn  = $entraDevice.LastSignInDateTime
                    IntuneLastSync   = $intuneDevice.LastSyncDateTime
                    # Report Metadata
                    Status           = "Matched"
                }
            } else {
                # Device only in Entra
                $entraOnlyDevices += [PSCustomObject]@{
                    # Device IDs
                    DeviceObjectId     = $entraDevice.DeviceObjectId
                    # Device Name
                    DeviceName         = $entraDevice.DeviceName
                    # User Info
                    PrimaryUserName    = $entraDevice.PrimaryUserName
                    PrimaryUser        = $entraDevice.PrimaryUserUPN
                    # Operating System
                    OperatingSystem    = $entraDevice.OperatingSystem
                    # Hardware Info
                    Manufacturer       = $entraDevice.Manufacturer
                    Model              = $entraDevice.Model
                    # Device Status/Config
                    JoinStatus         = $entraDevice.JoinStatus
                    # Activity Dates
                    LastSignInDateTime = $entraDevice.LastSignInDateTime
                }
            }
        }
    }

    # Check for Intune-only devices
    foreach ($intuneDevice in $IntuneDevices) {
        if ($intuneDevice.AzureADDeviceId) {
            $deviceId = $intuneDevice.AzureADDeviceId.ToLower()
            if (-not $entraLookup.ContainsKey($deviceId)) {
                # Device only in Intune
                $intuneOnlyDevices += [PSCustomObject]@{
                    # Device IDs
                    AzureADDeviceId   = $intuneDevice.AzureADDeviceId
                    # Device Name
                    DeviceName        = $intuneDevice.DeviceName
                    # User Info
                    UserDisplayName   = $intuneDevice.UserDisplayName
                    UserPrincipalName = $intuneDevice.UserPrincipalName
                    # Operating System
                    OperatingSystem   = $intuneDevice.OperatingSystem
                    # Hardware Info
                    Manufacturer      = $intuneDevice.Manufacturer
                    Model             = $intuneDevice.Model
                    SerialNumber      = $intuneDevice.SerialNumber
                    # Device Status/Config
                    ComplianceState   = $intuneDevice.ComplianceState
                    DeviceOwnership   = $intuneDevice.DeviceOwnership
                    # Activity Dates
                    LastSyncDateTime  = $intuneDevice.LastSyncDateTime
                }
            }
        }
    }

    Write-AuditLog "Device comparison completed: $($matchedDevices.Count) matched, $($entraOnlyDevices.Count) Entra-only, $($intuneOnlyDevices.Count) Intune-only" -Level "SUCCESS"

    return @{
        MatchedDevices    = $matchedDevices | Sort-Object IntuneUser, DeviceName
        EntraOnlyDevices  = $entraOnlyDevices | Sort-Object PrimaryUser, DeviceName
        IntuneOnlyDevices = $intuneOnlyDevices | Sort-Object UserPrincipalName, DeviceName
    }
}

function Export-Results {
    param(
        [array]$Users,
        [array]$EntraDevices,
        [array]$IntuneDevices,
        [object]$DeviceComparison
    )

    try {
        # Check if ImportExcel module is available, install if needed
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-AuditLog "Installing ImportExcel module..." -Level "INFO"
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
        }
        Import-Module ImportExcel -Force

        # Determine where to save the file based on OutputPath parameter
        if ($isDirectory) {
            $excelPath = Join-Path $outputDirectory $outputFilename
            Write-AuditLog "Creating timestamped file at: $excelPath" -Level "INFO"
        } else {
            $excelPath = $OutputPath
            Write-AuditLog "Creating file at: $excelPath" -Level "INFO"
        }

        # Export all data to different worksheets
        if ($Users.Count -gt 0) {
            $Users | Export-Excel -Path $excelPath -WorksheetName "Active Users" -TableName "Active Users Table" -TableStyle Medium2 -FreezeTopRow -AutoSize -ClearSheet
        }

        if ($EntraDevices.Count -gt 0) {
            $EntraDevices | Export-Excel -Path $excelPath -WorksheetName "Entra Devices" -TableName "Entra Devices Table" -TableStyle Medium2 -FreezeTopRow -AutoSize -ClearSheet
        }

        if ($IntuneDevices.Count -gt 0) {
            $IntuneDevices | Export-Excel -Path $excelPath -WorksheetName "Intune Devices" -TableName "Intune Devices Table" -TableStyle Medium2 -FreezeTopRow -AutoSize -ClearSheet
        }

        # Add device comparison worksheets
        if ($DeviceComparison) {
            if ($DeviceComparison.MatchedDevices.Count -gt 0) {
                $DeviceComparison.MatchedDevices | Export-Excel -Path $excelPath -WorksheetName "Matched Devices" -TableName "Matched Devices Table" -TableStyle Medium2 -FreezeTopRow -AutoSize -ClearSheet
            }
            if ($DeviceComparison.EntraOnlyDevices.Count -gt 0) {
                $DeviceComparison.EntraOnlyDevices | Export-Excel -Path $excelPath -WorksheetName "Entra-Only Devices" -TableName "Entra-Only Devices Table" -TableStyle Medium2 -FreezeTopRow -AutoSize -ClearSheet
            }
            if ($DeviceComparison.IntuneOnlyDevices.Count -gt 0) {
                $DeviceComparison.IntuneOnlyDevices | Export-Excel -Path $excelPath -WorksheetName "Intune-Only Devices" -TableName "Intune-Only Devices Table" -TableStyle Medium2 -FreezeTopRow -AutoSize -ClearSheet
            }
        }

        Write-AuditLog "Exported device audit to: $excelPath" -Level "SUCCESS"
        Write-AuditLog "Workbook contains: Users ($($Users.Count)), Entra ($($EntraDevices.Count)), Intune ($($IntuneDevices.Count)), Matched ($($DeviceComparison.MatchedDevices.Count)), Entra-Only ($($DeviceComparison.EntraOnlyDevices.Count)), Intune-Only ($($DeviceComparison.IntuneOnlyDevices.Count))" -Level "INFO"
    } catch {
        Write-AuditLog "Error during Excel export: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Show-AuditSummary {
    param(
        [array]$Users,
        [array]$EntraDevices,
        [array]$IntuneDevices,
        [object]$DeviceComparison
    )

    Write-Host "`n" -NoNewline
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host "DEVICE AUDIT SUMMARY" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor Cyan

    Write-Host "`nACTIVE USERS:" -ForegroundColor Yellow
    Write-Host "  Total Active Users: $($Users.Count)" -ForegroundColor Green

    Write-Host "`nENTRA DEVICES (active since cutoff):" -ForegroundColor Yellow
    Write-Host "  Total Active Devices: $($EntraDevices.Count)" -ForegroundColor Green

    if ($EntraDevices.Count -gt 0) {
        $joinTypes = $EntraDevices | Group-Object JoinStatus | Sort-Object Count -Descending
        foreach ($joinType in $joinTypes) {
            Write-Host "    $($joinType.Name): $($joinType.Count)" -ForegroundColor White
        }

        $withUsers = ($EntraDevices | Where-Object { $_.PrimaryUserUPN }).Count
        $withoutUsers = $EntraDevices.Count - $withUsers
        Write-Host "    With User Association: $withUsers" -ForegroundColor White
        Write-Host "    Without User Association: $withoutUsers" -ForegroundColor $(if ($withoutUsers -gt 0) { "Yellow" } else { "White" })
    }

    Write-Host "`nINTUNE DEVICES (active since cutoff):" -ForegroundColor Yellow
    Write-Host "  Total Active Devices: $($IntuneDevices.Count)" -ForegroundColor Green

    if ($IntuneDevices.Count -gt 0) {
        $osTypes = $IntuneDevices | Group-Object OperatingSystem | Sort-Object Count -Descending
        foreach ($osType in $osTypes) {
            Write-Host "    $($osType.Name): $($osType.Count)" -ForegroundColor White
        }

        $compliant = ($IntuneDevices | Where-Object { $_.ComplianceState -eq "compliant" }).Count
        $nonCompliant = ($IntuneDevices | Where-Object { $_.ComplianceState -eq "noncompliant" }).Count
        $unknown = $IntuneDevices.Count - $compliant - $nonCompliant

        Write-Host "    Compliant: $compliant" -ForegroundColor Green
        Write-Host "    Non-Compliant: $nonCompliant" -ForegroundColor $(if ($nonCompliant -gt 0) { "Red" } else { "White" })
        if ($unknown -gt 0) {
            Write-Host "    Unknown Compliance: $unknown" -ForegroundColor Yellow
        }

        $corporate = ($IntuneDevices | Where-Object { $_.DeviceOwnership -eq "company" }).Count
        $personal = ($IntuneDevices | Where-Object { $_.DeviceOwnership -eq "personal" }).Count
        Write-Host "    Corporate Owned: $corporate" -ForegroundColor White
        Write-Host "    Personal Owned: $personal" -ForegroundColor $(if ($personal -gt 0) { "Yellow" } else { "White" })
    }

    if ($DeviceComparison) {
        Write-Host "`nDEVICE COMPARISON RESULTS:" -ForegroundColor Yellow
        Write-Host "  Matched Devices: $($DeviceComparison.MatchedDevices.Count)" -ForegroundColor Green
        Write-Host "  Entra-Only Devices: $($DeviceComparison.EntraOnlyDevices.Count)" -ForegroundColor Yellow
        Write-Host "  Intune-Only Devices: $($DeviceComparison.IntuneOnlyDevices.Count)" -ForegroundColor Yellow
    }

    Write-Host "`n" -NoNewline
}

# Main execution
try {
    Write-AuditLog "Starting device audit for devices active since $($DeviceActivityCutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))." -Level "INFO"

    # Test Graph connection
    if (-not (Test-GraphConnection)) {
        Write-AuditLog "Cannot proceed without Graph connection." -Level "ERROR"
        exit 1
    }

    $activeUsers = Get-ActiveUsers
    $entraDevices = Get-EntraDevices
    $intuneDevices = Get-IntuneDevices
    $deviceComparison = Compare-EntraIntuneDevices -EntraDevices $entraDevices -IntuneDevices $intuneDevices

    # Export results
    Export-Results -Users $activeUsers -EntraDevices $entraDevices -IntuneDevices $intuneDevices -DeviceComparison $deviceComparison

    # Show summary
    Show-AuditSummary -Users $activeUsers -EntraDevices $entraDevices -IntuneDevices $intuneDevices -DeviceComparison $deviceComparison

    Write-AuditLog "Device audit completed successfully!" -Level "SUCCESS"
} catch {
    Write-AuditLog "Critical error during audit: $($_.Exception.Message)" -Level "ERROR"
    Write-AuditLog "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    exit 1
}
