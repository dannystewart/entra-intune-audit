# Entra Intune Audit Script

A comprehensive PowerShell script for auditing active users and devices across Microsoft Entra ID and Intune. This tool provides detailed insights into your organization's device inventory, user associations, and compliance status.

## Features

- **Active User Audit**: Retrieves all members from a specified user group with detailed profile information
- **Entra Device Inventory**: Captures all Azure AD joined/registered devices active within a configurable timeframe
- **Intune Device Analysis**: Gathers comprehensive Intune managed device data including compliance and enrollment details
- **Device Comparison**: Cross-references Entra and Intune inventories to identify mismatches and gaps
- **Orphaned Device Detection**: Identifies Entra devices without user associations
- **Excel Export**: Generates timestamped Excel workbooks with multiple worksheets for easy analysis
- **Comprehensive Logging**: Detailed logging with color-coded output and persistent log files

## How It Works

The script connects to Microsoft Graph API to gather data from both Entra ID and Intune, then performs the following analysis:

1. **User Collection**: Retrieves all members from a specified Active Directory group
2. **Device Discovery**: Collects devices from both Entra ID and Intune that have been active since a cutoff date
3. **Data Correlation**: Matches devices between systems using Azure AD device IDs
4. **Gap Analysis**: Identifies devices that exist in one system but not the other
5. **Orphan Detection**: Finds Entra devices without associated users
6. **Report Generation**: Exports results to Excel with separate worksheets for each data category

## Prerequisites

- PowerShell 5.1 or later (PowerShell 7 recommended)
- Microsoft Graph PowerShell SDK (`Connect-MgGraph`)
- Appropriate Graph API permissions:
  - `User.Read.All`
  - `Device.Read.All`
  - `DeviceManagementManagedDevices.Read.All`
  - `Group.Read.All`
  - `Reports.Read.All`
  - `AuditLog.Read.All`
  - `Directory.Read.All`

## Usage

### Basic Usage

```powershell
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff"
```

### Custom Cutoff Date

```powershell
# 7 days ago
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff" -DeviceActivityCutoffDate (Get-Date).AddDays(-7)

# Specific date
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff" -DeviceActivityCutoffDate "2025-01-15 00:00:00"
```

### Custom Output Location

```powershell
# Directory (creates timestamped file)
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff" -OutputPath "C:\Audit"

# Specific file
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff" -OutputPath "C:\OneDrive\MasterAudit.xlsx"
```

## Output

The script generates an Excel workbook with the following worksheets:

- **Active Users**: Complete user profile information
- **Entra Devices**: Azure AD device inventory with user associations
- **Intune Devices**: Intune managed device details and compliance status
- **Matched Devices**: Devices found in both systems with comparison data
- **Entra-Only Devices**: Devices only present in Entra ID
- **Intune-Only Devices**: Devices only present in Intune
- **Orphaned Entra Devices**: Entra devices without user associations

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `UserGroup` | string | Yes | - | Display name of the user group to audit |
| `DeviceActivityCutoffDate` | DateTime | No | 30 days ago | Cutoff date for device activity |
| `OutputPath` | string | No | Current directory | Output file path or directory |

## Logging

The script maintains detailed logs in `MasterAudit.log` with color-coded console output:

- **INFO**: General information and progress updates
- **SUCCESS**: Successful operations
- **WARNING**: Non-critical issues or limitations
- **ERROR**: Errors that may affect functionality

## Examples

### Standard Organizational Audit

```powershell
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff"
```

### Weekly Device Activity Check

```powershell
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff" -DeviceActivityCutoffDate (Get-Date).AddDays(-7)
```

### Custom Output with Specific Date

```powershell
.\Get-MasterAudit.ps1 -UserGroup "All Org Staff" -DeviceActivityCutoffDate "2025-01-01 00:00:00" -OutputPath "C:\Reports\WeeklyAudit.xlsx"
```

## Notes

- The script automatically installs the `ImportExcel` module if not present
- Device activity cutoff defaults to 30 days due to Graph API data availability limitations
- Hybrid Azure AD Joined devices (servers/VMs) are automatically filtered out
- All timestamps are in UTC format
- The script handles Graph API rate limiting and processes data in batches

## Author and License

Written by [Danny Stewart](https://github.com/dannystewart/entra-intune-audit) and published under the MIT License. See [LICENSE](LICENSE) file for details.
