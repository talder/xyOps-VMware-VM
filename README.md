

<p align="center"><img src="https://raw.githubusercontent.com/talder/xyOps-VMware-VM/refs/heads/main/logo.png" height="108" alt="Logo"/></p>
<h1 align="center">VMware VM Operations</h1>

# xyOps VMware VM Operations Plugin

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/talder/xyOps-VMware-VM/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-7.0+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![VCF.PowerCLI](https://img.shields.io/badge/VCF.PowerCLI-14.0+-green.svg)](https://developer.vmware.com/powercli)

A comprehensive VMware vCenter management plugin for xyOps that provides VM inventory, monitoring, and snapshot lifecycle management. Built with PowerShell and [VCF PowerCLI](https://developer.vmware.com/powercli), this plugin enables automated operations across your VMware infrastructure for your virtual machines.

## Disclaimer

**USE AT YOUR OWN RISK.** This software is provided "as is", without warranty of any kind, express or implied. The author and contributors are not responsible for any damages, data loss, system downtime, or other issues that may arise from the use of this software. Always test in non-production environments before running against production systems. By using this plugin, you acknowledge that you have read, understood, and accepted this disclaimer.

## Features

### Virtual Machine Management
- **List all virtual machines** with detailed information (CPU, memory, storage, IP addresses, uptime)
- **Power state monitoring** with separate views for powered on/off VMs
- **IP address tracking** with multi-line display for VMs with multiple IPs
- **Uptime calculation** for powered-on VMs
- **Create VM snapshots** with unique identifiers for multi-VM project tracking
- **Reboot VMs** (planned)
- **Stop VMs** (planned)

### Snapshot Management
- **List all VM snapshots** across the vCenter environment with size and creation date
- **Remove snapshots before date** - Delete snapshots older than a specific date
- **Remove snapshots before number of weeks** - Age-based snapshot cleanup (e.g., remove all snapshots older than 4 weeks)
- **Remove all snapshots** - Bulk snapshot removal with safety protections
- **Remove snapshots containing specific text** - Delete snapshots by name search (e.g., find and remove all snapshots with "ProjectX" in the name)
- **Smart snapshot protection** - Exception filters prevent accidental removal of backup snapshots (VEEAM, CommVault, etc.)
- **Comprehensive reporting** - Track removed, failed, and skipped snapshots with detailed reasons

### Output & Export
- **Multiple export formats**: JSON, CSV, Markdown (MD), and HTML
- **Professional HTML reports** with styled tables and color-coded status
- **Markdown display** in xyOps GUI with separate tables for different VM/snapshot states
- **File export** option for archival and external processing
- **Detailed captions** showing operation results at a glance

### Infrastructure
- **Configurable PowerCLI settings** (certificate validation, proxy, server mode)
- **Auto-installs VCF.PowerCLI module** if missing (with publisher verification skip)
- **Optimized module loading** - Imports only `VMware.VimAutomation.Core` at runtime for faster startup (uses `VCF.PowerCLI` only for install/update checks); skips redundant imports if already loaded (~5-10s savings)
- **Secure credential management** via xyOps secret vault
- **Error handling** with detailed error codes and messages
- **Performance-focused** - Update checks disabled by default for fast execution

## Requirements

### CLI Requirements

- **PowerShell Core (pwsh)** - Version 7.0 or later recommended
  - On macOS: `brew install powershell`
  - On Linux: [Install instructions](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux)
  - On Windows: Comes pre-installed or [download](https://aka.ms/powershell-release)

### Module Requirements

- **VCF.PowerCLI** - Automatically installed by the plugin if not present
  - The plugin will attempt to install PowerCLI using `Install-Module -Name VCF.PowerCLI -Scope CurrentUser -SkipPublisherCheck`
  - Requires internet connection for first-time installation
  - Manual installation: `Install-Module -Name VCF.PowerCLI -Force -AllowClobber -Scope CurrentUser -SkipPublisherCheck`
- **Performance**: At runtime only `VMware.VimAutomation.Core` is imported for speed; umbrella `VCF.PowerCLI` is used solely for installation and optional update checks. Import is skipped if already loaded (~5-10s savings)
  - **Note**: VCF.PowerCLI is the current module (VMware.PowerCLI is deprecated)

#### ⚠️ VCF.PowerCLI Update Check Warning

**IMPORTANT**: The "Check for VCF.PowerCLI update" parameter is **disabled by default** for performance reasons.

**Why it's disabled**:
- Checking for updates queries PowerShell Gallery over the internet on **every plugin execution**
- This adds **~10-15 seconds** to every run, even when no update is available
- The delay impacts all actions (List VMs, snapshots, etc.)

**When to enable**:
- ✅ **Recommended**: Create a **separate scheduled job** (e.g., monthly) that runs "List VMs" with update check enabled
- ✅ Enable manually once in a while to check for updates
- ❌ **NOT recommended**: Enable for production automation or frequent operations

**Best Practice**:
```yaml
# Example: Monthly update check job
Job: "VCF PowerCLI Update Check"
Schedule: "0 0 1 * *"  # First day of month at midnight
Action: "List VMs"
Parameters:
  - Check for VCF.PowerCLI update: ✓ (enabled)
  - All other normal settings
```

This approach keeps your regular operations fast while ensuring the module stays updated periodically.

### Secret Vault Configuration

**IMPORTANT**: This plugin requires vCenter credentials to be stored in an xyOps secret vault for secure authentication.

#### Setting Up the Secret Vault

1. **Create a Secret Vault** in xyOps (e.g., named `VMWARE-PLUGIN`)
2. **Add the following keys** to the vault:
   - `VMWARE_USERNAME` - Your vCenter username (e.g., `administrator@vsphere.local`)
   - `VMWARE_PASSWORD` - Your vCenter password

3. **Attach the vault** to the plugin when configuring it

The plugin will automatically read credentials from these environment variables at runtime.

**Note**: Username and password are NOT passed as plugin parameters. All authentication is handled securely through environment variables populated from the secret vault.

For detailed instructions on creating and managing secret vaults, see the [xyOps Secrets Documentation](https://github.com/pixlcore/xyops/blob/main/docs/secrets.md).

### vCenter Permissions

The vCenter user account must have sufficient permissions to perform the requested operations:
- **List VMs**: Read-only access to Virtual Machines
- **List Snapshots**: Read access to VM snapshots
- **Remove Snapshots**: Snapshot delete permissions on Virtual Machines
- **Reboot/Stop VMs**: Power operations on Virtual Machines (planned feature)
- **Create Snapshots**: Snapshot create permissions (planned feature)

**Recommended**: Use a vCenter user with at least the **"Virtual Machine Power User"** role.

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| **vCenter server** | text | Yes | - | vCenter server hostname or IP address (e.g., `vcenter.company.com`) |
| **Ignore certificates** | checkbox | No | `true` | Ignore SSL certificate validation warnings (useful for self-signed certificates) |
| **No proxy** | checkbox | No | `true` | Do not use system proxy settings (if unchecked, uses system proxy) |
| **Single VIserver mode** | checkbox | No | `true` | Connect to single vCenter only (if unchecked, allows multiple simultaneous connections) |
| **Selected action** | dropdown | Yes | - | Choose the operation to perform (see [Available Actions](#available-actions)) |
| **Date input** | text | Conditional* | - | Required for "Remove snapshots before date" action. Accepts multiple formats: `yyyy-MM-dd`, `yyyy-MM-dd HH:mm:ss`, `MM/dd/yyyy`, `dd/MM/yyyy`, `yyyy/MM/dd` |
| **Number of weeks** | number | Conditional** | - | Required for "Remove snapshots before number of weeks" action. Positive integer (e.g., `4` removes snapshots older than 4 weeks) |
| **Snapshot removal exception** | text | No | `VEEAM` | Comma-separated list of keywords to match in snapshot names/descriptions that should NOT be removed. Case-insensitive. Use `VEEAM,BACKUP,PROD` for multiple keywords. Leave empty to disable protection (⚠️ use with caution!) |
| **VM name** | text | Conditional*** | - | Required for "Create VM snapshot" action. Comma-separated list of VM names (e.g., `VM1,VM2,VM3`). All VMs receive the same snapshot name with unique identifier |
| **Snapshot name** | text | Conditional****/No | Auto-generated | Required for "Remove snapshots containing specific text" action (search text for snapshot names). For "Create VM snapshot" action: Optional - if empty, auto-generates `xyOps-VMware-VM-{timestamp}`. Unique identifiers are appended if selected |
| **Snapshot description** | textarea | No | Auto-generated | Optional snapshot description. If empty, auto-generates based on unique identifier selection. Custom descriptions override auto-generation |
| **Include VM memory in snapshot** | checkbox | No | `true` | Include VM memory state in snapshot (allows reverting to exact running state). Set to false for quicker snapshots without memory |
| **Use unique snapshot identifier** | dropdown | No | `No` | Append unique identifier to snapshot name: `No`, `Unique short UID` (8-char GUID), or `Timestamp` (yyyyMMddHHmmss) |
| **Check for VCF.PowerCLI update** | checkbox | No | `false` | **⚠️ WARNING**: Checks PowerShell Gallery for VCF.PowerCLI updates on **every run** (~10-15s delay). Leave **unchecked** for normal operations. Enable occasionally (e.g., monthly) or use a separate scheduled job with "List VMs" action |
| **Export format** | dropdown | No | `JSON` | Output format: `JSON`, `CSV`, `MD` (Markdown), or `HTML` |
| **Export to file** | checkbox | No | `false` | Also export the output to a timestamped file in addition to job output |
| **Enable debug mode** | checkbox | No | `false` | Write the JSON parameters to the job output for troubleshooting |

\* Required only when "Remove snapshots before date" action is selected  
\** Required only when "Remove snapshots before number of weeks" action is selected  
\*** Required only when "Create VM snapshot" action is selected  
\**** Required only when "Remove snapshots containing specific text" action is selected

**Authentication**: Credentials (`VMWARE_USERNAME` and `VMWARE_PASSWORD`) must be configured in a secret vault attached to this plugin. See [Secret Vault Configuration](#secret-vault-configuration) above.

## Available Actions

### List VMs
Retrieves a complete list of all virtual machines in the vCenter environment with detailed information.

**Output Information**:
- VM Name
- Power State (PoweredOn, PoweredOff, Suspended)
- IP Addresses (multi-line display for VMs with multiple IPs)
- Uptime (calculated for powered-on VMs)
- Number of CPUs
- Memory (GB)
- Used Space (GB)
- Provisioned Space (GB)
- Guest Operating System
- ESXi Host
- Folder
- Resource Pool
- Notes/Description

**Display**: Separate tables for Powered On and Powered Off VMs with color-coded backgrounds (light red for powered-off VMs in HTML export).

### List snapshots
Retrieves a complete list of all VM snapshots across the vCenter environment.

**Output Information**:
- VM Name
- Snapshot Name
- Created Date/Time
- Size (GB)
- Power State at snapshot time
- Description

**Use Cases**:
- Identify orphaned or forgotten snapshots
- Monitor snapshot storage consumption
- Compliance reporting for snapshot retention policies
- Pre-cleanup auditing before removal operations

### Remove snapshots before date
Removes all snapshots created before a specified date, with protection for critical snapshots.

**Required Parameter**: `dateinput` - Date in one of these formats:
- `2024-01-15` (yyyy-MM-dd)
- `2024-01-15 14:30:00` (yyyy-MM-dd HH:mm:ss)
- `01/15/2024` (MM/dd/yyyy)
- `15/01/2024` (dd/MM/yyyy)
- `2024/01/15` (yyyy/MM/dd)

**Protected Snapshots**: Snapshots matching exception criteria (default: "VEEAM") will be automatically skipped and reported separately.

**Output Sections**:
- **Successfully Removed Snapshots**: List of deleted snapshots with size information
- **Skipped Snapshots**: Protected snapshots that were not removed (yellow background in HTML)
- **Failed Removals**: Snapshots that encountered errors during deletion (red background in HTML)

**Use Cases**:
- Clean up snapshots before a specific maintenance window
- Remove snapshots older than a compliance date
- Targeted cleanup after project completion

### Remove snapshots before number of weeks
Removes all snapshots older than the specified number of weeks, automatically calculating the target date.

**Required Parameter**: `numberofweeks` - Positive integer (e.g., `4` for snapshots older than 4 weeks)

**Protected Snapshots**: Snapshots matching exception criteria (default: "VEEAM") will be automatically skipped.

**Age Calculation**: The plugin calculates the target date as `current_date - (weeks × 7)` and removes all snapshots created before that date.

**Output Sections**:
- Target date (calculated automatically)
- Successfully Removed Snapshots (with age in days)
- Skipped Snapshots
- Failed Removals

**Use Cases**:
- Implement automated snapshot retention policies (e.g., "keep snapshots for 4 weeks")
- Regular cleanup schedules via xyOps workflows
- Storage reclamation on aging snapshots

### Remove all snapshots
Removes ALL snapshots from all VMs in the vCenter environment.

**⚠️ WARNING**: This is a destructive operation! Use with extreme caution.

**Protected Snapshots**: Snapshots matching exception criteria (default: "VEEAM") will be automatically skipped and reported.

**Output Sections**:
- Successfully Removed Snapshots (with age in days)
- Skipped Snapshots (protected by exception filter)
- Failed Removals

**Use Cases**:
- Complete snapshot cleanup during infrastructure migrations
- Post-maintenance cleanup after testing
- Decommissioning environments while protecting production backups

**Best Practice**: Always run "List snapshots" first to audit what will be removed, then review the exception filter before executing.

### Remove snapshots containing specific text
Removes all snapshots whose names contain the specified search text, with protection for critical snapshots.

**Required Parameter**: `snapshotname` - Search text to match in snapshot names (case-insensitive, partial match)

**Search Behavior**: The action uses a **case-insensitive "contains" search** on snapshot names. For example:
- Search text `ProjectX` matches snapshots: `ProjectX-PreDeployment-a3f9b2c8`, `ProjectX-Test`, `my-ProjectX-snapshot`
- Search text `test` matches: `Test-Snapshot`, `My-Test-VM`, `TESTING-123`

**Protected Snapshots**: Snapshots matching exception criteria (default: "VEEAM") will be automatically skipped and reported.

**Output Sections**:
- Search text used
- Successfully Removed Snapshots (with age in days)
- Skipped Snapshots (protected by exception filter)
- Failed Removals

**Use Cases**:
- **Project Cleanup**: Remove all snapshots from a completed project (e.g., search for "ProjectX" to remove all ProjectX-related snapshots)
- **Unique Identifier Tracking**: Find and remove all snapshots with the same unique identifier created during multi-VM snapshot operations
- **Testing Cleanup**: Remove all test snapshots (e.g., search for "test" or "staging")
- **Department Cleanup**: Remove snapshots by department tag or naming convention
- **Coordinated Rollback**: Remove synchronized snapshots across multiple VMs created with the same identifier

**Example Workflow**:
1. Create multi-VM snapshot with unique ID: VMs get `ProjectX-PreDeployment-a3f9b2c8`
2. Test changes across all VMs
3. To clean up: Use "Remove snapshots containing specific text" with search `a3f9b2c8`
4. All related snapshots across all VMs are removed in one operation

**Best Practice**: Use "List snapshots" first to preview what will be matched before removal.

### Create VM snapshot
Creates snapshots for one or more virtual machines with optional unique identifiers for project tracking.

**Required Parameter**: `vmname` - Comma-separated list of VM names (e.g., `VM1,VM2,VM3`)

**Optional Parameters**:
- `snapshotname` - Custom snapshot name (default: auto-generated `xyOps-VMware-VM-{timestamp}`)
- `snapshotdescription` - Custom description (default: auto-generated based on unique identifier setting)
- `snapshotmemory` - Include VM memory state (default: `true`)
- `snapshotunique` - Append unique identifier: `No`, `Unique short UID` (8-char GUID), or `Timestamp`

**Snapshot Naming Logic**:
1. **Base Name**: If `snapshotname` is empty, generates `xyOps-VMware-VM-{timestamp}`. Otherwise uses provided name.
2. **Unique Identifier**: If `snapshotunique` is set:
   - `Unique short UID`: Appends `-{8-char-GUID}` (e.g., `MySnapshot-a3f9b2c8`)
   - `Timestamp`: Appends `-{yyyyMMddHHmmss}` (e.g., `MySnapshot-20260202173045`)
   - `No`: No identifier appended
3. **Final Name**: All specified VMs receive the **same snapshot name** (including the same unique identifier)

**Multi-VM Project Tracking**:
When creating snapshots for multiple VMs with the same unique identifier, you can easily track and manage related snapshots across a project. For example:
- **Project Snapshot**: VMs `APP01,APP02,DB01` all get snapshot `ProjectX-a3f9b2c8`
- **Easy Cleanup**: Search for snapshots containing `ProjectX-a3f9b2c8` to remove all related snapshots
- **Rollback Together**: Revert all VMs to the same checkpoint by identifying the shared unique ID

**Description Logic**:
- If `snapshotdescription` is provided: Uses custom description
- If `snapshotdescription` is empty AND unique identifier is used: "xyOps-VMware-VM added a custom snapshot with a unique identifier"
- If `snapshotdescription` is empty AND no unique identifier: "xyOps-VMware-VM snapshot"

**Output Sections**:
- **Successfully Created Snapshots**: List with VM name, snapshot name, description, created time, size, and memory indicator (✓/✗)
- **Failed Snapshot Creations**: VMs where snapshot creation failed with error messages (red background in HTML)

**Use Cases**:
- **Project Checkpoints**: Create synchronized snapshots across multi-VM applications before changes
- **Pre-Maintenance Snapshots**: Snapshot all VMs in a cluster with same unique ID for easy tracking
- **Testing Scenarios**: Create baseline snapshots for test environments with consistent naming
- **Rollback Points**: Establish coordinated rollback points across distributed systems

**Best Practices**:
1. Use unique identifiers for multi-VM projects to track related snapshots
2. Include memory for VMs that need exact state recovery (increases snapshot size)
3. Exclude memory for quick snapshots or when memory state is not critical
4. Use descriptive snapshot names for manual tracking
5. Document unique identifiers in change management tickets

### Reboot VM
*Coming soon* - Gracefully reboot a virtual machine (requires VMware Tools).

### Stop VM
*Coming soon* - Power off a virtual machine.

## Snapshot Removal Exception Filter

**IMPORTANT**: The `Snapshot removal exception` parameter protects critical snapshots from accidental deletion during removal operations.

### How It Works

1. **Dual Checking**: The plugin checks both snapshot **name** and **description** fields
2. **Case-Insensitive**: Matching ignores case (e.g., "veeam" matches "VEEAM BACKUP TEMPORARY SNAPSHOT")
3. **Partial Matching**: Any occurrence of the keyword triggers protection (e.g., "BACKUP" matches "Pre-BACKUP-Snapshot")
4. **Multiple Keywords**: Comma-separated values create multiple protection rules (OR logic)
5. **Reporting**: All skipped snapshots are reported separately in all output formats

### Default Protection

**Default Value**: `VEEAM`

This protects all Veeam backup snapshots by default, preventing removal during active backup operations:
- `VEEAM BACKUP TEMPORARY SNAPSHOT` (standard Veeam backup snapshot)
- `VEEAM REPLICATION SNAPSHOT`
- Any custom snapshot containing "VEEAM"

### Configuration Examples

| Use Case | Exception Value | Description |
|----------|----------------|-------------|
| **Veeam Backups Only** | `VEEAM` | Skip all Veeam backup snapshots (default) |
| **Multiple Backup Solutions** | `VEEAM,COMMVAULT,NBU,VERITAS` | Skip Veeam, CommVault, NetBackup, and Veritas snapshots |
| **Production Protection** | `PROD,PRODUCTION,CRITICAL` | Skip any snapshot marked as production or critical |
| **Custom Marker** | `DO_NOT_DELETE,KEEP,PERMANENT` | Skip snapshots with custom retention markers |
| **Backup + Production** | `VEEAM,BACKUP,PROD` | Comprehensive protection for backups and production |
| **No Protection** | ` ` (empty/space) | **⚠️ DANGER**: Remove ALL snapshots without exceptions |

### Output Reporting

Skipped snapshots are reported in all export formats with full transparency:

**Markdown Display** (xyOps GUI):
```markdown
## Skipped Snapshots (5)

| VM Name | Snapshot Name | Created | Size (GB) | Reason |
|---------|---------------|---------|-----------|--------|
| **VGRADIUS01** | VEEAM BACKUP TEMPORARY SNAPSHOT | 2026-02-02 17:18:00 | 2.5 | Matches exception criteria |
```

**CSV Export**:
```csv
VMName,SnapshotName,Created,SizeGB,Status,Message
"VGRADIUS01","VEEAM BACKUP TEMPORARY SNAPSHOT","2026-02-02 17:18:00",2.5,"Skipped","Matches exception criteria"
```

**JSON Export**:
```json
{
  "skippedCount": 5,
  "skippedSnapshots": [
    {
      "VMName": "VGRADIUS01",
      "SnapshotName": "VEEAM BACKUP TEMPORARY SNAPSHOT",
      "Created": "2026-02-02 17:18:00",
      "SizeGB": 2.5,
      "Reason": "Matches exception criteria"
    }
  ]
}
```

**HTML Export**:
- Yellow-highlighted table (`#ffffcc` background) for easy visual identification
- Separate section for skipped snapshots
- Full details including VM name, snapshot name, and reason

### Best Practices

1. **Always Test First**: Run "List snapshots" to see what exists before removing
2. **Review Exceptions**: Verify the exception filter matches your backup solution
3. **Start Conservative**: Use broad protection initially (e.g., `BACKUP,VEEAM,PROD`)
4. **Monitor Skipped**: Review skipped snapshot reports to ensure protection is working
5. **Audit Regularly**: Check if backup snapshots are being properly protected
6. **Never Empty**: Avoid leaving the exception field empty unless absolutely certain

## PowerCLI Configuration

The plugin automatically configures PowerCLI with the following settings at runtime:

| Setting | Value | Description |
|---------|-------|-------------|
| **ParticipateInCEIP** | `false` | Disable Customer Experience Improvement Program participation |
| **InvalidCertificateAction** | `Ignore` or `Warn` | Based on "Ignore certificates" checkbox |
| **ProxyPolicy** | `NoProxy` or `UseSystemProxy` | Based on "No proxy" checkbox |
| **DefaultVIServerMode** | `Single` or `Multiple` | Based on "Single VIserver mode" checkbox |
| **Scope** | `User` | Settings persist for the user account |

These settings are applied automatically before connecting to vCenter and persist for future PowerCLI sessions.

## Output Formats

The plugin supports four professional output formats:

### JSON Format
Structured data suitable for workflows, automation, and programmatic processing.

### CSV Format
Comma-separated values suitable for Excel, database import, or spreadsheet analysis. Includes combined output with removed, skipped, and failed snapshots with status indicators.

### Markdown (MD) Format
Human-readable tables with professional formatting for documentation and reports. Includes summary sections and separate tables for different snapshot states.

### HTML Format
Professionally styled HTML reports with color-coded tables:
- **Blue theme** with professional styling
- **Yellow background** (`#ffffcc`): Skipped snapshots
- **Red background** (`#ffcccc`): Failed removals
- **Light red** (`#ffebee`): Powered-off VMs
- Hover effects and shadow effects for better readability

## File Export

When **Export to file** is enabled, the plugin creates timestamped files:

**Filename Format**: `vmware_vm_{action}_{timestamp}.{ext}`

**Examples**:
- `vmware_vm_list vms_20260202_160530.json`
- `vmware_vm_list snapshots_20260202_160530.csv`
- `vmware_vm_remove snapshots before number of weeks_20260202_160530.md`
- `vmware_vm_remove all snapshots_20260202_160530.html`

Files are created in the current working directory and attached to the job output for download.

## Performance Metrics (xyOps Pie Chart)
The plugin emits runtime performance metrics on every execution. These appear as a pie chart on the xyOps Job Details page and as a time-series in the performance history.

- Metrics emitted (seconds):
  - `module_install` – time to locate/install/update VCF.PowerCLI (if enabled)
  - `module_import` – time to import PowerCLI (`VMware.VimAutomation.Core`)
  - `config` – time spent applying PowerCLI configuration
  - `connect` – time to connect to vCenter
  - `action` – time spent executing the selected action
  - `output` – time to build and emit report/export output
  - `t` – total runtime (excluded from pie, used for history)

Example payload:
```json
{ "xy": 1, "perf": { "module_install": 0.42, "module_import": 1.98, "config": 0.61, "connect": 3.44, "action": 2.21, "output": 0.34, "t": 9.00 } }
```

Notes:
- Metrics are always emitted (no debug required).
- If you enable the optional `debug` parameter, per-phase timing lines (e.g., `PERF connect=3.44s`) are printed to the console for additional visibility.

## Error Codes

| Code | Description | Resolution |
|------|-------------|------------|
| **1** | Failed to parse input JSON | Check job parameters are properly formatted |
| **2** | Missing required parameters | Ensure vCenter server, action, and conditional parameters are provided. Verify secret vault is attached |
| **3** | Failed to install VCF.PowerCLI module | Manually install: `Install-Module -Name VCF.PowerCLI -Force -SkipPublisherCheck` |
| **4** | Failed to connect to vCenter server | Verify vCenter hostname/IP, network connectivity, and credentials in secret vault |
| **5** | Action not yet implemented | The selected action is planned but not available yet |
| **6** | Unknown action specified | Check action name is valid |
| **7** | Invalid date format | Use yyyy-MM-dd, yyyy-MM-dd HH:mm:ss, MM/dd/yyyy, dd/MM/yyyy, or yyyy/MM/dd |
| **8** | Invalid number of weeks | Provide a positive integer (e.g., 1, 4, 12) |
| **9** | No valid VM names provided | Check vmname parameter contains at least one valid VM name |

## Usage Examples

**Note**: All examples require a secret vault with `VMWARE_USERNAME` and `VMWARE_PASSWORD` to be attached to the plugin.

### Example 1: List All VMs with IP Addresses

**Parameters**:
- vCenter server: `vcenter.company.com`
- Selected action: `List VMs`
- Export format: `MD`

**Result**: Markdown tables showing powered-on VMs (with uptime and IPs) and powered-off VMs separately.

---

### Example 2: Remove Old Snapshots (4 Week Retention)

**Parameters**:
- vCenter server: `vcenter.company.com`
- Selected action: `Remove snapshots before number of weeks`
- Number of weeks: `4`
- Snapshot removal exception: `VEEAM` (default)
- Export format: `JSON`
- Export to file: ✓

**Result**: Removes all snapshots older than 4 weeks, skips VEEAM snapshots, exports detailed JSON report.

---

### Example 3: Cleanup Before Specific Date

**Parameters**:
- vCenter server: `vcenter.company.com`
- Selected action: `Remove snapshots before date`
- Date input: `2026-01-15`
- Snapshot removal exception: `VEEAM,PROD,CRITICAL`
- Export format: `CSV`

**Result**: Removes all snapshots before January 15, 2026, protecting Veeam, production, and critical snapshots.

---

### Example 4: Create Multi-VM Project Snapshot

**Parameters**:
- vCenter server: `vcenter.company.com`
- Selected action: `Create VM snapshot`
- VM name: `APP01,APP02,DB01`
- Snapshot name: `ProjectX-PreDeployment`
- Snapshot description: `Snapshot before Project X deployment on 2026-02-02`
- Include VM memory in snapshot: ✓
- Use unique snapshot identifier: `Unique short UID`
- Export format: `MD`

**Result**: Creates snapshots on all three VMs with name `ProjectX-PreDeployment-a3f9b2c8` (same 8-char UID) for easy tracking and coordinated rollback.

---

### Example 5: Remove Project Snapshots by Unique Identifier

**Parameters**:
- vCenter server: `vcenter.company.com`
- Selected action: `Remove snapshots containing specific text`
- Snapshot name: `a3f9b2c8`
- Snapshot removal exception: `VEEAM` (default)
- Export format: `HTML`

**Result**: Removes all snapshots across all VMs containing the unique identifier `a3f9b2c8` in their names (e.g., `ProjectX-PreDeployment-a3f9b2c8`), protecting Veeam snapshots, with professional HTML report.

## Use Cases

### 1. Automated Snapshot Retention Policies
Implement company-wide snapshot retention with scheduled xyOps jobs.

### 2. VM Inventory and Capacity Planning
Track VM resource usage for capacity planning and license compliance.

### 3. Pre-Maintenance Snapshot Audit
Audit all snapshots before major infrastructure changes with professional reports.

### 4. Backup Snapshot Protection
Ensure backup snapshots are never accidentally deleted during cleanup operations.

### 5. Storage Reclamation
Reclaim storage from orphaned and old snapshots with data-driven decision making.

## Troubleshooting

### PowerCLI Installation Fails
Manually install: `Install-Module -Name VCF.PowerCLI -Force -AllowClobber -Scope CurrentUser -SkipPublisherCheck`

### Connection Fails with Certificate Error
Enable "Ignore certificates" checkbox or configure SSL certificates on vCenter.

### Permission Denied
Verify vCenter user has **Virtual Machine Power User** role or higher.

### Snapshots Not Being Removed
Check `Snapshot removal exception` parameter - may be too broad.

### Missing Credential Error
Verify secret vault is attached with `VMWARE_USERNAME` and `VMWARE_PASSWORD` keys.

### Plugin Execution is Slow (~20-30 seconds)
Disable "Check for VCF.PowerCLI update" parameter (should be unchecked by default). The update check queries PowerShell Gallery over the internet and adds 10-15 seconds to every run. Instead, create a separate monthly scheduled job with this option enabled.

## Best Practices

### Snapshot Management
1. Always audit first with "List snapshots"
2. Start with longer retention periods
3. Verify exception filter catches backup snapshots
4. Schedule removal jobs during maintenance windows
5. Enable export to file for audit trail

### Security
1. Use secret vault for credentials
2. Grant minimum vCenter permissions
3. Audit plugin execution access
4. Rotate credentials periodically

### Operations
1. Document exception keywords
2. Test in development first
3. Start with small VM sets
4. Automate via xyOps workflows
5. Store CSV exports for trend analysis

### Performance
1. **Keep "Check for VCF.PowerCLI update" disabled** for regular operations (default)
2. Create a separate monthly scheduled job for module updates
3. Use "List VMs" action with update check enabled for the update job
4. Monitor execution time - update checks add ~10-15 seconds per run
5. **Module auto-optimization**: The plugin automatically skips module import if VCF.PowerCLI is already loaded in the PowerShell session (saves ~5-10s on subsequent runs in the same session)

## Links

- [VMware PowerCLI Documentation](https://developer.vmware.com/powercli)
- [VCF PowerCLI Module]([https://www.powershellgallery.com/packages/VMware.PowerCLI](https://developer.broadcom.com/powercli/installation-guide))
- [PowerShell Installation](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [vCenter Server Documentation](https://docs.vmware.com/en/VMware-vSphere/index.html)
- [xyOps Secrets Documentation](https://github.com/pixlcore/xyops/blob/main/docs/secrets.md)

## License

MIT License

## Author

Tim Alderweireldt

## Version

1.0.1

---

**Support**: For issues, feature requests, or questions, please open an issue on the GitHub repository.
