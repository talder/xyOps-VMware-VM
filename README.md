# xyOps VMWare VM Operations Plugin

Manage VMWare virtual machines using PowerShell and [VCF PowerCLI](https://developer.vmware.com/powercli). This plugin provides operations for listing, rebooting, stopping, and snapshotting VMs in your vCenter environment.

## Features

- List all virtual machines with detailed information
- Reboot VMs (planned)
- Stop VMs (planned)
- Create VM snapshots (planned)
- Configurable PowerCLI settings (certificate validation, proxy, server mode)
- Multiple export formats: JSON, CSV, and Markdown
- Optional file export alongside job output
- Auto-installs VMware.PowerCLI module if missing
- Secure credential management via secret vault

## Requirements

### CLI Requirements

- **PowerShell Core (pwsh)** - Version 7.0 or later recommended
  - On macOS: `brew install powershell`
  - On Linux: [Install instructions](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux)
  - On Windows: Comes pre-installed or [download](https://aka.ms/powershell-release)

### Module Requirements

- **VCF.PowerCLI** - Automatically installed by the plugin if not present
  - The plugin will attempt to install PowerCLI using `Install-Module -Name VCF.PowerCLI -Scope CurrentUser`
  - Requires internet connection for first-time installation
  - Manual installation: `Install-Module -Name VCF.PowerCLI -Force -AllowClobber -Scope CurrentUser`
  - **Note**: VCF.PowerCLI is the current module (VMware.PowerCLI is deprecated)

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
- **Reboot/Stop VMs**: Power operations on Virtual Machines
- **Snapshots**: Snapshot management permissions

Recommended: Use a vCenter user with at least the "Virtual Machine Power User" role.

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| vCenter server | text | Yes | - | vCenter server hostname or IP address |
| Ignore certificates | checkbox | No | true | Ignore SSL certificate validation warnings |
| No proxy | checkbox | No | true | Do not use system proxy (if unchecked, uses system proxy) |
| Single VIserver mode | checkbox | No | true | Connect to single vCenter only (if unchecked, allows multiple connections) |
| Selected action | dropdown | Yes | - | Choose the operation to perform |
| Export format | dropdown | No | JSON | Export format: JSON, CSV, or MD (Markdown) |
| Export to file | checkbox | No | false | Export data to file(s) in addition to job output |
| Enable debug mode | checkbox | No | false | Enable debug output |

**Authentication**: Credentials (`VMWARE_USERNAME` and `VMWARE_PASSWORD`) must be configured in a secret vault attached to this plugin. See [Secret Vault Configuration](#secret-vault-configuration) above.

## Available Actions

### ListVMs
Retrieves a complete list of all virtual machines in the vCenter environment.

**Output Information**:
- VM Name
- Power State (PoweredOn, PoweredOff, Suspended)
- Number of CPUs
- Memory (GB)
- Used Space (GB)
- Provisioned Space (GB)
- Guest Operating System
- ESXi Host
- Folder
- Resource Pool
- Notes/Description

### RebootVM
*Coming soon* - Gracefully reboot a virtual machine (requires VMware Tools).

### StopVM
*Coming soon* - Power off a virtual machine.

### SnapshotVM
*Coming soon* - Create a snapshot of a virtual machine.

## PowerCLI Configuration

The plugin automatically configures PowerCLI with the following settings:
- **ParticipateInCEIP**: Always set to `false` (no participation in Customer Experience Improvement Program)
- **InvalidCertificateAction**: Set to `Ignore` if "Ignore certificates" is checked, otherwise `Warn`
- **ProxyPolicy**: Set to `NoProxy` if "No proxy" is checked, otherwise `UseSystemProxy`
- **DefaultVIServerMode**: Set to `Single` if "Single VIserver mode" is checked, otherwise `Multiple`

These settings are applied at the User scope and persist for future PowerCLI sessions.

## Output Formats

The plugin supports three export formats:

### JSON Format
Structured data suitable for workflows and automation:

```json
{
  "action": "ListVMs",
  "vcenterServer": "vcenter.company.com",
  "vmCount": 25,
  "vms": [
    {
      "Name": "Web-Server-01",
      "PowerState": "PoweredOn",
      "CPUs": 4,
      "MemoryGB": 16.00,
      "UsedSpaceGB": 120.50,
      "ProvisionedSpaceGB": 200.00,
      "GuestOS": "Microsoft Windows Server 2019 (64-bit)",
      "VMHost": "esxi-01.company.com",
      "Folder": "Production",
      "ResourcePool": "Resources",
      "Notes": "Production web server"
    }
  ]
}
```

### CSV Format
Comma-separated values suitable for Excel or database import:

```csv
Name,PowerState,CPUs,MemoryGB,UsedSpaceGB,ProvisionedSpaceGB,GuestOS,VMHost,Folder,ResourcePool,Notes
"Web-Server-01",PoweredOn,4,16.00,120.50,200.00,"Microsoft Windows Server 2019 (64-bit)","esxi-01.company.com","Production","Resources","Production web server"
```

### Markdown (MD) Format
Human-readable tables with formatted output:

```markdown
## Virtual Machines (25)

| Name | Power State | CPUs | Memory (GB) | Used Space (GB) | Guest OS | Host |
|------|-------------|------|-------------|-----------------|----------|------|
| **Web-Server-01** | PoweredOn | 4 | 16.00 | 120.50 | Microsoft Windows Server 2019 (64-bit) | esxi-01.company.com |
```

## File Export

When **Export to file** is enabled, the plugin creates timestamped files:

- **JSON**: `vmware_vm_listvms_20260202_095030.json`
- **CSV**: `vmware_vm_listvms_20260202_095030.csv`
- **MD**: `vmware_vm_listvms_20260202_095030.md`

Files are created in the job's working directory and also attached to the job output for download.

## Usage Examples

**Note**: All examples require a secret vault with `VMWARE_USERNAME` and `VMWARE_PASSWORD` to be attached to the plugin.

### List All VMs (Markdown Display)

Retrieve all VMs with Markdown formatting for easy reading:

**Parameters:**
- vCenter server: `vcenter.company.com`
- Selected action: `ListVMs`
- Export format: `MD`

### List VMs and Export to JSON File

Get VM inventory and save to JSON file for automation:

**Parameters:**
- vCenter server: `vcenter.company.com`
- Selected action: `ListVMs`
- Export format: `JSON`
- Export to file: ✓ (checked)

### List VMs with Custom PowerCLI Settings

Connect with specific PowerCLI configuration:

**Parameters:**
- vCenter server: `vcenter.company.com`
- Ignore certificates: ✓ (checked)
- No proxy: ✓ (checked)
- Single VIserver mode: ✓ (checked)
- Selected action: `ListVMs`
- Export format: `CSV`

## Error Codes

| Code | Description |
|------|-------------|
| 1 | Failed to parse input JSON |
| 2 | Missing required parameters |
| 3 | Failed to install VMware.PowerCLI module |
| 4 | Failed to connect to vCenter server |
| 5 | Action not yet implemented |
| 6 | Unknown action specified |
| 7 | General error during execution |

## Troubleshooting

### PowerCLI Installation Fails

If automatic installation fails, manually install PowerCLI:

```powershell
Install-Module -Name VCF.PowerCLI -Force -AllowClobber -Scope CurrentUser
```

### Connection Fails with Certificate Error

Enable the **Ignore certificates** checkbox to bypass certificate validation, or properly configure SSL certificates on your vCenter server.

### Permission Denied

Ensure the vCenter user has appropriate permissions:
- Virtual Machine Power User role (minimum)
- Read access to all folders and resource pools you want to query

### Multiple vCenter Connections

If you need to connect to multiple vCenter servers in sequence, uncheck the **Single VIserver mode** option. This allows PowerCLI to maintain multiple active connections.

## Use Cases

### 1. VM Inventory Automation
Run this plugin on a schedule to maintain an up-to-date inventory of all VMs, storing results in a bucket for reporting.

### 2. Capacity Planning
Collect VM resource usage regularly (CPU, memory, storage) to track trends and plan infrastructure expansion.

### 3. Compliance Reporting
Export VM lists with guest OS information for software licensing and compliance audits.

### 4. VM Lifecycle Management
Use in workflows to automate VM power operations based on schedules or events.

## Workflow Integration

The JSON output can be used in xyOps workflows:

1. **Trigger**: Schedule (daily at 6:00 AM)
2. **Step 1**: Run VMWare VM Operations plugin
   - Selected action: `ListVMs`
   - Export format: `JSON`
   - Export to file: ✓ (optional, for archival)
3. **Step 2**: Parse JSON data from `data` bucket
4. **Step 3**: Send alerts if:
   - VMs are powered off unexpectedly
   - Disk usage exceeds thresholds
   - Orphaned VMs detected (no notes, old snapshots, etc.)

## Links

- [VMware PowerCLI Documentation](https://developer.vmware.com/powercli)
- [VMware PowerCLI GitHub](https://github.com/vmware/powercli-core)
- [PowerShell Installation](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [vCenter Server Documentation](https://docs.vmware.com/en/VMware-vSphere/index.html)

## License

MIT License - See [LICENSE.md](LICENSE.md) for details.

## Author

Tim Alderweireldt

## Version

1.0.0
