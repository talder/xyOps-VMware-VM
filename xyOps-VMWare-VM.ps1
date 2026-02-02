# xyOps VMWare VM Operations Plugin - PowerShell Version
# Manages VMWare virtual machines using VMware PowerCLI

function Write-Output-JSON {
    param($Object)
    $json = $Object | ConvertTo-Json -Compress -Depth 100
    Write-Output $json
    [Console]::Out.Flush()
}

function Send-Progress {
    param([double]$Value)
    Write-Output-JSON @{ xy = 1; progress = $Value }
}

function Send-Success {
    param([string]$Description = "Operation completed successfully")
    Write-Output-JSON @{ xy = 1; code = 0; description = $Description }
}

function Send-Error {
    param([int]$Code, [string]$Description)
    Write-Output-JSON @{ xy = 1; code = $Code; description = $Description }
}

function Test-SnapshotException {
    param(
        [string]$SnapshotName,
        [string]$SnapshotDescription,
        [string]$ExceptionList
    )
    
    # If no exceptions provided, don't skip
    if ([string]::IsNullOrWhiteSpace($ExceptionList)) {
        return $false
    }
    
    # Parse comma-separated exception strings
    $exceptions = $ExceptionList -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    # Check if snapshot name or description contains any exception string (case-insensitive)
    foreach ($exception in $exceptions) {
        if ($SnapshotName -match [regex]::Escape($exception) -or 
            $SnapshotDescription -match [regex]::Escape($exception)) {
            return $true
        }
    }
    
    return $false
}

# Read input from STDIN
$inputJson = [Console]::In.ReadToEnd()

try {
    $jobData = $inputJson | ConvertFrom-Json -AsHashtable
}
catch {
    Send-Error -Code 1 -Description "Failed to parse input JSON: $($_.Exception.Message)"
    exit 1
}

# Extract parameters
$params = $jobData.params

# Helper function to get parameter value case-insensitively
function Get-ParamValue {
    param($ParamsObject, [string]$ParamName)
    if ($ParamsObject -is [hashtable]) {
        foreach ($key in $ParamsObject.Keys) {
            if ($key -ieq $ParamName) {
                return $ParamsObject[$key]
            }
        }
        return $null
    } else {
        $prop = $ParamsObject.PSObject.Properties | Where-Object { $_.Name -ieq $ParamName } | Select-Object -First 1
        if ($prop) { return $prop.Value }
        return $null
    }
}

# Check if debug mode is enabled
$debugRaw = Get-ParamValue -ParamsObject $params -ParamName 'debug'
$debug = if ($debugRaw -eq $true -or $debugRaw -eq "true" -or $debugRaw -eq "True") { $true } else { $false }

# Performance tracking
$perf = @{}
$overallSW = [System.Diagnostics.Stopwatch]::StartNew()

# If debug is enabled, output the incoming JSON
if ($debug) {
    Write-Host "=== DEBUG: Incoming JSON ==="
    $debugData = @{}
    if ($jobData -is [hashtable]) {
        foreach ($key in $jobData.Keys) {
            if ($key -ne 'script') {
                $debugData[$key] = $jobData[$key]
            }
        }
    } else {
        foreach ($prop in $jobData.PSObject.Properties) {
            if ($prop.Name -ne 'script') {
                $debugData[$prop.Name] = $prop.Value
            }
        }
    }
    $formattedJson = $debugData | ConvertTo-Json -Depth 10
    Write-Host $formattedJson
    Write-Host "=== END DEBUG ==="
}

# Get parameters
$vcenterServer = Get-ParamValue -ParamsObject $params -ParamName 'vcenterserver'
$username = $env:VMWARE_USERNAME
$password = $env:VMWARE_PASSWORD
$ignoreCertRaw = Get-ParamValue -ParamsObject $params -ParamName 'ignorecert'
$ignoreCert = if ($ignoreCertRaw -eq $true -or $ignoreCertRaw -eq "true" -or $ignoreCertRaw -eq "True") { $true } else { $false }
$proxyPolicyRaw = Get-ParamValue -ParamsObject $params -ParamName 'proxypolicy'
$noProxy = if ($proxyPolicyRaw -eq $true -or $proxyPolicyRaw -eq "true" -or $proxyPolicyRaw -eq "True") { $true } else { $false }
$viServerModeRaw = Get-ParamValue -ParamsObject $params -ParamName 'viservermode'
$singleMode = if ($viServerModeRaw -eq $true -or $viServerModeRaw -eq "true" -or $viServerModeRaw -eq "True") { $true } else { $false }
$selectedAction = Get-ParamValue -ParamsObject $params -ParamName 'selectedaction'
$exportFormatRaw = Get-ParamValue -ParamsObject $params -ParamName 'exportformat'
$exportFormat = if ([string]::IsNullOrWhiteSpace($exportFormatRaw)) { "JSON" } else { $exportFormatRaw.ToUpper() }
$exportToFileRaw = Get-ParamValue -ParamsObject $params -ParamName 'exporttofile'
$exportToFile = if ($exportToFileRaw -eq $true -or $exportToFileRaw -eq "true" -or $exportToFileRaw -eq "True") { $true } else { $false }
$dateInput = Get-ParamValue -ParamsObject $params -ParamName 'dateinput'
$numberOfWeeks = Get-ParamValue -ParamsObject $params -ParamName 'numberofweeks'
$snapshotRemovalException = Get-ParamValue -ParamsObject $params -ParamName 'snapshotremovalexception'
$vmName = Get-ParamValue -ParamsObject $params -ParamName 'vmname'
$snapshotName = Get-ParamValue -ParamsObject $params -ParamName 'snapshotname'
$snapshotDescription = Get-ParamValue -ParamsObject $params -ParamName 'snapshotdescription'
$snapshotMemoryRaw = Get-ParamValue -ParamsObject $params -ParamName 'snapshotmemory'
$snapshotMemory = if ($snapshotMemoryRaw -eq $true -or $snapshotMemoryRaw -eq "true" -or $snapshotMemoryRaw -eq "True") { $true } else { $false }
$snapshotUnique = Get-ParamValue -ParamsObject $params -ParamName 'snapshotunique'
$updateVCFPowerCLIRaw = Get-ParamValue -ParamsObject $params -ParamName 'updateVCFPowerCLI'
$updateVCFPowerCLI = if ($updateVCFPowerCLIRaw -eq $true -or $updateVCFPowerCLIRaw -eq "true" -or $updateVCFPowerCLIRaw -eq "True") { $true } else { $false }

# Validate required parameters
$missing = @()

# Check vCenter server parameter
if ([string]::IsNullOrWhiteSpace($vcenterServer)) {
    $missing += 'vcenterserver'
}

# Check username from environment variable
if ([string]::IsNullOrWhiteSpace($username)) {
    $missing += 'VMWARE_USERNAME (environment variable)'
}

# Check password from environment variable
if ([string]::IsNullOrWhiteSpace($password)) {
    $missing += 'VMWARE_PASSWORD (environment variable)'
}

# Check selected action
if ([string]::IsNullOrWhiteSpace($selectedAction)) {
    $missing += 'selectedaction'
}

# Check dateInput if action is 'Remove snapshots before date'
if ($selectedAction -and $selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE DATE") {
    if ([string]::IsNullOrWhiteSpace($dateInput)) {
        $missing += 'dateinput'
    }
}

# Check numberOfWeeks if action is 'Remove snapshots before number of weeks'
if ($selectedAction -and $selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS") {
    if ([string]::IsNullOrWhiteSpace($numberOfWeeks)) {
        $missing += 'numberofweeks'
    }
}

# Check vmname if action is 'Create VM snapshot'
if ($selectedAction -and $selectedAction.ToUpper() -eq "CREATE VM SNAPSHOT") {
    if ([string]::IsNullOrWhiteSpace($vmName)) {
        $missing += 'vmname'
    }
}

# Check snapshotname if action is 'Remove snapshots containing specific text'
if ($selectedAction -and $selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT") {
    if ([string]::IsNullOrWhiteSpace($snapshotName)) {
        $missing += 'snapshotname'
    }
}

if ($missing.Count -gt 0) {
    Send-Error -Code 2 -Description "Missing required parameters: $($missing -join ', '). Credentials must be provided via secret vault environment variables."
    exit 1
}

try {
    # Check if VCF.PowerCLI module is installed
    Send-Progress -Value 0.1
    
    $moduleInstallSW = [System.Diagnostics.Stopwatch]::StartNew()
    
    $powerCLIModule = Get-Module -ListAvailable -Name VCF.PowerCLI | Select-Object -First 1
    
    if (-not $powerCLIModule) {
        try {
            Write-Host "VCF.PowerCLI module not found, attempting to install..."
            # Use -SkipPublisherCheck to handle publisher changes (VMware -> Broadcom)
            Install-Module -Name VCF.PowerCLI -Force -AllowClobber -Scope CurrentUser -SkipPublisherCheck -ErrorAction Stop
            Write-Host "VCF.PowerCLI module installed successfully"
        }
        catch {
            Send-Error -Code 3 -Description "Failed to install required VCF.PowerCLI module. Please install it manually by running: Install-Module -Name VCF.PowerCLI -Force -SkipPublisherCheck (Install error: $($_.Exception.Message))"
            exit 1
        }
    } else {
        Write-Host "VCF.PowerCLI module found: Version $($powerCLIModule.Version)"
        
        # Check if an update is available (only if updateVCFPowerCLI parameter is enabled)
        if ($updateVCFPowerCLI) {
            Write-Host "Update check enabled - checking for VCF.PowerCLI updates..."
            try {
                $latestVersion = Find-Module -Name VCF.PowerCLI -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($latestVersion -and $latestVersion.Version -gt $powerCLIModule.Version) {
                    Write-Host "Updating VCF.PowerCLI from version $($powerCLIModule.Version) to $($latestVersion.Version)..."
                    Update-Module -Name VCF.PowerCLI -Force -SkipPublisherCheck -ErrorAction Stop
                    Write-Host "VCF.PowerCLI module updated successfully"
                } else {
                    Write-Host "VCF.PowerCLI is already up-to-date (version $($powerCLIModule.Version))"
                }
            }
            catch {
                # Update failed, but we can continue with existing version
                Write-Host "Could not update VCF.PowerCLI module, using existing version $($powerCLIModule.Version)"
            }
        } else {
            Write-Host "Update check disabled - using existing VCF.PowerCLI version $($powerCLIModule.Version)"
        }
    }
    
    $moduleInstallSW.Stop(); $perf.module_install = [math]::Round($moduleInstallSW.Elapsed.TotalSeconds,3); if ($debug) { Write-Host ("PERF module_install={0:n3}s" -f $perf.module_install) }
    
    # Import only the Core PowerCLI module for faster startup (keep VCF.PowerCLI for install/update checks)
    Send-Progress -Value 0.2
    $moduleImportSW = [System.Diagnostics.Stopwatch]::StartNew()
    
    $loadedCore = Get-Module -Name VMware.VimAutomation.Core
    if (-not $loadedCore) {
        Write-Host "Loading VMware.VimAutomation.Core module (optimized path)..."
        try {
            Import-Module VMware.VimAutomation.Core -ErrorAction Stop
            $loadedCore = Get-Module -Name VMware.VimAutomation.Core
            Write-Host "VMware.VimAutomation.Core loaded successfully (version $($loadedCore.Version))"
        }
        catch {
            Write-Host "Failed to load VMware.VimAutomation.Core directly: $($_.Exception.Message). Falling back to VCF.PowerCLI import..."
            # Fallback to umbrella module if direct Core import fails
            Import-Module VCF.PowerCLI -ErrorAction Stop
            Write-Host "VCF.PowerCLI module loaded successfully (fallback)"
        }
    } else {
        Write-Host "VMware.VimAutomation.Core already loaded (version $($loadedCore.Version))"
    }
    $moduleImportSW.Stop(); $perf.module_import = [math]::Round($moduleImportSW.Elapsed.TotalSeconds,3); if ($debug) { Write-Host ("PERF module_import={0:n3}s" -f $perf.module_import) }
    
    # Configure PowerCLI settings
    Send-Progress -Value 0.3
    $configSW = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "Configuring PowerCLI settings..."
    
    # Determine InvalidCertificateAction
    $certAction = if ($ignoreCert) { "Ignore" } else { "Warn" }
    
    # Determine ProxyPolicy
    $proxyPolicy = if ($noProxy) { "NoProxy" } else { "UseSystemProxy" }
    
    # Determine DefaultVIServerMode
    $viServerMode = if ($singleMode) { "Single" } else { "Multiple" }
    
    Write-Host "PowerCLI Config: InvalidCertificateAction=$certAction, ProxyPolicy=$proxyPolicy, DefaultVIServerMode=$viServerMode"
    
    # Only set configuration if values differ (to avoid slow disk writes every run)
    $currentCfg = Get-PowerCLIConfiguration -Scope User -ErrorAction SilentlyContinue | Select-Object -First 1
    $currentInvalid = $null; if ($currentCfg) { $currentInvalid = $currentCfg.InvalidCertificateAction }
    $currentProxy = $null; if ($currentCfg) { $currentProxy = $currentCfg.ProxyPolicy }
    $currentMode = $null; if ($currentCfg) { $currentMode = $currentCfg.DefaultVIServerMode }
    $currentCeip = $null; if ($currentCfg) {
        if ($currentCfg.PSObject.Properties.Name -contains 'ParticipateInCEIP') { $currentCeip = $currentCfg.ParticipateInCEIP }
        elseif ($currentCfg.PSObject.Properties.Name -contains 'ParticipatingInCEIP') { $currentCeip = $currentCfg.ParticipatingInCEIP }
    }
    $needsCfgUpdate = $false
    if ($currentInvalid -ne $certAction) { $needsCfgUpdate = $true }
    if ($currentProxy -ne $proxyPolicy) { $needsCfgUpdate = $true }
    if ($currentMode -ne $viServerMode) { $needsCfgUpdate = $true }
    if ($null -ne $currentCeip -and $currentCeip -ne $false) { $needsCfgUpdate = $true }

    if ($needsCfgUpdate) {
        Set-PowerCLIConfiguration -Scope User `
            -ParticipateInCEIP $false `
            -InvalidCertificateAction $certAction `
            -Confirm:$false `
            -ProxyPolicy $proxyPolicy `
            -DefaultVIServerMode $viServerMode `
            -ErrorAction Stop | Out-Null
        Write-Host "PowerCLI configuration updated"
    } else {
        Write-Host "PowerCLI configuration already matches desired settings (no changes)"
    }

    $configSW.Stop(); $perf.config = [math]::Round($configSW.Elapsed.TotalSeconds,3); if ($debug) { Write-Host ("PERF config={0:n3}s" -f $perf.config) }
    
    # Connect to vCenter
    Send-Progress -Value 0.4
    $connectSW = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "Connecting to vCenter: $vcenterServer"
    
    $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
    
    try {
        $viConnection = Connect-VIServer -Server $vcenterServer -Credential $credential -ErrorAction Stop
        Write-Host "Successfully connected to vCenter: $($viConnection.Name)"
    }
    catch {
        Send-Error -Code 4 -Description "Failed to connect to vCenter server '$vcenterServer': $($_.Exception.Message)"
        exit 1
    }
    $connectSW.Stop(); $perf.connect = [math]::Round($connectSW.Elapsed.TotalSeconds,3); if ($debug) { Write-Host ("PERF connect={0:n3}s" -f $perf.connect) }
    
    Send-Progress -Value 0.5
    
    # Execute selected action
    Write-Host "Executing action: $selectedAction"
    
    $actionData = @{}
    $actionSW = [System.Diagnostics.Stopwatch]::StartNew()
    
    switch ($selectedAction.ToUpper()) {
        "LIST VMS" {
            Write-Host "Retrieving list of all VMs..."
            $vms = Get-VM -Server $viConnection | Sort-Object Name
            
            $vmList = @()
            foreach ($vm in $vms) {
                # Get IP addresses (comma-separated if multiple)
                $ipAddresses = "N/A"
                if ($vm.Guest.IPAddress -and $vm.Guest.IPAddress.Count -gt 0) {
                    $ipAddresses = ($vm.Guest.IPAddress -join ", ")
                }
                
                # Calculate uptime
                $uptime = "N/A"
                if ($vm.PowerState -eq "PoweredOn" -and $vm.ExtensionData.Runtime.BootTime) {
                    $bootTime = $vm.ExtensionData.Runtime.BootTime
                    $timeSpan = (Get-Date) - $bootTime
                    if ($timeSpan.Days -gt 0) {
                        $uptime = "$($timeSpan.Days)d $($timeSpan.Hours)h $($timeSpan.Minutes)m"
                    } elseif ($timeSpan.Hours -gt 0) {
                        $uptime = "$($timeSpan.Hours)h $($timeSpan.Minutes)m"
                    } else {
                        $uptime = "$($timeSpan.Minutes)m"
                    }
                }
                
                $vmInfo = @{
                    Name = $vm.Name
                    PowerState = $vm.PowerState.ToString()
                    CPUs = $vm.NumCpu
                    MemoryGB = [math]::Round($vm.MemoryGB, 2)
                    UsedSpaceGB = [math]::Round($vm.UsedSpaceGB, 2)
                    ProvisionedSpaceGB = [math]::Round($vm.ProvisionedSpaceGB, 2)
                    IPAddresses = $ipAddresses
                    Uptime = $uptime
                    GuestOS = $vm.Guest.OSFullName
                    VMHost = $vm.VMHost.Name
                    Folder = $vm.Folder.Name
                    ResourcePool = $vm.ResourcePool.Name
                    Notes = $vm.Notes
                }
                $vmList += $vmInfo
            }
            
            $actionData = @{
                action = "ListVMs"
                vcenterServer = $vcenterServer
                vmCount = $vmList.Count
                vms = $vmList
            }
            
            Write-Host "Retrieved $($vmList.Count) VMs"
        }
        
        "LIST SNAPSHOTS" {
            Write-Host "Retrieving list of VMs with snapshots..."
            $allVMs = Get-VM -Server $viConnection | Sort-Object Name
            
            $snapshotList = @()
            foreach ($vm in $allVMs) {
                $snapshots = Get-Snapshot -VM $vm -ErrorAction SilentlyContinue
                
                if ($snapshots) {
                    foreach ($snapshot in $snapshots) {
                        $snapshotInfo = @{
                            VMName = $vm.Name
                            SnapshotName = $snapshot.Name
                            Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                            SizeGB = [math]::Round($snapshot.SizeGB, 2)
                            Description = $snapshot.Description
                            PowerState = $snapshot.PowerState.ToString()
                        }
                        $snapshotList += $snapshotInfo
                    }
                }
            }
            
            $actionData = @{
                action = "ListSnapshots"
                vcenterServer = $vcenterServer
                snapshotCount = $snapshotList.Count
                vmWithSnapshotsCount = ($snapshotList | Select-Object -ExpandProperty VMName -Unique).Count
                snapshots = $snapshotList
            }
            
            Write-Host "Retrieved $($snapshotList.Count) snapshot(s) from $($actionData.vmWithSnapshotsCount) VM(s)"
        }
        
        "REBOOT VM" {
            # TODO: Implement reboot VM functionality
            Send-Error -Code 5 -Description "Action 'Reboot VM' is not yet implemented"
            exit 1
        }
        
        "STOP VM" {
            # TODO: Implement stop VM functionality
            Send-Error -Code 5 -Description "Action 'Stop VM' is not yet implemented"
            exit 1
        }
        
        "CREATE VM SNAPSHOT" {
            Write-Host "Creating snapshot(s) for VM(s): $vmName..."
            
            # Parse VM names (comma-separated)
            $vmNames = $vmName -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            if ($vmNames.Count -eq 0) {
                Send-Error -Code 9 -Description "No valid VM names provided. Please specify at least one VM name."
                exit 1
            }
            
            # Generate base snapshot name if not provided
            $baseSnapshotName = if ([string]::IsNullOrWhiteSpace($snapshotName)) {
                $timestamp = Get-Date -Format "yyyyMMddHHmmss"
                "xyOps-VMWare-VM-$timestamp"
            } else {
                $snapshotName
            }
            
            # Generate unique identifier based on selection
            $uniqueIdentifier = ""
            if (-not [string]::IsNullOrWhiteSpace($snapshotUnique) -and $snapshotUnique -ne "No") {
                if ($snapshotUnique -eq "Unique short UID") {
                    # Generate GUID and take first 8 characters
                    $guid = [System.Guid]::NewGuid().ToString("N")
                    $uniqueIdentifier = "-" + $guid.Substring(0, 8)
                } elseif ($snapshotUnique -eq "Timestamp") {
                    # Generate timestamp
                    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
                    $uniqueIdentifier = "-$timestamp"
                }
            }
            
            # Build final snapshot name
            $finalSnapshotName = $baseSnapshotName + $uniqueIdentifier
            
            # Determine snapshot description
            $finalDescription = if ([string]::IsNullOrWhiteSpace($snapshotDescription)) {
                if ($uniqueIdentifier -ne "") {
                    "xyOps-VMWare-VM added a custom snapshot with a unique identifier"
                } else {
                    "xyOps-VMWare-VM snapshot"
                }
            } else {
                $snapshotDescription
            }
            
            Write-Host "Snapshot Name: $finalSnapshotName"
            Write-Host "Description: $finalDescription"
            Write-Host "Memory Snapshot: $snapshotMemory"
            Write-Host "Target VMs: $($vmNames -join ', ')"
            
            $createdSnapshots = @()
            $failedSnapshots = @()
            
            foreach ($vmNameItem in $vmNames) {
                try {
                    Write-Host "Processing VM: $vmNameItem..."
                    
                    # Get the VM object
                    $vm = Get-VM -Name $vmNameItem -Server $viConnection -ErrorAction Stop
                    
                    # Create the snapshot
                    Write-Host "Creating snapshot '$finalSnapshotName' for VM '$vmNameItem'..."
                    $snapshot = New-Snapshot -VM $vm -Name $finalSnapshotName -Description $finalDescription -Memory:$snapshotMemory -Quiesce:$false -ErrorAction Stop
                    
                    $snapshotInfo = @{
                        VMName = $vm.Name
                        SnapshotName = $snapshot.Name
                        Description = $snapshot.Description
                        Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                        SizeGB = [math]::Round($snapshot.SizeGB, 2)
                        Memory = $snapshotMemory
                        Status = "Success"
                        Message = "Snapshot created successfully"
                    }
                    $createdSnapshots += $snapshotInfo
                    Write-Host "Successfully created snapshot '$finalSnapshotName' for VM '$vmNameItem'"
                }
                catch {
                    $failedInfo = @{
                        VMName = $vmNameItem
                        SnapshotName = $finalSnapshotName
                        Description = $finalDescription
                        Memory = $snapshotMemory
                        Status = "Failed"
                        Message = $_.Exception.Message
                    }
                    $failedSnapshots += $failedInfo
                    Write-Host "Failed to create snapshot for VM '$vmNameItem': $($_.Exception.Message)"
                }
            }
            
            $actionData = @{
                action = "CreateVMSnapshot"
                vcenterServer = $vcenterServer
                snapshotName = $finalSnapshotName
                description = $finalDescription
                memory = $snapshotMemory
                uniqueIdentifier = $snapshotUnique
                targetVMs = $vmNames
                createdCount = $createdSnapshots.Count
                failedCount = $failedSnapshots.Count
                totalProcessed = $createdSnapshots.Count + $failedSnapshots.Count
                createdSnapshots = $createdSnapshots
                failedSnapshots = $failedSnapshots
            }
            
            Write-Host "Snapshot creation completed: $($createdSnapshots.Count) created, $($failedSnapshots.Count) failed"
        }
        
        "REMOVE SNAPSHOTS BEFORE DATE" {
            Write-Host "Removing snapshots created before $dateInput..."
            
            # Parse the date input with explicit format validation
            $targetDate = $null
            $validFormats = @(
                'yyyy-MM-dd',
                'yyyy-MM-dd HH:mm:ss',
                'MM/dd/yyyy',
                'dd/MM/yyyy',
                'yyyy/MM/dd'
            )
            
            # Try parsing with specific formats first
            $parseSuccess = $false
            foreach ($format in $validFormats) {
                try {
                    $targetDate = [DateTime]::ParseExact($dateInput, $format, $null)
                    $parseSuccess = $true
                    Write-Host "Target date parsed using format '$format': $($targetDate.ToString('yyyy-MM-dd HH:mm:ss'))"
                    break
                }
                catch {
                    # Try next format
                }
            }
            
            # If specific formats failed, try general parsing as fallback
            if (-not $parseSuccess) {
                try {
                    $targetDate = [DateTime]::Parse($dateInput)
                    $parseSuccess = $true
                    Write-Host "Target date parsed: $($targetDate.ToString('yyyy-MM-dd HH:mm:ss'))"
                }
                catch {
                    Send-Error -Code 7 -Description "Invalid date format: '$dateInput'. Please use one of these formats: yyyy-MM-dd, yyyy-MM-dd HH:mm:ss, MM/dd/yyyy, dd/MM/yyyy, or yyyy/MM/dd"
                    exit 1
                }
            }
            
            $allVMs = Get-VM -Server $viConnection | Sort-Object Name
            
            $removedSnapshots = @()
            $failedRemovals = @()
            $skippedSnapshots = @()
            
            foreach ($vm in $allVMs) {
                $snapshots = Get-Snapshot -VM $vm -ErrorAction SilentlyContinue
                
                if ($snapshots) {
                    foreach ($snapshot in $snapshots) {
                        if ($snapshot.Created -lt $targetDate) {
                            # Check if snapshot matches exception criteria
                            $shouldSkip = Test-SnapshotException -SnapshotName $snapshot.Name -SnapshotDescription $snapshot.Description -ExceptionList $snapshotRemovalException
                            
                            if ($shouldSkip) {
                                Write-Host "Skipping snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (matches exception criteria)"
                                $skippedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Reason = "Matches exception criteria"
                                }
                                $skippedSnapshots += $skippedInfo
                                continue
                            }
                            
                            Write-Host "Removing snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (Created: $($snapshot.Created.ToString('yyyy-MM-dd HH:mm:ss')))..."
                            
                            try {
                                Remove-Snapshot -Snapshot $snapshot -Confirm:$false -ErrorAction Stop
                                
                                $removedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Status = "Success"
                                    Message = "Successfully removed"
                                }
                                $removedSnapshots += $removedInfo
                                Write-Host "Successfully removed snapshot '$($snapshot.Name)' from VM '$($vm.Name)'"
                            }
                            catch {
                                $failedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Status = "Failed"
                                    Message = $_.Exception.Message
                                }
                                $failedRemovals += $failedInfo
                                Write-Host "Failed to remove snapshot '$($snapshot.Name)' from VM '$($vm.Name)': $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            
            $actionData = @{
                action = "RemoveSnapshotsBeforeDate"
                vcenterServer = $vcenterServer
                targetDate = $targetDate.ToString("yyyy-MM-dd HH:mm:ss")
                removedCount = $removedSnapshots.Count
                failedCount = $failedRemovals.Count
                skippedCount = $skippedSnapshots.Count
                totalProcessed = $removedSnapshots.Count + $failedRemovals.Count
                removedSnapshots = $removedSnapshots
                failedRemovals = $failedRemovals
                skippedSnapshots = $skippedSnapshots
            }
            
            Write-Host "Snapshot removal completed: $($removedSnapshots.Count) removed, $($failedRemovals.Count) failed, $($skippedSnapshots.Count) skipped"
        }
        
        "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS" {
            Write-Host "Removing snapshots older than $numberOfWeeks weeks..."
            
            # Validate numberOfWeeks is a valid number
            try {
                $weeksValue = [int]$numberOfWeeks
                if ($weeksValue -le 0) {
                    Send-Error -Code 8 -Description "Invalid number of weeks: '$numberOfWeeks'. Please provide a positive number."
                    exit 1
                }
            }
            catch {
                Send-Error -Code 8 -Description "Invalid number of weeks: '$numberOfWeeks'. Please provide a valid positive integer."
                exit 1
            }
            
            # Calculate the target date (current date minus number of weeks)
            $targetDate = (Get-Date).AddDays(-($weeksValue * 7))
            Write-Host "Target date calculated: $($targetDate.ToString('yyyy-MM-dd HH:mm:ss')) (snapshots older than $weeksValue week(s) will be removed)"
            
            $allVMs = Get-VM -Server $viConnection | Sort-Object Name
            
            $removedSnapshots = @()
            $failedRemovals = @()
            $skippedSnapshots = @()
            
            foreach ($vm in $allVMs) {
                $snapshots = Get-Snapshot -VM $vm -ErrorAction SilentlyContinue
                
                if ($snapshots) {
                    foreach ($snapshot in $snapshots) {
                        if ($snapshot.Created -lt $targetDate) {
                            # Check if snapshot matches exception criteria
                            $shouldSkip = Test-SnapshotException -SnapshotName $snapshot.Name -SnapshotDescription $snapshot.Description -ExceptionList $snapshotRemovalException
                            
                            if ($shouldSkip) {
                                Write-Host "Skipping snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (matches exception criteria)"
                                $ageInDays = [math]::Round(((Get-Date) - $snapshot.Created).TotalDays, 1)
                                $skippedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    AgeDays = $ageInDays
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Reason = "Matches exception criteria"
                                }
                                $skippedSnapshots += $skippedInfo
                                continue
                            }
                            
                            $ageInDays = [math]::Round(((Get-Date) - $snapshot.Created).TotalDays, 1)
                            Write-Host "Removing snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (Created: $($snapshot.Created.ToString('yyyy-MM-dd HH:mm:ss')), Age: $ageInDays days)..."
                            
                            try {
                                Remove-Snapshot -Snapshot $snapshot -Confirm:$false -ErrorAction Stop
                                
                                $removedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    AgeDays = $ageInDays
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Status = "Success"
                                    Message = "Successfully removed"
                                }
                                $removedSnapshots += $removedInfo
                                Write-Host "Successfully removed snapshot '$($snapshot.Name)' from VM '$($vm.Name)'"
                            }
                            catch {
                                $failedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    AgeDays = $ageInDays
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Status = "Failed"
                                    Message = $_.Exception.Message
                                }
                                $failedRemovals += $failedInfo
                                Write-Host "Failed to remove snapshot '$($snapshot.Name)' from VM '$($vm.Name)': $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            
            $actionData = @{
                action = "RemoveVMSnapshotsBeforeNumberOfWeeks"
                vcenterServer = $vcenterServer
                numberOfWeeks = $weeksValue
                targetDate = $targetDate.ToString("yyyy-MM-dd HH:mm:ss")
                removedCount = $removedSnapshots.Count
                failedCount = $failedRemovals.Count
                skippedCount = $skippedSnapshots.Count
                totalProcessed = $removedSnapshots.Count + $failedRemovals.Count
                removedSnapshots = $removedSnapshots
                failedRemovals = $failedRemovals
                skippedSnapshots = $skippedSnapshots
            }
            
            Write-Host "Snapshot removal completed: $($removedSnapshots.Count) removed, $($failedRemovals.Count) failed, $($skippedSnapshots.Count) skipped"
        }
        
        "REMOVE ALL SNAPSHOTS" {
            Write-Host "Removing ALL snapshots from all VMs..."
            
            $allVMs = Get-VM -Server $viConnection | Sort-Object Name
            
            $removedSnapshots = @()
            $failedRemovals = @()
            $skippedSnapshots = @()
            
            foreach ($vm in $allVMs) {
                $snapshots = Get-Snapshot -VM $vm -ErrorAction SilentlyContinue
                
                if ($snapshots) {
                    foreach ($snapshot in $snapshots) {
                        # Check if snapshot matches exception criteria
                        $shouldSkip = Test-SnapshotException -SnapshotName $snapshot.Name -SnapshotDescription $snapshot.Description -ExceptionList $snapshotRemovalException
                        
                        if ($shouldSkip) {
                            Write-Host "Skipping snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (matches exception criteria)"
                            $ageInDays = [math]::Round(((Get-Date) - $snapshot.Created).TotalDays, 1)
                            $skippedInfo = @{
                                VMName = $vm.Name
                                SnapshotName = $snapshot.Name
                                Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                AgeDays = $ageInDays
                                SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                Reason = "Matches exception criteria"
                            }
                            $skippedSnapshots += $skippedInfo
                            continue
                        }
                        
                        $ageInDays = [math]::Round(((Get-Date) - $snapshot.Created).TotalDays, 1)
                        Write-Host "Removing snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (Created: $($snapshot.Created.ToString('yyyy-MM-dd HH:mm:ss')), Age: $ageInDays days)..."
                        
                        try {
                            Remove-Snapshot -Snapshot $snapshot -Confirm:$false -ErrorAction Stop
                            
                            $removedInfo = @{
                                VMName = $vm.Name
                                SnapshotName = $snapshot.Name
                                Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                AgeDays = $ageInDays
                                SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                Status = "Success"
                                Message = "Successfully removed"
                            }
                            $removedSnapshots += $removedInfo
                            Write-Host "Successfully removed snapshot '$($snapshot.Name)' from VM '$($vm.Name)'"
                        }
                        catch {
                            $failedInfo = @{
                                VMName = $vm.Name
                                SnapshotName = $snapshot.Name
                                Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                AgeDays = $ageInDays
                                SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                Status = "Failed"
                                Message = $_.Exception.Message
                            }
                            $failedRemovals += $failedInfo
                            Write-Host "Failed to remove snapshot '$($snapshot.Name)' from VM '$($vm.Name)': $($_.Exception.Message)"
                        }
                    }
                }
            }
            
            $actionData = @{
                action = "RemoveAllSnapshots"
                vcenterServer = $vcenterServer
                removedCount = $removedSnapshots.Count
                failedCount = $failedRemovals.Count
                skippedCount = $skippedSnapshots.Count
                totalProcessed = $removedSnapshots.Count + $failedRemovals.Count
                removedSnapshots = $removedSnapshots
                failedRemovals = $failedRemovals
                skippedSnapshots = $skippedSnapshots
            }
            
            Write-Host "Snapshot removal completed: $($removedSnapshots.Count) removed, $($failedRemovals.Count) failed, $($skippedSnapshots.Count) skipped"
        }
        
        "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT" {
            Write-Host "Removing snapshots containing text: '$snapshotName'..."
            
            $allVMs = Get-VM -Server $viConnection | Sort-Object Name
            
            $removedSnapshots = @()
            $failedRemovals = @()
            $skippedSnapshots = @()
            
            foreach ($vm in $allVMs) {
                $snapshots = Get-Snapshot -VM $vm -ErrorAction SilentlyContinue
                
                if ($snapshots) {
                    foreach ($snapshot in $snapshots) {
                        # Check if snapshot name contains the search text (case-insensitive)
                        if ($snapshot.Name -like "*$snapshotName*") {
                            # Check if snapshot matches exception criteria
                            $shouldSkip = Test-SnapshotException -SnapshotName $snapshot.Name -SnapshotDescription $snapshot.Description -ExceptionList $snapshotRemovalException
                            
                            if ($shouldSkip) {
                                Write-Host "Skipping snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (matches exception criteria)"
                                $ageInDays = [math]::Round(((Get-Date) - $snapshot.Created).TotalDays, 1)
                                $skippedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    AgeDays = $ageInDays
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Reason = "Matches exception criteria"
                                }
                                $skippedSnapshots += $skippedInfo
                                continue
                            }
                            
                            $ageInDays = [math]::Round(((Get-Date) - $snapshot.Created).TotalDays, 1)
                            Write-Host "Removing snapshot '$($snapshot.Name)' from VM '$($vm.Name)' (matches text: '$snapshotName')..."
                            
                            try {
                                Remove-Snapshot -Snapshot $snapshot -Confirm:$false -ErrorAction Stop
                                
                                $removedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    AgeDays = $ageInDays
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Status = "Success"
                                    Message = "Successfully removed"
                                }
                                $removedSnapshots += $removedInfo
                                Write-Host "Successfully removed snapshot '$($snapshot.Name)' from VM '$($vm.Name)'"
                            }
                            catch {
                                $failedInfo = @{
                                    VMName = $vm.Name
                                    SnapshotName = $snapshot.Name
                                    Created = $snapshot.Created.ToString("yyyy-MM-dd HH:mm:ss")
                                    AgeDays = $ageInDays
                                    SizeGB = [math]::Round($snapshot.SizeGB, 2)
                                    Status = "Failed"
                                    Message = $_.Exception.Message
                                }
                                $failedRemovals += $failedInfo
                                Write-Host "Failed to remove snapshot '$($snapshot.Name)' from VM '$($vm.Name)': $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            
            $actionData = @{
                action = "RemoveSnapshotsContainingText"
                vcenterServer = $vcenterServer
                searchText = $snapshotName
                removedCount = $removedSnapshots.Count
                failedCount = $failedRemovals.Count
                skippedCount = $skippedSnapshots.Count
                totalProcessed = $removedSnapshots.Count + $failedRemovals.Count
                removedSnapshots = $removedSnapshots
                failedRemovals = $failedRemovals
                skippedSnapshots = $skippedSnapshots
            }
            
            Write-Host "Snapshot removal completed: $($removedSnapshots.Count) removed, $($failedRemovals.Count) failed, $($skippedSnapshots.Count) skipped"
        }
        
        default {
            Send-Error -Code 6 -Description "Unknown action: $selectedAction. Valid actions are: ListVMs, ListSnapshots, CreateVMSnapshot, RemoveSnapshotsBeforeDate, RemoveSnapshotsBeforeNumberOfWeeks, RemoveAllSnapshots, RemoveSnapshotsContainingSpecificText, RebootVM, StopVM"
            exit 1
        }
    }
    $actionSW.Stop(); $perf.action = [math]::Round($actionSW.Elapsed.TotalSeconds,3); if ($debug) { Write-Host ("PERF action={0:n3}s" -f $perf.action) }
    
    Send-Progress -Value 0.7
    
    # Build Markdown output for display in xyOps GUI
    Write-Host "Building Markdown output for display..."
    $markdownContent = "# VMWare VM Operations Report`n`n"
    $markdownContent += "## Operation Details`n`n"
    $markdownContent += "| Property | Value |`n"
    $markdownContent += "|----------|-------|`n"
    $markdownContent += "| **Action** | $($actionData.action) |`n"
    $markdownContent += "| **vCenter Server** | $($actionData.vcenterServer) |`n"
    $markdownContent += "| **Timestamp** | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') |`n"
    $markdownContent += "`n"
    
    if ($selectedAction.ToUpper() -eq "LIST SNAPSHOTS") {
        # Display snapshots grouped by VM
        if ($actionData.snapshots.Count -gt 0) {
            $markdownContent += "## VM Snapshots Overview`n`n"
            $markdownContent += "**Total Snapshots**: $($actionData.snapshotCount) | **VMs with Snapshots**: $($actionData.vmWithSnapshotsCount)`n`n"
            $markdownContent += "## Snapshot Details`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Size (GB) | Power State | Description |`n"
            $markdownContent += "|---------|---------------|---------|-----------|-------------|-------------|`n"
            
            foreach ($snapshot in $actionData.snapshots) {
                $desc = if ([string]::IsNullOrWhiteSpace($snapshot.Description)) { "-" } else { $snapshot.Description }
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.SizeGB) | $($snapshot.PowerState) | $desc |`n"
            }
            $markdownContent += "`n"
        } else {
            $markdownContent += "## No Snapshots Found`n`n"
            $markdownContent += "There are no snapshots in the vCenter environment.`n`n"
        }
    } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE DATE") {
        # Display snapshot removal results
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Target Date**: $($actionData.targetDate)`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Skipped**: $($actionData.skippedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.skippedSnapshots.Count -gt 0) {
            $markdownContent += "## Skipped Snapshots ($($actionData.skippedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Size (GB) | Reason |`n"
            $markdownContent += "|---------|---------------|---------|-----------|--------|`n"
            
            foreach ($skipped in $actionData.skippedSnapshots) {
                $markdownContent += "| **$($skipped.VMName)** | $($skipped.SnapshotName) | $($skipped.Created) | $($skipped.SizeGB) | $($skipped.Reason) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
            $markdownContent += "## No Snapshots Found`n`n"
            $markdownContent += "No snapshots were found before the target date $($actionData.targetDate).`n`n"
        }
    } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS") {
        # Display snapshot removal results for weeks-based removal
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Number of Weeks**: $($actionData.numberOfWeeks) week(s)`n`n"
        $markdownContent += "**Target Date**: $($actionData.targetDate) (snapshots older than this are removed)`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Skipped**: $($actionData.skippedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.AgeDays) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.skippedSnapshots.Count -gt 0) {
            $markdownContent += "## Skipped Snapshots ($($actionData.skippedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Reason |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|--------|`n"
            
            foreach ($skipped in $actionData.skippedSnapshots) {
                $markdownContent += "| **$($skipped.VMName)** | $($skipped.SnapshotName) | $($skipped.Created) | $($skipped.AgeDays) | $($skipped.SizeGB) | $($skipped.Reason) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.AgeDays) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
            $markdownContent += "## No Snapshots Found`n`n"
            $markdownContent += "No snapshots older than $($actionData.numberOfWeeks) week(s) were found.`n`n"
        }
    } elseif ($selectedAction.ToUpper() -eq "REMOVE ALL SNAPSHOTS") {
        # Display snapshot removal results for all snapshots
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Action**: Remove ALL Snapshots`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Skipped**: $($actionData.skippedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.AgeDays) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.skippedSnapshots.Count -gt 0) {
            $markdownContent += "## Skipped Snapshots ($($actionData.skippedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Reason |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|--------|`n"
            
            foreach ($skipped in $actionData.skippedSnapshots) {
                $markdownContent += "| **$($skipped.VMName)** | $($skipped.SnapshotName) | $($skipped.Created) | $($skipped.AgeDays) | $($skipped.SizeGB) | $($skipped.Reason) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.AgeDays) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
            $markdownContent += "## No Snapshots Found`n`n"
            $markdownContent += "No snapshots were found in the vCenter environment.`n`n"
        }
    } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT") {
        # Display snapshot removal results for text-based search
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Search Text**: $($actionData.searchText)`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Skipped**: $($actionData.skippedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.AgeDays) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.skippedSnapshots.Count -gt 0) {
            $markdownContent += "## Skipped Snapshots ($($actionData.skippedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Reason |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|--------|`n"
            
            foreach ($skipped in $actionData.skippedSnapshots) {
                $markdownContent += "| **$($skipped.VMName)** | $($skipped.SnapshotName) | $($skipped.Created) | $($skipped.AgeDays) | $($skipped.SizeGB) | $($skipped.Reason) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.AgeDays) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
            $markdownContent += "## No Snapshots Found`n`n"
            $markdownContent += "No snapshots containing the text '$($actionData.searchText)' were found.`n`n"
        }
    } elseif ($selectedAction.ToUpper() -eq "CREATE VM SNAPSHOT") {
        # Display snapshot creation results
        $markdownContent += "## Snapshot Creation Summary`n`n"
        $markdownContent += "**Snapshot Name**: $($actionData.snapshotName)`n`n"
        $markdownContent += "**Description**: $($actionData.description)`n`n"
        $markdownContent += "**Memory Snapshot**: $($actionData.memory)`n`n"
        $markdownContent += "**Unique Identifier**: $($actionData.uniqueIdentifier)`n`n"
        $markdownContent += "**Created**: $($actionData.createdCount) | **Failed**: $($actionData.failedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.createdSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Created Snapshots ($($actionData.createdCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Description | Created | Size (GB) | Memory |`n"
            $markdownContent += "|---------|---------------|-------------|---------|-----------|--------|`n"
            
            foreach ($snapshot in $actionData.createdSnapshots) {
                $memoryIcon = if ($snapshot.Memory) { "✓" } else { "✗" }
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Description) | $($snapshot.Created) | $($snapshot.SizeGB) | $memoryIcon |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedSnapshots.Count -gt 0) {
            $markdownContent += "## Failed Snapshot Creations ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Error Message |`n"
            $markdownContent += "|---------|---------------|---------------|`n"
            
            foreach ($failed in $actionData.failedSnapshots) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.createdCount -eq 0 -and $actionData.failedCount -gt 0) {
            $markdownContent += "## No Snapshots Created`n`n"
            $markdownContent += "All snapshot creation attempts failed. Please check the error messages above.`n`n"
        }
    } elseif ($selectedAction.ToUpper() -eq "LIST VMS") {
        # Separate VMs by power state
        $poweredOnVMs = $actionData.vms | Where-Object { $_.PowerState -eq "PoweredOn" }
        $poweredOffVMs = $actionData.vms | Where-Object { $_.PowerState -eq "PoweredOff" }
        
        # PoweredOn VMs table
        if ($poweredOnVMs.Count -gt 0) {
            $markdownContent += "## Powered On Virtual Machines ($($poweredOnVMs.Count))`n`n"
            $markdownContent += "| Name | IP Addresses | Uptime | CPUs | Memory (GB) | Used Space (GB) | Guest OS | Host |`n"
            $markdownContent += "|------|--------------|--------|------|-------------|-----------------|----------|------|`n"
            
            foreach ($vm in $poweredOnVMs) {
                # Format IP addresses with line breaks for Markdown
                $ipFormatted = if ($vm.IPAddresses -ne "N/A" -and $vm.IPAddresses -match ",") {
                    ($vm.IPAddresses -split ", ") -join "<br>"
                } else {
                    $vm.IPAddresses
                }
                $markdownContent += "| **$($vm.Name)** | $ipFormatted | $($vm.Uptime) | $($vm.CPUs) | $($vm.MemoryGB) | $($vm.UsedSpaceGB) | $($vm.GuestOS) | $($vm.VMHost) |`n"
            }
            $markdownContent += "`n"
        }
        
        # PoweredOff VMs table
        if ($poweredOffVMs.Count -gt 0) {
            $markdownContent += "## Powered Off Virtual Machines ($($poweredOffVMs.Count))`n`n"
            $markdownContent += "| Name | CPUs | Memory (GB) | Used Space (GB) | Guest OS | Host |`n"
            $markdownContent += "|------|------|-------------|-----------------|----------|------|`n"
            
            foreach ($vm in $poweredOffVMs) {
                $markdownContent += "| **$($vm.Name)** | $($vm.CPUs) | $($vm.MemoryGB) | $($vm.UsedSpaceGB) | $($vm.GuestOS) | $($vm.VMHost) |`n"
            }
            $markdownContent += "`n"
        }
    }
    
    $outputSW = [System.Diagnostics.Stopwatch]::StartNew()

    # Always output the Markdown for display in xyOps GUI
    if ($selectedAction.ToUpper() -eq "LIST SNAPSHOTS") {
        $caption = "vCenter: $($actionData.vcenterServer) | Action: $($actionData.action) | Snapshots: $($actionData.snapshotCount) from $($actionData.vmWithSnapshotsCount) VM(s)"
    } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE DATE") {
        $caption = "vCenter: $($actionData.vcenterServer) | Action: $($actionData.action) | Removed: $($actionData.removedCount), Failed: $($actionData.failedCount), Skipped: $($actionData.skippedCount)"
    } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS") {
        $caption = "vCenter: $($actionData.vcenterServer) | Action: $($actionData.action) | Weeks: $($actionData.numberOfWeeks) | Removed: $($actionData.removedCount), Failed: $($actionData.failedCount), Skipped: $($actionData.skippedCount)"
    } elseif ($selectedAction.ToUpper() -eq "REMOVE ALL SNAPSHOTS") {
        $caption = "vCenter: $($actionData.vcenterServer) | Action: $($actionData.action) | Removed: $($actionData.removedCount), Failed: $($actionData.failedCount), Skipped: $($actionData.skippedCount)"
    } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT") {
        $caption = "vCenter: $($actionData.vcenterServer) | Action: $($actionData.action) | Search: '$($actionData.searchText)' | Removed: $($actionData.removedCount), Failed: $($actionData.failedCount), Skipped: $($actionData.skippedCount)"
    } elseif ($selectedAction.ToUpper() -eq "CREATE VM SNAPSHOT") {
        $caption = "vCenter: $($actionData.vcenterServer) | Action: $($actionData.action) | Created: $($actionData.createdCount), Failed: $($actionData.failedCount)"
    } else {
        $caption = "vCenter: $($actionData.vcenterServer) | Action: $($actionData.action) | VMs: $($actionData.vmCount)"
    }
    Write-Output-JSON @{
        xy = 1
        markdown = @{
            title = "VMWare VM Operations Report"
            content = $markdownContent
            caption = $caption
        }
    }
    
    Send-Progress -Value 0.9
    
    # Output data in requested format
    Write-Host "Outputting data in $exportFormat format..."
    
    switch ($exportFormat) {
        "JSON" {
            if ($exportToFile) {
                # Generate filename with timestamp
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $filename = "vmware_vm_$($selectedAction.ToLower())_$timestamp.json"
                
                # Write JSON to file
                $actionData | ConvertTo-Json -Depth 100 | Out-File -FilePath $filename -Encoding UTF8
                Write-Host "JSON file created: $filename"
                
                # Output JSON data AND file reference
                $jsonData = @{
                    xy = 1
                    data = $actionData
                    files = @(
                        @{
                            path = $filename
                            name = $filename
                        }
                    )
                }
                Write-Output-JSON $jsonData
            } else {
                # Output JSON data only (no file)
                $jsonData = @{
                    xy = 1
                    data = $actionData
                }
                Write-Output-JSON $jsonData
            }
        }
        
        "CSV" {
            # Build CSV output
            $csvContent = ""
            
            if ($selectedAction.ToUpper() -eq "LIST VMS") {
                # CSV for VM list
                $csvContent = "Name,PowerState,IPAddresses,Uptime,CPUs,MemoryGB,UsedSpaceGB,ProvisionedSpaceGB,GuestOS,VMHost,Folder,ResourcePool,Notes`n"
                foreach ($vm in $actionData.vms) {
                    $csvContent += "`"$($vm.Name)`",$($vm.PowerState),`"$($vm.IPAddresses)`",`"$($vm.Uptime)`",$($vm.CPUs),$($vm.MemoryGB),$($vm.UsedSpaceGB),$($vm.ProvisionedSpaceGB),`"$($vm.GuestOS)`",`"$($vm.VMHost)`",`"$($vm.Folder)`",`"$($vm.ResourcePool)`",`"$($vm.Notes)`"`n"
                }
                $csvContent = $csvContent.TrimEnd("`n")
            } elseif ($selectedAction.ToUpper() -eq "LIST SNAPSHOTS") {
                # CSV for snapshot list
                $csvContent = "VMName,SnapshotName,Created,SizeGB,PowerState,Description`n"
                foreach ($snapshot in $actionData.snapshots) {
                    $desc = if ([string]::IsNullOrWhiteSpace($snapshot.Description)) { "" } else { $snapshot.Description }
                    $csvContent += "`"$($snapshot.VMName)`",`"$($snapshot.SnapshotName)`",`"$($snapshot.Created)`",$($snapshot.SizeGB),`"$($snapshot.PowerState)`",`"$desc`"`n"
                }
                $csvContent = $csvContent.TrimEnd("`n")
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE DATE") {
                # CSV for removed snapshots
                $csvContent = "VMName,SnapshotName,Created,SizeGB,Status,Message`n"
                foreach ($snapshot in $actionData.removedSnapshots) {
                    $csvContent += "`"$($snapshot.VMName)`",`"$($snapshot.SnapshotName)`",`"$($snapshot.Created)`",$($snapshot.SizeGB),`"$($snapshot.Status)`",`"$($snapshot.Message)`"`n"
                }
                foreach ($skipped in $actionData.skippedSnapshots) {
                    $csvContent += "`"$($skipped.VMName)`",`"$($skipped.SnapshotName)`",`"$($skipped.Created)`",$($skipped.SizeGB),`"Skipped`",`"$($skipped.Reason)`"`n"
                }
                foreach ($failed in $actionData.failedRemovals) {
                    $csvContent += "`"$($failed.VMName)`",`"$($failed.SnapshotName)`",`"$($failed.Created)`",$($failed.SizeGB),`"$($failed.Status)`",`"$($failed.Message)`"`n"
                }
                $csvContent = $csvContent.TrimEnd("`n")
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS") {
                # CSV for removed snapshots (weeks-based)
                $csvContent = "VMName,SnapshotName,Created,AgeDays,SizeGB,Status,Message`n"
                foreach ($snapshot in $actionData.removedSnapshots) {
                    $csvContent += "`"$($snapshot.VMName)`",`"$($snapshot.SnapshotName)`",`"$($snapshot.Created)`",$($snapshot.AgeDays),$($snapshot.SizeGB),`"$($snapshot.Status)`",`"$($snapshot.Message)`"`n"
                }
                foreach ($skipped in $actionData.skippedSnapshots) {
                    $csvContent += "`"$($skipped.VMName)`",`"$($skipped.SnapshotName)`",`"$($skipped.Created)`",$($skipped.AgeDays),$($skipped.SizeGB),`"Skipped`",`"$($skipped.Reason)`"`n"
                }
                foreach ($failed in $actionData.failedRemovals) {
                    $csvContent += "`"$($failed.VMName)`",`"$($failed.SnapshotName)`",`"$($failed.Created)`",$($failed.AgeDays),$($failed.SizeGB),`"$($failed.Status)`",`"$($failed.Message)`"`n"
                }
                $csvContent = $csvContent.TrimEnd("`n")
            } elseif ($selectedAction.ToUpper() -eq "REMOVE ALL SNAPSHOTS") {
                # CSV for removed all snapshots
                $csvContent = "VMName,SnapshotName,Created,AgeDays,SizeGB,Status,Message`n"
                foreach ($snapshot in $actionData.removedSnapshots) {
                    $csvContent += "`"$($snapshot.VMName)`",`"$($snapshot.SnapshotName)`",`"$($snapshot.Created)`",$($snapshot.AgeDays),$($snapshot.SizeGB),`"$($snapshot.Status)`",`"$($snapshot.Message)`"`n"
                }
                foreach ($skipped in $actionData.skippedSnapshots) {
                    $csvContent += "`"$($skipped.VMName)`",`"$($skipped.SnapshotName)`",`"$($skipped.Created)`",$($skipped.AgeDays),$($skipped.SizeGB),`"Skipped`",`"$($skipped.Reason)`"`n"
                }
                foreach ($failed in $actionData.failedRemovals) {
                    $csvContent += "`"$($failed.VMName)`",`"$($failed.SnapshotName)`",`"$($failed.Created)`",$($failed.AgeDays),$($failed.SizeGB),`"$($failed.Status)`",`"$($failed.Message)`"`n"
                }
                $csvContent = $csvContent.TrimEnd("`n")
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT") {
                # CSV for removed snapshots containing text
                $csvContent = "VMName,SnapshotName,Created,AgeDays,SizeGB,Status,Message`n"
                foreach ($snapshot in $actionData.removedSnapshots) {
                    $csvContent += "`"$($snapshot.VMName)`",`"$($snapshot.SnapshotName)`",`"$($snapshot.Created)`",$($snapshot.AgeDays),$($snapshot.SizeGB),`"$($snapshot.Status)`",`"$($snapshot.Message)`"`n"
                }
                foreach ($skipped in $actionData.skippedSnapshots) {
                    $csvContent += "`"$($skipped.VMName)`",`"$($skipped.SnapshotName)`",`"$($skipped.Created)`",$($skipped.AgeDays),$($skipped.SizeGB),`"Skipped`",`"$($skipped.Reason)`"`n"
                }
                foreach ($failed in $actionData.failedRemovals) {
                    $csvContent += "`"$($failed.VMName)`",`"$($failed.SnapshotName)`",`"$($failed.Created)`",$($failed.AgeDays),$($failed.SizeGB),`"$($failed.Status)`",`"$($failed.Message)`"`n"
                }
                $csvContent = $csvContent.TrimEnd("`n")
            } elseif ($selectedAction.ToUpper() -eq "CREATE VM SNAPSHOT") {
                # CSV for created snapshots
                $csvContent = "VMName,SnapshotName,Description,Created,SizeGB,Memory,Status,Message`n"
                foreach ($snapshot in $actionData.createdSnapshots) {
                    $memoryStr = if ($snapshot.Memory) { "Yes" } else { "No" }
                    $csvContent += "`"$($snapshot.VMName)`",`"$($snapshot.SnapshotName)`",`"$($snapshot.Description)`",`"$($snapshot.Created)`",$($snapshot.SizeGB),`"$memoryStr`",`"$($snapshot.Status)`",`"$($snapshot.Message)`"`n"
                }
                foreach ($failed in $actionData.failedSnapshots) {
                    $memoryStr = if ($failed.Memory) { "Yes" } else { "No" }
                    $csvContent += "`"$($failed.VMName)`",`"$($failed.SnapshotName)`",`"$($failed.Description)`",`"`",$memoryStr,`"$($failed.Status)`",`"$($failed.Message)`"`n"
                }
                $csvContent = $csvContent.TrimEnd("`n")
            }
            
            if ($exportToFile) {
                # Generate filename with timestamp
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $filename = "vmware_vm_$($selectedAction.ToLower())_$timestamp.csv"
                
                # Write CSV to file
                $csvContent | Out-File -FilePath $filename -Encoding UTF8 -NoNewline
                Write-Host "CSV file created: $filename"
                
                # Output CSV data AND file reference
                Write-Output-JSON @{
                    xy = 1
                    data = $csvContent
                    files = @(
                        @{
                            path = $filename
                            name = $filename
                        }
                    )
                }
            } else {
                # Output CSV data only
                Write-Output-JSON @{
                    xy = 1
                    data = $csvContent
                }
            }
        }
        
        "MD" {
            # Build Markdown output
            $markdownContent = "# VMWare VM Operations Report`n`n"
            $markdownContent += "## Operation Details`n`n"
            $markdownContent += "| Property | Value |`n"
            $markdownContent += "|----------|-------|`n"
            $markdownContent += "| **Action** | $($actionData.action) |`n"
            $markdownContent += "| **vCenter Server** | $($actionData.vcenterServer) |`n"
            $markdownContent += "| **Timestamp** | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') |`n"
            $markdownContent += "`n"
            
            if ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE DATE") {
        # Display snapshot removal results
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Target Date**: $($actionData.targetDate)`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
            if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                $markdownContent += "## No Snapshots Found`n`n"
                $markdownContent += "No snapshots were found before the target date $($actionData.targetDate).`n`n"
            }
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS") {
        # Display snapshot removal results for weeks-based removal
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Number of Weeks**: $($actionData.numberOfWeeks) week(s)`n`n"
        $markdownContent += "**Target Date**: $($actionData.targetDate) (snapshots older than this are removed)`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.AgeDays) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.AgeDays) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
            if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                $markdownContent += "## No Snapshots Found`n`n"
                $markdownContent += "No snapshots older than $($actionData.numberOfWeeks) week(s) were found.`n`n"
            }
            } elseif ($selectedAction.ToUpper() -eq "REMOVE ALL SNAPSHOTS") {
        # Display snapshot removal results for all snapshots
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Action**: Remove ALL Snapshots`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.AgeDays) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.AgeDays) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
            if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                $markdownContent += "## No Snapshots Found`n`n"
                $markdownContent += "No snapshots were found in the vCenter environment.`n`n"
            }
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT") {
        # Display snapshot removal results for text-based search
        $markdownContent += "## Snapshot Removal Summary`n`n"
        $markdownContent += "**Search Text**: $($actionData.searchText)`n`n"
        $markdownContent += "**Removed**: $($actionData.removedCount) | **Failed**: $($actionData.failedCount) | **Skipped**: $($actionData.skippedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.removedSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Removed Snapshots ($($actionData.removedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|`n"
            
            foreach ($snapshot in $actionData.removedSnapshots) {
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Created) | $($snapshot.AgeDays) | $($snapshot.SizeGB) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.skippedSnapshots.Count -gt 0) {
            $markdownContent += "## Skipped Snapshots ($($actionData.skippedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Reason |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|--------|`n"
            
            foreach ($skipped in $actionData.skippedSnapshots) {
                $markdownContent += "| **$($skipped.VMName)** | $($skipped.SnapshotName) | $($skipped.Created) | $($skipped.AgeDays) | $($skipped.SizeGB) | $($skipped.Reason) |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedRemovals.Count -gt 0) {
            $markdownContent += "## Failed Removals ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Created | Age (Days) | Size (GB) | Error Message |`n"
            $markdownContent += "|---------|---------------|---------|------------|-----------|---------------|`n"
            
            foreach ($failed in $actionData.failedRemovals) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Created) | $($failed.AgeDays) | $($failed.SizeGB) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
            if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                $markdownContent += "## No Snapshots Found`n`n"
                $markdownContent += "No snapshots containing the text '$($actionData.searchText)' were found.`n`n"
            }
            } elseif ($selectedAction.ToUpper() -eq "CREATE VM SNAPSHOT") {
        # Display snapshot creation results
        $markdownContent += "## Snapshot Creation Summary`n`n"
        $markdownContent += "**Snapshot Name**: $($actionData.snapshotName)`n`n"
        $markdownContent += "**Description**: $($actionData.description)`n`n"
        $markdownContent += "**Memory Snapshot**: $($actionData.memory)`n`n"
        $markdownContent += "**Unique Identifier**: $($actionData.uniqueIdentifier)`n`n"
        $markdownContent += "**Created**: $($actionData.createdCount) | **Failed**: $($actionData.failedCount) | **Total Processed**: $($actionData.totalProcessed)`n`n"
        
        if ($actionData.createdSnapshots.Count -gt 0) {
            $markdownContent += "## Successfully Created Snapshots ($($actionData.createdCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Description | Created | Size (GB) | Memory |`n"
            $markdownContent += "|---------|---------------|-------------|---------|-----------|--------|`n"
            
            foreach ($snapshot in $actionData.createdSnapshots) {
                $memoryIcon = if ($snapshot.Memory) { "✓" } else { "✗" }
                $markdownContent += "| **$($snapshot.VMName)** | $($snapshot.SnapshotName) | $($snapshot.Description) | $($snapshot.Created) | $($snapshot.SizeGB) | $memoryIcon |`n"
            }
            $markdownContent += "`n"
        }
        
        if ($actionData.failedSnapshots.Count -gt 0) {
            $markdownContent += "## Failed Snapshot Creations ($($actionData.failedCount))`n`n"
            $markdownContent += "| VM Name | Snapshot Name | Error Message |`n"
            $markdownContent += "|---------|---------------|---------------|`n"
            
            foreach ($failed in $actionData.failedSnapshots) {
                $markdownContent += "| **$($failed.VMName)** | $($failed.SnapshotName) | $($failed.Message) |`n"
            }
            $markdownContent += "`n"
        }
        
            if ($actionData.createdCount -eq 0 -and $actionData.failedCount -gt 0) {
                $markdownContent += "## No Snapshots Created`n`n"
                $markdownContent += "All snapshot creation attempts failed. Please check the error messages above.`n`n"
            }
            } elseif ($selectedAction.ToUpper() -eq "LIST VMS") {
                $markdownContent += "## Virtual Machines ($($actionData.vmCount))`n`n"
                $markdownContent += "| Name | Power State | IP Addresses | Uptime | CPUs | Memory (GB) | Used Space (GB) | Guest OS | Host |`n"
                $markdownContent += "|------|-------------|--------------|--------|------|-------------|-----------------|----------|------|`n"
                
                foreach ($vm in $actionData.vms) {
                    # Format IP addresses with line breaks for Markdown
                    $ipFormatted = if ($vm.IPAddresses -ne "N/A" -and $vm.IPAddresses -match ",") {
                        ($vm.IPAddresses -split ", ") -join "<br>"
                    } else {
                        $vm.IPAddresses
                    }
                    $markdownContent += "| **$($vm.Name)** | $($vm.PowerState) | $ipFormatted | $($vm.Uptime) | $($vm.CPUs) | $($vm.MemoryGB) | $($vm.UsedSpaceGB) | $($vm.GuestOS) | $($vm.VMHost) |`n"
                }
            }
            
            if ($exportToFile) {
                # Generate filename with timestamp
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $filename = "vmware_vm_$($selectedAction.ToLower())_$timestamp.md"
                
                # Write Markdown to file
                $markdownContent | Out-File -FilePath $filename -Encoding UTF8 -NoNewline
                Write-Host "Markdown file created: $filename"
                
                # Output markdown content directly as data AND file reference
                $mdData = @{
                    xy = 1
                    data = $markdownContent
                    files = @(
                        @{
                            path = $filename
                            name = $filename
                        }
                    )
                }
                Write-Output-JSON $mdData
            } else {
                # Output markdown directly (no file)
                $mdData = @{
                    xy = 1
                    data = $markdownContent
                }
                Write-Output-JSON $mdData
            }
        }
        
        "HTML" {
            # Build HTML output
            $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>VMWare VM Operations Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #333; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #0078d4; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; background-color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background-color: #0078d4; color: white; padding: 12px; text-align: left; font-weight: bold; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f0f0f0; }
        .info-table { width: auto; max-width: 500px; }
        .powered-off { background-color: #ffebee; }
        .status { font-weight: bold; }
        .ip-list { line-height: 1.6; }
    </style>
</head>
<body>
    <h1>VMWare VM Operations Report</h1>
    
    <h2>Operation Details</h2>
    <table class="info-table">
        <tr><th>Property</th><th>Value</th></tr>
        <tr><td><strong>Action</strong></td><td>$($actionData.action)</td></tr>
        <tr><td><strong>vCenter Server</strong></td><td>$($actionData.vcenterServer)</td></tr>
        <tr><td><strong>Timestamp</strong></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
"@
            
            # Add action-specific info to table
            if ($selectedAction.ToUpper() -eq "LIST VMS") {
                $htmlContent += "        <tr><td><strong>Total VMs</strong></td><td>$($actionData.vmCount)</td></tr>`n"
            } elseif ($selectedAction.ToUpper() -eq "LIST SNAPSHOTS") {
                $htmlContent += "        <tr><td><strong>Total Snapshots</strong></td><td>$($actionData.snapshotCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>VMs with Snapshots</strong></td><td>$($actionData.vmWithSnapshotsCount)</td></tr>`n"
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE DATE") {
                $htmlContent += "        <tr><td><strong>Target Date</strong></td><td>$($actionData.targetDate)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Removed</strong></td><td>$($actionData.removedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Failed</strong></td><td>$($actionData.failedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Total Processed</strong></td><td>$($actionData.totalProcessed)</td></tr>`n"
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS") {
                $htmlContent += "        <tr><td><strong>Number of Weeks</strong></td><td>$($actionData.numberOfWeeks)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Target Date</strong></td><td>$($actionData.targetDate)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Removed</strong></td><td>$($actionData.removedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Failed</strong></td><td>$($actionData.failedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Total Processed</strong></td><td>$($actionData.totalProcessed)</td></tr>`n"
            } elseif ($selectedAction.ToUpper() -eq "REMOVE ALL SNAPSHOTS") {
                $htmlContent += "        <tr><td><strong>Removed</strong></td><td>$($actionData.removedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Failed</strong></td><td>$($actionData.failedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Total Processed</strong></td><td>$($actionData.totalProcessed)</td></tr>`n"
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT") {
                $htmlContent += "        <tr><td><strong>Search Text</strong></td><td>$($actionData.searchText)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Removed</strong></td><td>$($actionData.removedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Failed</strong></td><td>$($actionData.failedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Skipped</strong></td><td>$($actionData.skippedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Total Processed</strong></td><td>$($actionData.totalProcessed)</td></tr>`n"
            } elseif ($selectedAction.ToUpper() -eq "CREATE VM SNAPSHOT") {
                $htmlContent += "        <tr><td><strong>Snapshot Name</strong></td><td>$($actionData.snapshotName)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Description</strong></td><td>$($actionData.description)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Memory Snapshot</strong></td><td>$($actionData.memory)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Unique Identifier</strong></td><td>$($actionData.uniqueIdentifier)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Created</strong></td><td>$($actionData.createdCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Failed</strong></td><td>$($actionData.failedCount)</td></tr>`n"
                $htmlContent += "        <tr><td><strong>Total Processed</strong></td><td>$($actionData.totalProcessed)</td></tr>`n"
            }
            
            $htmlContent += "    </table>`n"
            
            if ($selectedAction.ToUpper() -eq "LIST VMS") {
                # Separate VMs by power state
                $poweredOnVMs = $actionData.vms | Where-Object { $_.PowerState -eq "PoweredOn" }
                $poweredOffVMs = $actionData.vms | Where-Object { $_.PowerState -eq "PoweredOff" }
                
                # PoweredOn VMs table
                if ($poweredOnVMs.Count -gt 0) {
                    $htmlContent += "`n    <h2>Powered On Virtual Machines ($($poweredOnVMs.Count))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>Name</th><th>IP Addresses</th><th>Uptime</th><th>CPUs</th><th>Memory (GB)</th><th>Used Space (GB)</th><th>Guest OS</th><th>Host</th></tr>`n"
                    
                    foreach ($vm in $poweredOnVMs) {
                        # Format IP addresses with line breaks for HTML
                        $ipFormatted = if ($vm.IPAddresses -ne "N/A" -and $vm.IPAddresses -match ",") {
                            ($vm.IPAddresses -split ", ") -join "<br>"
                        } else {
                            $vm.IPAddresses
                        }
                        $htmlContent += "        <tr><td><strong>$($vm.Name)</strong></td><td class='ip-list'>$ipFormatted</td><td>$($vm.Uptime)</td><td>$($vm.CPUs)</td><td>$($vm.MemoryGB)</td><td>$($vm.UsedSpaceGB)</td><td>$($vm.GuestOS)</td><td>$($vm.VMHost)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # PoweredOff VMs table
                if ($poweredOffVMs.Count -gt 0) {
                    $htmlContent += "`n    <h2>Powered Off Virtual Machines ($($poweredOffVMs.Count))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>Name</th><th>CPUs</th><th>Memory (GB)</th><th>Used Space (GB)</th><th>Guest OS</th><th>Host</th></tr>`n"
                    
                    foreach ($vm in $poweredOffVMs) {
                        $htmlContent += "        <tr class='powered-off'><td><strong>$($vm.Name)</strong></td><td>$($vm.CPUs)</td><td>$($vm.MemoryGB)</td><td>$($vm.UsedSpaceGB)</td><td>$($vm.GuestOS)</td><td>$($vm.VMHost)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
            } elseif ($selectedAction.ToUpper() -eq "LIST SNAPSHOTS") {
                # List snapshots table
                if ($actionData.snapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>VM Snapshots ($($actionData.snapshotCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Size (GB)</th><th>Power State</th><th>Description</th></tr>`n"
                    
                    foreach ($snapshot in $actionData.snapshots) {
                        $desc = if ([string]::IsNullOrWhiteSpace($snapshot.Description)) { "-" } else { $snapshot.Description }
                        $htmlContent += "        <tr><td><strong>$($snapshot.VMName)</strong></td><td>$($snapshot.SnapshotName)</td><td>$($snapshot.Created)</td><td>$($snapshot.SizeGB)</td><td>$($snapshot.PowerState)</td><td>$desc</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                } else {
                    $htmlContent += "`n    <h2>No Snapshots Found</h2>`n"
                    $htmlContent += "    <p>There are no snapshots in the vCenter environment.</p>`n"
                }
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE DATE") {
                # Removed snapshots table
                if ($actionData.removedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Successfully Removed Snapshots ($($actionData.removedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Size (GB)</th></tr>`n"
                    
                    foreach ($snapshot in $actionData.removedSnapshots) {
                        $htmlContent += "        <tr><td><strong>$($snapshot.VMName)</strong></td><td>$($snapshot.SnapshotName)</td><td>$($snapshot.Created)</td><td>$($snapshot.SizeGB)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Skipped snapshots table
                if ($actionData.skippedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Skipped Snapshots ($($actionData.skippedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Size (GB)</th><th>Reason</th></tr>`n"
                    
                    foreach ($skipped in $actionData.skippedSnapshots) {
                        $htmlContent += "        <tr style='background-color: #ffffcc;'><td><strong>$($skipped.VMName)</strong></td><td>$($skipped.SnapshotName)</td><td>$($skipped.Created)</td><td>$($skipped.SizeGB)</td><td>$($skipped.Reason)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Failed removals table
                if ($actionData.failedRemovals.Count -gt 0) {
                    $htmlContent += "`n    <h2>Failed Removals ($($actionData.failedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Size (GB)</th><th>Error Message</th></tr>`n"
                    
                    foreach ($failed in $actionData.failedRemovals) {
                        $htmlContent += "        <tr style='background-color: #ffcccc;'><td><strong>$($failed.VMName)</strong></td><td>$($failed.SnapshotName)</td><td>$($failed.Created)</td><td>$($failed.SizeGB)</td><td>$($failed.Message)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                    $htmlContent += "`n    <h2>No Snapshots Found</h2>`n"
                    $htmlContent += "    <p>No snapshots were found before the target date $($actionData.targetDate).</p>`n"
                }
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS BEFORE NUMBER OF WEEKS") {
                # Removed snapshots table (weeks-based)
                if ($actionData.removedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Successfully Removed Snapshots ($($actionData.removedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th></tr>`n"
                    
                    foreach ($snapshot in $actionData.removedSnapshots) {
                        $htmlContent += "        <tr><td><strong>$($snapshot.VMName)</strong></td><td>$($snapshot.SnapshotName)</td><td>$($snapshot.Created)</td><td>$($snapshot.AgeDays)</td><td>$($snapshot.SizeGB)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Skipped snapshots table
                if ($actionData.skippedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Skipped Snapshots ($($actionData.skippedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th><th>Reason</th></tr>`n"
                    
                    foreach ($skipped in $actionData.skippedSnapshots) {
                        $htmlContent += "        <tr style='background-color: #ffffcc;'><td><strong>$($skipped.VMName)</strong></td><td>$($skipped.SnapshotName)</td><td>$($skipped.Created)</td><td>$($skipped.AgeDays)</td><td>$($skipped.SizeGB)</td><td>$($skipped.Reason)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Failed removals table
                if ($actionData.failedRemovals.Count -gt 0) {
                    $htmlContent += "`n    <h2>Failed Removals ($($actionData.failedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th><th>Error Message</th></tr>`n"
                    
                    foreach ($failed in $actionData.failedRemovals) {
                        $htmlContent += "        <tr style='background-color: #ffcccc;'><td><strong>$($failed.VMName)</strong></td><td>$($failed.SnapshotName)</td><td>$($failed.Created)</td><td>$($failed.AgeDays)</td><td>$($failed.SizeGB)</td><td>$($failed.Message)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                    $htmlContent += "`n    <h2>No Snapshots Found</h2>`n"
                    $htmlContent += "    <p>No snapshots older than $($actionData.numberOfWeeks) week(s) were found.</p>`n"
                }
            } elseif ($selectedAction.ToUpper() -eq "REMOVE ALL SNAPSHOTS") {
                # Removed all snapshots tables
                if ($actionData.removedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Successfully Removed Snapshots ($($actionData.removedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th></tr>`n"
                    
                    foreach ($snapshot in $actionData.removedSnapshots) {
                        $htmlContent += "        <tr><td><strong>$($snapshot.VMName)</strong></td><td>$($snapshot.SnapshotName)</td><td>$($snapshot.Created)</td><td>$($snapshot.AgeDays)</td><td>$($snapshot.SizeGB)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Skipped snapshots table
                if ($actionData.skippedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Skipped Snapshots ($($actionData.skippedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th><th>Reason</th></tr>`n"
                    
                    foreach ($skipped in $actionData.skippedSnapshots) {
                        $htmlContent += "        <tr style='background-color: #ffffcc;'><td><strong>$($skipped.VMName)</strong></td><td>$($skipped.SnapshotName)</td><td>$($skipped.Created)</td><td>$($skipped.AgeDays)</td><td>$($skipped.SizeGB)</td><td>$($skipped.Reason)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Failed removals table
                if ($actionData.failedRemovals.Count -gt 0) {
                    $htmlContent += "`n    <h2>Failed Removals ($($actionData.failedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th><th>Error Message</th></tr>`n"
                    
                    foreach ($failed in $actionData.failedRemovals) {
                        $htmlContent += "        <tr style='background-color: #ffcccc;'><td><strong>$($failed.VMName)</strong></td><td>$($failed.SnapshotName)</td><td>$($failed.Created)</td><td>$($failed.AgeDays)</td><td>$($failed.SizeGB)</td><td>$($failed.Message)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                    $htmlContent += "`n    <h2>No Snapshots Found</h2>`n"
                    $htmlContent += "    <p>No snapshots were found in the vCenter environment.</p>`n"
                }
            } elseif ($selectedAction.ToUpper() -eq "REMOVE SNAPSHOTS CONTAINING SPECIFIC TEXT") {
                # Removed snapshots containing text tables
                if ($actionData.removedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Successfully Removed Snapshots ($($actionData.removedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th></tr>`n"
                    
                    foreach ($snapshot in $actionData.removedSnapshots) {
                        $htmlContent += "        <tr><td><strong>$($snapshot.VMName)</strong></td><td>$($snapshot.SnapshotName)</td><td>$($snapshot.Created)</td><td>$($snapshot.AgeDays)</td><td>$($snapshot.SizeGB)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Skipped snapshots table
                if ($actionData.skippedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Skipped Snapshots ($($actionData.skippedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th><th>Reason</th></tr>`n"
                    
                    foreach ($skipped in $actionData.skippedSnapshots) {
                        $htmlContent += "        <tr style='background-color: #ffffcc;'><td><strong>$($skipped.VMName)</strong></td><td>$($skipped.SnapshotName)</td><td>$($skipped.Created)</td><td>$($skipped.AgeDays)</td><td>$($skipped.SizeGB)</td><td>$($skipped.Reason)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Failed removals table
                if ($actionData.failedRemovals.Count -gt 0) {
                    $htmlContent += "`n    <h2>Failed Removals ($($actionData.failedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Created</th><th>Age (Days)</th><th>Size (GB)</th><th>Error Message</th></tr>`n"
                    
                    foreach ($failed in $actionData.failedRemovals) {
                        $htmlContent += "        <tr style='background-color: #ffcccc;'><td><strong>$($failed.VMName)</strong></td><td>$($failed.SnapshotName)</td><td>$($failed.Created)</td><td>$($failed.AgeDays)</td><td>$($failed.SizeGB)</td><td>$($failed.Message)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                if ($actionData.removedCount -eq 0 -and $actionData.failedCount -eq 0) {
                    $htmlContent += "`n    <h2>No Snapshots Found</h2>`n"
                    $htmlContent += "    <p>No snapshots containing the text '$($actionData.searchText)' were found.</p>`n"
                }
            } elseif ($selectedAction.ToUpper() -eq "CREATE VM SNAPSHOT") {
                # Created snapshots table
                if ($actionData.createdSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Successfully Created Snapshots ($($actionData.createdCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Description</th><th>Created</th><th>Size (GB)</th><th>Memory</th></tr>`n"
                    
                    foreach ($snapshot in $actionData.createdSnapshots) {
                        $memoryIcon = if ($snapshot.Memory) { "✓" } else { "✗" }
                        $htmlContent += "        <tr><td><strong>$($snapshot.VMName)</strong></td><td>$($snapshot.SnapshotName)</td><td>$($snapshot.Description)</td><td>$($snapshot.Created)</td><td>$($snapshot.SizeGB)</td><td>$memoryIcon</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                # Failed snapshot creations table
                if ($actionData.failedSnapshots.Count -gt 0) {
                    $htmlContent += "`n    <h2>Failed Snapshot Creations ($($actionData.failedCount))</h2>`n"
                    $htmlContent += "    <table>`n"
                    $htmlContent += "        <tr><th>VM Name</th><th>Snapshot Name</th><th>Error Message</th></tr>`n"
                    
                    foreach ($failed in $actionData.failedSnapshots) {
                        $htmlContent += "        <tr style='background-color: #ffcccc;'><td><strong>$($failed.VMName)</strong></td><td>$($failed.SnapshotName)</td><td>$($failed.Message)</td></tr>`n"
                    }
                    
                    $htmlContent += "    </table>`n"
                }
                
                if ($actionData.createdCount -eq 0 -and $actionData.failedCount -gt 0) {
                    $htmlContent += "`n    <h2>No Snapshots Created</h2>`n"
                    $htmlContent += "    <p>All snapshot creation attempts failed. Please check the error messages above.</p>`n"
                }
            }
            
            $htmlContent += @"
</body>
</html>
"@
            
            if ($exportToFile) {
                # Generate filename with timestamp
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $filename = "vmware_vm_$($selectedAction.ToLower())_$timestamp.html"
                
                # Write HTML to file
                $htmlContent | Out-File -FilePath $filename -Encoding UTF8 -NoNewline
                Write-Host "HTML file created: $filename"
                
                # Output HTML content as data AND file reference
                $htmlData = @{
                    xy = 1
                    data = $htmlContent
                    files = @(
                        @{
                            path = $filename
                            name = $filename
                        }
                    )
                }
                Write-Output-JSON $htmlData
            } else {
                # Output HTML directly (no file)
                $htmlData = @{
                    xy = 1
                    data = $htmlContent
                }
                Write-Output-JSON $htmlData
            }
        }
        
        default {
            Write-Host "Unknown export format: $exportFormat, defaulting to JSON"
            $jsonData = @{
                xy = 1
                data = $actionData
            }
            Write-Output-JSON $jsonData
        }
    }

    $outputSW.Stop(); $perf.output = [math]::Round($outputSW.Elapsed.TotalSeconds,3); if ($debug) { Write-Host ("PERF output={0:n3}s" -f $perf.output) }
    
    # Disconnect from vCenter
    Disconnect-VIServer -Server $viConnection -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Disconnected from vCenter"

    # Output performance metrics for xyOps pie chart
    $overallSW.Stop(); $perf.t = [math]::Round($overallSW.Elapsed.TotalSeconds,3)
    Write-Output-JSON @{ xy = 1; perf = $perf }
    if ($debug) {
        Write-Host ("PERF total={0:n3}s" -f $perf.t)
    }
    
    # Success message
    $summary = "Operation '$selectedAction' completed successfully"
    if ($selectedAction.ToUpper() -eq "LISTVMS") {
        $summary += ": $($actionData.vmCount) VM(s) retrieved"
    }
    
    Send-Success -Description $summary
}
catch {
    Send-Error -Code 7 -Description "Error: $($_.Exception.Message)"
    exit 1
}
