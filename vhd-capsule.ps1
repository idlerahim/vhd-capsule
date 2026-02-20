<#
.SYNOPSIS
    Robust VHD/VHDX Manager and Capsule Launcher.
.DESCRIPTION
    Manages the lifecycle of Virtual Hard Disks (VHD/VHDX) and provides a specialized "Capsule Mode" 
    for isolating applications with tracking of filesystem changes.
    
    Features:
    - Create VHD/VHDX (Fixed/Dynamic)
    - Browse and Select VHDs
    - Operations: Mount, Compact, Defrag, Clean Junk
    - Capsule Mode: Mount -> Snapshot -> Execute -> Diff -> Maintenance
    
.PARAMETER VHDPath
    (Optional) Path to a VHD/VHDX file to pre-select or launch.
.PARAMETER Mode
    (Optional) Start directly in specific mode: 'Manager', 'Capsule'. Default is 'Menu'.
.PARAMETER AppPath
    (Optional) For Capsule Mode: The relative path to the executable inside the VHD.
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$VHDPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Menu", "Manager", "Capsule")]
    [string]$Mode = "Menu",

    [Parameter(Mandatory = $false)]
    [string]$AppPath,

    [Parameter(Mandatory = $false)]
    [string]$InitialDir,

    [Parameter(Mandatory = $false)]
    [string]$SourceFolder,

    [Parameter(Mandatory = $false)]
    [string]$DestinationPath,

    [Parameter(Mandatory = $false)]
    [string]$SizeGB,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# -------------------------------------------------------------------------
# 0. INITIALIZATION & SAFETY
# -------------------------------------------------------------------------

if ($InitialDir -and (Test-Path $InitialDir)) {
    Set-Location $InitialDir
}

# Formatting for consistency
$Host.UI.RawUI.WindowTitle = "VHD Capsule Manager"
$ErrorActionPreference = "Stop"

function Assert-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "[INFO] Requesting Administrator privileges..." -ForegroundColor Yellow
        $params = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        if ($VHDPath) { $params += " -VHDPath `"$VHDPath`"" }
        if ($Mode -ne "Menu") { $params += " -Mode $Mode" }
        if ($AppPath) { $params += " -AppPath `"$AppPath`"" }
        if ($SourceFolder) { $params += " -SourceFolder `"$SourceFolder`"" }
        if ($DestinationPath) { $params += " -DestinationPath `"$DestinationPath`"" }
        if ($SizeGB) { $params += " -SizeGB $SizeGB" }
        if ($Force) { $params += " -Force" }
        $params += " -InitialDir `"$($pwd.Path)`""
        
        Start-Process powershell.exe $params -Verb RunAs
        exit
    }
}
Assert-Admin

# -------------------------------------------------------------------------
# 1. HELPER FUNCTIONS
# -------------------------------------------------------------------------

function Show-Header {
    param([string]$Title)
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   VHD CAPSULE: $Title" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-UserChoice {
    param([string]$Prompt, [int]$Max, [scriptblock]$UIBlock, [int]$Default = 0)
    
    # Initial Draw
    if ($UIBlock) { & $UIBlock }

    while ($true) {
        $inputVal = Read-Host $Prompt
        
        if ([string]::IsNullOrWhiteSpace($inputVal) -and $Default -gt 0) {
            return $Default
        }

        if ($inputVal -match "^\d+$" -and [int]$inputVal -ge 0 -and [int]$inputVal -le $Max) {
            return [int]$inputVal
        }
        
        # Invalid Input: Refresh
        if ($UIBlock) {
            & $UIBlock
            Write-Host "Invalid selection. Please enter a number between 0 and $Max." -ForegroundColor Red
        }
        else {
            Write-Host "Invalid selection." -ForegroundColor Red
        }
    }
}

function Get-VHDXPhysicalSize {
    param([string]$Path)
    if (Test-Path $Path) {
        $file = Get-Item $Path
        return [Math]::Round($file.Length / 1MB, 2)
    }
    return 0
}

function Get-FreeDriveLetter {
    $used = (Get-PSDrive -PSProvider FileSystem).Name
    foreach ($code in 68..90) {
        # D to Z
        $letter = [char]$code
        if ($letter -notin $used) { return "${letter}:" }
    }
    throw "No free drive letters available."
}

# Wrapper to handle DiskPart scripting
function Invoke-DiskPartScript {
    param([string]$ScriptContent)
    $dpScript = Join-Path $env:TEMP "vhd_manager_script.txt"
    $ScriptContent | Out-File -FilePath $dpScript -Encoding ASCII
    
    try {
        $output = diskpart /s $dpScript
        return $output
    }
    finally {
        if (Test-Path $dpScript) { Remove-Item $dpScript }
    }
}

# Robust Mounting Logic
function Mount-VHDNative {
    param([string]$Path)
    
    Write-Host "Mounting $Path..." -ForegroundColor Yellow
    
    try {
        $mount = Mount-DiskImage -ImagePath $Path -PassThru -ErrorAction Stop
        
        # Wait for volume to become available
        Start-Sleep -Seconds 2
        
        # Check if a drive letter was auto-assigned
        $volume = $mount | Get-Disk | Get-Partition | Where-Object { $_.DriveLetter } | Select-Object -First 1
        if ($volume) {
            return "$($volume.DriveLetter):"
        }
        
        # No letter auto-assigned — pick a free one and assign explicitly via diskpart
        Write-Host "No drive letter found. Attempting to assign..." -ForegroundColor Yellow
        $freeLetter = Get-FreeDriveLetter
        $script = @"
select vdisk file="$Path"
select partition 1
assign letter=$($freeLetter[0])
"@
        Invoke-DiskPartScript -ScriptContent $script | Out-Null
        
        # Verify assignment
        Start-Sleep -Seconds 1
        $mount = Get-DiskImage -ImagePath $Path
        $volume = $mount | Get-Disk | Get-Partition | Where-Object { $_.DriveLetter } | Select-Object -First 1
        if ($volume) { return "$($volume.DriveLetter):" }
        
        # Fallback: return the letter we requested
        if (Test-Path $freeLetter) { return $freeLetter }
    }
    catch {
        Write-Host "Mount failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    return $null
}

function Dismount-VHDNative {
    param([string]$Path)
    Write-Host "Dismounting $Path..." -ForegroundColor Yellow
    Dismount-DiskImage -ImagePath $Path -ErrorAction SilentlyContinue | Out-Null
}

# -------------------------------------------------------------------------
# 2. CORE LOGIC BLOCKS
# -------------------------------------------------------------------------

function New-VHDItem {
    Show-Header "Create New VHD/VHDX: Step 1/7"
    
    # 0. Format Selection
    Write-Host "Virtual Hard Disk Type:"
    Write-Host "1. VHDX [Default]"
    Write-Host "2. VHD"
    Write-Host "`n0. Cancel"
    $fmtChoice = Read-Host "`nSelect Type"
    if ($fmtChoice -eq "0") { return }
    $extension = if ($fmtChoice -eq "2") { ".vhd" } else { ".vhdx" }
    
    # 1. Filename
    Show-Header "Create New VHD/VHDX: Step 2/7"
    Write-Host "Enter Filename (e.g. MyApp$extension)"
    Write-Host "`n0. Cancel"
    $name = Read-Host "`nFilename"
    if ($name -eq "0") { return }
    
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    if ($name -match "\.(vhd|vhdx)$") { $name = $name -replace "\.(vhd|vhdx)$", "" }
    $name += $extension
    $targetPath = Join-Path $pwd $name

    # 2. Size
    Show-Header "Create New VHD/VHDX: Step 3/7"
    Write-Host "Enter Size in GB"
    Write-Host "`n0. Cancel"
    $sizeGB = Read-Host "`nSize"
    if ($sizeGB -eq "0") { return }
    
    if ($sizeGB -notmatch "^\d+$") { $sizeGB = 10 }
    $sizeMB = [int]$sizeGB * 1024

    # 3. Type
    Show-Header "Create New VHD/VHDX: Step 4/7"
    Write-Host "Virtual Disk Type:"
    Write-Host "1. Dynamic (Saves Space) [Default]"
    Write-Host "2. Fixed (Better Performance)"
    Write-Host "`n0. Cancel"
    $typeChoice = Read-Host "`nSelect (1/2)"
    if ($typeChoice -eq "0") { return }
    $type = if ($typeChoice -eq "2") { "fixed" } else { "expandable" }

    # 4. Partition Style
    Show-Header "Create New VHD/VHDX: Step 5/7"
    Write-Host "Initialize Disk as:"
    Write-Host "1. GPT (GUID Partition Table) [Default]"
    Write-Host "2. MBR (Master Boot Record)"
    Write-Host "`n0. Cancel"
    $partChoice = Read-Host "`nSelect Style (1/2)"
    if ($partChoice -eq "0") { return }
    $partStyle = if ($partChoice -eq "2") { "mbr" } else { "gpt" }

    # 5. File System
    Show-Header "Create New VHD/VHDX: Step 6/7"
    Write-Host "Format Disk as:"
    Write-Host "1. NTFS [Default]"
    Write-Host "2. FAT32"
    Write-Host "`n0. Cancel"
    $fsChoice = Read-Host "`nSelect File System (1/2)"
    if ($fsChoice -eq "0") { return }
    $fs = if ($fsChoice -eq "2") { "fat32" } else { "ntfs" }
    
    $enableCompression = $false
    if ($fs -eq "ntfs") {
        Show-Header "Create New VHD/VHDX: Step 6/7 (Compression)"
        Write-Host "Enable file and folder compression:"
        Write-Host "1. No [Default]"
        Write-Host "2. Yes"
        Write-Host "`n0. Cancel"
        $compChoice = Read-Host "`nSelect (1/2)"
        if ($compChoice -eq "0") { return }
        if ($compChoice -eq "2") { $enableCompression = $true }
    }

    # 6. Allocation Unit
    Show-Header "Create New VHD/VHDX: Step 7/7"
    Write-Host "Allocation Unit Size:"
    $allocUnits = @(
        @{ L = "Default 4K (General Windows usage)"; V = "default" },
        @{ L = "512 B (Legacy systems)"; V = "512" },
        @{ L = "1 KB (Embedded or special workloads)"; V = "1024" },
        @{ L = "2 KB (Niche workloads)"; V = "2048" },
        @{ L = "4 KB (OS drives, apps, mixed data, VHDX)"; V = "4096" },
        @{ L = "8 KB (Moderate archive workloads)"; V = "8192" },
        @{ L = "16 KB (ISOs, VHDX storage, video archives)"; V = "16K" },
        @{ L = "32 KB (Media servers, large backups)"; V = "32K" },
        @{ L = "64 KB (Large databases, VM storage)"; V = "64K" },
        @{ L = "128 KB (Enterprise workloads)"; V = "128K" },
        @{ L = "256 KB (Specialized storage appliances)"; V = "256K" },
        @{ L = "512 KB (Large streaming storage)"; V = "512K" },
        @{ L = "1 MB (Backup volumes)"; V = "1024K" },
        @{ L = "2 MB (High performance storage arrays)"; V = "2048K" }
    )
    for ($i = 0; $i -lt $allocUnits.Count; $i++) {
        Write-Host "$($i+1). $($allocUnits[$i].L)"
    }
    Write-Host "`n0. Cancel"
    $allocChoice = Read-Host "`nSelect Allocation Unit (Default: 1)"
    if ($allocChoice -eq "0") { return }
    
    if ($allocChoice -notmatch "^\d+$" -or [int]$allocChoice -lt 1 -or [int]$allocChoice -gt $allocUnits.Count) { $allocChoice = 1 }
    $allocUnit = $allocUnits[[int]$allocChoice - 1].V

    # 7. Volume Label & Confirmation
    Show-Header "Create New VHD/VHDX: Confirmation"
    Write-Host "Enter Volume Label (Default: `"New Volume`")"
    Write-Host "`n0. Cancel"
    $label = Read-Host "`nLabel"
    if ($label -eq "0") { return }
    if ([string]::IsNullOrWhiteSpace($label)) { $label = "New Volume" }

    # Summary
    Write-Host "`nReady to create VHD:" -ForegroundColor Cyan
    Write-Host "Path      : $targetPath"
    Write-Host "Size      : $sizeGB GB"
    Write-Host "Type      : $type"
    Write-Host "Style     : $partStyle"
    Write-Host "Format    : $fs (Unit: $allocUnit)"
    Write-Host "Compress  : $enableCompression"
    Write-Host "Label     : $label"
    
    $confirm = Read-Host "`nPress Enter to Create or '0' to Cancel"
    if ($confirm -eq "0") { return }

    # Execute
    Write-Host "`nCreating..." -ForegroundColor Yellow
    
    $script = @"
create vdisk file="$targetPath" maximum=$sizeMB type=$type
select vdisk file="$targetPath"
attach vdisk
convert $partStyle
create partition primary
"@
    
    # Build format command
    $fmtCmd = "format fs=$fs label=`"$label`" quick"
    if ($allocUnit -ne "default") { $fmtCmd += " unit=$allocUnit" }
    if ($enableCompression) { $fmtCmd += " compress" }
    
    $script += "`n$fmtCmd"
    $script += "`nassign`ndetach vdisk"
    
    Invoke-DiskPartScript -ScriptContent $script | Out-Null
    
    Show-Header "Create New VHD/VHDX: Complete"
    if (Test-Path $targetPath) {
        Write-Host "Successfully created: $targetPath" -ForegroundColor Green
    }
    else {
        Write-Host "Failed to create VHD." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to menu"
}

function New-VHDCapsuleFromFolder {
    param(
        [string]$InputSource,
        [string]$InputDest,
        [string]$InputSize
    )

    Show-Header "Create VHD Capsule from Folder: Step 1/4"
    
    # 1. Source Selection
    if (-not [string]::IsNullOrWhiteSpace($InputSource)) {
        $sourcePath = $InputSource
    }
    else {
        $sourcePath = Read-Host "Enter Source Folder Path (e.g. C:\Apps\MyApp)"
        if ([string]::IsNullOrWhiteSpace($sourcePath)) { return }
    }
    
    $sourcePath = $sourcePath.Trim('"').TrimEnd('\')
    if (-not (Test-Path $sourcePath -PathType Container)) {
        Write-Host "Invalid folder path: $sourcePath" -ForegroundColor Red
        if ($InputSource) { Read-Host "Press Enter to Exit"; exit }
        Read-Host "Press Enter"
        return
    }
    $sourceName = Split-Path $sourcePath -Leaf

    # 2. Destination Selection
    Show-Header "Create VHD Capsule from Folder: Step 2/4"
    $defaultDest = Split-Path $sourcePath -Parent
    
    if (-not [string]::IsNullOrWhiteSpace($InputDest)) {
        $destDir = $InputDest
    }
    else {
        Write-Host "Source: $sourcePath"
        Write-Host "Enter Destination Directory for VHDX"
        Write-Host "Default: $defaultDest" -ForegroundColor Gray
        
        if (-not [string]::IsNullOrWhiteSpace($InputSource)) {
            $destDir = $defaultDest
            Write-Host "Destination: $destDir (Default)"
        }
        else {
            $destDir = Read-Host "Destination Path"
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($destDir)) { $destDir = $defaultDest }
    $destDir = $destDir.Trim('"')
    if (-not (Test-Path $destDir)) {
        Write-Host "Invalid destination path." -ForegroundColor Red
        if ($InputDest) { Read-Host "Press Enter to Exit"; exit }
        Read-Host "Press Enter"
        return
    }

    # 3. Size Calculation
    Show-Header "Create VHD Capsule from Folder: Step 3/4"
    Write-Host "Analyzing source folder..." -ForegroundColor Yellow
    $stats = Get-ChildItem -Path $sourcePath -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
    $sourceSizeGB = [Math]::Round($stats.Sum / 1GB, 2)
    $minSizeGB = [Math]::Ceiling($sourceSizeGB + 2)
    $defaultSizeGB = [Math]::Ceiling($sourceSizeGB + 5)
    
    Write-Host "Source Size      : $sourceSizeGB GB"
    Write-Host "Minimum Required : $minSizeGB GB (Source + 2GB)"
    
    if (-not [string]::IsNullOrWhiteSpace($InputSize)) {
        $sizeGB = $InputSize
    }
    else {
        if (-not [string]::IsNullOrWhiteSpace($InputSource)) {
            $sizeGB = $defaultSizeGB
            Write-Host "Size: $sizeGB GB (Default)"
        }
        else {
            $sizeGB = Read-Host "Enter VHDX Size in GB (Default: $defaultSizeGB)"
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($sizeGB)) { $sizeGB = $defaultSizeGB }
    
    # Validate Size
    if ($sizeGB -notmatch "^\d+$") { $sizeGB = $defaultSizeGB }
    if ([int]$sizeGB -lt $minSizeGB) {
        Write-Host "Size too small. Setting to minimum: $minSizeGB GB" -ForegroundColor Yellow
        $sizeGB = $minSizeGB
    }
    $sizeMB = [int]$sizeGB * 1024

    # 4. Configuration & Confirmation
    $config = @{
        Extension = ".vhdx" # Default
        Type      = "expandable" # Dynamic
        PartStyle = "gpt"
        Fs        = "ntfs"
        AllocUnit = "4096"
        Label     = $sourceName
        Compress  = $false
    }
    
    # Loop needs to be handled carefully with UserChoice
    while ($true) {
        $targetName = "$sourceName$($config.Extension)"
        $targetPath = Join-Path $destDir $targetName
        
        $ui = {
            Show-Header "Review VHD Capsule Configuration"
            Write-Host "Source      : $sourcePath"
            Write-Host "Target      : $targetPath"
            Write-Host "Size        : $sizeGB GB"
            Write-Host "Type        : $($config.Type)"
            Write-Host "Style       : $($config.PartStyle)"
            Write-Host "Format      : $($config.Fs) (Unit: $($config.AllocUnit))"
            Write-Host "Compression : $($config.Compress)"
            Write-Host "Label       : $($config.Label)"
            Write-Host "--------------------------------"
            Write-Host "1. Proceed"
            Write-Host "2. Modify Settings"
            Write-Host "3. Cancel"
        }
        
        # PowerShell closures capture variables, so this should work fine locally.
        
        if ($Force) {
            $choice = 1
            Write-Host "Force mode enabled. Proceeding with default settings..." -ForegroundColor Yellow
        }
        else {
            $choice = Get-UserChoice -Prompt "`nSelect Option (Default: 1)" -Max 3 -UIBlock $ui -Default 1
        }
        
        if ($choice -eq 3) { return }
        if ($choice -eq 1) { break }
        
        # Modify Logic
        if ($choice -eq 2) {
            Show-Header "Modify Settings"
            
            # Format
            Write-Host "1. VHDX [Default], 2. VHD"
            if ((Read-Host "Select") -eq "2") { $config.Extension = ".vhd" } else { $config.Extension = ".vhdx" }
            
            # Type
            Write-Host "1. Dynamic [Default], 2. Fixed"
            if ((Read-Host "Select") -eq "2") { $config.Type = "fixed" } else { $config.Type = "expandable" }
            
            # Partition
            Write-Host "1. GPT [Default], 2. MBR"
            if ((Read-Host "Select") -eq "2") { $config.PartStyle = "mbr" } else { $config.PartStyle = "gpt" }
            
            # FS
            Write-Host "1. NTFS [Default], 2. FAT32"
            if ((Read-Host "Select") -eq "2") { 
                $config.Fs = "fat32"; $config.Compress = $false 
            }
            else { 
                $config.Fs = "ntfs"
                Write-Host "Enable Compression? 1. No [Default], 2. Yes"
                if ((Read-Host "Select") -eq "2") { $config.Compress = $true } else { $config.Compress = $false }
            }
            
            # Label
            $lbl = Read-Host "Volume Label (Default: $($config.Label))"
            if (-not [string]::IsNullOrWhiteSpace($lbl)) { $config.Label = $lbl }
        }
    }

    # 5. Execution
    Show-Header "Creating Capsule..."
    Write-Host "Creating VHD..." -ForegroundColor Yellow
    
    # Build diskpart script with format options
    $fmtCmd = "format fs=$($config.Fs) label=`"$($config.Label)`" quick"
    if ($config.AllocUnit -ne "default") { $fmtCmd += " unit=$($config.AllocUnit)" }
    if ($config.Compress) { $fmtCmd += " compress" }

    $script = @"
create vdisk file="$targetPath" maximum=$sizeMB type=$($config.Type)
select vdisk file="$targetPath"
attach vdisk
convert $($config.PartStyle)
create partition primary
$fmtCmd
assign
detach vdisk
"@
    
    Invoke-DiskPartScript -ScriptContent $script | Out-Null
    
    if (-not (Test-Path $targetPath)) {
        Write-Host "Failed to create VHD." -ForegroundColor Red; Read-Host "Press Enter"; return
    }

    # 6. Copying
    Write-Host "Mounting for data transfer..." -ForegroundColor Yellow
    # Mounting can be tricky if quick removal/re-add happens. Wait a bit.
    Start-Sleep 2
    $drive = Mount-VHDNative -Path $targetPath
    if (-not $drive) { Write-Host "Mount failed." -ForegroundColor Red; return }
    try {
        Write-Host "Copying files from source (Robocopy)..." -ForegroundColor Cyan
        Write-Host "Source: $sourcePath" 
        Write-Host "Target: $drive"
        
        $sw = [Diagnostics.Stopwatch]::StartNew()
        
        # Robocopy Argument Handling
        # Robocopy requires explicit quotes around paths with spaces when valid via Start-Process ArgumentList in PS 5.1
        $srcArg = "`"$sourcePath`""
        $dstArg = "`"$drive`""
        
        $argsList = @($srcArg, $dstArg, "/E", "/COPY:DAT", "/J", "/MT:8", "/R:3", "/W:1", "/NFL", "/NDL", "/XJD", "/XJF")
        # Added /XJD /XJF to exclude junction points which can cause loops or access issues
        
        $p = Start-Process -FilePath "robocopy.exe" -ArgumentList $argsList -PassThru -NoNewWindow -Wait
        
        $sw.Stop()
        
        # 7. Final Report
        Write-Host "`nAnalysis & Verification..." -ForegroundColor Yellow
        # Since we are mounted, we can measure destination
        # Use SilentlyContinue to ignore 'System Volume Information' access denied errors
        $destStats = Get-ChildItem -Path $drive -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
        $destSizeGB = 0; if ($destStats.Sum) { $destSizeGB = [Math]::Round($destStats.Sum / 1GB, 4) }
        $destCount = (Get-ChildItem -Path $drive -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count
        $sourceCount = (Get-ChildItem -Path $sourcePath -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count
        
        Dismount-VHDNative -Path $targetPath
        
        Show-Header "Capsule Creation Complete"
        Write-Host "Source" -ForegroundColor Cyan
        Write-Host "  Path   : $sourcePath"
        Write-Host "  Size   : $sourceSizeGB GB"
        Write-Host "  Items  : $sourceCount"
        
        Write-Host "`nVHD Capsule" -ForegroundColor Cyan
        Write-Host "  Path   : $targetPath"
        Write-Host "  Size   : $destSizeGB GB (Content Total)"
        Write-Host "  Items  : $destCount"
        Write-Host "  Time   : $($sw.Elapsed.ToString('hh\:mm\:ss'))"
        
        if ($p.ExitCode -ge 8) {
            Write-Host "`nWarning: Robocopy reported errors (Exit Code $($p.ExitCode))" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "`n[ERROR] An error occurred during copy/verification:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        # Attempt emergency dismount
        Dismount-VHDNative -Path $targetPath
    }
    
    Read-Host "`nPress Enter to return to menu"
}

function Invoke-VHDManager {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-Host "Error: File not found: $Path" -ForegroundColor Red
        return
    }

    $exitOps = $false
    while (-not $exitOps) {
        $img = Get-DiskImage -ImagePath $Path -ErrorAction SilentlyContinue
        $isMounted = $null
        if ($img) { $isMounted = $img.Attached }

        $ui = {
            Show-Header "Operations: $(Split-Path $Path -Leaf)"
            Write-Host "Selected: $Path"
            Write-Host "Size: $(Get-VHDXPhysicalSize $Path) MB" -ForegroundColor Gray
            Write-Host "-------------------"
            
            Write-Host "1. Launch in Capsule Mode"
            
            if ($isMounted) { Write-Host "2. Dismount" } else { Write-Host "2. Mount" }
            Write-Host "3. Current State (Size, Files, Fragmentation)"
            Write-Host "4. Compact (Shrink File)"
            Write-Host "5. Defragment Inside"
            Write-Host "6. Clean Junk"
            Write-Host "`n0. Go back to main menu"
        }
        
        $choice = Get-UserChoice -Prompt "`nSelect Option (Default: 1)" -Max 6 -UIBlock $ui -Default 1
        
        switch ($choice) {
            1 {
                # Capsule Mode
                $exitOps = $true
                Invoke-CapsuleMode -Path $Path -AppPath $null
            }
            2 { 
                if ($isMounted) {
                    Dismount-VHDNative -Path $Path
                    Write-Host "Dismounted." -ForegroundColor Yellow
                }
                else {
                    $drive = Mount-VHDNative -Path $Path
                    if ($drive) { 
                        Write-Host "Mounted at $drive" -ForegroundColor Green
                        Invoke-Item $drive
                    }
                }
                Read-Host "Press Enter"
            }
            3 {
                # Current State
                $drive = Mount-VHDNative -Path $Path
                if ($drive) {
                    $stats = @{
                        PhysicalSize = "$(Get-VHDXPhysicalSize $Path) MB"
                        UsedSpace    = "$([Math]::Round(((Get-Volume -DriveLetter $drive[0]).SizeRemaining / 1GB), 2)) GB Free"
                        FileCount    = (Get-ChildItem $drive -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
                    }
                    $stats | Format-Table -AutoSize
                    Optimize-Volume -DriveLetter $drive[0] -Analyze -Verbose
                    Dismount-VHDNative -Path $Path
                    Read-Host "Press Enter"
                }
            }
            4 {
                # Compact
                Dismount-VHDNative -Path $Path
                Write-Host "Compacting..." -ForegroundColor Yellow
                $script = @"
select vdisk file="$Path"
attach vdisk readonly
compact vdisk
detach vdisk
"@
                Invoke-DiskPartScript -ScriptContent $script | Out-Null
                Write-Host "Compaction Complete. New Size: $(Get-VHDXPhysicalSize $Path) MB" -ForegroundColor Green
                Read-Host "Press Enter"
            }
            5 {
                # Defrag
                $drive = Mount-VHDNative -Path $Path
                if ($drive) {
                    Optimize-Volume -DriveLetter $drive[0] -Defrag -Verbose
                    Read-Host "Press Enter"
                }
            }
            6 {
                # Clean Junk
                $drive = Mount-VHDNative -Path $Path
                if ($drive) {
                    $junk = @("$drive\`$RECYCLE.BIN", "$drive\System Volume Information")
                    foreach ($j in $junk) {
                        if (Test-Path $j) {
                            Remove-Item $j -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Host "Cleaned $j" -ForegroundColor Yellow
                        }
                    }
                    Read-Host "Cleanup Done. Press Enter"
                }
            }
            0 { $exitOps = $true; Dismount-VHDNative -Path $Path }
        }
    }
}

function Invoke-CapsuleMode {
    param([string]$Path, [string]$AppPath)
    
    if (-not $Path) {
        $Path = Read-Host "Drag and drop VHD file here"
        $Path = $Path.Trim('"')
    }
    if (-not (Test-Path $Path)) {
        Write-Host "File not found." -ForegroundColor Red; Start-Sleep 2; return
    }

    Show-Header "CAPSULE MODE: $(Split-Path $Path -Leaf)"
    
    # Step 1: Virtualization
    Write-Host "[1/4] Mounting..." -ForegroundColor Cyan
    $drive = Mount-VHDNative -Path $Path
    if (-not $drive) { Write-Host "Failed to mount." -ForegroundColor Red; return }
    
    if (-not $AppPath) {
        $defaultLnk = "launch_app.lnk"
        $isDefaultFound = Test-Path (Join-Path $drive $defaultLnk)
        $defaultMsg = "Enter Relative Path"
        if ($isDefaultFound) { $defaultMsg += " (Default: $defaultLnk)" }
        
        $inputPath = Read-Host $defaultMsg
        
        if ([string]::IsNullOrWhiteSpace($inputPath)) {
            if ($isDefaultFound) { $AppPath = $defaultLnk }
            else { 
                Write-Host "No path provided." -ForegroundColor Red; 
                Dismount-VHDNative -Path $Path; return 
            }
        }
        else {
            $AppPath = $inputPath.Trim('"')
        }
    }
    
    # Step 2: Snapshot
    Write-Host "[2/4] Taking filesystem snapshot..." -ForegroundColor Cyan
    $Before = Get-ChildItem -Path $drive -Recurse -File -ErrorAction SilentlyContinue | Select-Object FullName, LastWriteTime, Length
    
    # Step 3: Execution
    $FullPath = Join-Path $drive $AppPath
    if (Test-Path $FullPath) {
        Write-Host "[3/4] Launching $AppPath..." -ForegroundColor Green
        $proc = Start-Process -FilePath $FullPath -PassThru
        $proc.WaitForExit()
        Write-Host "Execution Finished." -ForegroundColor Yellow
        Start-Sleep -Seconds 3 # Flush buffers
    }
    else {
        Write-Host "Executable not found at $FullPath" -ForegroundColor Red
        Read-Host "Press Enter to continue to Maintenance"
    }
    
    # Analysis
    Write-Host "Analyzing changes..." -ForegroundColor Cyan
    $After = Get-ChildItem -Path $drive -Recurse -File -ErrorAction SilentlyContinue | Select-Object FullName, LastWriteTime, Length
    $Diffs = Compare-Object -ReferenceObject $Before -DifferenceObject $After -Property FullName, LastWriteTime, Length -PassThru
    
    if ($Diffs) {
        $Diffs | Format-Table -AutoSize
    }
    else {
        Write-Host "No filesystem changes detected." -ForegroundColor Gray
    }
    
    # Step 4: Maintenance
    Write-Host "[4/4] Maintenance Menu" -ForegroundColor Cyan
    
    $exitMaint = $false
    while (-not $exitMaint) {
        Write-Host "`n1. Current State"
        Write-Host "2. Compact"
        Write-Host "3. Defrag"
        Write-Host "4. Clean Junk"
        Write-Host "`n0. Go back to main menu"
    
        $c = Get-UserChoice -Prompt "`nSelection" -Max 4 -UIBlock $null
        switch ($c) {
            0 { $exitMaint = $true }
            1 { 
                Write-Host "Size: $(Get-VHDXPhysicalSize $Path) MB" 
            }
            2 {
                Dismount-VHDNative -Path $Path
                $script = "select vdisk file=`"$Path`"`nattach vdisk readonly`ncompact vdisk`ndetach vdisk"
                Invoke-DiskPartScript $script | Out-Null
                Write-Host "Compacted." -ForegroundColor Green
                $drive = Mount-VHDNative -Path $Path # Remount for continued maintenance if needed
            }
            3 { 
                if (-not $drive) { Write-Host "VHD is not mounted." -ForegroundColor Red }
                else { Optimize-Volume -DriveLetter $drive[0] -Defrag -Verbose }
            }
            4 { 
                if (-not $drive) { Write-Host "VHD is not mounted." -ForegroundColor Red }
                else {
                    $junk = @("$drive\`$RECYCLE.BIN", "$drive\System Volume Information")
                    foreach ($j in $junk) {
                        if (Test-Path $j) {
                            Remove-Item $j -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Host "Cleaned $j" -ForegroundColor Yellow
                        }
                    }
                    Write-Host "Cleanup Done." -ForegroundColor Green
                }
            }
        }
    }
    
    Dismount-VHDNative -Path $Path
    Write-Host "Capsule Closed." -ForegroundColor Green
    Start-Sleep 2
}

# -------------------------------------------------------------------------
# 3. FILE BROWSER
# -------------------------------------------------------------------------

function Select-VHDFile {
    Show-Header "Browse VHDs"
    Write-Host "Scanning directory: $pwd" -ForegroundColor Gray
    
    $files = Get-ChildItem -Path $pwd -File | Where-Object { $_.Extension -match "\.vhd(x)?$" }
    
    if (-not $files) {
        Write-Host "No VHD/VHDX files found in current directory." -ForegroundColor Yellow
        Read-Host "Press Enter"
        return $null
    }
    
    $ui = {
        Show-Header "Browse VHDs"
        Write-Host "Found $($files.Count) VHDs in $pwd`n"
        for ($i = 0; $i -lt $files.Count; $i++) {
            Write-Host "$($i+1). $($files[$i].Name) ($([Math]::Round($files[$i].Length/1MB, 0)) MB)"
        }
        Write-Host "`n0. Go back to main menu"
    }
    
    $sel = Get-UserChoice -Prompt "`nSelect Number" -Max $files.Count -UIBlock $ui
    if ($sel -eq 0) { return $null }
    return $files[$sel - 1].FullName
}

# -------------------------------------------------------------------------
# 4. MAIN ENTRY POINT
# -------------------------------------------------------------------------

if ($SourceFolder) {
    New-VHDCapsuleFromFolder -InputSource $SourceFolder -InputDest $DestinationPath -InputSize $SizeGB
    exit
}
if ($Mode -eq "Capsule" -and $VHDPath) {
    Invoke-CapsuleMode -Path $VHDPath -AppPath $AppPath
    exit
}
elseif ($Mode -eq "Manager" -and $VHDPath) {
    Invoke-VHDManager -Path $VHDPath
    exit
}

# Interactive Menu Loop
while ($true) {
    $ui = {
        Show-Header "Main Menu"
        Write-Host "1. Create VHD (Initialize & Format)"
        Write-Host "2. Create VHD Capsule from folder"
        Write-Host "3. Browse VHD (Select from list)"
        Write-Host "4. Manual Select VHD (Path input)"
        Write-Host "5. Launch VHDX in Capsule Mode"
        Write-Host "`n0. Exit"
    }
    
    $mainChoice = Get-UserChoice -Prompt "`nSelect Option" -Max 5 -UIBlock $ui
    
    switch ($mainChoice) {
        1 { New-VHDItem }
        2 { New-VHDCapsuleFromFolder }
        3 { 
            $p = Select-VHDFile
            if ($p) { Invoke-VHDManager -Path $p }
        }
        4 {
            $p = Read-Host "Enter full path to VHD/VHDX (Press Enter to go back)"
            if ([string]::IsNullOrWhiteSpace($p)) { continue }
            
            $p = $p.Trim('"')
            if (Test-Path $p) { Invoke-VHDManager -Path $p }
            else { Write-Host "File not found." -ForegroundColor Red; Read-Host "Press Enter" }
        }
        5 {
            $filesCheck = Get-ChildItem -Path $pwd -File | Where-Object { $_.Extension -match "\.vhd(x)?$" }
            
            if ($filesCheck) {
                # If files exist, browse them. If Select-VHDFile returns null (0), loop back (continue).
                $p = Select-VHDFile
                if (-not $p) { continue }
            }
            else {
                # No files found, go straight to manual
                Write-Host "No VHD/VHDX files found in current directory." -ForegroundColor Yellow
                $p = Read-Host "Enter full path to VHD/VHDX (Press Enter to go back)"
                if ([string]::IsNullOrWhiteSpace($p)) { continue }
                $p = $p.Trim('"')
            }

            if ($p -and (Test-Path $p)) { Invoke-CapsuleMode -Path $p -AppPath $null }
        }
        0 { exit }
    }
}
