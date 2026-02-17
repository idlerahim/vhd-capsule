<#
.SYNOPSIS
    Robust VHD/VHDX Manager and Capsule Launcher.
.DESCRIPTION
    Manages the lifecycle of Virtual Hard Disks (VHD/VHDX) and provides a specialized "Capsule Mode" 
    for isolating applications/games with tracking of filesystem changes.
    
    Features:
    - Create VHD/VHDX (Fixed/Dynamic)
    - Browse and Select VHDs
    - Operations: Mount, Compact, Defrag, Clean Junk
    - Capsule Mode: Mount -> Snapshot -> Execute -> Diff -> Maintenance
    
.PARAMETER VHDPath
    (Optional) Path to a VHD/VHDX file to pre-select or launch.
.PARAMETER Mode
    (Optional) Start directly in specific mode: 'Manager', 'Capsule'. Default is 'Menu'.
.PARAMETER GamePath
    (Optional) For Capsule Mode: The relative path to the executable inside the VHD.
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$VHDPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Menu", "Manager", "Capsule")]
    [string]$Mode = "Menu",

    [Parameter(Mandatory = $false)]
    [string]$GamePath,

    [Parameter(Mandatory = $false)]
    [string]$InitialDir
)

# -------------------------------------------------------------------------
# 0. INITIALIZATION & SAFETY
# -------------------------------------------------------------------------

if ($InitialDir -and (Test-Path $InitialDir)) {
    Set-Location $InitialDir
}

# Formatting for consistency
$Host.UI.RawUI.WindowTitle = "VHD Manager & Capsule Launcher"
$ErrorActionPreference = "Stop"

function Assert-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "[INFO] Requesting Administrator privileges..." -ForegroundColor Yellow
        $params = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        if ($VHDPath) { $params += " -VHDPath `"$VHDPath`"" }
        if ($Mode -ne "Menu") { $params += " -Mode $Mode" }
        if ($GamePath) { $params += " -GamePath `"$GamePath`"" }
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
    Write-Host "   VHD MANAGER: $Title" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-UserChoice {
    param([string]$Prompt, [int]$Max)
    while ($true) {
        $inputVal = Read-Host $Prompt
        if ($inputVal -match "^\d+$" -and [int]$inputVal -ge 1 -and [int]$inputVal -le $Max) {
            return [int]$inputVal
        }
        Write-Host "Invalid selection. Please enter a number between 1 and $Max." -ForegroundColor Red
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
    $letters = 68..90 | ForEach-Object { [char]$_ + ":" } # D: to Z:
    $used = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Name
    foreach ($letter in $letters) {
        if ($used -notcontains $letter[0]) { return $letter }
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
    
    # Try generic Mount-DiskImage
    try {
        $mount = Mount-DiskImage -ImagePath $Path -PassThru -ErrorAction Stop
        
        # Wait for volume
        Start-Sleep -Seconds 2
        
        $volume = $mount | Get-Disk | Get-Partition | Where-Object { $_.DriveLetter } | Select-Object -First 1
        if ($volume) {
            return "$($volume.DriveLetter):"
        }
        
        # If no letter assigned, try to assign one via diskpart
        Write-Host "No drive letter found. Attempting to assign..." -ForegroundColor Yellow
        $script = @"
select vdisk file="$Path"
attach vdisk
select partition 1
assign
"@
        Invoke-DiskPartScript -ScriptContent $script | Out-Null
        
        # Check again
        $mount = Get-DiskImage -ImagePath $Path
        $volume = $mount | Get-Disk | Get-Partition | Where-Object { $_.DriveLetter } | Select-Object -First 1
        if ($volume) { return "$($volume.DriveLetter):" }
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
    Write-Host "0. Cancel"
    $fmtChoice = Read-Host "Select Type"
    if ($fmtChoice -eq "0") { return }
    $extension = if ($fmtChoice -eq "2") { ".vhd" } else { ".vhdx" }
    
    # 1. Filename
    Show-Header "Create New VHD/VHDX: Step 2/7"
    Write-Host "Enter Filename (e.g. MyGames$extension)"
    Write-Host "0. Cancel"
    $name = Read-Host "Filename"
    if ($name -eq "0") { return }
    
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    if ($name -match "\.(vhd|vhdx)$") { $name = $name -replace "\.(vhd|vhdx)$", "" }
    $name += $extension
    $targetPath = Join-Path $pwd $name

    # 2. Size
    Show-Header "Create New VHD/VHDX: Step 3/7"
    Write-Host "Enter Size in GB"
    Write-Host "0. Cancel"
    $sizeGB = Read-Host "Size"
    if ($sizeGB -eq "0") { return }
    
    if ($sizeGB -notmatch "^\d+$") { $sizeGB = 10 }
    $sizeMB = [int]$sizeGB * 1024

    # 3. Type
    Show-Header "Create New VHD/VHDX: Step 4/7"
    Write-Host "Virtual Disk Type:"
    Write-Host "1. Dynamic (Saves Space) [Default]"
    Write-Host "2. Fixed (Better Performance)"
    Write-Host "0. Cancel"
    $typeChoice = Read-Host "Select (1/2)"
    if ($typeChoice -eq "0") { return }
    $type = if ($typeChoice -eq "2") { "fixed" } else { "expandable" }

    # 4. Partition Style
    Show-Header "Create New VHD/VHDX: Step 5/7"
    Write-Host "Initialize Disk as:"
    Write-Host "1. GPT (GUID Partition Table) [Default]"
    Write-Host "2. MBR (Master Boot Record)"
    Write-Host "0. Cancel"
    $partChoice = Read-Host "Select Style (1/2)"
    if ($partChoice -eq "0") { return }
    $partStyle = if ($partChoice -eq "2") { "mbr" } else { "gpt" }

    # 5. File System
    Show-Header "Create New VHD/VHDX: Step 6/7"
    Write-Host "Format Disk as:"
    Write-Host "1. NTFS [Default]"
    Write-Host "2. FAT32"
    Write-Host "0. Cancel"
    $fsChoice = Read-Host "Select File System (1/2)"
    if ($fsChoice -eq "0") { return }
    $fs = if ($fsChoice -eq "2") { "fat32" } else { "ntfs" }
    
    $enableCompression = $false
    if ($fs -eq "ntfs") {
        Write-Host "`nEnable file and folder compression:"
        Write-Host "1. No [Default]"
        Write-Host "2. Yes"
        Write-Host "0. Cancel"
        $compChoice = Read-Host "Select (1/2)"
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
        @{ L = "4 KB (OS drives, games, mixed data, VHDX)"; V = "4096" },
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
    Write-Host "0. Cancel"
    $allocChoice = Read-Host "Select Allocation Unit (Default: 1)"
    if ($allocChoice -eq "0") { return }
    
    if ($allocChoice -notmatch "^\d+$" -or [int]$allocChoice -lt 1 -or [int]$allocChoice -gt $allocUnits.Count) { $allocChoice = 1 }
    $allocUnit = $allocUnits[[int]$allocChoice - 1].V

    # 7. Volume Label & Confirmation
    Show-Header "Create New VHD/VHDX: Confirmation"
    Write-Host "Enter Volume Label (Default: `"New Volume`")"
    Write-Host "0. Cancel"
    $label = Read-Host "Label"
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

        Show-Header "Operations: $(Split-Path $Path -Leaf)"
        Write-Host "Selected: $Path"
        Write-Host "Size: $(Get-VHDXPhysicalSize $Path) MB" -ForegroundColor Gray
        Write-Host "-------------------"
        
        if ($isMounted) { Write-Host "1. Dismount" } else { Write-Host "1. Mount" }
        Write-Host "2. Current State (Size, Files, Fragmentation)"
        Write-Host "3. Compact (Shrink File)"
        Write-Host "4. Defragment Inside"
        Write-Host "5. Clean Junk"
        Write-Host "6. Back to Main Menu"
        
        $choice = Get-UserChoice "`nSelect Option" 6
        
        switch ($choice) {
            1 { 
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
            2 {
                # Current State
                $drive = Mount-VHDNative -Path $Path
                if ($drive) {
                    $stats = @{
                        PhysicalSize = "$(Get-VHDXPhysicalSize $Path) MB"
                        UsedSpace    = "$([Math]::Round(((Get-Volume -DriveLetter $drive[0]).SizeRemaining / 1GB), 2)) GB Free"
                        FileCount    = (Get-ChildItem $drive -Recurse -File | Measure-Object).Count
                    }
                    $stats | Format-Table -AutoSize
                    Optimize-Volume -DriveLetter $drive[0] -Analyze -Verbose
                    Read-Host "Press Enter"
                }
            }
            3 {
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
            4 {
                # Defrag
                $drive = Mount-VHDNative -Path $Path
                if ($drive) {
                    Optimize-Volume -DriveLetter $drive[0] -Defrag -Verbose
                    Read-Host "Press Enter"
                }
            }
            5 {
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
            6 { $exitOps = $true; Dismount-VHDNative -Path $Path }
        }
    }
}

function Invoke-CapsuleMode {
    param([string]$Path, [string]$RelPath)
    
    if (-not $Path) {
        $Path = Read-Host "Drag and drop VHD file here"
        $Path = $Path.Trim('"')
    }
    if (-not (Test-Path $Path)) {
        Write-Host "File not found." -ForegroundColor Red; Start-Sleep 2; return
    }

    if (-not $RelPath) {
        $RelPath = Read-Host "Enter Relative Game Path (e.g. GameFolder\Game.exe)"
    }
    
    Show-Header "CAPSULE MODE: $(Split-Path $Path -Leaf)"
    
    # Step 1: Virtualization
    Write-Host "[1/4] Mounting..." -ForegroundColor Cyan
    $drive = Mount-VHDNative -Path $Path
    if (-not $drive) { Write-Host "Failed to mount." -ForegroundColor Red; return }
    
    # Step 2: Snapshot
    Write-Host "[2/4] Taking filesystem snapshot..." -ForegroundColor Cyan
    $Before = Get-ChildItem -Path $drive -Recurse -File | Select-Object FullName, LastWriteTime, Length
    
    # Step 3: Execution
    $FullPath = Join-Path $drive $RelPath
    if (Test-Path $FullPath) {
        Write-Host "[3/4] Launching $RelPath..." -ForegroundColor Green
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
    $After = Get-ChildItem -Path $drive -Recurse -File | Select-Object FullName, LastWriteTime, Length
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
        Write-Host "`n1. Dismount and Exit"
        Write-Host "2. Current State"
        Write-Host "3. Compact"
        Write-Host "4. Defrag"
        Write-Host "5. Clean Junk"
        
        $c = Read-Host "Selection"
        switch ($c) {
            1 { $exitMaint = $true }
            2 { 
                Write-Host "Size: $(Get-VHDXPhysicalSize $Path) MB" 
            }
            3 {
                Dismount-VHDNative -Path $Path
                $script = "select vdisk file=`"$Path`"`nattach vdisk readonly`ncompact vdisk`ndetach vdisk"
                Invoke-DiskPartScript $script | Out-Null
                Write-Host "Compacted." -ForegroundColor Green
                $drive = Mount-VHDNative -Path $Path # Remount for continued maintenance if needed
            }
            4 { Optimize-Volume -DriveLetter $drive[0] -Defrag -Verbose }
            5 { 
                Remove-Item "$drive\`$RECYCLE.BIN" -Recurse -Force -ErrorAction SilentlyContinue 
                Write-Host "Cleaned."
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
    Write-Host "Scanning directory: $pwd" -ForegroundColor Gray
    $files = Get-ChildItem -Path $pwd -File | Where-Object { $_.Extension -match "\.vhd(x)?$" }
    if (-not $files) {
        Write-Host "No VHD/VHDX files found in current directory." -ForegroundColor Yellow
        return $null
    }
    
    Write-Host "Found VHDs:"
    for ($i = 0; $i -lt $files.Count; $i++) {
        Write-Host "$($i+1). $($files[$i].Name) ($([Math]::Round($files[$i].Length/1MB, 0)) MB)"
    }
    
    $sel = Get-UserChoice "`nSelect Number" $files.Count
    return $files[$sel - 1].FullName
}

# -------------------------------------------------------------------------
# 4. MAIN ENTRY POINT
# -------------------------------------------------------------------------

if ($Mode -eq "Capsule" -and $VHDPath) {
    Invoke-CapsuleMode -Path $VHDPath -RelPath $GamePath
    exit
}
elseif ($Mode -eq "Manager" -and $VHDPath) {
    Invoke-VHDManager -Path $VHDPath
    exit
}

# Interactive Menu Loop
while ($true) {
    Show-Header "Main Menu"
    Write-Host "1. Create VHD (Initialize & Format)"
    Write-Host "2. Browse VHD (Select from list)"
    Write-Host "3. Manual Select VHD (Path input)"
    Write-Host "4. Launch VHDX in Capsule Mode"
    Write-Host "5. Exit"
    
    $mainChoice = Get-UserChoice "`nSelect Option" 5
    
    switch ($mainChoice) {
        1 { New-VHDItem }
        2 { 
            $p = Select-VHDFile
            if ($p) { Invoke-VHDManager -Path $p }
            else { Read-Host "Press Enter" }
        }
        3 {
            $p = Read-Host "Enter full path to VHD/VHDX"
            $p = $p.Trim('"')
            if (Test-Path $p) { Invoke-VHDManager -Path $p }
            else { Write-Host "File not found." -ForegroundColor Red; Read-Host "Press Enter" }
        }
        4 {
            $p = Select-VHDFile
            if (-not $p) {
                $p = Read-Host "Enter full path to VHD/VHDX"
                $p = $p.Trim('"')
            }
            if ($p -and (Test-Path $p)) { Invoke-CapsuleMode -Path $p -RelPath $null }
        }
        5 { exit }
    }
}
