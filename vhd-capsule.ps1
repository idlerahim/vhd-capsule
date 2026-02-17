$VHDXFileName = "Super Street Fighter IV Arcade Edition.vhdx"
$RelativeGamePath = "SSFIV.exe"
$VHDXPath = Join-Path $PSScriptRoot $VHDXFileName

# 0. Safety Check & Cleanup
if (-not (Test-Path $VHDXPath)) {
    Write-Host "[ERROR] VHDX file not found at: $VHDXPath" -ForegroundColor Red
    Read-Host "Press any key to exit"
    exit
}

# Attempt to clear any "stuck" mounts from previous failed runs
$ExistingMount = Get-DiskImage -ImagePath $VHDXPath
if ($ExistingMount.Attached) {
    Write-Host "[INFO] VHDX is already attached. Attempting to cycle connection..." -ForegroundColor Yellow
    Dismount-DiskImage -ImagePath $VHDXPath -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# 1. Self-Elevation Block
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Get-VHDXPhysicalSize {
    $file = Get-Item $VHDXPath
    return [Math]::Round($file.Length / 1MB, 2)
}

Write-Host "--- Starting Game Launch Sequence ---" -ForegroundColor White
Write-Host "VHDX Name: $VHDXFileName"
Write-Host "Game Path: $RelativeGamePath"
Write-Host "Current Physical Size: $(Get-VHDXPhysicalSize) MB`n"

# STEP 1: VIRTUALIZATION
Write-Host "[STEP 1/4] STATUS: MOUNTING VHDX [$VHDXFileName]..." -ForegroundColor Cyan
try {
    $DiskImage = Mount-DiskImage -ImagePath $VHDXPath -PassThru -ErrorAction Stop
    Start-Sleep -Seconds 3
    $Partition = $DiskImage | Get-Disk | Get-Partition | Where-Object { $_.DriveLetter }
    
    if (-not $Partition.DriveLetter) {
        throw "Disk mounted but no drive letter assigned."
    }
    
    $Drive = "$($Partition.DriveLetter):"
    Write-Host "                    COMPLETE: Mounted on $Drive" -ForegroundColor Green
}
catch {
    Write-Host "`n[FATAL ERROR] Could not mount VHDX." -ForegroundColor Red
    if ($_.Exception.HResult -eq -2147024864) { # 0x80070020
        Write-Host "REASON: The file is locked by another process." -ForegroundColor Yellow
        Write-Host "FIX: Check if the VHDX is already open in Disk Management, another script, or an Emulator." -ForegroundColor White
    } else {
        Write-Host "REASON: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    Write-Host "`nScript cannot continue."
    Read-Host "Press any key to exit"
    exit
}

# STEP 2: PRE-FLIGHT SNAPSHOT
Write-Host "[STEP 2/4] STATUS: TAKING BEFORE SNAPSHOT..." -ForegroundColor Cyan
$Before = Get-ChildItem -Path $Drive -Recurse -File | Select-Object FullName, LastWriteTime, Length
Write-Host "                    COMPLETE: Baseline established." -ForegroundColor Green

# STEP 3: EXECUTION & TRACKING
Write-Host "[STEP 3/4] STATUS: LAUNCHING GAME [$RelativeGamePath]..." -ForegroundColor Cyan
$FullGamePath = Join-Path $Drive $RelativeGamePath

if (Test-Path $FullGamePath) {
    $gameProcess = Start-Process -FilePath $FullGamePath -PassThru
    
    # Wait for the process to exit
    $gameProcess.WaitForExit()
    
    Write-Host "                    COMPLETE: Game process closed." -ForegroundColor Green
    Write-Host "                    STATUS: Waiting 3s for file system to settle..." -ForegroundColor Gray
    
    # The requested 3-second wait for file flushing
    Start-Sleep -Seconds 3
}
else {
    Write-Host "[ERROR] Could not find $RelativeGamePath on $Drive" -ForegroundColor Red
    Read-Host "Press any key to continue to maintenance..."
}

# ANALYSIS
Write-Host "`n[ANALYSIS] File System Changes:" -ForegroundColor Yellow
$After = Get-ChildItem -Path $Drive -Recurse -File | Select-Object FullName, LastWriteTime, Length
$Diffs = Compare-Object -ReferenceObject $Before -DifferenceObject $After -Property FullName, LastWriteTime, Length -PassThru
$Results = foreach ($Item in $Diffs) {
    $RelativeName = $Item.FullName.Replace($Drive, "")
    $Status = if ($Item.SideIndicator -eq "=>") { if ($Before.FullName -contains $Item.FullName) { "MODIFIED" } else { "ADDED" } } else { if ($After.FullName -notcontains $Item.FullName) { "DELETED" } else { $null } }
    if ($Status) { [PSCustomObject]@{ Status = $Status; File = $RelativeName; Size = "$([Math]::Round($Item.Length / 1KB, 2)) KB"; Date = $Item.LastWriteTime.ToString("dd-MMM-yy HH:mm:ss") } }
}
if ($Results) { $Results | Sort-Object Status, Date | Format-Table -AutoSize } else { Write-Host "No changes detected." -ForegroundColor Gray }

# STEP 4: MAINTENANCE & COMPACTION
Write-Host "[STEP 4/4] STATUS: ENTERING MAINTENANCE MODE..." -ForegroundColor Cyan
$ExitMenu = $false
while (-not $ExitMenu) {
    Write-Host "`n--- VHDX MANAGEMENT MENU ---" -ForegroundColor Magenta
    Write-Host "1. Dismount and Exit (Press 1 or Enter)"
    Write-Host "2. Current State (Size, Files, Fragmentation)"
    Write-Host "3. Compact (Shrink VHDX File)"
    Write-Host "4. Defragment Inside VHDX"
    Write-Host "5. Clean Junk (.BIN & System Volume Info)"
    $choice = Read-Host "`nSelect an option"

    switch ($choice) {
        {($_ -eq "1") -or ($_ -eq "")} { $ExitMenu = $true }
        "2" {
            $allFiles = Get-ChildItem -Path $Drive -Recurse
            $fileCount = ($allFiles | Where-Object { -not $_.PSIsContainer }).Count
            $usedSpace = (Get-Volume -DriveLetter $Partition.DriveLetter).SizeRemaining
            Write-Host "`nInternal Files: $fileCount | Physical Size: $(Get-VHDXPhysicalSize) MB" -ForegroundColor Cyan
            Optimize-Volume -DriveLetter $Partition.DriveLetter -Analyze -Verbose
        }
		"3" { 
            $OldSize = Get-VHDXPhysicalSize
            Write-Host "Dismounting for compaction..." -ForegroundColor Cyan
            Dismount-DiskImage -ImagePath $VHDXPath
            Start-Sleep -Seconds 2

            # Create a temporary DiskPart script
            $dpScript = Join-Path $env:TEMP "compact_vhd.txt"
            # The closing tag below MUST stay at the far left margin
            $scriptContent = @"
select vdisk file="$VHDXPath"
attach vdisk readonly
compact vdisk
detach vdisk
"@
            $scriptContent | Out-File -FilePath $dpScript -Encoding ASCII

            Write-Host "Compacting VHDX via DiskPart (this may take a minute)..." -ForegroundColor Yellow
            diskpart /s $dpScript
            Remove-Item $dpScript

            $NewSize = Get-VHDXPhysicalSize
            Write-Host "Reclaimed: $([Math]::Round($OldSize - $NewSize, 2)) MB" -ForegroundColor Green
            
            # Remount so the script can continue tracking or exit
            $DiskImage = Mount-DiskImage -ImagePath $VHDXPath -PassThru
            Start-Sleep -Seconds 2
        }
        "4" { Optimize-Volume -DriveLetter $Partition.DriveLetter -Defrag -Verbose }
		"5" {
            Write-Host "`n[EXEC] Analyzing junk files..." -ForegroundColor Cyan
            $junkPaths = @("$Drive\`$RECYCLE.BIN", "$Drive\System Volume Information")
            $totalFreedBytes = 0
            $filesDeleted = 0
            $foldersDeleted = 0
            
            foreach ($path in $junkPaths) {
                if (Test-Path $path) {
                    # Get all items including those in hidden SID subfolders
                    $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                    
                    foreach ($item in $items) {
                        if ($item.PSIsContainer) {
                            $foldersDeleted++
                        } else {
                            $shortName = $item.FullName.Replace($Drive, '')
                            Write-Host "Deleting: $shortName ($([Math]::Round($item.Length / 1KB, 2)) KB)" -ForegroundColor Gray
                            $totalFreedBytes += $item.Length
                            $filesDeleted++
                        }
                    }
                    # Force removal of the root junk directories and their contents
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                    $foldersDeleted++ # Count the parent junk folder itself
                }
            }
            
            if ($filesDeleted -eq 0 -and $foldersDeleted -eq 0) {
                Write-Host "No junk files found. The Recycle Bin is already empty." -ForegroundColor Gray
            } else {
                Write-Host "`nCleanup Complete!" -ForegroundColor Green
                Write-Host "Items Removed: $filesDeleted Files, $foldersDeleted Folders"
                Write-Host "Total Freed  : $totalFreedBytes Bytes ($([Math]::Round($totalFreedBytes / 1MB, 4)) MB)" -ForegroundColor Yellow
            }
        }
    }
}

# FINAL DISMOUNT
Dismount-DiskImage -ImagePath $VHDXPath
Write-Host "[STEP 4/4] COMPLETE: VHDX Disconnected." -ForegroundColor Green
Write-Host "--- Sequence Complete ---"
Read-Host "Press any key to exit"