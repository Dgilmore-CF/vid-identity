<#
.SYNOPSIS
    Organizes video files by resolution - sort into folders, move files, or delete low-res videos.

.DESCRIPTION
    Analyzes video files and performs organization tasks:
    - Sort videos into resolution-based folders (4K, 1080p, 720p, etc.)
    - Intelligently move files with disk space checking and queuing
    - Mass delete videos below a specified resolution threshold
    
    Cross-platform compatible: Windows, macOS, and Linux.

.PARAMETER Path
    One or more directories to scan for video files.

.PARAMETER Recurse
    Include subdirectories in the scan.

.PARAMETER Action
    The action to perform: 'Sort', 'Delete', or 'Report' (dry-run).

.PARAMETER DestinationRoot
    Root folder where resolution subfolders will be created (for Sort action).

.PARAMETER MinResolution
    Minimum resolution to keep. Videos below this will be deleted (for Delete action).
    Options: 4K, 1440p, 1080p, 720p, 480p, 360p

.PARAMETER Force
    Skip confirmation prompts for destructive operations.

.PARAMETER FFprobePath
    Optional path to ffprobe executable.

.EXAMPLE
    # Sort videos into resolution folders
    .\Organize-Videos.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -Recurse

.EXAMPLE
    # Delete all videos below 720p (with confirmation)
    .\Organize-Videos.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -Recurse

.EXAMPLE
    # Preview what would happen (dry-run)
    .\Organize-Videos.ps1 -Path "D:\Videos" -Action Report -MinResolution 1080p -Recurse

.NOTES
    Requires FFprobe (part of FFmpeg).
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Position = 0, Mandatory = $true)]
    [string[]]$Path,
    
    [Parameter(Mandatory = $true)]
    [ValidateSet("Sort", "Delete", "Report")]
    [string]$Action,
    
    [switch]$Recurse,
    
    [string]$DestinationRoot,
    
    [ValidateSet("4K", "1440p", "1080p", "720p", "480p", "360p")]
    [string]$MinResolution,
    
    [switch]$Force,
    
    [string]$FFprobePath
)

#region Configuration

$VideoExtensions = @(
    "*.mp4", "*.mkv", "*.avi", "*.mov", "*.wmv", "*.flv", "*.webm",
    "*.m4v", "*.mpg", "*.mpeg", "*.3gp", "*.3g2", "*.mts", "*.m2ts",
    "*.ts", "*.vob", "*.ogv", "*.divx", "*.xvid", "*.asf", "*.rm",
    "*.rmvb", "*.f4v", "*.hevc", "*.264", "*.265"
)

$ResolutionThresholds = @{
    "4K"    = 2160
    "1440p" = 1440
    "1080p" = 1080
    "720p"  = 720
    "480p"  = 480
    "360p"  = 360
}

$ResolutionFolderNames = @{
    "4K UHD"    = "4K_UHD"
    "1440p QHD" = "1440p_QHD"
    "1080p FHD" = "1080p_FHD"
    "720p HD"   = "720p_HD"
    "480p SD"   = "480p_SD"
    "360p"      = "360p"
    "Low"       = "Low_Resolution"
}

#endregion

#region Platform Detection

$IsWindowsOS = $false
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $IsWindowsOS = $IsWindows
} else {
    $IsWindowsOS = $true
}

#endregion

#region FFprobe Functions

$script:FFprobeCmd = "ffprobe"

function Find-FFprobe {
    param([string]$CustomPath)
    
    if ($CustomPath) {
        if (Test-Path $CustomPath) {
            $script:FFprobeCmd = $CustomPath
            return $true
        }
        Write-Warning "Specified FFprobe path not found: $CustomPath"
        return $false
    }
    
    try {
        $null = & ffprobe -version 2>&1
        $script:FFprobeCmd = "ffprobe"
        return $true
    } catch { }
    
    if ($IsWindowsOS) {
        $commonPaths = @(
            "$env:ProgramFiles\ffmpeg\bin\ffprobe.exe",
            "$env:ProgramFiles(x86)\ffmpeg\bin\ffprobe.exe",
            "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\*\ffprobe.exe",
            "$env:ChocolateyInstall\bin\ffprobe.exe",
            "C:\ffmpeg\bin\ffprobe.exe",
            "$env:USERPROFILE\ffmpeg\bin\ffprobe.exe",
            "$env:USERPROFILE\scoop\apps\ffmpeg\current\bin\ffprobe.exe"
        )
        
        foreach ($path in $commonPaths) {
            $resolved = Resolve-Path $path -ErrorAction SilentlyContinue
            if ($resolved) {
                $script:FFprobeCmd = $resolved.Path | Select-Object -First 1
                return $true
            }
        }
    }
    
    return $false
}

function Get-VideoDetails {
    param([string]$FilePath)
    
    try {
        $json = & $script:FFprobeCmd -v quiet -print_format json -show_streams -show_format "$FilePath" 2>&1
        $info = $json | ConvertFrom-Json
        
        $videoStream = $info.streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
        
        if (-not $videoStream) {
            return $null
        }
        
        $height = [int]$videoStream.height
        $resolutionCategory = switch ($height) {
            { $_ -ge 2160 } { "4K UHD" }
            { $_ -ge 1440 } { "1440p QHD" }
            { $_ -ge 1080 } { "1080p FHD" }
            { $_ -ge 720 }  { "720p HD" }
            { $_ -ge 480 }  { "480p SD" }
            { $_ -ge 360 }  { "360p" }
            default         { "Low" }
        }
        
        $fileInfo = Get-Item $FilePath
        
        return [PSCustomObject]@{
            FileName           = $fileInfo.Name
            FullPath           = $FilePath
            Width              = [int]$videoStream.width
            Height             = $height
            Resolution         = "$($videoStream.width)x$($videoStream.height)"
            ResolutionCategory = $resolutionCategory
            FileSizeBytes      = $fileInfo.Length
            FileSizeMB         = [math]::Round($fileInfo.Length / 1MB, 2)
            Drive              = if ($IsWindowsOS) { [System.IO.Path]::GetPathRoot($FilePath) } else { "/" }
        }
    }
    catch {
        Write-Warning "Failed to analyze: $FilePath - $_"
        return $null
    }
}

#endregion

#region Disk Space Functions

function Get-DriveInfo {
    param([string]$Path)
    
    try {
        if ($IsWindowsOS) {
            $root = [System.IO.Path]::GetPathRoot($Path)
            $drive = Get-PSDrive -Name $root.TrimEnd(':\') -ErrorAction SilentlyContinue
            if ($drive) {
                return [PSCustomObject]@{
                    Path          = $root
                    FreeBytes     = $drive.Free
                    FreeMB        = [math]::Round($drive.Free / 1MB, 2)
                    FreeGB        = [math]::Round($drive.Free / 1GB, 2)
                    UsedBytes     = $drive.Used
                    TotalBytes    = $drive.Free + $drive.Used
                }
            }
            
            # Fallback to WMI/CIM
            $driveLetter = $root.TrimEnd('\')
            $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
            if ($disk) {
                return [PSCustomObject]@{
                    Path          = $root
                    FreeBytes     = $disk.FreeSpace
                    FreeMB        = [math]::Round($disk.FreeSpace / 1MB, 2)
                    FreeGB        = [math]::Round($disk.FreeSpace / 1GB, 2)
                    UsedBytes     = $disk.Size - $disk.FreeSpace
                    TotalBytes    = $disk.Size
                }
            }
        }
        else {
            # macOS/Linux - use df command
            $mountPoint = $Path
            while ($mountPoint -ne "/" -and -not (Test-Path (Join-Path $mountPoint ".." -Resolve) -PathType Container)) {
                $mountPoint = Split-Path $mountPoint -Parent
            }
            
            $dfOutput = & df -k $Path 2>&1 | Select-Object -Last 1
            $parts = $dfOutput -split '\s+' | Where-Object { $_ }
            
            if ($parts.Count -ge 4) {
                $totalKB = [long]$parts[1]
                $usedKB = [long]$parts[2]
                $availKB = [long]$parts[3]
                
                return [PSCustomObject]@{
                    Path          = $Path
                    FreeBytes     = $availKB * 1024
                    FreeMB        = [math]::Round(($availKB * 1024) / 1MB, 2)
                    FreeGB        = [math]::Round(($availKB * 1024) / 1GB, 2)
                    UsedBytes     = $usedKB * 1024
                    TotalBytes    = $totalKB * 1024
                }
            }
        }
    }
    catch {
        Write-Warning "Could not get drive info for $Path : $_"
    }
    
    return $null
}

function Test-SufficientSpace {
    param(
        [string]$DestinationPath,
        [long]$RequiredBytes,
        [double]$SafetyMarginPercent = 5
    )
    
    $driveInfo = Get-DriveInfo -Path $DestinationPath
    
    if (-not $driveInfo) {
        Write-Warning "Could not determine disk space for: $DestinationPath"
        return $false
    }
    
    # Add safety margin
    $safetyMargin = $driveInfo.TotalBytes * ($SafetyMarginPercent / 100)
    $requiredWithMargin = $RequiredBytes + $safetyMargin
    
    return $driveInfo.FreeBytes -ge $requiredWithMargin
}

#endregion

#region File Operations

function Move-VideoFile {
    param(
        [PSCustomObject]$VideoInfo,
        [string]$DestinationFolder,
        [switch]$WhatIf
    )
    
    $destPath = Join-Path $DestinationFolder $VideoInfo.FileName
    
    # Handle duplicate filenames
    if (Test-Path $destPath) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($VideoInfo.FileName)
        $extension = [System.IO.Path]::GetExtension($VideoInfo.FileName)
        $counter = 1
        
        do {
            $newName = "${baseName}_$counter$extension"
            $destPath = Join-Path $DestinationFolder $newName
            $counter++
        } while (Test-Path $destPath)
    }
    
    if ($WhatIf) {
        Write-Host "  [WhatIf] Would move: $($VideoInfo.FileName) -> $DestinationFolder" -ForegroundColor DarkYellow
        return [PSCustomObject]@{
            Success     = $true
            Source      = $VideoInfo.FullPath
            Destination = $destPath
            WhatIf      = $true
        }
    }
    
    try {
        # Ensure destination folder exists
        if (-not (Test-Path $DestinationFolder)) {
            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
        }
        
        Move-Item -Path $VideoInfo.FullPath -Destination $destPath -Force
        
        return [PSCustomObject]@{
            Success     = $true
            Source      = $VideoInfo.FullPath
            Destination = $destPath
            WhatIf      = $false
        }
    }
    catch {
        Write-Warning "Failed to move $($VideoInfo.FileName): $_"
        return [PSCustomObject]@{
            Success     = $false
            Source      = $VideoInfo.FullPath
            Destination = $destPath
            Error       = $_.Exception.Message
            WhatIf      = $false
        }
    }
}

function Remove-VideoFile {
    param(
        [PSCustomObject]$VideoInfo,
        [switch]$WhatIf
    )
    
    if ($WhatIf) {
        Write-Host "  [WhatIf] Would delete: $($VideoInfo.FileName) ($($VideoInfo.FileSizeMB) MB)" -ForegroundColor DarkYellow
        return [PSCustomObject]@{
            Success   = $true
            Path      = $VideoInfo.FullPath
            SizeBytes = $VideoInfo.FileSizeBytes
            WhatIf    = $true
        }
    }
    
    try {
        Remove-Item -Path $VideoInfo.FullPath -Force
        
        return [PSCustomObject]@{
            Success   = $true
            Path      = $VideoInfo.FullPath
            SizeBytes = $VideoInfo.FileSizeBytes
            WhatIf    = $false
        }
    }
    catch {
        Write-Warning "Failed to delete $($VideoInfo.FileName): $_"
        return [PSCustomObject]@{
            Success   = $false
            Path      = $VideoInfo.FullPath
            Error     = $_.Exception.Message
            WhatIf    = $false
        }
    }
}

#endregion

#region Queue Management

function Invoke-QueuedMoves {
    param(
        [array]$VideoList,
        [string]$DestinationRoot,
        [switch]$WhatIf
    )
    
    $moveQueue = [System.Collections.Generic.List[PSCustomObject]]::new()
    $results = @{
        Moved   = @()
        Queued  = @()
        Failed  = @()
        Skipped = @()
    }
    
    # Group videos by destination folder
    $groupedVideos = $VideoList | Group-Object ResolutionCategory
    
    foreach ($group in $groupedVideos) {
        $folderName = $ResolutionFolderNames[$group.Name]
        $destFolder = Join-Path $DestinationRoot $folderName
        
        Write-Host "`nProcessing $($group.Name) videos -> $destFolder" -ForegroundColor Cyan
        
        # Calculate total size needed for this group
        $totalSizeNeeded = ($group.Group | Measure-Object -Property FileSizeBytes -Sum).Sum
        $totalSizeGB = [math]::Round($totalSizeNeeded / 1GB, 2)
        Write-Host "  Total size: $totalSizeGB GB" -ForegroundColor DarkGray
        
        # Check if we have enough space for all files
        $driveInfo = Get-DriveInfo -Path $DestinationRoot
        
        if (-not $driveInfo) {
            Write-Warning "Cannot determine disk space. Proceeding with caution..."
        }
        
        foreach ($video in $group.Group) {
            # Skip if source and destination are same location
            $destPath = Join-Path $destFolder $video.FileName
            if ($video.FullPath -eq $destPath) {
                Write-Host "  Skipped (already in place): $($video.FileName)" -ForegroundColor DarkGray
                $results.Skipped += $video
                continue
            }
            
            # Check if same drive (move is instant) or different drive (copy+delete)
            $sourceDrive = if ($IsWindowsOS) { [System.IO.Path]::GetPathRoot($video.FullPath) } else { "/" }
            $destDrive = if ($IsWindowsOS) { [System.IO.Path]::GetPathRoot($DestinationRoot) } else { "/" }
            $sameDrive = $sourceDrive -eq $destDrive
            
            if ($sameDrive) {
                # Same drive - move is instant, no space check needed
                $moveResult = Move-VideoFile -VideoInfo $video -DestinationFolder $destFolder -WhatIf:$WhatIf
                
                if ($moveResult.Success) {
                    $results.Moved += $moveResult
                    Write-Host "  Moved: $($video.FileName)" -ForegroundColor Green
                }
                else {
                    $results.Failed += $moveResult
                }
            }
            else {
                # Different drive - need space check
                if (Test-SufficientSpace -DestinationPath $DestinationRoot -RequiredBytes $video.FileSizeBytes) {
                    $moveResult = Move-VideoFile -VideoInfo $video -DestinationFolder $destFolder -WhatIf:$WhatIf
                    
                    if ($moveResult.Success) {
                        $results.Moved += $moveResult
                        Write-Host "  Moved: $($video.FileName)" -ForegroundColor Green
                    }
                    else {
                        $results.Failed += $moveResult
                    }
                }
                else {
                    # Queue for later when space becomes available
                    $queueItem = [PSCustomObject]@{
                        Video       = $video
                        Destination = $destFolder
                        Reason      = "Insufficient disk space"
                    }
                    $moveQueue.Add($queueItem)
                    $results.Queued += $queueItem
                    Write-Host "  Queued (low space): $($video.FileName) ($($video.FileSizeMB) MB)" -ForegroundColor Yellow
                }
            }
        }
    }
    
    # Process queued items if any files were moved (freeing up space)
    if ($moveQueue.Count -gt 0 -and $results.Moved.Count -gt 0) {
        Write-Host "`nRetrying queued files..." -ForegroundColor Cyan
        
        $retryQueue = [System.Collections.Generic.List[PSCustomObject]]::new($moveQueue)
        $moveQueue.Clear()
        
        foreach ($queueItem in $retryQueue) {
            if (Test-SufficientSpace -DestinationPath $DestinationRoot -RequiredBytes $queueItem.Video.FileSizeBytes) {
                $moveResult = Move-VideoFile -VideoInfo $queueItem.Video -DestinationFolder $queueItem.Destination -WhatIf:$WhatIf
                
                if ($moveResult.Success) {
                    $results.Moved += $moveResult
                    $results.Queued = $results.Queued | Where-Object { $_.Video.FullPath -ne $queueItem.Video.FullPath }
                    Write-Host "  Moved (from queue): $($queueItem.Video.FileName)" -ForegroundColor Green
                }
                else {
                    $results.Failed += $moveResult
                }
            }
            else {
                Write-Host "  Still queued: $($queueItem.Video.FileName)" -ForegroundColor Yellow
            }
        }
    }
    
    # Report remaining queued items
    if ($results.Queued.Count -gt 0) {
        Write-Host "`n[!] $($results.Queued.Count) file(s) remain queued due to insufficient disk space:" -ForegroundColor Yellow
        $totalQueuedSize = ($results.Queued.Video | Measure-Object -Property FileSizeBytes -Sum).Sum
        Write-Host "    Total size needed: $([math]::Round($totalQueuedSize / 1GB, 2)) GB" -ForegroundColor Yellow
        
        $driveInfo = Get-DriveInfo -Path $DestinationRoot
        if ($driveInfo) {
            Write-Host "    Available space: $($driveInfo.FreeGB) GB" -ForegroundColor Yellow
            Write-Host "    Space deficit: $([math]::Round(($totalQueuedSize - $driveInfo.FreeBytes) / 1GB, 2)) GB" -ForegroundColor Yellow
        }
    }
    
    return $results
}

function Invoke-MassDeletion {
    param(
        [array]$VideoList,
        [string]$MinResolution,
        [switch]$WhatIf,
        [switch]$Force
    )
    
    $minHeight = $ResolutionThresholds[$MinResolution]
    
    # Filter videos below minimum resolution
    $toDelete = $VideoList | Where-Object { $_.Height -lt $minHeight }
    
    if ($toDelete.Count -eq 0) {
        Write-Host "No videos found below $MinResolution resolution." -ForegroundColor Green
        return @{ Deleted = @(); Skipped = @(); Failed = @() }
    }
    
    $totalSize = ($toDelete | Measure-Object -Property FileSizeBytes -Sum).Sum
    $totalSizeGB = [math]::Round($totalSize / 1GB, 2)
    
    Write-Host "`nFiles to delete (below $MinResolution / ${minHeight}p):" -ForegroundColor Yellow
    Write-Host "  Count: $($toDelete.Count) file(s)" -ForegroundColor White
    Write-Host "  Total size: $totalSizeGB GB" -ForegroundColor White
    Write-Host ""
    
    # Group by resolution for summary
    $grouped = $toDelete | Group-Object ResolutionCategory | Sort-Object { $ResolutionThresholds[$_.Name] } -Descending
    foreach ($g in $grouped) {
        $groupSize = [math]::Round(($g.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
        Write-Host "  $($g.Name): $($g.Count) file(s), $groupSize GB" -ForegroundColor Gray
    }
    
    # Confirmation
    if (-not $WhatIf -and -not $Force) {
        Write-Host ""
        Write-Host "WARNING: This will permanently delete $($toDelete.Count) file(s) ($totalSizeGB GB)!" -ForegroundColor Red
        $confirm = Read-Host "Type 'DELETE' to confirm, or anything else to cancel"
        
        if ($confirm -ne "DELETE") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return @{ Deleted = @(); Skipped = $toDelete; Failed = @() }
        }
    }
    
    $results = @{
        Deleted = @()
        Skipped = @()
        Failed  = @()
    }
    
    $processed = 0
    foreach ($video in $toDelete) {
        $processed++
        $percent = [math]::Round(($processed / $toDelete.Count) * 100, 0)
        Write-Progress -Activity "Deleting videos" -Status "$processed of $($toDelete.Count)" -PercentComplete $percent
        
        $deleteResult = Remove-VideoFile -VideoInfo $video -WhatIf:$WhatIf
        
        if ($deleteResult.Success) {
            $results.Deleted += $deleteResult
            if (-not $WhatIf) {
                Write-Host "  Deleted: $($video.FileName) ($($video.Resolution))" -ForegroundColor Red
            }
        }
        else {
            $results.Failed += $deleteResult
        }
    }
    
    Write-Progress -Activity "Deleting videos" -Completed
    
    return $results
}

#endregion

#region Main Execution

Write-Host "Video Organizer" -ForegroundColor Cyan
Write-Host "===============" -ForegroundColor Cyan
Write-Host "Action: $Action" -ForegroundColor White

# Validate parameters
if ($Action -eq "Sort" -and -not $DestinationRoot) {
    Write-Error "DestinationRoot is required for Sort action."
    exit 1
}

if ($Action -eq "Delete" -and -not $MinResolution) {
    Write-Error "MinResolution is required for Delete action."
    exit 1
}

# Find FFprobe
if (-not (Find-FFprobe -CustomPath $FFprobePath)) {
    Write-Error "FFprobe not found. Please install FFmpeg."
    exit 1
}

Write-Host "Using FFprobe: $script:FFprobeCmd" -ForegroundColor DarkGray

# Resolve paths
$resolvedPaths = @()
foreach ($p in $Path) {
    try {
        $resolved = Resolve-Path $p -ErrorAction Stop
        $resolvedPaths += $resolved.Path
    }
    catch {
        Write-Warning "Path not found: $p"
    }
}

if ($resolvedPaths.Count -eq 0) {
    Write-Error "No valid paths specified."
    exit 1
}

# Show disk space info
if ($DestinationRoot) {
    $destInfo = Get-DriveInfo -Path $DestinationRoot
    if ($destInfo) {
        Write-Host "Destination drive free space: $($destInfo.FreeGB) GB" -ForegroundColor White
    }
}

# Find video files
Write-Host "`nScanning for video files..." -ForegroundColor White
$videoFiles = @()
foreach ($scanPath in $resolvedPaths) {
    $searchParams = @{
        Path    = $scanPath
        Include = $VideoExtensions
        File    = $true
    }
    if ($Recurse) {
        $searchParams.Recurse = $true
    }
    
    $found = Get-ChildItem @searchParams -ErrorAction SilentlyContinue
    if ($found) {
        $videoFiles += $found
    }
}

if ($videoFiles.Count -eq 0) {
    Write-Warning "No video files found."
    exit 0
}

Write-Host "Found $($videoFiles.Count) video file(s)" -ForegroundColor Green

# Analyze all videos
Write-Host "Analyzing video files..." -ForegroundColor White
$analyzedVideos = @()
$processed = 0

foreach ($file in $videoFiles) {
    $processed++
    $percent = [math]::Round(($processed / $videoFiles.Count) * 100, 0)
    Write-Progress -Activity "Analyzing videos" -Status "$processed of $($videoFiles.Count)" -PercentComplete $percent
    
    $details = Get-VideoDetails -FilePath $file.FullName
    if ($details) {
        $analyzedVideos += $details
    }
}

Write-Progress -Activity "Analyzing videos" -Completed

Write-Host "Successfully analyzed $($analyzedVideos.Count) video(s)" -ForegroundColor Green

# Show resolution summary
Write-Host "`nResolution Summary:" -ForegroundColor Cyan
$analyzedVideos | Group-Object ResolutionCategory | Sort-Object { 
    switch ($_.Name) {
        "4K UHD" { 1 }
        "1440p QHD" { 2 }
        "1080p FHD" { 3 }
        "720p HD" { 4 }
        "480p SD" { 5 }
        "360p" { 6 }
        "Low" { 7 }
        default { 99 }
    }
} | ForEach-Object {
    $groupSize = [math]::Round(($_.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
    Write-Host "  $($_.Name): $($_.Count) file(s), $groupSize GB" -ForegroundColor White
}

# Execute action
$isWhatIf = $Action -eq "Report"

switch ($Action) {
    "Sort" {
        Write-Host "`nSorting videos into resolution folders..." -ForegroundColor Cyan
        $results = Invoke-QueuedMoves -VideoList $analyzedVideos -DestinationRoot $DestinationRoot -WhatIf:$isWhatIf
        
        Write-Host "`n--- Sort Results ---" -ForegroundColor Cyan
        Write-Host "  Moved: $($results.Moved.Count)" -ForegroundColor Green
        Write-Host "  Queued: $($results.Queued.Count)" -ForegroundColor Yellow
        Write-Host "  Skipped: $($results.Skipped.Count)" -ForegroundColor Gray
        Write-Host "  Failed: $($results.Failed.Count)" -ForegroundColor Red
    }
    
    "Delete" {
        Write-Host "`nPreparing mass deletion..." -ForegroundColor Cyan
        $results = Invoke-MassDeletion -VideoList $analyzedVideos -MinResolution $MinResolution -WhatIf:$isWhatIf -Force:$Force
        
        $deletedSize = [math]::Round(($results.Deleted | Measure-Object -Property SizeBytes -Sum).Sum / 1GB, 2)
        
        Write-Host "`n--- Delete Results ---" -ForegroundColor Cyan
        Write-Host "  Deleted: $($results.Deleted.Count) file(s), $deletedSize GB freed" -ForegroundColor $(if ($isWhatIf) { "Yellow" } else { "Green" })
        Write-Host "  Failed: $($results.Failed.Count)" -ForegroundColor Red
    }
    
    "Report" {
        Write-Host "`n--- Report Mode (No Changes Made) ---" -ForegroundColor Yellow
        
        if ($MinResolution) {
            $minHeight = $ResolutionThresholds[$MinResolution]
            $belowThreshold = $analyzedVideos | Where-Object { $_.Height -lt $minHeight }
            $belowSize = [math]::Round(($belowThreshold | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
            
            Write-Host "`nVideos below $MinResolution (would be deleted):" -ForegroundColor Yellow
            Write-Host "  Count: $($belowThreshold.Count) file(s)" -ForegroundColor White
            Write-Host "  Size: $belowSize GB" -ForegroundColor White
            
            if ($belowThreshold.Count -gt 0 -and $belowThreshold.Count -le 20) {
                Write-Host "`n  Files:" -ForegroundColor Gray
                foreach ($v in $belowThreshold) {
                    Write-Host "    - $($v.FileName) ($($v.Resolution))" -ForegroundColor Gray
                }
            }
        }
        
        if ($DestinationRoot) {
            Write-Host "`nSort preview:" -ForegroundColor Yellow
            $analyzedVideos | Group-Object ResolutionCategory | ForEach-Object {
                $folderName = $ResolutionFolderNames[$_.Name]
                Write-Host "  $($_.Name) -> $DestinationRoot\$folderName ($($_.Count) files)" -ForegroundColor Gray
            }
        }
    }
}

Write-Host "`nDone!" -ForegroundColor Green

#endregion
