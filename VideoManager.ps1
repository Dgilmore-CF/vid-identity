<#
.SYNOPSIS
    Analyze, organize, and manage video files by resolution and codec.

.DESCRIPTION
    A comprehensive video file management tool that:
    - Analyzes video files and exports resolution/codec info to Excel
    - Sorts videos into resolution-based folders
    - Mass deletes videos below a resolution threshold
    - Intelligently handles disk space with move queuing
    
    Cross-platform compatible: Windows, macOS, and Linux.

.PARAMETER Path
    One or more directories to scan for video files. Defaults to current directory.

.PARAMETER Recurse
    Include subdirectories in the scan.

.PARAMETER Action
    The action to perform:
    - Analyze: Scan and export video info to spreadsheet (default)
    - Sort: Move videos into resolution-based folders
    - Delete: Delete videos below minimum resolution
    - Report: Dry-run preview of Sort or Delete actions

.PARAMETER OutputFile
    Path for Excel/CSV output (Analyze action). Defaults to VideoInfo.xlsx.

.PARAMETER DestinationRoot
    Root folder for resolution subfolders (Sort action).

.PARAMETER MinResolution
    Minimum resolution to keep (Delete/Report actions).
    Options: 4K, 1440p, 1080p, 720p, 480p, 360p

.PARAMETER Force
    Skip confirmation prompts for destructive operations.

.PARAMETER FFprobePath
    Optional path to ffprobe executable.

.EXAMPLE
    # Analyze videos and export to Excel
    .\VideoManager.ps1 -Path "D:\Videos" -Recurse

.EXAMPLE
    # Sort videos into resolution folders
    .\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -Recurse

.EXAMPLE
    # Delete all videos below 720p
    .\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -Recurse

.EXAMPLE
    # Preview what would be deleted (dry-run)
    .\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 1080p -Recurse

.NOTES
    Requires FFprobe (part of FFmpeg). Install via:
    - Windows: winget install FFmpeg
    - macOS:   brew install ffmpeg
    - Linux:   sudo apt install ffmpeg
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
    [string[]]$Path = @("."),
    
    [ValidateSet("Analyze", "Sort", "Delete", "Report")]
    [string]$Action = "Analyze",
    
    [switch]$Recurse,
    
    [string]$OutputFile,
    
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

function Test-FFprobe {
    try {
        $null = & $script:FFprobeCmd -version 2>&1
        return $true
    }
    catch {
        return $false
    }
}

function Get-VideoDetails {
    param(
        [string]$FilePath,
        [switch]$Extended
    )
    
    try {
        $json = & $script:FFprobeCmd -v quiet -print_format json -show_streams -show_format "$FilePath" 2>&1
        $info = $json | ConvertFrom-Json
        
        $videoStream = $info.streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
        $audioStream = $info.streams | Where-Object { $_.codec_type -eq "audio" } | Select-Object -First 1
        
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
        
        # Basic info for all actions
        $result = [PSCustomObject]@{
            FileName           = $fileInfo.Name
            FullPath           = $FilePath
            Directory          = $fileInfo.DirectoryName
            Width              = [int]$videoStream.width
            Height             = $height
            Resolution         = "$($videoStream.width)x$($videoStream.height)"
            ResolutionCategory = $resolutionCategory
            VideoCodec         = $videoStream.codec_name
            FileSizeBytes      = $fileInfo.Length
            FileSizeMB         = [math]::Round($fileInfo.Length / 1MB, 2)
            Drive              = if ($IsWindowsOS) { [System.IO.Path]::GetPathRoot($FilePath) } else { "/" }
        }
        
        # Extended info for Analyze action
        if ($Extended) {
            $durationSec = 0
            if ($info.format.duration) {
                $durationSec = [double]($info.format.duration)
            }
            $duration = [TimeSpan]::FromSeconds($durationSec)
            $durationStr = "{0:D2}:{1:D2}:{2:D2}" -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds
            
            $bitrateMbps = 0
            if ($info.format.bit_rate) {
                $bitrateMbps = [math]::Round([double]($info.format.bit_rate) / 1000000, 2)
            }
            
            $frameRate = $null
            if ($videoStream.r_frame_rate) {
                $parts = $videoStream.r_frame_rate -split "/"
                if ($parts.Count -eq 2 -and [int]$parts[1] -ne 0) {
                    $frameRate = [math]::Round([double]$parts[0] / [double]$parts[1], 2)
                }
            }
            
            $result | Add-Member -NotePropertyName "VideoCodecLong" -NotePropertyValue $videoStream.codec_long_name
            $result | Add-Member -NotePropertyName "AudioCodec" -NotePropertyValue $(if ($audioStream) { $audioStream.codec_name } else { "None" })
            $result | Add-Member -NotePropertyName "AudioCodecLong" -NotePropertyValue $(if ($audioStream) { $audioStream.codec_long_name } else { "None" })
            $result | Add-Member -NotePropertyName "Duration" -NotePropertyValue $durationStr
            $result | Add-Member -NotePropertyName "DurationSeconds" -NotePropertyValue ([math]::Round($durationSec, 2))
            $result | Add-Member -NotePropertyName "FrameRate" -NotePropertyValue $frameRate
            $result | Add-Member -NotePropertyName "BitrateMbps" -NotePropertyValue $bitrateMbps
            $result | Add-Member -NotePropertyName "PixelFormat" -NotePropertyValue $videoStream.pix_fmt
            $result | Add-Member -NotePropertyName "ColorSpace" -NotePropertyValue $videoStream.color_space
            $result | Add-Member -NotePropertyName "HDR" -NotePropertyValue $(if ($videoStream.color_transfer -match "smpte2084|arib-std-b67|bt2020") { "Yes" } else { "No" })
        }
        
        return $result
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
            if ($drive -and $drive.Free) {
                return [PSCustomObject]@{
                    Path       = $root
                    FreeBytes  = $drive.Free
                    FreeMB     = [math]::Round($drive.Free / 1MB, 2)
                    FreeGB     = [math]::Round($drive.Free / 1GB, 2)
                    TotalBytes = $drive.Free + $drive.Used
                }
            }
            
            $driveLetter = $root.TrimEnd('\')
            $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
            if ($disk) {
                return [PSCustomObject]@{
                    Path       = $root
                    FreeBytes  = $disk.FreeSpace
                    FreeMB     = [math]::Round($disk.FreeSpace / 1MB, 2)
                    FreeGB     = [math]::Round($disk.FreeSpace / 1GB, 2)
                    TotalBytes = $disk.Size
                }
            }
        }
        else {
            $dfOutput = & df -k $Path 2>&1 | Select-Object -Last 1
            $parts = $dfOutput -split '\s+' | Where-Object { $_ }
            
            if ($parts.Count -ge 4) {
                $totalKB = [long]$parts[1]
                $availKB = [long]$parts[3]
                
                return [PSCustomObject]@{
                    Path       = $Path
                    FreeBytes  = $availKB * 1024
                    FreeMB     = [math]::Round(($availKB * 1024) / 1MB, 2)
                    FreeGB     = [math]::Round(($availKB * 1024) / 1GB, 2)
                    TotalBytes = $totalKB * 1024
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
        return $true  # Assume OK if we can't check
    }
    
    $safetyMargin = $driveInfo.TotalBytes * ($SafetyMarginPercent / 100)
    $requiredWithMargin = $RequiredBytes + $safetyMargin
    
    return $driveInfo.FreeBytes -ge $requiredWithMargin
}

#endregion

#region Export Functions

function Export-ToExcel {
    param(
        [array]$Data,
        [string]$OutputPath
    )
    
    $useImportExcel = Get-Module -ListAvailable -Name ImportExcel
    
    if ($useImportExcel) {
        $Data | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Video Info"
        Write-Host "Excel file created: $OutputPath" -ForegroundColor Green
    }
    else {
        $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, ".csv")
        $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV file created: $csvPath" -ForegroundColor Yellow
        Write-Host "Tip: Install ImportExcel for .xlsx output: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor DarkGray
    }
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
        return [PSCustomObject]@{ Success = $true; WhatIf = $true }
    }
    
    try {
        if (-not (Test-Path $DestinationFolder)) {
            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
        }
        
        Move-Item -Path $VideoInfo.FullPath -Destination $destPath -Force
        return [PSCustomObject]@{ Success = $true; WhatIf = $false }
    }
    catch {
        Write-Warning "Failed to move $($VideoInfo.FileName): $_"
        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
    }
}

function Remove-VideoFile {
    param(
        [PSCustomObject]$VideoInfo,
        [switch]$WhatIf
    )
    
    if ($WhatIf) {
        Write-Host "  [WhatIf] Would delete: $($VideoInfo.FileName) ($($VideoInfo.FileSizeMB) MB)" -ForegroundColor DarkYellow
        return [PSCustomObject]@{ Success = $true; SizeBytes = $VideoInfo.FileSizeBytes; WhatIf = $true }
    }
    
    try {
        Remove-Item -Path $VideoInfo.FullPath -Force
        return [PSCustomObject]@{ Success = $true; SizeBytes = $VideoInfo.FileSizeBytes; WhatIf = $false }
    }
    catch {
        Write-Warning "Failed to delete $($VideoInfo.FileName): $_"
        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
    }
}

#endregion

#region Action Handlers

function Invoke-AnalyzeAction {
    param([array]$Videos, [string]$OutputPath)
    
    Write-Host "`nExporting $($Videos.Count) video(s) to spreadsheet..." -ForegroundColor White
    Export-ToExcel -Data $Videos -OutputPath $OutputPath
    
    Write-Host "`nResolution Summary:" -ForegroundColor Cyan
    $Videos | Group-Object ResolutionCategory | Sort-Object Count -Descending | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count) file(s)" -ForegroundColor White
    }
    
    Write-Host "`nCodec Summary:" -ForegroundColor Cyan
    $Videos | Group-Object VideoCodec | Sort-Object Count -Descending | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count) file(s)" -ForegroundColor White
    }
}

function Invoke-SortAction {
    param([array]$Videos, [string]$DestRoot, [switch]$WhatIf)
    
    $results = @{ Moved = 0; Queued = 0; Skipped = 0; Failed = 0 }
    $moveQueue = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    $groupedVideos = $Videos | Group-Object ResolutionCategory
    
    foreach ($group in $groupedVideos) {
        $folderName = $ResolutionFolderNames[$group.Name]
        $destFolder = Join-Path $DestRoot $folderName
        
        $totalSizeGB = [math]::Round(($group.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
        Write-Host "`nProcessing $($group.Name) -> $destFolder ($totalSizeGB GB)" -ForegroundColor Cyan
        
        foreach ($video in $group.Group) {
            $destPath = Join-Path $destFolder $video.FileName
            
            if ($video.FullPath -eq $destPath) {
                $results.Skipped++
                continue
            }
            
            $sourceDrive = $video.Drive
            $destDrive = if ($IsWindowsOS) { [System.IO.Path]::GetPathRoot($DestRoot) } else { "/" }
            $sameDrive = $sourceDrive -eq $destDrive
            
            if ($sameDrive -or (Test-SufficientSpace -DestinationPath $DestRoot -RequiredBytes $video.FileSizeBytes)) {
                $moveResult = Move-VideoFile -VideoInfo $video -DestinationFolder $destFolder -WhatIf:$WhatIf
                
                if ($moveResult.Success) {
                    $results.Moved++
                    if (-not $WhatIf) {
                        Write-Host "  Moved: $($video.FileName)" -ForegroundColor Green
                    }
                } else {
                    $results.Failed++
                }
            } else {
                $moveQueue.Add([PSCustomObject]@{ Video = $video; Destination = $destFolder })
                $results.Queued++
                Write-Host "  Queued (low space): $($video.FileName)" -ForegroundColor Yellow
            }
        }
    }
    
    # Retry queued items
    if ($moveQueue.Count -gt 0 -and $results.Moved -gt 0) {
        Write-Host "`nRetrying queued files..." -ForegroundColor Cyan
        
        foreach ($item in $moveQueue) {
            if (Test-SufficientSpace -DestinationPath $DestRoot -RequiredBytes $item.Video.FileSizeBytes) {
                $moveResult = Move-VideoFile -VideoInfo $item.Video -DestinationFolder $item.Destination -WhatIf:$WhatIf
                
                if ($moveResult.Success) {
                    $results.Moved++
                    $results.Queued--
                    Write-Host "  Moved (from queue): $($item.Video.FileName)" -ForegroundColor Green
                }
            }
        }
    }
    
    Write-Host "`n--- Sort Results ---" -ForegroundColor Cyan
    Write-Host "  Moved: $($results.Moved)" -ForegroundColor Green
    Write-Host "  Queued: $($results.Queued)" -ForegroundColor $(if ($results.Queued -gt 0) { "Yellow" } else { "Gray" })
    Write-Host "  Skipped: $($results.Skipped)" -ForegroundColor Gray
    Write-Host "  Failed: $($results.Failed)" -ForegroundColor $(if ($results.Failed -gt 0) { "Red" } else { "Gray" })
}

function Invoke-DeleteAction {
    param([array]$Videos, [string]$MinRes, [switch]$WhatIf, [switch]$ForceDelete)
    
    $minHeight = $ResolutionThresholds[$MinRes]
    $toDelete = $Videos | Where-Object { $_.Height -lt $minHeight }
    
    if ($toDelete.Count -eq 0) {
        Write-Host "`nNo videos found below $MinRes resolution." -ForegroundColor Green
        return
    }
    
    $totalSize = ($toDelete | Measure-Object -Property FileSizeBytes -Sum).Sum
    $totalSizeGB = [math]::Round($totalSize / 1GB, 2)
    
    Write-Host "`nFiles to delete (below $MinRes / ${minHeight}p):" -ForegroundColor Yellow
    Write-Host "  Count: $($toDelete.Count) file(s)" -ForegroundColor White
    Write-Host "  Total size: $totalSizeGB GB" -ForegroundColor White
    
    $toDelete | Group-Object ResolutionCategory | ForEach-Object {
        $groupSize = [math]::Round(($_.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
        Write-Host "    $($_.Name): $($_.Count) file(s), $groupSize GB" -ForegroundColor Gray
    }
    
    if (-not $WhatIf -and -not $ForceDelete) {
        Write-Host "`nWARNING: This will permanently delete $($toDelete.Count) file(s) ($totalSizeGB GB)!" -ForegroundColor Red
        $confirm = Read-Host "Type 'DELETE' to confirm"
        
        if ($confirm -ne "DELETE") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return
        }
    }
    
    $deleted = 0
    $failed = 0
    $freedBytes = 0
    
    foreach ($video in $toDelete) {
        $result = Remove-VideoFile -VideoInfo $video -WhatIf:$WhatIf
        
        if ($result.Success) {
            $deleted++
            $freedBytes += $result.SizeBytes
            if (-not $WhatIf) {
                Write-Host "  Deleted: $($video.FileName)" -ForegroundColor Red
            }
        } else {
            $failed++
        }
    }
    
    $freedGB = [math]::Round($freedBytes / 1GB, 2)
    Write-Host "`n--- Delete Results ---" -ForegroundColor Cyan
    Write-Host "  Deleted: $deleted file(s), $freedGB GB freed" -ForegroundColor $(if ($WhatIf) { "Yellow" } else { "Green" })
    if ($failed -gt 0) {
        Write-Host "  Failed: $failed" -ForegroundColor Red
    }
}

function Invoke-ReportAction {
    param([array]$Videos, [string]$MinRes, [string]$DestRoot)
    
    Write-Host "`n--- Report Mode (No Changes) ---" -ForegroundColor Yellow
    
    if ($MinRes) {
        $minHeight = $ResolutionThresholds[$MinRes]
        $belowThreshold = $Videos | Where-Object { $_.Height -lt $minHeight }
        $belowSize = [math]::Round(($belowThreshold | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
        
        Write-Host "`nVideos below $MinRes (would be deleted):" -ForegroundColor Yellow
        Write-Host "  Count: $($belowThreshold.Count) file(s)" -ForegroundColor White
        Write-Host "  Size: $belowSize GB" -ForegroundColor White
        
        if ($belowThreshold.Count -gt 0 -and $belowThreshold.Count -le 20) {
            foreach ($v in $belowThreshold) {
                Write-Host "    - $($v.FileName) ($($v.Resolution))" -ForegroundColor Gray
            }
        }
    }
    
    if ($DestRoot) {
        Write-Host "`nSort preview:" -ForegroundColor Yellow
        $Videos | Group-Object ResolutionCategory | ForEach-Object {
            $folderName = $ResolutionFolderNames[$_.Name]
            $sizeGB = [math]::Round(($_.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
            Write-Host "  $($_.Name) -> $DestRoot/$folderName ($($_.Count) files, $sizeGB GB)" -ForegroundColor Gray
        }
    }
}

#endregion

#region Main Execution

Write-Host "Video Manager" -ForegroundColor Cyan
Write-Host "=============" -ForegroundColor Cyan
Write-Host "Action: $Action" -ForegroundColor White

# Validate parameters
if ($Action -eq "Sort" -and -not $DestinationRoot) {
    Write-Error "DestinationRoot is required for Sort action. Use -DestinationRoot <path>"
    exit 1
}

if ($Action -eq "Delete" -and -not $MinResolution) {
    Write-Error "MinResolution is required for Delete action. Use -MinResolution <4K|1440p|1080p|720p|480p|360p>"
    exit 1
}

# Find FFprobe
if (-not (Find-FFprobe -CustomPath $FFprobePath)) {
    $msg = "FFprobe not found. "
    if ($IsWindowsOS) {
        $msg += "Install with: winget install FFmpeg"
    } else {
        $msg += "Install with: brew install ffmpeg (macOS) or sudo apt install ffmpeg (Linux)"
    }
    Write-Error $msg
    exit 1
}

if (-not (Test-FFprobe)) {
    Write-Error "FFprobe found but failed to execute."
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

# Show disk info for sort
if ($DestinationRoot) {
    $destInfo = Get-DriveInfo -Path $DestinationRoot
    if ($destInfo) {
        Write-Host "Destination free space: $($destInfo.FreeGB) GB" -ForegroundColor White
    }
}

# Find video files
Write-Host "`nScanning $($resolvedPaths.Count) path(s)..." -ForegroundColor White
$videoFiles = @()
foreach ($scanPath in $resolvedPaths) {
    $searchParams = @{ Path = $scanPath; Include = $VideoExtensions; File = $true }
    if ($Recurse) { $searchParams.Recurse = $true }
    
    $found = Get-ChildItem @searchParams -ErrorAction SilentlyContinue
    if ($found) { $videoFiles += $found }
}

if ($videoFiles.Count -eq 0) {
    Write-Warning "No video files found."
    exit 0
}

Write-Host "Found $($videoFiles.Count) video file(s)" -ForegroundColor Green

# Analyze videos
Write-Host "Analyzing..." -ForegroundColor White
$analyzedVideos = @()
$useExtended = $Action -eq "Analyze"
$processed = 0

foreach ($file in $videoFiles) {
    $processed++
    $percent = [math]::Round(($processed / $videoFiles.Count) * 100, 0)
    Write-Progress -Activity "Analyzing videos" -Status "$processed of $($videoFiles.Count)" -PercentComplete $percent
    
    $details = Get-VideoDetails -FilePath $file.FullName -Extended:$useExtended
    if ($details) {
        $analyzedVideos += $details
        if ($Action -eq "Analyze") {
            Write-Host "[$processed/$($videoFiles.Count)] $($file.Name) - $($details.Resolution) - $($details.VideoCodec)" -ForegroundColor Gray
        }
    }
}

Write-Progress -Activity "Analyzing videos" -Completed
Write-Host "Analyzed $($analyzedVideos.Count) video(s)" -ForegroundColor Green

# Show resolution summary
Write-Host "`nResolution Summary:" -ForegroundColor Cyan
$analyzedVideos | Group-Object ResolutionCategory | Sort-Object { $ResolutionThresholds[$_.Name] } -Descending | ForEach-Object {
    $sizeGB = [math]::Round(($_.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
    Write-Host "  $($_.Name): $($_.Count) file(s), $sizeGB GB" -ForegroundColor White
}

# Set default output file
if ($Action -eq "Analyze" -and -not $OutputFile) {
    $OutputFile = Join-Path $resolvedPaths[0] "VideoInfo.xlsx"
}

# Execute action
$isWhatIf = $Action -eq "Report"

switch ($Action) {
    "Analyze" { Invoke-AnalyzeAction -Videos $analyzedVideos -OutputPath $OutputFile }
    "Sort"    { Invoke-SortAction -Videos $analyzedVideos -DestRoot $DestinationRoot -WhatIf:$isWhatIf }
    "Delete"  { Invoke-DeleteAction -Videos $analyzedVideos -MinRes $MinResolution -WhatIf:$isWhatIf -ForceDelete:$Force }
    "Report"  { Invoke-ReportAction -Videos $analyzedVideos -MinRes $MinResolution -DestRoot $DestinationRoot }
}

Write-Host "`nDone!" -ForegroundColor Green

#endregion
