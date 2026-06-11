<#
.SYNOPSIS
    Analyze, organize, and manage video files by resolution and codec.

.DESCRIPTION
    A comprehensive video file management tool that:
    - Analyzes video files and exports resolution/codec info to Excel
    - Sorts videos into resolution-based folders
    - Mass deletes videos below a resolution threshold
    - Logs deleted files so higher-quality versions can be re-acquired
    - Intelligently handles disk space with move queuing
    - Offers interactive drive/volume selection

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

.PARAMETER SelectDrive
    Interactively select one or more drives/volumes to scan (and, for Sort, a single
    destination drive). Multiple scan drives can be chosen by entering a comma/space
    separated list of numbers, or 'all'.

.PARAMETER DeleteLog
    Path to the CSV deletion log written before any files are deleted. This log
    captures the metadata needed to re-acquire a higher-quality version of each
    removed file. Defaults to VideoManager_DeletedLog_<timestamp>.csv.

.EXAMPLE
    # Analyze videos and export to Excel
    .\VideoManager.ps1 -Path "D:\Videos" -Recurse

.EXAMPLE
    # Sort videos into resolution folders
    .\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -Recurse

.EXAMPLE
    # Delete all videos below 720p (logs removed files for re-acquisition)
    .\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -Recurse

.EXAMPLE
    # Interactively pick a drive to scan
    .\VideoManager.ps1 -SelectDrive -Recurse

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
    
    [string]$FFprobePath,

    [switch]$SelectDrive,

    [string]$DeleteLog
)

#region Configuration

$VideoExtensions = @(
    "*.mp4", "*.mkv", "*.avi", "*.mov", "*.wmv", "*.flv", "*.webm",
    "*.m4v", "*.mpg", "*.mpeg", "*.3gp", "*.3g2", "*.mts", "*.m2ts",
    "*.ts", "*.vob", "*.ogv", "*.divx", "*.xvid", "*.asf", "*.rm",
    "*.rmvb", "*.f4v", "*.hevc", "*.264", "*.265"
)

# Build a fast, case-insensitive lookup of bare extensions (".mp4", ".mkv", ...)
$VideoExtensionSet = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)
foreach ($ext in $VideoExtensions) { [void]$VideoExtensionSet.Add($ext.TrimStart('*')) }

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

# Numeric rank for resolution categories so summaries/sorts order correctly.
$ResolutionSortOrder = @{
    "4K UHD"    = 6
    "1440p QHD" = 5
    "1080p FHD" = 4
    "720p HD"   = 3
    "480p SD"   = 2
    "360p"      = 1
    "Low"       = 0
}

#endregion

#region Platform Detection

$IsWindowsOS = $false
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $IsWindowsOS = $IsWindows
} else {
    $IsWindowsOS = $true
}

# Cache volume identifiers per directory (avoids repeated df calls on Unix).
$script:VolumeCache = @{}

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
            Drive              = Get-VolumeId -Path $FilePath
        }
        
        # Extended info for Analyze / Delete / Report actions
        if ($Extended) {
            $durationSec = 0
            if ($info.format.duration) {
                $durationSec = [double]($info.format.duration)
            }
            $duration = [TimeSpan]::FromSeconds($durationSec)
            $durationStr = '{0:00}:{1:00}:{2:00}' -f [int][math]::Floor($duration.TotalHours), $duration.Minutes, $duration.Seconds
            
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
            
            # HDR: a true HDR transfer function is the reliable signal. bt2020 color
            # space alone is not sufficient (can be SDR wide-gamut).
            $isHdr = $videoStream.color_transfer -match "smpte2084|arib-std-b67"
            
            $result | Add-Member -NotePropertyName "VideoCodecLong" -NotePropertyValue $videoStream.codec_long_name
            $result | Add-Member -NotePropertyName "AudioCodec" -NotePropertyValue $(if ($audioStream) { $audioStream.codec_name } else { "None" })
            $result | Add-Member -NotePropertyName "AudioCodecLong" -NotePropertyValue $(if ($audioStream) { $audioStream.codec_long_name } else { "None" })
            $result | Add-Member -NotePropertyName "AudioChannels" -NotePropertyValue $(if ($audioStream) { $audioStream.channels } else { $null })
            $result | Add-Member -NotePropertyName "Duration" -NotePropertyValue $durationStr
            $result | Add-Member -NotePropertyName "DurationSeconds" -NotePropertyValue ([math]::Round($durationSec, 2))
            $result | Add-Member -NotePropertyName "FrameRate" -NotePropertyValue $frameRate
            $result | Add-Member -NotePropertyName "BitrateMbps" -NotePropertyValue $bitrateMbps
            $result | Add-Member -NotePropertyName "PixelFormat" -NotePropertyValue $videoStream.pix_fmt
            $result | Add-Member -NotePropertyName "ColorSpace" -NotePropertyValue $videoStream.color_space
            $result | Add-Member -NotePropertyName "HDR" -NotePropertyValue $(if ($isHdr) { "Yes" } else { "No" })
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

function Get-VolumeId {
    <#
        Returns a stable identifier for the volume/filesystem that a path lives on.
        Windows: the path root (e.g. "D:\").
        Unix:    the backing filesystem device reported by df (distinguishes mounts).
    #>
    param([string]$Path)
    
    if ($IsWindowsOS) {
        return [System.IO.Path]::GetPathRoot($Path)
    }
    
    # Resolve to an existing directory for df (file may sit in a dir that exists).
    $lookup = $Path
    if (-not (Test-Path $lookup)) {
        $lookup = Split-Path $Path -Parent
        if (-not $lookup) { $lookup = "/" }
    }
    
    if ($script:VolumeCache.ContainsKey($lookup)) {
        return $script:VolumeCache[$lookup]
    }
    
    $volId = "/"
    try {
        $dfOutput = & df -k -P $lookup 2>$null | Select-Object -Last 1
        $parts = $dfOutput -split '\s+' | Where-Object { $_ }
        if ($parts.Count -ge 1) {
            # First column is the filesystem device; uniquely identifies the volume.
            $volId = $parts[0]
        }
    }
    catch { }
    
    $script:VolumeCache[$lookup] = $volId
    return $volId
}

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
            $dfOutput = & df -k -P $Path 2>&1 | Select-Object -Last 1
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

function Get-AvailableDrives {
    <# Enumerate mounted filesystem drives/volumes for interactive selection. #>
    $drives = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    if ($IsWindowsOS) {
        Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | ForEach-Object {
            $free = if ($_.Free) { [math]::Round($_.Free / 1GB, 2) } else { $null }
            $total = if ($_.Free -and $_.Used) { [math]::Round(($_.Free + $_.Used) / 1GB, 2) } else { $null }
            $drives.Add([PSCustomObject]@{
                Name   = $_.Name
                Root   = $_.Root
                FreeGB = $free
                TotalGB = $total
            })
        }
    }
    else {
        $lines = & df -k -P 2>$null | Select-Object -Skip 1
        foreach ($line in $lines) {
            $parts = $line -split '\s+' | Where-Object { $_ }
            if ($parts.Count -ge 6) {
                $availKB = [long]$parts[3]
                $totalKB = [long]$parts[1]
                # Mount point is everything after the 5th column (may contain spaces).
                $mount = ($parts[5..($parts.Count - 1)] -join ' ')
                # Skip pseudo-filesystems that aren't useful targets.
                if ($parts[0] -match '^(devfs|tmpfs|map|none|overlay)$') { continue }
                $drives.Add([PSCustomObject]@{
                    Name    = $parts[0]
                    Root    = $mount
                    FreeGB  = [math]::Round(($availKB * 1024) / 1GB, 2)
                    TotalGB = [math]::Round(($totalKB * 1024) / 1GB, 2)
                })
            }
        }
    }
    
    return $drives
}

function Select-Drive {
    <#
        Interactively select one or more drives/volumes.
        Returns an array of selected roots (use -Multiple to allow more than one).
        Returns $null if nothing is selected.
    #>
    param(
        [string]$Prompt = "Select a drive",
        [switch]$Multiple
    )
    
    $drives = @(Get-AvailableDrives)
    if ($drives.Count -eq 0) {
        Write-Warning "No drives detected."
        return $null
    }
    
    Write-Host "`n$Prompt :" -ForegroundColor Cyan
    for ($i = 0; $i -lt $drives.Count; $i++) {
        $d = $drives[$i]
        $freeStr = if ($null -ne $d.FreeGB) { "$($d.FreeGB) GB free" } else { "size unknown" }
        Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $d.Root, $freeStr) -ForegroundColor White
    }
    
    if ($Multiple) {
        Write-Host "Enter one or more numbers separated by commas or spaces (e.g. 1,3,4)," -ForegroundColor DarkGray
        Write-Host "type 'all' to select every drive, or leave blank to cancel." -ForegroundColor DarkGray
    }
    
    while ($true) {
        $promptText = if ($Multiple) { "Selection" } else { "Enter number (1-$($drives.Count)) or blank to cancel" }
        $choice = Read-Host $promptText
        if ([string]::IsNullOrWhiteSpace($choice)) { return $null }
        
        if (-not $Multiple) {
            $index = 0
            if ([int]::TryParse($choice, [ref]$index) -and $index -ge 1 -and $index -le $drives.Count) {
                return @($drives[$index - 1].Root)
            }
            Write-Host "Invalid selection. Try again." -ForegroundColor Yellow
            continue
        }
        
        # Multiple selection: accept "all" or a comma/space-separated list of numbers.
        if ($choice.Trim() -eq "all") {
            return @($drives | ForEach-Object { $_.Root })
        }
        
        $tokens = $choice -split '[,\s]+' | Where-Object { $_ }
        $selected = [System.Collections.Generic.List[string]]::new()
        $valid = $true
        
        foreach ($token in $tokens) {
            $index = 0
            if ([int]::TryParse($token, [ref]$index) -and $index -ge 1 -and $index -le $drives.Count) {
                $root = $drives[$index - 1].Root
                if (-not $selected.Contains($root)) { $selected.Add($root) }
            }
            else {
                Write-Host "Invalid entry: '$token'. Use numbers 1-$($drives.Count), 'all', or blank to cancel." -ForegroundColor Yellow
                $valid = $false
                break
            }
        }
        
        if ($valid -and $selected.Count -gt 0) {
            return @($selected)
        }
    }
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
    
    # Ensure the module is actually loaded (availability alone is not enough).
    if (-not (Get-Command Export-Excel -ErrorAction SilentlyContinue)) {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
        }
    }
    
    if (Get-Command Export-Excel -ErrorAction SilentlyContinue) {
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

function Write-DeletionLog {
    <#
        Records the metadata of files that are about to be (or would be) deleted so a
        higher-quality replacement can be sourced later. Written BEFORE deletion.
    #>
    param(
        [array]$Items,
        [string]$LogPath,
        [switch]$Preview
    )
    
    if (-not $Items -or $Items.Count -eq 0) { return }
    
    $records = $Items | ForEach-Object {
        [PSCustomObject]@{
            FileName           = $_.FileName
            SearchTitle        = [System.IO.Path]::GetFileNameWithoutExtension($_.FileName)
            Resolution         = $_.Resolution
            ResolutionCategory = $_.ResolutionCategory
            Width              = $_.Width
            Height             = $_.Height
            VideoCodec         = $_.VideoCodec
            Duration           = $_.Duration
            DurationSeconds    = $_.DurationSeconds
            FrameRate          = $_.FrameRate
            BitrateMbps        = $_.BitrateMbps
            FileSizeMB         = $_.FileSizeMB
            OriginalDirectory  = $_.Directory
            OriginalFullPath   = $_.FullPath
            LoggedAt           = (Get-Date).ToString("s")
        }
    }
    
    try {
        $records | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
        $label = if ($Preview) { "Re-acquisition preview log" } else { "Deletion log" }
        Write-Host "$label written: $LogPath" -ForegroundColor Cyan
        Write-Host "  Use this list to source higher-quality replacements." -ForegroundColor DarkGray
    }
    catch {
        Write-Warning "Failed to write deletion log to $LogPath : $_"
    }
}

#endregion

#region File Operations

function Move-VideoFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [PSCustomObject]$VideoInfo,
        [string]$DestinationFolder
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
    
    if (-not $PSCmdlet.ShouldProcess("$($VideoInfo.FileName) -> $DestinationFolder", "Move")) {
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
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [PSCustomObject]$VideoInfo
    )
    
    if (-not $PSCmdlet.ShouldProcess("$($VideoInfo.FileName) ($($VideoInfo.FileSizeMB) MB)", "Delete")) {
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
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([array]$Videos, [string]$DestRoot, [switch]$WhatIf)
    
    $results = @{ Moved = 0; Queued = 0; Skipped = 0; Failed = 0 }
    $moveQueue = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    $destDrive = Get-VolumeId -Path $DestRoot
    $groupedVideos = $Videos | Group-Object ResolutionCategory
    
    foreach ($group in $groupedVideos) {
        $folderName = $ResolutionFolderNames[$group.Name]
        if (-not $folderName) {
            Write-Warning "Unknown resolution category '$($group.Name)'; using 'Low_Resolution'."
            $folderName = "Low_Resolution"
        }
        $destFolder = Join-Path $DestRoot $folderName
        
        $totalSizeGB = [math]::Round(($group.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
        Write-Host "`nProcessing $($group.Name) -> $destFolder ($totalSizeGB GB)" -ForegroundColor Cyan
        
        foreach ($video in $group.Group) {
            $destPath = Join-Path $destFolder $video.FileName
            
            if ($video.FullPath -eq $destPath) {
                $results.Skipped++
                continue
            }
            
            $sameDrive = $video.Drive -eq $destDrive
            
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
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([array]$Videos, [string]$MinRes, [switch]$WhatIf, [switch]$ForceDelete, [string]$LogPath)
    
    $minHeight = $ResolutionThresholds[$MinRes]
    $toDelete = @($Videos | Where-Object { $_.Height -lt $minHeight })
    
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
    
    # Always log what is about to be removed BEFORE deleting, so a higher-quality
    # version can be re-acquired later.
    Write-DeletionLog -Items $toDelete -LogPath $LogPath -Preview:$WhatIf
    
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
    param([array]$Videos, [string]$MinRes, [string]$DestRoot, [string]$LogPath)
    
    Write-Host "`n--- Report Mode (No Changes) ---" -ForegroundColor Yellow
    
    if (-not $MinRes -and -not $DestRoot) {
        Write-Host "`nNothing to preview. Provide -MinResolution (delete preview)" -ForegroundColor Yellow
        Write-Host "and/or -DestinationRoot (sort preview)." -ForegroundColor Yellow
        return
    }
    
    if ($MinRes) {
        $minHeight = $ResolutionThresholds[$MinRes]
        $belowThreshold = @($Videos | Where-Object { $_.Height -lt $minHeight })
        $belowSize = [math]::Round(($belowThreshold | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
        
        Write-Host "`nVideos below $MinRes (would be deleted):" -ForegroundColor Yellow
        Write-Host "  Count: $($belowThreshold.Count) file(s)" -ForegroundColor White
        Write-Host "  Size: $belowSize GB" -ForegroundColor White
        
        if ($belowThreshold.Count -gt 0 -and $belowThreshold.Count -le 20) {
            foreach ($v in $belowThreshold) {
                Write-Host "    - $($v.FileName) ($($v.Resolution))" -ForegroundColor Gray
            }
        }
        
        # Write a re-acquisition preview log for the would-be-deleted set.
        if ($belowThreshold.Count -gt 0) {
            Write-DeletionLog -Items $belowThreshold -LogPath $LogPath -Preview
        }
    }
    
    if ($DestRoot) {
        Write-Host "`nSort preview:" -ForegroundColor Yellow
        $Videos | Group-Object ResolutionCategory | Sort-Object { $ResolutionSortOrder[$_.Name] } -Descending | ForEach-Object {
            $folderName = $ResolutionFolderNames[$_.Name]
            if (-not $folderName) { $folderName = "Low_Resolution" }
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

# Interactive drive selection
if ($SelectDrive) {
    $selectedScan = @(Select-Drive -Prompt "Select one or more drives/volumes to scan" -Multiple)
    if ($selectedScan.Count -gt 0) {
        $Path = $selectedScan
        Write-Host "Scan drive(s): $($selectedScan -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "No drive selected; using default path." -ForegroundColor Yellow
    }
    
    # Destination must be a single volume for sorting.
    if ($Action -eq "Sort" -and -not $DestinationRoot) {
        $selectedDest = @(Select-Drive -Prompt "Select a destination drive/volume for sorted videos")
        if ($selectedDest.Count -gt 0) {
            $DestinationRoot = $selectedDest[0]
            Write-Host "Destination drive: $DestinationRoot" -ForegroundColor Green
        }
    }
}

# Validate parameters
if ($Action -eq "Sort" -and -not $DestinationRoot) {
    Write-Error "DestinationRoot is required for Sort action. Use -DestinationRoot <path> (or -SelectDrive)."
    exit 1
}

if ($Action -eq "Delete" -and -not $MinResolution) {
    Write-Error "MinResolution is required for Delete action. Use -MinResolution <4K|1440p|1080p|720p|480p|360p>"
    exit 1
}

# Validate / prepare DestinationRoot for Sort
if ($Action -eq "Sort" -and $DestinationRoot) {
    if (-not (Test-Path $DestinationRoot)) {
        if ($PSCmdlet.ShouldProcess($DestinationRoot, "Create destination root directory")) {
            try {
                New-Item -Path $DestinationRoot -ItemType Directory -Force | Out-Null
                Write-Host "Created destination root: $DestinationRoot" -ForegroundColor Green
            }
            catch {
                Write-Error "Could not create DestinationRoot '$DestinationRoot': $_"
                exit 1
            }
        }
    }
    elseif (-not (Test-Path $DestinationRoot -PathType Container)) {
        Write-Error "DestinationRoot '$DestinationRoot' exists but is not a directory."
        exit 1
    }
}

# Resolve default deletion log path
if (-not $DeleteLog) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $DeleteLog = Join-Path (Get-Location).Path "VideoManager_DeletedLog_$timestamp.csv"
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

# Find video files (extension filtering via HashSet works on PS 5.1 and 7+,
# recursive or not, unlike Get-ChildItem -Include on a bare directory path).
Write-Host "`nScanning $($resolvedPaths.Count) path(s)..." -ForegroundColor White
$videoFiles = @()
foreach ($scanPath in $resolvedPaths) {
    $found = Get-ChildItem -Path $scanPath -File -Recurse:$Recurse -ErrorAction SilentlyContinue |
        Where-Object { $VideoExtensionSet.Contains($_.Extension) }
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
# Extended metadata is needed for Analyze output and for the deletion/re-acquisition
# log (duration, codec, bitrate help find a better replacement).
$useExtended = $Action -in @("Analyze", "Delete", "Report")
$processed = 0
$failedCount = 0

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
    else {
        $failedCount++
    }
}

Write-Progress -Activity "Analyzing videos" -Completed
Write-Host "Analyzed $($analyzedVideos.Count) video(s)" -ForegroundColor Green
if ($failedCount -gt 0) {
    Write-Host "Failed to analyze $failedCount file(s) (skipped; not a valid video or unreadable)." -ForegroundColor Yellow
}

if ($analyzedVideos.Count -eq 0) {
    Write-Warning "No analyzable videos found."
    exit 0
}

# Show resolution summary (ranked highest -> lowest)
Write-Host "`nResolution Summary:" -ForegroundColor Cyan
$analyzedVideos | Group-Object ResolutionCategory | Sort-Object { $ResolutionSortOrder[$_.Name] } -Descending | ForEach-Object {
    $sizeGB = [math]::Round(($_.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
    Write-Host "  $($_.Name): $($_.Count) file(s), $sizeGB GB" -ForegroundColor White
}

# Set default output file
if ($Action -eq "Analyze" -and -not $OutputFile) {
    $OutputFile = Join-Path $resolvedPaths[0] "VideoInfo.xlsx"
}

# Execute action
$isWhatIf = ($Action -eq "Report") -or $WhatIfPreference

switch ($Action) {
    "Analyze" { Invoke-AnalyzeAction -Videos $analyzedVideos -OutputPath $OutputFile }
    "Sort"    { Invoke-SortAction -Videos $analyzedVideos -DestRoot $DestinationRoot -WhatIf:$isWhatIf }
    "Delete"  { Invoke-DeleteAction -Videos $analyzedVideos -MinRes $MinResolution -WhatIf:$isWhatIf -ForceDelete:$Force -LogPath $DeleteLog }
    "Report"  { Invoke-ReportAction -Videos $analyzedVideos -MinRes $MinResolution -DestRoot $DestinationRoot -LogPath $DeleteLog }
}

Write-Host "`nDone!" -ForegroundColor Green

#endregion
