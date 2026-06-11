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
    Root folder for resolution subfolders (Sort action). Used as the fallback
    destination for any quality not explicitly mapped via -QualityMap.

.PARAMETER QualityMap
    Hashtable mapping resolution qualities to destination roots so different
    qualities can be sorted onto different drives in a single session. Keys may
    be the short names (4K, 1440p, 1080p, 720p, 480p, 360p, Low) or the full
    category names (e.g. "4K UHD"). Example:
        -QualityMap @{ "4K" = "D:\4K"; "1080p" = "E:\HD"; "720p" = "E:\HD" }
    When a target drive runs low on space, files intelligently overflow to the
    destination drive with the most available space (largest files placed first).

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
    # Sort different qualities to different drives, overflowing intelligently
    .\VideoManager.ps1 -Path "D:\Videos" -Action Sort -Recurse `
        -QualityMap @{ "4K" = "X:\4K"; "1080p" = "Y:\HD"; "720p" = "Y:\HD" } `
        -DestinationRoot "Z:\Other"

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
    
    [hashtable]$QualityMap,
    
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

# Map between short quality names (used in -QualityMap) and full category names.
$ShortToCategory = @{
    "4K"    = "4K UHD"
    "1440p" = "1440p QHD"
    "1080p" = "1080p FHD"
    "720p"  = "720p HD"
    "480p"  = "480p SD"
    "360p"  = "360p"
    "Low"   = "Low"
}

$CategoryToShort = @{
    "4K UHD"    = "4K"
    "1440p QHD" = "1440p"
    "1080p FHD" = "1080p"
    "720p HD"   = "720p"
    "480p SD"   = "480p"
    "360p"      = "360p"
    "Low"       = "Low"
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

function Resolve-OutputPath {
    <#
        Turn a user-supplied output path into an absolute file path.
        - Relative paths resolve against the current directory.
        - A directory (existing, or trailing separator) gets a default filename.
        Returns the absolute path (file is not created here).
    #>
    param(
        [string]$Path,
        [string]$DefaultFileName = "VideoInfo.xlsx"
    )
    
    if ([string]::IsNullOrWhiteSpace($Path)) { $Path = $DefaultFileName }
    
    # Treat as a folder if it exists as a container or ends with a separator.
    $endsWithSep = $Path -match '[\\/]\s*$'
    if ($endsWithSep -or (Test-Path $Path -PathType Container)) {
        $Path = Join-Path $Path $DefaultFileName
    }
    
    # Make absolute relative to the current working directory.
    if (-not [System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path (Get-Location).Path $Path
    }
    
    return [System.IO.Path]::GetFullPath($Path)
}

function Export-ToExcel {
    param(
        [array]$Data,
        [string]$OutputPath
    )
    
    $OutputPath = Resolve-OutputPath -Path $OutputPath
    
    # Ensure the destination directory exists.
    $outDir = Split-Path $OutputPath -Parent
    if ($outDir -and -not (Test-Path $outDir)) {
        try { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
        catch {
            Write-Warning "Could not create output directory '$outDir': $_"
            return
        }
    }
    
    # Ensure the module is actually loaded (availability alone is not enough).
    if (-not (Get-Command Export-Excel -ErrorAction SilentlyContinue)) {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
        }
    }
    
    if (Get-Command Export-Excel -ErrorAction SilentlyContinue) {
        try {
            $Data | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Video Info"
            Write-Host "Excel file saved to: $OutputPath" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to write Excel file '$OutputPath': $_"
        }
    }
    else {
        $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, ".csv")
        try {
            $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "ImportExcel module not found - saved CSV instead of XLSX." -ForegroundColor Yellow
            Write-Host "CSV file saved to: $csvPath" -ForegroundColor Yellow
            Write-Host "Tip: for .xlsx output run: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor DarkGray
        }
        catch {
            Write-Warning "Failed to write CSV file '$csvPath': $_"
        }
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

function Resolve-CategoryRootMap {
    <#
        Build a map of full resolution category -> destination root for the
        categories actually present. -QualityMap keys may be short names
        (4K, 1080p, ...) or full categories ("4K UHD"). Unmapped categories fall
        back to $DefaultRoot. Returns the map plus any categories left unmapped.
    #>
    param(
        [hashtable]$QualityMap,
        [string]$DefaultRoot,
        [string[]]$Categories
    )
    
    $map = @{}
    $unmapped = [System.Collections.Generic.List[string]]::new()
    
    foreach ($cat in $Categories) {
        $short = $CategoryToShort[$cat]
        $root = $null
        
        if ($QualityMap) {
            if ($short -and $QualityMap.ContainsKey($short)) {
                $root = $QualityMap[$short]
            }
            elseif ($QualityMap.ContainsKey($cat)) {
                $root = $QualityMap[$cat]
            }
        }
        
        if (-not $root) { $root = $DefaultRoot }
        
        if ($root) {
            $map[$cat] = $root
        }
        else {
            $unmapped.Add($cat)
        }
    }
    
    return [PSCustomObject]@{ Map = $map; Unmapped = @($unmapped) }
}

function Get-InteractiveQualityAssignment {
    <#
        Interactively assign each present resolution category to a drive/volume.
        Returns a hashtable of full category -> destination root, or $null if
        the user cancels.
    #>
    param([string[]]$Categories)
    
    $drives = @(Get-AvailableDrives)
    if ($drives.Count -eq 0) {
        Write-Warning "No drives detected for assignment."
        return $null
    }
    
    Write-Host "`nAssign each quality to a destination drive:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $drives.Count; $i++) {
        $d = $drives[$i]
        $freeStr = if ($null -ne $d.FreeGB) { "$($d.FreeGB) GB free" } else { "size unknown" }
        Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $d.Root, $freeStr) -ForegroundColor White
    }
    Write-Host "Tip: multiple qualities can share the same drive." -ForegroundColor DarkGray
    
    $map = @{}
    $lastIndex = 1
    
    foreach ($cat in $Categories) {
        while ($true) {
            $choice = Read-Host "Drive number for [$cat] (default $lastIndex)"
            if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "$lastIndex" }
            
            $index = 0
            if ([int]::TryParse($choice, [ref]$index) -and $index -ge 1 -and $index -le $drives.Count) {
                $map[$cat] = $drives[$index - 1].Root
                $lastIndex = $index
                break
            }
            Write-Host "Invalid selection. Enter 1-$($drives.Count)." -ForegroundColor Yellow
        }
    }
    
    return $map
}

function New-SortPlan {
    <#
        Produce a placement plan for sorting videos onto (potentially) multiple
        drives. Honours the preferred drive per quality, and intelligently
        overflows to the destination drive with the most available space when a
        preferred drive is full. Largest files are placed first (first-fit
        decreasing) for better packing. Same-volume moves consume no extra space.
    #>
    param(
        [array]$Videos,
        [hashtable]$CategoryRootMap,
        [double]$SafetyMarginPercent = 5
    )
    
    $roots = @($CategoryRootMap.Values | Select-Object -Unique)
    
    # Gather per-volume capacity once; track projected free space as we plan.
    $rootVol = @{}
    $volInfo = @{}
    foreach ($root in $roots) {
        $vid = Get-VolumeId -Path $root
        $rootVol[$root] = $vid
        if (-not $volInfo.ContainsKey($vid)) {
            $di = Get-DriveInfo -Path $root
            if ($di) {
                $free = [double]$di.FreeBytes
                $total = [double]$di.TotalBytes
                $reserve = $total * ($SafetyMarginPercent / 100)
            }
            else {
                # Unknown capacity: assume effectively unlimited (matches prior behaviour).
                $free = [double]::MaxValue
                $total = 0
                $reserve = 0
            }
            $volInfo[$vid] = [PSCustomObject]@{
                VolumeId      = $vid
                FreeBytes     = $free
                TotalBytes    = $total
                Reserve       = $reserve
                ProjectedFree = $free
                Roots         = [System.Collections.Generic.List[string]]::new()
            }
        }
        $volInfo[$vid].Roots.Add($root)
    }
    
    $assignments = [System.Collections.Generic.List[PSCustomObject]]::new()
    $unplaceable = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    # First-fit decreasing: place the largest files while space is plentiful.
    $ordered = $Videos | Sort-Object FileSizeBytes -Descending
    
    foreach ($video in $ordered) {
        $category = $video.ResolutionCategory
        $preferredRoot = $CategoryRootMap[$category]
        if (-not $preferredRoot) { $preferredRoot = $roots | Select-Object -First 1 }
        
        # Try preferred drive first, then the remaining drives by most free space.
        $others = @($roots |
            Where-Object { $_ -ne $preferredRoot } |
            Sort-Object { $volInfo[$rootVol[$_]].ProjectedFree } -Descending)
        $candidateRoots = @($preferredRoot) + $others
        
        $placed = $false
        foreach ($cand in $candidateRoots) {
            $vid = $rootVol[$cand]
            $vi = $volInfo[$vid]
            $isOverflow = ($cand -ne $preferredRoot)
            $sameVolAsSource = ($video.Drive -eq $vid)
            
            if ($sameVolAsSource) {
                # Move within the same volume: instant, no extra space used.
                $assignments.Add([PSCustomObject]@{
                    Video      = $video
                    Root       = $cand
                    VolumeId   = $vid
                    IsOverflow = $isOverflow
                    SameVolume = $true
                })
                $placed = $true
                break
            }
            
            if (($vi.ProjectedFree - $video.FileSizeBytes) -ge $vi.Reserve) {
                $vi.ProjectedFree -= $video.FileSizeBytes
                $assignments.Add([PSCustomObject]@{
                    Video      = $video
                    Root       = $cand
                    VolumeId   = $vid
                    IsOverflow = $isOverflow
                    SameVolume = $false
                })
                $placed = $true
                break
            }
        }
        
        if (-not $placed) { $unplaceable.Add($video) }
    }
    
    return [PSCustomObject]@{
        Assignments = $assignments
        Unplaceable = $unplaceable
        Volumes     = @($volInfo.Values)
        RootVolume  = $rootVol
    }
}

function Invoke-SortAction {
    param([array]$Videos, [hashtable]$CategoryRootMap, [switch]$WhatIf)
    
    $results = @{ Moved = 0; Skipped = 0; Failed = 0; Overflowed = 0; Unplaceable = 0 }
    
    $plan = New-SortPlan -Videos $Videos -CategoryRootMap $CategoryRootMap
    
    # Pre-create distinct destination roots (real runs only).
    if (-not $WhatIf) {
        foreach ($root in @($CategoryRootMap.Values | Select-Object -Unique)) {
            if (-not (Test-Path $root)) {
                try { New-Item -Path $root -ItemType Directory -Force | Out-Null }
                catch { Write-Warning "Could not create destination root '$root': $_" }
            }
        }
    }
    
    # Group assignments by destination volume for readable, batched output.
    $byVolume = $plan.Assignments | Group-Object VolumeId
    
    foreach ($volGroup in $byVolume) {
        $sample = $volGroup.Group[0]
        $volSizeGB = [math]::Round(($volGroup.Group | ForEach-Object { $_.Video.FileSizeBytes } | Measure-Object -Sum).Sum / 1GB, 2)
        Write-Host "`nVolume $($volGroup.Name) <- $($volGroup.Count) file(s), $volSizeGB GB" -ForegroundColor Cyan
        
        foreach ($a in $volGroup.Group) {
            $video = $a.Video
            $folderName = $ResolutionFolderNames[$video.ResolutionCategory]
            if (-not $folderName) { $folderName = "Low_Resolution" }
            $destFolder = Join-Path $a.Root $folderName
            $destPath = Join-Path $destFolder $video.FileName
            
            if ($video.FullPath -eq $destPath) {
                $results.Skipped++
                continue
            }
            
            $moveResult = Move-VideoFile -VideoInfo $video -DestinationFolder $destFolder -WhatIf:$WhatIf
            
            if ($moveResult.Success) {
                $results.Moved++
                if ($a.IsOverflow) { $results.Overflowed++ }
                $tag = if ($a.IsOverflow) { " [overflow]" } else { "" }
                if (-not $WhatIf) {
                    Write-Host "  Moved$tag`: $($video.FileName) -> $folderName" -ForegroundColor $(if ($a.IsOverflow) { "Yellow" } else { "Green" })
                }
            }
            else {
                $results.Failed++
            }
        }
    }
    
    # Report anything that could not be placed on any destination drive.
    if ($plan.Unplaceable.Count -gt 0) {
        $results.Unplaceable = $plan.Unplaceable.Count
        $unSizeGB = [math]::Round(($plan.Unplaceable | ForEach-Object { $_.FileSizeBytes } | Measure-Object -Sum).Sum / 1GB, 2)
        Write-Host "`nUnplaceable (no drive has room): $($plan.Unplaceable.Count) file(s), $unSizeGB GB" -ForegroundColor Red
        foreach ($v in $plan.Unplaceable | Select-Object -First 20) {
            Write-Host "  - $($v.FileName) ($($v.FileSizeMB) MB, $($v.ResolutionCategory))" -ForegroundColor DarkYellow
        }
        Write-Host "  Free up space or add another drive to -QualityMap." -ForegroundColor DarkGray
    }
    
    Write-Host "`n--- Sort Results ---" -ForegroundColor Cyan
    Write-Host "  Moved: $($results.Moved) (of which overflow: $($results.Overflowed))" -ForegroundColor Green
    Write-Host "  Skipped (already in place): $($results.Skipped)" -ForegroundColor Gray
    Write-Host "  Unplaceable: $($results.Unplaceable)" -ForegroundColor $(if ($results.Unplaceable -gt 0) { "Red" } else { "Gray" })
    Write-Host "  Failed: $($results.Failed)" -ForegroundColor $(if ($results.Failed -gt 0) { "Red" } else { "Gray" })
}

function Invoke-DeleteAction {
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
    param([array]$Videos, [string]$MinRes, [hashtable]$CategoryRootMap, [string]$LogPath)
    
    Write-Host "`n--- Report Mode (No Changes) ---" -ForegroundColor Yellow
    
    $hasSort = $CategoryRootMap -and $CategoryRootMap.Count -gt 0
    
    if (-not $MinRes -and -not $hasSort) {
        Write-Host "`nNothing to preview. Provide -MinResolution (delete preview)" -ForegroundColor Yellow
        Write-Host "and/or -DestinationRoot / -QualityMap (sort preview)." -ForegroundColor Yellow
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
    
    if ($hasSort) {
        Write-Host "`nSort preview (quality -> drive):" -ForegroundColor Yellow
        $Videos | Group-Object ResolutionCategory | Sort-Object { $ResolutionSortOrder[$_.Name] } -Descending | ForEach-Object {
            $folderName = $ResolutionFolderNames[$_.Name]
            if (-not $folderName) { $folderName = "Low_Resolution" }
            $root = $CategoryRootMap[$_.Name]
            $sizeGB = [math]::Round(($_.Group | Measure-Object -Property FileSizeBytes -Sum).Sum / 1GB, 2)
            Write-Host "  $($_.Name) -> $root/$folderName ($($_.Count) files, $sizeGB GB)" -ForegroundColor Gray
        }
        
        # Show how the intelligent planner would actually distribute across drives.
        $plan = New-SortPlan -Videos $Videos -CategoryRootMap $CategoryRootMap
        
        Write-Host "`nProjected distribution per drive (after overflow balancing):" -ForegroundColor Yellow
        $plan.Assignments | Group-Object VolumeId | ForEach-Object {
            $sizeGB = [math]::Round(($_.Group | ForEach-Object { $_.Video.FileSizeBytes } | Measure-Object -Sum).Sum / 1GB, 2)
            $overflow = @($_.Group | Where-Object { $_.IsOverflow }).Count
            Write-Host "  $($_.Name): $($_.Count) file(s), $sizeGB GB (overflow: $overflow)" -ForegroundColor Gray
        }
        
        if ($plan.Unplaceable.Count -gt 0) {
            $unSizeGB = [math]::Round(($plan.Unplaceable | ForEach-Object { $_.FileSizeBytes } | Measure-Object -Sum).Sum / 1GB, 2)
            Write-Host "`n  Unplaceable (no drive has room): $($plan.Unplaceable.Count) file(s), $unSizeGB GB" -ForegroundColor Red
        }
    }
}

#endregion

#region Main Execution

Write-Host "Video Manager" -ForegroundColor Cyan
Write-Host "=============" -ForegroundColor Cyan
Write-Host "Action: $Action" -ForegroundColor White

# Will the destination drives be chosen interactively (per quality, after scan)?
$interactiveSort = ($SelectDrive -and $Action -eq "Sort" -and -not $DestinationRoot -and (-not $QualityMap -or $QualityMap.Count -eq 0))

# Interactive scan-drive selection
if ($SelectDrive) {
    $selectedScan = @(Select-Drive -Prompt "Select one or more drives/volumes to scan" -Multiple)
    if ($selectedScan.Count -gt 0) {
        $Path = $selectedScan
        Write-Host "Scan drive(s): $($selectedScan -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "No drive selected; using default path." -ForegroundColor Yellow
    }
}

# Validate parameters
if ($Action -eq "Sort" -and -not $DestinationRoot -and (-not $QualityMap -or $QualityMap.Count -eq 0) -and -not $interactiveSort) {
    Write-Error "Sort requires -DestinationRoot, -QualityMap, or -SelectDrive to choose destinations."
    exit 1
}

if ($Action -eq "Delete" -and -not $MinResolution) {
    Write-Error "MinResolution is required for Delete action. Use -MinResolution <4K|1440p|1080p|720p|480p|360p>"
    exit 1
}

# Light validation for an explicit DestinationRoot (per-root creation happens in the
# sort engine; quality-mapped roots are created/validated there too).
if ($Action -eq "Sort" -and $DestinationRoot -and (Test-Path $DestinationRoot) -and -not (Test-Path $DestinationRoot -PathType Container)) {
    Write-Error "DestinationRoot '$DestinationRoot' exists but is not a directory."
    exit 1
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

# Find video files (extension filtering via HashSet works on PS 5.1 and 7+,
# recursive or not, unlike Get-ChildItem -Include on a bare directory path).
# Stream the results so a live progress indicator shows the scan is working
# (recursive scans of large drives can otherwise look frozen).
Write-Host "`nScanning $($resolvedPaths.Count) path(s)..." -ForegroundColor White
$videoFiles = [System.Collections.Generic.List[object]]::new()
$examined = 0
$matched = 0
$lastUpdate = [DateTime]::MinValue
$spinner = '|', '/', '-', '\'
$spinIndex = 0

foreach ($scanPath in $resolvedPaths) {
    Get-ChildItem -Path $scanPath -File -Recurse:$Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $examined++
        if ($VideoExtensionSet.Contains($_.Extension)) {
            $videoFiles.Add($_)
            $matched++
        }
        
        # Throttle UI updates to ~7/sec so the indicator stays smooth without
        # slowing the scan itself.
        $now = [DateTime]::Now
        if (($now - $lastUpdate).TotalMilliseconds -ge 150) {
            $lastUpdate = $now
            $spinIndex = ($spinIndex + 1) % $spinner.Count
            Write-Progress -Activity "Scanning for video files $($spinner[$spinIndex])" `
                -Status "$matched video(s) found  |  $examined item(s) scanned" `
                -CurrentOperation $_.DirectoryName
        }
    }
}

Write-Progress -Activity "Scanning for video files" -Completed

if ($videoFiles.Count -eq 0) {
    Write-Warning "No video files found ($examined item(s) scanned)."
    exit 0
}

Write-Host "Found $($videoFiles.Count) video file(s) ($examined item(s) scanned)" -ForegroundColor Green

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
    Write-Progress -Activity "Analyzing videos ($percent%)" -Status "$processed of $($videoFiles.Count): $($file.Name)" -PercentComplete $percent
    
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

# Build the quality -> destination drive map (Sort and Report only).
$categoryRootMap = @{}
if ($Action -in @("Sort", "Report")) {
    $presentCategories = @($analyzedVideos |
        Select-Object -ExpandProperty ResolutionCategory -Unique |
        Sort-Object { $ResolutionSortOrder[$_] } -Descending)
    
    if ($interactiveSort) {
        $categoryRootMap = Get-InteractiveQualityAssignment -Categories $presentCategories
        if (-not $categoryRootMap -or $categoryRootMap.Count -eq 0) {
            Write-Error "No destination drives were assigned."
            exit 1
        }
    }
    elseif (($QualityMap -and $QualityMap.Count -gt 0) -or $DestinationRoot) {
        $resolved = Resolve-CategoryRootMap -QualityMap $QualityMap -DefaultRoot $DestinationRoot -Categories $presentCategories
        $categoryRootMap = $resolved.Map
        
        if ($resolved.Unmapped.Count -gt 0) {
            if ($Action -eq "Sort") {
                Write-Error "No destination for quality: $($resolved.Unmapped -join ', '). Add it to -QualityMap or set -DestinationRoot as a fallback."
                exit 1
            }
            else {
                Write-Warning "No destination mapped for: $($resolved.Unmapped -join ', '); excluded from sort preview."
            }
        }
    }
    
    # Show free space for each distinct destination drive.
    if ($categoryRootMap.Count -gt 0) {
        Write-Host "`nDestination drive(s):" -ForegroundColor White
        foreach ($root in @($categoryRootMap.Values | Select-Object -Unique)) {
            $di = Get-DriveInfo -Path $root
            $freeStr = if ($di) { "$($di.FreeGB) GB free" } else { "free space unknown" }
            Write-Host "  $root ($freeStr)" -ForegroundColor Gray
        }
    }
}

# Execute action
$isWhatIf = ($Action -eq "Report") -or $WhatIfPreference

switch ($Action) {
    "Analyze" { Invoke-AnalyzeAction -Videos $analyzedVideos -OutputPath $OutputFile }
    "Sort"    { Invoke-SortAction -Videos $analyzedVideos -CategoryRootMap $categoryRootMap -WhatIf:$isWhatIf }
    "Delete"  { Invoke-DeleteAction -Videos $analyzedVideos -MinRes $MinResolution -WhatIf:$isWhatIf -ForceDelete:$Force -LogPath $DeleteLog }
    "Report"  { Invoke-ReportAction -Videos $analyzedVideos -MinRes $MinResolution -CategoryRootMap $categoryRootMap -LogPath $DeleteLog }
}

Write-Host "`nDone!" -ForegroundColor Green

#endregion
