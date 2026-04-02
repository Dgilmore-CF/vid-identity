<#
.SYNOPSIS
    Scans video files in a directory and exports resolution/codec info to Excel.

.DESCRIPTION
    Uses FFprobe to analyze video files and extract actual resolution and codec 
    information from the video stream (not just metadata). Outputs to a sortable
    Excel spreadsheet.
    
    Cross-platform compatible: Windows, macOS, and Linux.

.PARAMETER Path
    One or more directories to scan for video files. Accepts multiple paths.
    Defaults to current directory.

.PARAMETER Recurse
    Include subdirectories in the scan.

.PARAMETER OutputFile
    Path for the Excel output file. Defaults to VideoInfo.xlsx in the scan directory.

.PARAMETER FFprobePath
    Optional path to ffprobe executable. Use if ffprobe is not in PATH.

.EXAMPLE
    # Windows - Single path
    .\Get-VideoInfo.ps1 -Path "D:\Videos" -Recurse -OutputFile "D:\VideoReport.xlsx"

.EXAMPLE
    # Windows - Multiple paths
    .\Get-VideoInfo.ps1 -Path "D:\Videos", "E:\Movies", "F:\Downloads" -Recurse

.EXAMPLE
    # macOS/Linux - Multiple paths
    ./Get-VideoInfo.ps1 -Path "/home/user/Videos", "/mnt/media" -Recurse

.EXAMPLE
    # Custom FFprobe path (Windows)
    .\Get-VideoInfo.ps1 -Path "D:\Videos" -FFprobePath "C:\ffmpeg\bin\ffprobe.exe"

.NOTES
    Requires FFprobe (part of FFmpeg). Install via:
    - Windows: winget install FFmpeg  OR  choco install ffmpeg
    - macOS:   brew install ffmpeg
    - Linux:   sudo apt install ffmpeg  OR  sudo dnf install ffmpeg
    - Or download from https://ffmpeg.org/download.html
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
    [string[]]$Path = @("."),
    
    [switch]$Recurse,
    
    [string]$OutputFile,
    
    [string]$FFprobePath
)

# Video file extensions to scan
$VideoExtensions = @(
    "*.mp4", "*.mkv", "*.avi", "*.mov", "*.wmv", "*.flv", "*.webm",
    "*.m4v", "*.mpg", "*.mpeg", "*.3gp", "*.3g2", "*.mts", "*.m2ts",
    "*.ts", "*.vob", "*.ogv", "*.divx", "*.xvid", "*.asf", "*.rm",
    "*.rmvb", "*.f4v", "*.hevc", "*.264", "*.265"
)

# Detect platform
$IsWindowsOS = $false
if ($PSVersionTable.PSVersion.Major -ge 6) {
    # PowerShell Core 6+ has $IsWindows automatic variable
    $IsWindowsOS = $IsWindows
} else {
    # Windows PowerShell 5.1 only runs on Windows
    $IsWindowsOS = $true
}

# Global variable for ffprobe command
$script:FFprobeCmd = "ffprobe"

function Find-FFprobe {
    param([string]$CustomPath)
    
    # If custom path provided, use it
    if ($CustomPath) {
        if (Test-Path $CustomPath) {
            $script:FFprobeCmd = $CustomPath
            return $true
        }
        Write-Warning "Specified FFprobe path not found: $CustomPath"
        return $false
    }
    
    # Try ffprobe in PATH
    try {
        $null = & ffprobe -version 2>&1
        $script:FFprobeCmd = "ffprobe"
        return $true
    } catch { }
    
    # Windows: Check common installation locations
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
    param([string]$FilePath)
    
    try {
        # Use FFprobe to get JSON output with stream and format info
        $json = & $script:FFprobeCmd -v quiet -print_format json -show_streams -show_format "$FilePath" 2>&1
        $info = $json | ConvertFrom-Json
        
        # Find the video stream
        $videoStream = $info.streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
        $audioStream = $info.streams | Where-Object { $_.codec_type -eq "audio" } | Select-Object -First 1
        
        if (-not $videoStream) {
            return $null
        }
        
        # Calculate duration in readable format
        $durationSec = [double]($info.format.duration)
        $duration = [TimeSpan]::FromSeconds($durationSec)
        $durationStr = "{0:D2}:{1:D2}:{2:D2}" -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds
        
        # Get file size in MB
        $fileSizeMB = [math]::Round([double]($info.format.size) / 1MB, 2)
        
        # Calculate bitrate in Mbps
        $bitrateMbps = [math]::Round([double]($info.format.bit_rate) / 1000000, 2)
        
        # Determine resolution category
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
        
        # Get frame rate
        $frameRate = $null
        if ($videoStream.r_frame_rate) {
            $parts = $videoStream.r_frame_rate -split "/"
            if ($parts.Count -eq 2 -and [int]$parts[1] -ne 0) {
                $frameRate = [math]::Round([double]$parts[0] / [double]$parts[1], 2)
            }
        }
        
        return [PSCustomObject]@{
            FileName           = [System.IO.Path]::GetFileName($FilePath)
            Directory          = [System.IO.Path]::GetDirectoryName($FilePath)
            Width              = [int]$videoStream.width
            Height             = [int]$videoStream.height
            Resolution         = "$($videoStream.width)x$($videoStream.height)"
            ResolutionCategory = $resolutionCategory
            VideoCodec         = $videoStream.codec_name
            VideoCodecLong     = $videoStream.codec_long_name
            AudioCodec         = if ($audioStream) { $audioStream.codec_name } else { "None" }
            AudioCodecLong     = if ($audioStream) { $audioStream.codec_long_name } else { "None" }
            Duration           = $durationStr
            DurationSeconds    = [math]::Round($durationSec, 2)
            FrameRate          = $frameRate
            BitrateMbps        = $bitrateMbps
            FileSizeMB         = $fileSizeMB
            PixelFormat        = $videoStream.pix_fmt
            ColorSpace         = $videoStream.color_space
            HDR                = if ($videoStream.color_transfer -match "smpte2084|arib-std-b67|bt2020") { "Yes" } else { "No" }
            FullPath           = $FilePath
        }
    }
    catch {
        Write-Warning "Failed to process: $FilePath - $_"
        return $null
    }
}

function Export-ToExcel {
    param(
        [array]$Data,
        [string]$OutputPath
    )
    
    # Check if ImportExcel module is available
    $useImportExcel = Get-Module -ListAvailable -Name ImportExcel
    
    if ($useImportExcel) {
        # Use ImportExcel module for proper Excel file
        $Data | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Video Info"
        Write-Host "Excel file created: $OutputPath" -ForegroundColor Green
    }
    else {
        # Fallback to CSV (can be opened and sorted in Excel)
        $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, ".csv")
        $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV file created: $csvPath" -ForegroundColor Yellow
        Write-Host "Note: Install ImportExcel module for native .xlsx output:" -ForegroundColor Yellow
        Write-Host "  Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Cyan
    }
}

# Main execution
Write-Host "Video File Analyzer" -ForegroundColor Cyan
Write-Host "===================" -ForegroundColor Cyan

# Find and verify FFprobe
if (-not (Find-FFprobe -CustomPath $FFprobePath)) {
    $errorMsg = "FFprobe not found."
    
    if ($IsWindowsOS) {
        $errorMsg += @"


Please install FFmpeg using one of these methods:

  Option 1 - WinGet (Windows 10/11):
    winget install FFmpeg

  Option 2 - Chocolatey:
    choco install ffmpeg

  Option 3 - Scoop:
    scoop install ffmpeg

  Option 4 - Manual Install:
    1. Download from https://ffmpeg.org/download.html (Windows builds)
    2. Extract to C:\ffmpeg
    3. Add C:\ffmpeg\bin to your PATH environment variable
       OR use -FFprobePath parameter:
       .\Get-VideoInfo.ps1 -FFprobePath "C:\ffmpeg\bin\ffprobe.exe"
"@
    } else {
        $errorMsg += @"


Please install FFmpeg:
  macOS:   brew install ffmpeg
  Ubuntu:  sudo apt install ffmpeg
  Fedora:  sudo dnf install ffmpeg
  Or download from: https://ffmpeg.org/download.html
"@
    }
    
    Write-Error $errorMsg
    exit 1
}

if (-not (Test-FFprobe)) {
    Write-Error "FFprobe found but failed to execute. Check permissions or installation."
    exit 1
}

Write-Host "Using FFprobe: $script:FFprobeCmd" -ForegroundColor DarkGray

# Resolve all paths and validate
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

# Set default output file if not specified
if (-not $OutputFile) {
    # Use first path for default output location
    $OutputFile = Join-Path $resolvedPaths[0] "VideoInfo.xlsx"
}

Write-Host "Scanning $($resolvedPaths.Count) path(s):" -ForegroundColor White
foreach ($p in $resolvedPaths) {
    Write-Host "  - $p" -ForegroundColor White
}
Write-Host "Recursive: $Recurse" -ForegroundColor White

# Find all video files across all paths
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

$videoFiles = $videoFiles | Sort-Object FullName
$totalFiles = $videoFiles.Count

if ($totalFiles -eq 0) {
    Write-Warning "No video files found in specified path(s)."
    exit 0
}

Write-Host "Found $totalFiles video file(s)" -ForegroundColor Green
Write-Host ""

# Process each file
$results = @()
$processed = 0

foreach ($file in $videoFiles) {
    $processed++
    $percent = [math]::Round(($processed / $totalFiles) * 100, 0)
    Write-Progress -Activity "Analyzing videos" -Status "$processed of $totalFiles - $($file.Name)" -PercentComplete $percent
    
    $details = Get-VideoDetails -FilePath $file.FullName
    if ($details) {
        $results += $details
        Write-Host "[$processed/$totalFiles] $($file.Name) - $($details.Resolution) - $($details.VideoCodec)" -ForegroundColor Gray
    }
}

Write-Progress -Activity "Analyzing videos" -Completed

if ($results.Count -eq 0) {
    Write-Warning "No valid video files could be analyzed."
    exit 0
}

# Export results
Write-Host ""
Write-Host "Exporting $($results.Count) video(s) to spreadsheet..." -ForegroundColor White
Export-ToExcel -Data $results -OutputPath $OutputFile

# Summary statistics
Write-Host ""
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "-------" -ForegroundColor Cyan
$results | Group-Object ResolutionCategory | Sort-Object Count -Descending | ForEach-Object {
    Write-Host "  $($_.Name): $($_.Count) file(s)" -ForegroundColor White
}
Write-Host ""
$results | Group-Object VideoCodec | Sort-Object Count -Descending | ForEach-Object {
    Write-Host "  $($_.Name): $($_.Count) file(s)" -ForegroundColor White
}

Write-Host ""
Write-Host "Done!" -ForegroundColor Green
