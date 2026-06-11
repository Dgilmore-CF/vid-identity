# VideoManager

A comprehensive PowerShell tool for analyzing, organizing, and managing video files by resolution and codec.

**Supports:** Windows, macOS, and Linux

## Features

- **Analyze** - Scan videos and export resolution/codec info to Excel/CSV
- **Sort** - Organize videos into resolution-based folders (4K, 1080p, 720p, etc.)
- **Delete** - Mass delete videos below a resolution threshold
- **Report** - Dry-run preview of any action
- **Disk Space Management** - Intelligent queuing when space is limited
- **Drive Selection** - Interactively pick a drive/volume to scan or sort to
- **Deletion Log** - Every removed file is logged first so you can re-acquire a higher-quality version

## Requirements

- **PowerShell** 5.1+ (Windows) or PowerShell Core 7+ (macOS/Linux)
- **FFmpeg/FFprobe** - Analyzes actual video stream data (not just metadata)

---

## Installing FFmpeg

### Windows

```powershell
# Option 1 - WinGet (Recommended)
winget install FFmpeg

# Option 2 - Chocolatey
choco install ffmpeg

# Option 3 - Scoop
scoop install ffmpeg
```

**Manual Installation:**
1. Download from https://www.gyan.dev/ffmpeg/builds/
2. Extract to `C:\ffmpeg`
3. Add `C:\ffmpeg\bin` to PATH, or use `-FFprobePath` parameter

### macOS

```bash
brew install ffmpeg
```

### Linux

**Ubuntu/Debian:**
```bash
sudo apt install ffmpeg
```

**Fedora:**
```bash
sudo dnf install ffmpeg
```

**Arch:**
```bash
sudo pacman -S ffmpeg
```

---

## Installing PowerShell (macOS/Linux)

```bash
# macOS
brew install powershell/tap/powershell

# Ubuntu
sudo apt-get update && sudo apt-get install -y powershell
```

Run with: `pwsh ./VideoManager.ps1`

---

## Optional: ImportExcel Module

For native `.xlsx` output (otherwise exports to CSV):

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

---

## Usage

### Analyze Videos (Default)

Export video info to Excel/CSV spreadsheet:

```powershell
# Windows
.\VideoManager.ps1 -Path "D:\Videos" -Recurse

# Multiple directories
.\VideoManager.ps1 -Path "D:\Videos", "E:\Movies" -Recurse -OutputFile "D:\Report.xlsx"

# macOS/Linux
pwsh ./VideoManager.ps1 -Path "/home/user/Videos" -Recurse
```

### Sort Videos by Resolution

Organize into folders: `4K_UHD`, `1080p_FHD`, `720p_HD`, etc.

```powershell
# Sort videos into resolution folders on a single drive
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -Recurse

# macOS/Linux
pwsh ./VideoManager.ps1 -Path "/home/user/Videos" -Action Sort -DestinationRoot "/home/user/Sorted" -Recurse
```

### Sort Different Qualities to Different Drives

Use `-QualityMap` to send each quality to its own drive in a single session.
Keys can be short names (`4K`, `1440p`, `1080p`, `720p`, `480p`, `360p`, `Low`)
or full category names (`"4K UHD"`). Anything not listed falls back to
`-DestinationRoot` (if provided).

```powershell
# 4K to a big archive drive, HD to a fast drive, everything else to a catch-all
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -Recurse `
    -QualityMap @{ "4K" = "X:\4K"; "1080p" = "Y:\HD"; "720p" = "Y:\HD" } `
    -DestinationRoot "Z:\Other"
```

**Intelligent multi-drive balancing:** when a target drive runs low on space, files
automatically overflow to the destination drive with the most free space. Largest
files are placed first (first-fit-decreasing) for efficient packing, a 5% free-space
buffer is kept on every drive, and files already on a destination volume are moved
without an unnecessary cross-drive copy. Use `-Action Report` first to preview the
exact distribution per drive.

### Delete Low-Resolution Videos

```powershell
# Delete videos below 720p (prompts for confirmation)
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -Recurse

# Delete without confirmation
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 1080p -Recurse -Force

# Thresholds: 4K, 1440p, 1080p, 720p, 480p, 360p
```

### Preview Changes (Dry-Run)

```powershell
# Preview what would be deleted (also writes a re-acquisition log)
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 1080p -Recurse

# Preview how files would be sorted
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -DestinationRoot "D:\Sorted" -Recurse
```

### Interactive Drive Selection

Pick one or more drives/volumes from a numbered list instead of typing paths.
At the prompt, enter a single number, a comma/space-separated list (e.g. `1,3,4`),
or `all` to select every drive:

```powershell
# Choose one or more drives to scan
.\VideoManager.ps1 -SelectDrive -Recurse

# Choose multiple source drives and a single destination drive for sorting
.\VideoManager.ps1 -SelectDrive -Action Sort -Recurse
```

You can also use standard PowerShell `-WhatIf` with `Sort` and `Delete` to preview
changes without modifying anything.

---

## Deletion Log (Re-Acquisition)

Before **any** file is deleted (and during a `Delete`/`Report` dry-run), VideoManager
writes a CSV log capturing the metadata needed to source a better copy later. The log
is written *first*, so the record survives even if the deletion is interrupted.

```powershell
# Custom log path
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -DeleteLog "D:\reacquire.csv"
```

Default location: `VideoManager_DeletedLog_<timestamp>.csv` in the current directory.

**Logged columns:** `FileName`, `SearchTitle`, `Resolution`, `ResolutionCategory`,
`Width`, `Height`, `VideoCodec`, `Duration`, `DurationSeconds`, `FrameRate`,
`BitrateMbps`, `FileSizeMB`, `OriginalDirectory`, `OriginalFullPath`, `LoggedAt`.

---

## Actions

| Action | Description |
|--------|-------------|
| `Analyze` | Export video info to Excel/CSV (default) |
| `Sort` | Move videos into resolution folders |
| `Delete` | Delete videos below minimum resolution |
| `Report` | Dry-run preview of Sort or Delete (also exports the analysis spreadsheet when `-OutputFile` is given) |

---

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-Path` | One or more directories to scan (default: current directory) |
| `-Recurse` | Include subdirectories |
| `-Action` | `Analyze` (default), `Sort`, `Delete`, or `Report` |
| `-OutputFile` | Excel/CSV output path (Analyze) |
| `-DestinationRoot` | Root folder for resolution subfolders (Sort); fallback for unmapped qualities |
| `-QualityMap` | Hashtable mapping qualities to destination drives, e.g. `@{ "4K"="X:\4K"; "1080p"="Y:\HD" }` |
| `-MinResolution` | Minimum resolution to keep: `4K`, `1440p`, `1080p`, `720p`, `480p`, `360p` |
| `-Force` | Skip confirmation prompts for destructive operations |
| `-FFprobePath` | Explicit path to the ffprobe executable |
| `-SelectDrive` | Interactively select one or more drives/volumes to scan (single destination for Sort) |
| `-DeleteLog` | Path to the deletion/re-acquisition CSV log |
| `-WhatIf` | Standard PowerShell dry-run for Sort/Delete |

---

## Examples Reference

A complete reference of every option. Windows examples use `.\VideoManager.ps1`;
on macOS/Linux substitute `pwsh ./VideoManager.ps1`.

### `-Path` (scan location)

```powershell
# Default: scan the current directory
.\VideoManager.ps1

# Single directory
.\VideoManager.ps1 -Path "D:\Videos"

# Multiple directories (comma-separated)
.\VideoManager.ps1 -Path "D:\Videos", "E:\Movies", "F:\TV"

# Positional (the -Path name is optional)
.\VideoManager.ps1 "D:\Videos"

# Piped in from another command
Get-ChildItem "D:\Media" -Directory | .\VideoManager.ps1
```

### `-Recurse` (include subfolders)

```powershell
# Scan the top level only (default)
.\VideoManager.ps1 -Path "D:\Videos"

# Scan all subfolders too
.\VideoManager.ps1 -Path "D:\Videos" -Recurse
```

### `-Action Analyze` (default) — export metadata

```powershell
# Explicitly request Analyze
.\VideoManager.ps1 -Path "D:\Videos" -Action Analyze -Recurse

# Analyze is the default, so this is equivalent
.\VideoManager.ps1 -Path "D:\Videos" -Recurse
```

### `-OutputFile` (Analyze output path)

```powershell
# Default: <first scan path>\VideoInfo.xlsx
.\VideoManager.ps1 -Path "D:\Videos" -Recurse

# Custom Excel output (requires ImportExcel module)
.\VideoManager.ps1 -Path "D:\Videos" -Recurse -OutputFile "D:\Reports\Library.xlsx"

# Without ImportExcel, output automatically falls back to .csv
.\VideoManager.ps1 -Path "D:\Videos" -Recurse -OutputFile "D:\Reports\Library.csv"
```

### `-Action Sort` + `-DestinationRoot` — organize by resolution

```powershell
# Move videos into resolution subfolders under the destination root
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -Recurse

# The destination root is created automatically if it does not exist
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\New\Sorted"

# Preview a sort without moving anything (standard dry-run)
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -WhatIf
```

### `-QualityMap` — send each quality to a different drive

```powershell
# Per-quality drive targets; unmapped qualities use -DestinationRoot as fallback
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -Recurse `
    -QualityMap @{ "4K" = "X:\4K"; "1440p" = "X:\4K"; "1080p" = "Y:\HD"; "720p" = "Y:\HD" } `
    -DestinationRoot "Z:\Other"

# Full category names also work as keys
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -QualityMap @{ "4K UHD" = "X:\4K" } -DestinationRoot "Y:\Rest"

# Preview the exact per-drive distribution (including overflow) before moving
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -Recurse `
    -QualityMap @{ "4K" = "X:\4K"; "1080p" = "Y:\HD" } -DestinationRoot "Z:\Other"

# Interactively assign each detected quality to a drive
.\VideoManager.ps1 -SelectDrive -Action Sort -Recurse
```

### `-Action Delete` + `-MinResolution` — remove low-res videos

```powershell
# Delete everything below 720p (prompts: type DELETE to confirm)
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -Recurse

# Valid thresholds: 4K, 1440p, 1080p, 720p, 480p, 360p
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 1080p -Recurse
```

### `-Force` (skip the confirmation prompt)

```powershell
# Delete below 1080p with no confirmation prompt
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 1080p -Recurse -Force
```

### `-Action Report` — dry-run preview (no changes)

```powershell
# Preview what a Delete would remove (also writes a re-acquisition log)
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 1080p -Recurse

# Preview how a Sort would distribute files
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -DestinationRoot "D:\Sorted" -Recurse

# Preview both at once
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 720p -DestinationRoot "D:\Sorted" -Recurse

# Preview AND save the full analysis spreadsheet (add -OutputFile)
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 1080p -Recurse -OutputFile "D:\Report.xlsx"
```

> In `Report` mode the analysis spreadsheet is written only when you pass
> `-OutputFile` (so a plain dry-run never creates files). The `Analyze` action
> always writes the spreadsheet.

### `-WhatIf` (standard PowerShell dry-run)

```powershell
# Equivalent to a Sort dry-run, using the built-in -WhatIf switch
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -WhatIf

# Delete dry-run via -WhatIf (still writes the re-acquisition log, deletes nothing)
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -WhatIf
```

### `-SelectDrive` (interactive drive picker)

At the prompt you can enter a single number, a comma/space-separated list
(e.g. `1,3,4`), or `all`. Scan selection supports multiple drives; the Sort
destination is always a single drive.

```powershell
# Pick one or more drives/volumes to scan from a numbered list
.\VideoManager.ps1 -SelectDrive -Recurse

# Pick multiple source drives AND a single destination drive for sorting
.\VideoManager.ps1 -SelectDrive -Action Sort -Recurse

# Pick drives, then delete below 480p across all of them
.\VideoManager.ps1 -SelectDrive -Action Delete -MinResolution 480p -Recurse
```

### `-DeleteLog` (re-acquisition log path)

```powershell
# Default: VideoManager_DeletedLog_<timestamp>.csv in the current directory
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p

# Custom log path
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -DeleteLog "D:\reacquire.csv"

# The log is also produced in Report mode for the would-be-deleted set
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 1080p -DeleteLog "D:\preview.csv"
```

### `-FFprobePath` (explicit ffprobe location)

```powershell
# Point at a specific ffprobe binary if it is not on PATH
.\VideoManager.ps1 -Path "D:\Videos" -FFprobePath "C:\ffmpeg\bin\ffprobe.exe"

# macOS/Linux
pwsh ./VideoManager.ps1 -Path "/home/user/Videos" -FFprobePath "/usr/local/bin/ffprobe"
```

### Combined / real-world examples

```powershell
# Full analysis of multiple drives into one custom Excel report
.\VideoManager.ps1 -Path "D:\Videos", "E:\Movies" -Recurse -OutputFile "D:\Reports\AllMedia.xlsx"

# Safely preview, then perform a cleanup of sub-720p files with a named log
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 720p -Recurse -DeleteLog "D:\cleanup.csv"
.\VideoManager.ps1 -Path "D:\Videos" -Action Delete -MinResolution 720p -Recurse -Force -DeleteLog "D:\cleanup.csv"

# Interactive: choose a drive, sort it, and dry-run first
.\VideoManager.ps1 -SelectDrive -Action Sort -Recurse -WhatIf

# Get the built-in help (auto-generated from comment-based help)
Get-Help .\VideoManager.ps1 -Full
Get-Help .\VideoManager.ps1 -Examples
```

---

## Disk Space Handling

- **Per-quality targeting**: With `-QualityMap`, each quality goes to its assigned drive
- **Intelligent overflow**: When a target drive is full, files move to the destination drive with the most free space
- **First-fit decreasing**: Largest files are placed first for efficient packing
- **Same-volume moves**: Instant, no cross-drive copy and no extra space consumed
- **Safety margin**: Maintains a 5% free-space buffer on every destination drive
- **Unplaceable reporting**: Files that fit nowhere are listed so you can free space or add a drive
- **Preview**: `-Action Report` shows the projected distribution per drive before any move

---

## Resolution Categories

| Category | Min Height | Sort Folder |
|----------|------------|-------------|
| 4K UHD | 2160p | `4K_UHD` |
| 1440p QHD | 1440p | `1440p_QHD` |
| 1080p FHD | 1080p | `1080p_FHD` |
| 720p HD | 720p | `720p_HD` |
| 480p SD | 480p | `480p_SD` |
| 360p | 360p | `360p` |
| Low | <360p | `Low_Resolution` |

---

## Output Columns (Analyze Action)

| Column | Description |
|--------|-------------|
| FileName | Video file name |
| Resolution | Width x Height (e.g., 1920x1080) |
| ResolutionCategory | 4K UHD, 1080p FHD, 720p HD, etc. |
| VideoCodec | h264, hevc, vp9, etc. |
| AudioCodec | aac, ac3, etc. |
| Duration | HH:MM:SS format |
| FrameRate | Frames per second |
| BitrateMbps | Overall bitrate |
| FileSizeMB | File size in megabytes |
| HDR | Yes/No |

---

## Supported Formats

MP4, MKV, AVI, MOV, WMV, FLV, WebM, M4V, MPG, MPEG, 3GP, MTS, M2TS, TS, VOB, OGV, HEVC, and more
