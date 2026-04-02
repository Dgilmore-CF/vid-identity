# VideoManager

A comprehensive PowerShell tool for analyzing, organizing, and managing video files by resolution and codec.

**Supports:** Windows, macOS, and Linux

## Features

- **Analyze** - Scan videos and export resolution/codec info to Excel/CSV
- **Sort** - Organize videos into resolution-based folders (4K, 1080p, 720p, etc.)
- **Delete** - Mass delete videos below a resolution threshold
- **Report** - Dry-run preview of any action
- **Disk Space Management** - Intelligent queuing when space is limited

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
# Sort videos into resolution folders
.\VideoManager.ps1 -Path "D:\Videos" -Action Sort -DestinationRoot "D:\Sorted" -Recurse

# macOS/Linux
pwsh ./VideoManager.ps1 -Path "/home/user/Videos" -Action Sort -DestinationRoot "/home/user/Sorted" -Recurse
```

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
# Preview what would be deleted
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -MinResolution 1080p -Recurse

# Preview how files would be sorted
.\VideoManager.ps1 -Path "D:\Videos" -Action Report -DestinationRoot "D:\Sorted" -Recurse
```

---

## Actions

| Action | Description |
|--------|-------------|
| `Analyze` | Export video info to Excel/CSV (default) |
| `Sort` | Move videos into resolution folders |
| `Delete` | Delete videos below minimum resolution |
| `Report` | Dry-run preview of Sort or Delete |

---

## Disk Space Handling

- **Same-drive moves**: Instant, no space check needed
- **Cross-drive moves**: Checks available space before each move
- **Queuing**: Files queued when space is low, auto-retried as space frees up
- **Safety margin**: Maintains 5% free space buffer

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
