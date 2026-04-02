# Video Identity - Video File Analyzer

Cross-platform PowerShell script to scan video files in a directory, extract resolution and codec information using FFprobe, and export to a sortable Excel spreadsheet.

**Supports:** Windows, macOS, and Linux

## Requirements

- **PowerShell** 5.1+ (Windows) or PowerShell Core 7+ (macOS/Linux)
- **FFmpeg/FFprobe** - Used to analyze actual video stream data (not just metadata)

---

## Installing FFmpeg

### Windows

Choose one of these methods:

**Option 1 - WinGet (Recommended for Windows 10/11):**
```powershell
winget install FFmpeg
```

**Option 2 - Chocolatey:**
```powershell
choco install ffmpeg
```

**Option 3 - Scoop:**
```powershell
scoop install ffmpeg
```

**Option 4 - Manual Installation:**
1. Download from https://www.gyan.dev/ffmpeg/builds/ (choose "ffmpeg-release-essentials.zip")
2. Extract to `C:\ffmpeg`
3. Add `C:\ffmpeg\bin` to your PATH:
   - Press `Win + X` → System → Advanced system settings → Environment Variables
   - Under "User variables", edit `Path` and add `C:\ffmpeg\bin`
4. Restart PowerShell

**Or use the `-FFprobePath` parameter** if you don't want to modify PATH:
```powershell
.\Get-VideoInfo.ps1 -Path "D:\Videos" -FFprobePath "C:\ffmpeg\bin\ffprobe.exe"
```

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

## Installing PowerShell (if needed)

### macOS/Linux

PowerShell Core is required on non-Windows systems:

**macOS:**
```bash
brew install powershell/tap/powershell
```

**Ubuntu:**
```bash
sudo apt-get install -y wget apt-transport-https software-properties-common
wget -q "https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb"
sudo dpkg -i packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y powershell
```

Run with: `pwsh ./Get-VideoInfo.ps1`

---

## Optional: ImportExcel Module

For native `.xlsx` output with formatting:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

Without this module, the script outputs CSV files (still sortable in Excel).

---

## Usage

### Windows (PowerShell)

```powershell
# Scan current directory
.\Get-VideoInfo.ps1

# Scan specific directory
.\Get-VideoInfo.ps1 -Path "D:\Videos"

# Scan recursively (include subdirectories)
.\Get-VideoInfo.ps1 -Path "D:\Videos" -Recurse

# Scan MULTIPLE directories at once
.\Get-VideoInfo.ps1 -Path "D:\Videos", "E:\Movies", "F:\Downloads" -Recurse

# Specify output file
.\Get-VideoInfo.ps1 -Path "D:\Videos" -Recurse -OutputFile "D:\VideoReport.xlsx"

# Use custom FFprobe location
.\Get-VideoInfo.ps1 -Path "D:\Videos" -FFprobePath "C:\ffmpeg\bin\ffprobe.exe"
```

### macOS/Linux (PowerShell Core)

```bash
# Single path
pwsh ./Get-VideoInfo.ps1 -Path "/home/user/Videos" -Recurse

# Multiple paths
pwsh ./Get-VideoInfo.ps1 -Path "/home/user/Videos", "/mnt/media", "/mnt/nas/movies" -Recurse

# With output file
pwsh ./Get-VideoInfo.ps1 -Path "/home/user/Videos" -Recurse -OutputFile "/home/user/report.xlsx"
```

## Output Columns

| Column | Description |
|--------|-------------|
| FileName | Video file name |
| Directory | Parent directory path |
| Width | Video width in pixels |
| Height | Video height in pixels |
| Resolution | Width x Height (e.g., 1920x1080) |
| ResolutionCategory | Category: 4K UHD, 1080p FHD, 720p HD, etc. |
| VideoCodec | Video codec short name (h264, hevc, vp9, etc.) |
| VideoCodecLong | Full codec name |
| AudioCodec | Audio codec short name |
| AudioCodecLong | Full audio codec name |
| Duration | Duration in HH:MM:SS format |
| DurationSeconds | Duration in seconds |
| FrameRate | Frames per second |
| BitrateMbps | Overall bitrate in Mbps |
| FileSizeMB | File size in megabytes |
| PixelFormat | Pixel format (yuv420p, etc.) |
| ColorSpace | Color space information |
| HDR | Yes/No - HDR detection based on color transfer |
| FullPath | Complete file path |

## Supported Video Formats

MP4, MKV, AVI, MOV, WMV, FLV, WebM, M4V, MPG, MPEG, 3GP, MTS, M2TS, TS, VOB, OGV, HEVC, and more.

## Example Output

```
Video File Analyzer
===================
Scanning: D:\Videos
Recursive: True
Found 42 video file(s)

[1/42] movie.mp4 - 1920x1080 - h264
[2/42] clip.mkv - 3840x2160 - hevc
...

Summary
-------
  1080p FHD: 25 file(s)
  4K UHD: 10 file(s)
  720p HD: 7 file(s)

  h264: 30 file(s)
  hevc: 12 file(s)

Done!
```
