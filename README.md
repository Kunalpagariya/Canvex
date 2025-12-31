# Canvex - Canva Compatible Excel Creator

A PyQt5-based desktop application that automatically downloads images from Bing and inserts them into Excel files based on column mappings.

## Features

- **Image Search & Download**: Searches Bing Images for headshots based on column data
- **Smart Filtering**: Filters out watermarked/premium images from restricted sites
- **Column Mapping**: Flexible mapping between input and output columns
- **Batch Processing**: Process entire Excel sheets with progress tracking
- **Dark/Light Mode**: Native macOS theme support
- **Recent Files**: Quick access to recently used Excel files
- **Settings Management**: Persistent settings for resolution, format, and search preferences

## Supported Platforms

- **macOS** (Intel & Apple Silicon): `.app` bundle
- **Windows**: Automated `.exe` builds via GitHub Actions

## Installation

### macOS

Download the latest `Canvex.app` from Releases and run it.

### Windows

The Windows `.exe` is automatically built and available as an artifact:

1. Go to **Actions** tab → **Build Windows EXE**
2. Select the latest successful workflow run
3. Download the **Canvex-Windows** artifact
4. Extract and run `Canvex.exe`

## Building Locally

### Prerequisites

```bash
pip install -r requirements.txt
pip install pyinstaller
```

### macOS Build

```bash
pyinstaller Canvex.spec -y
```

Output: `dist/Canvex.app`

### Windows Build

```bash
pyinstaller Canvex.spec -y
```

Output: `dist/Canvex.exe`

## Automated Builds

Every push to `main` or `master` branch automatically triggers a Windows build via GitHub Actions. The `.exe` is available as an artifact for 30 days.

To create a release build:
```bash
git tag v1.0.0
git push --tags
```

This automatically creates a GitHub Release with the Windows `.exe`.

## Usage

1. **Select Excel File**: Click "Select Excel File" to load your spreadsheet
2. **Add Mappings**: Create column mappings:
   - **Input Column**: Column containing search terms
   - **Output Column**: Column to insert images (or create new)
3. **Configure Settings**: Adjust resolution, format, and search engine
4. **Generate**: Click "Generate Canva-Compatible Excel" and wait for processing

## Project Structure

```
FinalBuildMac/
├── Canvex.py              # Main application
├── Canvex.spec            # PyInstaller spec file
├── requirements.txt       # Python dependencies
├── .github/
│   └── workflows/
│       └── build-windows.yml  # GitHub Actions workflow
└── dist/
    ├── Canvex.app         # macOS bundle
    └── Canvex.exe         # Windows executable
```

## Technical Details

- **Framework**: PyQt5
- **Image Source**: Bing Images
- **Browser**: Selenium + Chrome WebDriver
- **Excel Output**: xlsxwriter
- **Image Processing**: Pillow

## License

Private Project
