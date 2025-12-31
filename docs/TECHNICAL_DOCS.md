# Canvex — Functional & Technical Documentation

<div align="center">
  <h2>🖼️ Image Excel Creator</h2>
  <p><em>Complete Technical Reference</em></p>
  <p>Version 1.0 | Last Updated: December 2025</p>
</div>

---

## 📋 Table of Contents

### Part I: Functional Specification
1. [Product Overview](#1-product-overview)
2. [User Stories & Use Cases](#2-user-stories--use-cases)
3. [Functional Requirements](#3-functional-requirements)
4. [User Interface Specification](#4-user-interface-specification)
5. [Workflow & Process Flows](#5-workflow--process-flows)

### Part II: Technical Specification
6. [Architecture Overview](#6-architecture-overview)
7. [Module Documentation](#7-module-documentation)
8. [Data Flow & Processing](#8-data-flow--processing)
9. [External Dependencies](#9-external-dependencies)
10. [Configuration & Settings](#10-configuration--settings)
11. [Error Handling & Logging](#11-error-handling--logging)
12. [Performance Optimization](#12-performance-optimization)
13. [Security Considerations](#13-security-considerations)
14. [Deployment & Packaging](#14-deployment--packaging)
15. [API Reference](#15-api-reference)
16. [Recent Enhancements](#16-recent-enhancements-december-2025)

### Appendices
- [A. File Structure](#appendix-a-file-structure)
- [B. Settings Schema](#appendix-b-settings-schema)
- [C. Supported Formats](#appendix-c-supported-formats)
- [D. Recent Changes Summary](#appendix-d-recent-changes-summary)
- [E. Future Improvements](#appendix-e-future-improvements)

---

# Part I: Functional Specification

---

## 1. Product Overview

### 1.1 Purpose

Canvex is a desktop application designed to automate the process of searching for images on the web and inserting them into Excel spreadsheets based on text data in specified columns.

### 1.2 Problem Statement

Manually searching for images and inserting them into Excel files is:
- **Time-consuming:** Each image requires multiple steps
- **Error-prone:** Easy to mix up images with wrong entries
- **Tedious:** Repetitive for large datasets

### 1.3 Solution

Canvex automates this workflow:
```
Excel Data → Image Search → Download → Filter → Resize → Insert → Export
```

### 1.4 Target Users

| User Type | Description | Primary Use Case |
|-----------|-------------|------------------|
| **HR Professionals** | Create employee directories | Headshot directories |
| **Content Creators** | Build media catalogs | Actor/character sheets |
| **Marketers** | Product catalogs | Product image galleries |
| **Researchers** | Data visualization | Image-rich datasets |

### 1.5 Key Value Propositions

1. **Automation** — Reduces hours of manual work to minutes
2. **Intelligence** — Smart filtering removes unwanted images
3. **Flexibility** — Multiple search engines, themes, and output options
4. **Reliability** — Checkpoint saving, error recovery, and logging

---

## 2. User Stories & Use Cases

### 2.1 User Stories

#### US-001: Basic Image Insertion
> **As a** user with an Excel file containing names,  
> **I want to** automatically find and insert headshot images,  
> **So that** I can create a visual directory without manual searching.

**Acceptance Criteria:**
- User can load an Excel file
- User can specify which column contains search terms
- Application searches for images and inserts them
- Output Excel contains embedded images

#### US-002: Custom Search Configuration
> **As a** user processing different types of content,  
> **I want to** customize search parameters (theme, resolution, format),  
> **So that** I get appropriate images for my specific use case.

**Acceptance Criteria:**
- User can select from preset themes or enter custom
- User can choose resolution from presets or enter custom
- User can select output format (PNG/JPG/WEBP)
- Settings persist between sessions

#### US-003: Multiple Column Processing
> **As a** user with multiple columns needing images,  
> **I want to** create multiple mappings in one run,  
> **So that** I don't have to process the same file multiple times.

**Acceptance Criteria:**
- User can add multiple input→output column mappings
- Each mapping processes independently
- All mappings execute in a single processing run

#### US-004: Progress Monitoring
> **As a** user processing large files,  
> **I want to** see real-time progress and be able to cancel,  
> **So that** I know the status and can stop if needed.

**Acceptance Criteria:**
- Progress bar shows completion percentage
- Cancel button safely stops processing
- Partial results are saved on cancellation

### 2.2 Use Case Diagram

```
                    ┌─────────────────────────────────────────┐
                    │              CANVEX SYSTEM              │
                    │                                         │
  ┌─────┐           │   ┌─────────────────────────────┐      │
  │     │──Load────▶│   │     Load Excel File         │      │
  │     │           │   └─────────────────────────────┘      │
  │     │           │                                         │
  │     │──Config──▶│   ┌─────────────────────────────┐      │
  │     │           │   │   Configure Settings        │      │
  │ User│           │   └─────────────────────────────┘      │
  │     │           │                                         │
  │     │──Map────▶ │   ┌─────────────────────────────┐      │
  │     │           │   │   Create Column Mappings    │      │
  │     │           │   └─────────────────────────────┘      │
  │     │           │                                         │
  │     │──Process─▶│   ┌─────────────────────────────┐      │    ┌─────────────┐
  │     │           │   │   Process & Generate Output │──────│───▶│ Search APIs │
  └─────┘           │   └─────────────────────────────┘      │    └─────────────┘
                    │                                         │
                    └─────────────────────────────────────────┘
```

### 2.3 Use Case Details

#### UC-001: Load Excel File

| Field | Value |
|-------|-------|
| **Name** | Load Excel File |
| **Actor** | User |
| **Precondition** | Application is running, user has .xlsx file |
| **Trigger** | User clicks "Select Excel File" or drags file |
| **Main Flow** | 1. User initiates file selection<br>2. System displays file dialog<br>3. User selects .xlsx file<br>4. System validates file format<br>5. If multiple sheets, system prompts for selection<br>6. System loads column headers<br>7. System displays column mapping interface |
| **Alternative Flow** | 3a. User cancels → Return to initial state<br>4a. Invalid file → Display error message |
| **Postcondition** | Column headers loaded, mapping UI visible |

#### UC-002: Configure Image Settings

| Field | Value |
|-------|-------|
| **Name** | Configure Image Settings |
| **Actor** | User |
| **Precondition** | Excel file loaded |
| **Main Flow** | 1. User selects Image Theme from dropdown<br>2. User selects Search Browser<br>3. User selects Resolution<br>4. User selects Format<br>5. If JPG selected, user sets quality |
| **Postcondition** | Settings configured for processing |

#### UC-003: Create Column Mappings

| Field | Value |
|-------|-------|
| **Name** | Create Column Mappings |
| **Actor** | User |
| **Precondition** | Excel file loaded with columns |
| **Main Flow** | 1. User clicks "+ Add Mapping"<br>2. System adds row to mapping table<br>3. User selects Input Column<br>4. User selects Output Column or "Create New..."<br>5. If new column, user enters name<br>6. Repeat for additional mappings |
| **Postcondition** | At least one mapping configured |

#### UC-004: Process Excel File

| Field | Value |
|-------|-------|
| **Name** | Process Excel File |
| **Actor** | User |
| **Precondition** | File loaded, settings configured, mappings created |
| **Main Flow** | 1. User clicks "Start Processing"<br>2. System prompts for save location<br>3. User selects location and filename<br>4. System locks UI, shows progress<br>5. For each row and mapping:<br>&nbsp;&nbsp;a. Search for images<br>&nbsp;&nbsp;b. Download best candidates<br>&nbsp;&nbsp;c. Apply filters and resize<br>&nbsp;&nbsp;d. Insert into output Excel<br>6. System saves output file<br>7. System unlocks UI<br>8. System prompts to open file |
| **Alternative Flow** | 4a. User clicks Cancel → Save checkpoint, unlock UI |
| **Postcondition** | Output Excel created with images |

---

## 3. Functional Requirements

### 3.1 File Operations

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-001 | System shall accept .xlsx files as input | Must |
| FR-002 | System shall support drag-and-drop file loading | Must |
| FR-003 | System shall detect and list multiple sheets | Must |
| FR-004 | System shall export output as .xlsx with embedded images | Must |
| FR-005 | System shall create processing log files | Must |
| FR-006 | System shall maintain list of recently opened files | Should |

### 3.2 Image Search

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-010 | System shall search Bing Images | Must |
| FR-011 | System shall search Google Images | Must |
| FR-012 | System shall search DuckDuckGo Images | Should |
| FR-013 | System shall append theme keywords to search queries | Must |
| FR-014 | System shall filter out stock photo websites | Must |
| FR-015 | System shall retry failed searches with fallback engine | Should |

### 3.3 Image Processing

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-020 | System shall download images in parallel | Must |
| FR-021 | System shall resize images to target resolution | Must |
| FR-022 | System shall filter low-quality images | Must |
| FR-023 | System shall filter black/white images (optional) | Should |
| FR-024 | System shall filter cartoon/graphic images (optional) | Should |
| FR-025 | System shall prefer portrait-oriented images (optional) | Should |
| FR-026 | System shall support PNG, JPG, WEBP output formats | Must |

### 3.4 User Interface

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-030 | System shall provide light and dark themes | Must |
| FR-031 | System shall follow system theme automatically | Should |
| FR-032 | System shall display real-time progress | Must |
| FR-033 | System shall allow safe cancellation | Must |
| FR-034 | System shall persist user settings | Must |
| FR-035 | System shall restore settings on startup | Must |

### 3.5 Non-Functional Requirements

| ID | Requirement | Priority |
|----|-------------|----------|
| NFR-001 | System shall process at least 2-5 rows per second | Should |
| NFR-002 | System shall support files with 10,000+ rows | Should |
| NFR-003 | System shall run on macOS 10.14+ and Windows 10+ | Must |
| NFR-004 | System shall provide clear error messages | Must |
| NFR-005 | System shall checkpoint progress on cancellation | Must |

---

## 4. User Interface Specification

### 4.1 Main Window

**Window Properties:**
| Property | Value |
|----------|-------|
| Title | "Canvex" |
| Default Size | 900 × 700 pixels |
| Minimum Size | 700 × 600 pixels |
| Resizable | Yes |
| Accept Drops | Yes (.xlsx files) |

### 4.2 UI Components

#### Toolbar
| Component | Type | Function |
|-----------|------|----------|
| File Button | QPushButton | Opens file menu dialog |
| Settings Button | QPushButton | Opens settings dialog |
| Help Button | QPushButton | Opens help dialog |
| About Button | QPushButton | Opens about dialog |
| Theme Button | QPushButton | Opens theme selector |

#### Configuration Area
| Component | Type | Options |
|-----------|------|---------|
| Image Theme | QComboBox | 8 presets + "Custom Theme..." |
| Search Browser | QComboBox | "Bing Images", "Google Images", "DuckDuckGo" |
| Resolution | QComboBox | "240p" - "3840p" + "Custom..." |
| Format | QComboBox | "PNG", "JPG", "WEBP" |
| JPG Quality | QComboBox | "60 (Low)" - "100 (Ultra)" |

#### Mapping Table
| Column | Width | Type |
|--------|-------|------|
| # | 40px | QTableWidgetItem (read-only) |
| Input Column | Stretch | QComboBox |
| Output Column | Stretch | QComboBox |
| New Column Name | 140px | QLineEdit (conditional) |
| Delete | 50px | QPushButton |

### 4.3 Dialog Specifications

#### Settings Dialog
- **Size:** 480 × 580 pixels
- **Sections:** Output, Search, Performance, Image Filters
- **Scrollable:** Yes

#### Help Dialog
- **Size:** 700 × 550 pixels
- **Layout:** Sidebar navigation + content stack
- **Pages:** Getting Started, Features, Settings, Tips, Reference

#### Theme Dialog
- **Size:** 300 × 150 pixels
- **Options:** Light, Dark, System

### 4.4 Style Sheets

**Theme Colors:**

| Element | Dark Theme | Light Theme |
|---------|------------|-------------|
| Background | #1e1e1e | #f5f5f7 |
| Card Background | #2d2d2d | #ffffff |
| Text | #ffffff | #1d1d1f |
| Accent | #0a84ff | #0a84ff |
| Border | #404040 | #d2d2d7 |
| Error | #ff3b30 | #ff3b30 |
| Success | #34c759 | #34c759 |

---

## 5. Workflow & Process Flows

### 5.1 High-Level Workflow

```
┌──────────────┐    ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
│   STARTUP    │───▶│   LOAD FILE  │───▶│  CONFIGURE   │───▶│   PROCESS    │
│              │    │              │    │              │    │              │
│ • Init UI    │    │ • Select file│    │ • Theme      │    │ • Search     │
│ • Load prefs │    │ • Read cols  │    │ • Browser    │    │ • Download   │
│ • Apply theme│    │ • Sheet sel  │    │ • Resolution │    │ • Filter     │
└──────────────┘    └──────────────┘    │ • Mappings   │    │ • Insert     │
                                        └──────────────┘    └──────────────┘
                                                                    │
                                                                    ▼
                                                            ┌──────────────┐
                                                            │   COMPLETE   │
                                                            │              │
                                                            │ • Save file  │
                                                            │ • Write log  │
                                                            │ • Prompt open│
                                                            └──────────────┘
```

### 5.2 Processing Flowchart

```
                        ┌─────────────────┐
                        │  START SESSION  │
                        └────────┬────────┘
                                 │
                        ┌────────▼────────┐
                        │  Validate Input │
                        │  (mappings, file)│
                        └────────┬────────┘
                                 │
                        ┌────────▼────────┐
                        │ Select Save Path │
                        └────────┬────────┘
                                 │
                        ┌────────▼────────┐
                        │ Create Workbook  │
                        │ Start Selenium   │
                        └────────┬────────┘
                                 │
              ┌─────────────────────────────────────┐
              │        FOR EACH ROW IN EXCEL        │
              │  ┌──────────────────────────────┐   │
              │  │   Write text data to output  │   │
              │  └─────────────┬────────────────┘   │
              │                │                    │
              │  ┌─────────────▼────────────────┐   │
              │  │  FOR EACH MAPPING            │   │
              │  │  ┌────────────────────────┐  │   │
              │  │  │ Search images (browser)│  │   │
              │  │  └───────────┬────────────┘  │   │
              │  │              │               │   │
              │  │  ┌───────────▼────────────┐  │   │
              │  │  │ Download candidates    │  │   │
              │  │  │ (parallel, max 8)      │  │   │
              │  │  └───────────┬────────────┘  │   │
              │  │              │               │   │
              │  │  ┌───────────▼────────────┐  │   │
              │  │  │ Apply filters:         │  │   │
              │  │  │ • Brightness check     │  │   │
              │  │  │ • Color variance       │  │   │
              │  │  │ • Unique colors        │  │   │
              │  │  │ • Portrait preference  │  │   │
              │  │  └───────────┬────────────┘  │   │
              │  │              │               │   │
              │  │  ┌───────────▼────────────┐  │   │
              │  │  │ Resize & save temp     │  │   │
              │  │  └───────────┬────────────┘  │   │
              │  │              │               │   │
              │  │  ┌───────────▼────────────┐  │   │
              │  │  │ Insert into Excel cell │  │   │
              │  │  └────────────────────────┘  │   │
              │  └─────────────────────────────┘   │
              │                                    │
              │  ┌──────────────────────────────┐   │
              │  │   Update progress bar        │   │
              │  └──────────────────────────────┘   │
              │                                    │
              │  ┌──────────────────────────────┐   │
              │  │   Check cancel requested?    │───┼──▶ CANCEL
              │  └──────────────────────────────┘   │
              └────────────────┬────────────────────┘
                               │
                      ┌────────▼────────┐
                      │  Close Workbook  │
                      │  Quit Selenium   │
                      │  Cleanup temps   │
                      │  Write log       │
                      └────────┬────────┘
                               │
                      ┌────────▼────────┐
                      │    COMPLETE     │
                      └─────────────────┘
```

### 5.3 Image Search Flow

```
Search Term + Theme
       │
       ▼
┌─────────────────────────────────────────────────────────┐
│               URL Construction                          │
│  Bing:  bing.com/images/search?q={term}+{theme}        │
│  Google: google.com/search?tbm=isch&q={term}+{theme}   │
│  DDG:   duckduckgo.com/?q={term}+{theme}&ia=images     │
└─────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────┐
│               Selenium Scraping                         │
│  • Load page with Chrome WebDriver                     │
│  • Wait for images to load (1 second)                  │
│  • Extract image URLs from DOM                         │
│  • Parse JSON metadata (Bing: a.iusc elements)         │
└─────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────┐
│               URL Validation                            │
│  • Check URL starts with http/https                    │
│  • Block stock photo sites (BAD_SITES list)            │
│  • Verify common image extensions                       │
│  • Shuffle and limit results (max 36)                  │
└─────────────────────────────────────────────────────────┘
       │
       ▼
    Return URL List
```

---

# Part II: Technical Specification

---

## 6. Architecture Overview

### 6.1 High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                         CANVEX APPLICATION                          │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │                    PRESENTATION LAYER                        │   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────┐  │   │
│  │  │ Main Window │  │  Dialogs    │  │  Theme Management   │  │   │
│  │  │ (GUI)       │  │  (Settings, │  │  (Light/Dark/System)│  │   │
│  │  │             │  │   Help...)  │  │                     │  │   │
│  │  └─────────────┘  └─────────────┘  └─────────────────────┘  │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                              │                                      │
│                              ▼                                      │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │                    BUSINESS LOGIC LAYER                      │   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────┐  │   │
│  │  │ WorkerUltra │  │  Settings   │  │  Column Mapping     │  │   │
│  │  │ (QThread)   │  │  Manager    │  │  Handler            │  │   │
│  │  └─────────────┘  └─────────────┘  └─────────────────────┘  │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                              │                                      │
│                              ▼                                      │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │                    DATA/SERVICE LAYER                        │   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────┐  │   │
│  │  │ Image Search│  │  Image      │  │  Excel I/O          │  │   │
│  │  │ (Selenium)  │  │  Processing │  │  (pandas/xlsxwriter)│  │   │
│  │  │             │  │  (Pillow)   │  │                     │  │   │
│  │  └─────────────┘  └─────────────┘  └─────────────────────┘  │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      EXTERNAL SERVICES                              │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────────┐ │
│  │ Bing Images │  │ Google Imgs │  │ DuckDuckGo Images           │ │
│  └─────────────┘  └─────────────┘  └─────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────┘
```

### 6.2 Threading Model

```
┌────────────────────────────────────────────────────────────┐
│                     MAIN THREAD (GUI)                      │
│  • PyQt5 event loop                                        │
│  • User interaction handling                               │
│  • UI updates via signals                                  │
└────────────────────────────────┬───────────────────────────┘
                                 │
                                 │ spawn
                                 ▼
┌────────────────────────────────────────────────────────────┐
│                  WORKER THREAD (WorkerUltra)               │
│  • Heavy I/O operations                                    │
│  • Selenium browser control                                │
│  • Network requests                                        │
│  • File writes                                             │
│                                                            │
│  ┌────────────────────────────────────────────────────┐   │
│  │          THREAD POOL (concurrent.futures)          │   │
│  │  • Parallel image downloads                        │   │
│  │  • Max workers: min(20, CPU_count * 2)            │   │
│  └────────────────────────────────────────────────────┘   │
└────────────────────────────────────────────────────────────┘

Signal Communication:
  Worker ──sig_overall──▶ UI (progress percentage)
  Worker ──sig_step────▶ UI (per-download progress)
  Worker ──sig_log─────▶ UI (log messages)
  Worker ──sig_done────▶ UI (success with path)
  Worker ──sig_error───▶ UI (error message)
```

### 6.3 Component Diagram

```
┌──────────────────────────────────────────────────────────────────────────┐
│                              Canvex.py                                    │
│                                                                          │
│  ┌───────────────────────┐      ┌────────────────────────────────────┐  │
│  │ CanvaImageExcelCreator│      │           WorkerUltra              │  │
│  │ (QWidget)             │      │           (QThread)                │  │
│  │                       │      │                                    │  │
│  │ • __init__()          │      │ • __init__()                       │  │
│  │ • load_excel()        │◀────▶│ • run()                            │  │
│  │ • add_mapping()       │      │ • log()                            │  │
│  │ • start_session()     │      │ • get_resolution()                 │  │
│  │ • cancel_session()    │      │                                    │  │
│  │ • load_settings()     │      │ Signals:                           │  │
│  │ • save_settings()     │      │ • sig_overall                      │  │
│  │ • show_*_dialog()     │      │ • sig_step                         │  │
│  └───────────────────────┘      │ • sig_log                          │  │
│                                 │ • sig_done                          │  │
│                                 │ • sig_error                         │  │
│                                 └────────────────────────────────────┘  │
│                                                                          │
│  ┌───────────────────────────────────────────────────────────────────┐  │
│  │                        Utility Functions                           │  │
│  │                                                                    │  │
│  │  Image Search:           Image Processing:     Helpers:           │  │
│  │  • bing_urls()           • dl_resize()         • create_driver()  │  │
│  │  • google_urls()         • is_valid_image_url()• wait()           │  │
│  │  • ddg_urls()            • DOWNLOAD_CACHE      • system_dark_mode()│
│  │  • fetch_image_urls()                          • resource_path()  │  │
│  │                                                • get_writable_dir()│
│  │                                                • get_temp_dir()   │  │
│  └───────────────────────────────────────────────────────────────────┘  │
│                                                                          │
│  ┌───────────────────────────────────────────────────────────────────┐  │
│  │                         Constants                                  │  │
│  │  • THEMES (list)         • DARK_STYLE (CSS)    • LIGHT_STYLE (CSS)│  │
│  │  • BAD_SITES (list)      • DOWNLOAD_CACHE (dict)                  │  │
│  └───────────────────────────────────────────────────────────────────┘  │
│                                                                          │
└──────────────────────────────────────────────────────────────────────────┘
```

---

## 7. Module Documentation

### 7.1 Main Application Class

#### `CanvaImageExcelCreator(QWidget)`

**Purpose:** Main GUI window and application controller

**Attributes:**

| Attribute | Type | Description |
|-----------|------|-------------|
| `excel_path` | `str` | Path to loaded Excel file |
| `columns` | `list[str]` | Column names from Excel |
| `worker` | `WorkerUltra` | Background processing thread |
| `session_running` | `bool` | Processing state flag |
| `manual_theme_override` | `str\|None` | Theme setting ("light"/"dark"/None) |
| `filter_portrait` | `bool` | Portrait filter setting |
| `filter_bw` | `bool` | B&W filter setting |
| `filter_cartoon` | `bool` | Cartoon filter setting |
| `settings_path` | `str` | Path to settings JSON file |
| `last_dir` | `str` | Last used directory |
| `app_icon` | `QIcon` | Application icon |

**Key Methods:**

```python
def load_excel(self):
    """Open file dialog, load Excel, populate columns"""
    
def add_mapping(self):
    """Add row to mapping table with dropdowns"""
    
def start_session(self):
    """Validate, create WorkerUltra, start processing"""
    
def cancel_session(self):
    """Request worker cancellation"""
    
def load_settings(self):
    """Load settings from JSON file"""
    
def save_settings(self, mappings):
    """Save settings to JSON file"""
    
def lock_ui(self) / def unlock_ui(self):
    """Enable/disable UI during processing"""
```

### 7.2 Worker Thread Class

#### `WorkerUltra(QThread)`

**Purpose:** Background thread for heavy processing operations

**Constructor Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `excel_path` | `str` | Input Excel file path |
| `mappings` | `list[tuple]` | [(input_col, output_col), ...] |
| `save_path` | `str` | Output Excel file path |
| `theme` | `str` | Selected theme preset |
| `custom_theme` | `str` | Custom theme text |
| `resolution` | `str` | Resolution preset |
| `custom_res` | `str` | Custom resolution value |
| `fmt` | `str` | Output format (png/jpg/webp) |
| `jpg_quality` | `int` | JPG quality (60-100) |
| `browser` | `str` | Search engine name |
| `selected_sheet` | `str` | Excel sheet name |

**Signals:**

| Signal | Type | Purpose |
|--------|------|---------|
| `sig_overall` | `pyqtSignal(int)` | Overall progress (0-100) |
| `sig_step` | `pyqtSignal(int)` | Per-image progress (0-100) |
| `sig_log` | `pyqtSignal(str)` | Log message |
| `sig_done` | `pyqtSignal(str)` | Completion with output path |
| `sig_error` | `pyqtSignal(str)` | Error message |

**Processing Flow:**
```python
def run(self):
    # 1. Load Excel with pandas
    # 2. Create xlsxwriter workbook
    # 3. Initialize Selenium driver
    # 4. Create thread pool for downloads
    # 5. For each row:
    #    - Write text data
    #    - For each mapping:
    #      - Search images
    #      - Download in parallel
    #      - Filter and select best
    #      - Save temp file
    #      - Insert into Excel
    # 6. Close workbook
    # 7. Cleanup and emit done/error
```

### 7.3 Image Search Functions

#### `bing_urls(driver, term, theme, limit=36)`
```python
"""Scrape Bing Images for candidate URLs.

Args:
    driver: Selenium WebDriver instance
    term: Search term string
    theme: Theme string to append to query
    limit: Maximum URLs to return

Returns:
    list[str]: Image URLs (may be fewer than limit)

Implementation:
    - Constructs URL: bing.com/images/search?q={term}+{theme}
    - Parses JSON from 'a.iusc' elements
    - Filters through is_valid_image_url()
    - Retries once if no results
"""
```

#### `google_urls(driver, term, theme, limit=36)`
```python
"""Scrape Google Images for candidate URLs.

Note: Google DOM changes frequently. This is best-effort.
Falls back to Bing if no results.
"""
```

#### `ddg_urls(driver, term, theme, limit=36)`
```python
"""Scrape DuckDuckGo Images for candidate URLs.

Alternative search engine, useful when others are rate-limited.
"""
```

#### `fetch_image_urls(driver, term, theme, browser, limit=36)`
```python
"""Dispatch to appropriate search function based on browser string.

Args:
    browser: "Bing Images", "Google Images", or "DuckDuckGo"
    
Returns:
    URL list from selected search engine
"""
```

### 7.4 Image Processing Functions

#### `dl_resize(url, target)`
```python
"""Download image and resize with quality filtering.

Args:
    url: Image URL to download
    target: Target size (longest side in pixels)

Returns:
    PIL.Image on success, None on failure

Filtering Logic:
    1. Download with 7-second timeout
    2. Reject if mean brightness < 10 or > 245
    3. Reject if max channel stddev < 12 (grayscale)
    4. Reject if unique colors < 24 (cartoon/graphic)
    5. Reject if resized dimensions < 80px
    6. Resize using LANCZOS for quality

Caching:
    Uses DOWNLOAD_CACHE dict for URLs < 1MB
"""
```

#### `is_valid_image_url(url)`
```python
"""Validate image URL for processing.

Checks:
    - Starts with http/https
    - Not from BAD_SITES (stock photo domains)
    - Has common image extension (.jpg, .png, .webp)
    - Or contains imgres?url= (Bing redirect)
"""
```

### 7.5 Utility Functions

#### `create_driver()`
```python
"""Create configured Selenium Chrome WebDriver.

Options applied:
    - disable-blink-features=AutomationControlled
    - disable-gpu, no-sandbox
    - incognito mode
    - window-size=1280,1000
    - page_load_timeout=12

Returns:
    webdriver.Chrome instance
"""
```

#### `system_dark_mode()`
```python
"""Detect OS dark mode setting.

Windows: Read registry AppsUseLightTheme
macOS: Check AppleInterfaceStyle
Linux: Default to dark

Returns:
    bool: True if dark mode
"""
```

#### `resource_path(filename)`
```python
"""Get path to bundled resource file.

Handles both development and PyInstaller frozen modes.
"""
```

#### `get_writable_dir()`
```python
"""Get writable directory for settings/temp files.

Returns:
    macOS: ~/Library/Application Support/Canvex
    Windows: %APPDATA%/Canvex
    Development: Script directory
"""
```

---

## 8. Data Flow & Processing

### 8.1 Input Data Flow

```
Excel File (.xlsx)
       │
       ▼
┌─────────────────────────────────┐
│  pandas.read_excel()            │
│  • Parse worksheet              │
│  • Read headers → self.columns  │
│  • Data stored in DataFrame     │
└─────────────────────────────────┘
       │
       ▼
User Configuration
(Theme, Browser, Resolution, Format, Mappings)
       │
       ▼
┌─────────────────────────────────┐
│  WorkerUltra.__init__()         │
│  • Store all parameters         │
│  • Initialize cancel flag       │
└─────────────────────────────────┘
```

### 8.2 Processing Data Flow

```
For each row in DataFrame:
       │
       ▼
┌─────────────────────────────────┐
│  Text Data                      │
│  • Copy all cell values to      │
│    output workbook              │
└─────────────────────────────────┘
       │
       ▼
For each (input_col, output_col) mapping:
       │
       ▼
┌─────────────────────────────────┐
│  Get Search Term                │
│  cell_value = row[input_col]    │
│  search_query = value + theme   │
└─────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────┐
│  Image Search                   │
│  urls = fetch_image_urls(       │
│      driver, term, theme,       │
│      browser, limit=24          │
│  )                              │
└─────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────┐
│  Parallel Download              │
│  urls = urls[:8]  # limit       │
│  futures = [pool.submit(        │
│      dl_resize, u, resolution   │
│  ) for u in urls]               │
└─────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────┐
│  Select Best Image              │
│  • Prefer portrait (h > w)      │
│  • Use first valid as fallback  │
│  final_img = portrait or fallback│
└─────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────┐
│  Save & Insert                  │
│  • Save to temp file            │
│  • Insert via xlsxwriter        │
│    sheet.insert_image(          │
│        row, col, path,          │
│        {'x_scale': 0.2,         │
│         'y_scale': 0.2}         │
│    )                            │
└─────────────────────────────────┘
```

### 8.3 Output Data Structure

**Excel Output:**
```
┌─────────────────────────────────────────────────────────────────┐
│                    Output Workbook Structure                     │
├─────────────────────────────────────────────────────────────────┤
│  Sheet Name: Based on output filename (max 31 chars)            │
│                                                                 │
│  Headers (Row 0):                                               │
│  [Original Col 1] [Original Col 2] ... [New Image Col 1] ...    │
│                                                                 │
│  Data Rows (Row 1+):                                            │
│  [Text Data] [Text Data] ... [Embedded Image] ...               │
│                                                                 │
│  Image Properties:                                              │
│  • object_position: 1 (move with cells)                         │
│  • x_scale: 0.20                                                │
│  • y_scale: 0.20                                                │
│                                                                 │
│  Column Width: 22 (for image columns)                           │
│  Row Height: 120px (default for all rows)                       │
└─────────────────────────────────────────────────────────────────┘
```

---

## 9. External Dependencies

### 9.1 Python Package Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `PyQt5` | ≥5.15 | GUI framework |
| `pandas` | ≥1.3 | Excel reading |
| `xlsxwriter` | ≥3.0 | Excel writing with images |
| `Pillow` | ≥9.0 | Image processing |
| `selenium` | ≥4.0 | Web scraping |
| `webdriver-manager` | ≥3.8 | ChromeDriver management |
| `requests` | ≥2.28 | HTTP requests |
| `qtawesome` | ≥1.0 | (Optional) Font Awesome icons |

### 9.2 System Dependencies

| Dependency | Purpose | Required |
|------------|---------|----------|
| Chrome Browser | Selenium web scraping | Yes |
| ChromeDriver | Chrome automation | Auto-installed |

### 9.3 Blocked Domains (BAD_SITES)

Images from these domains are filtered out:
```python
BAD_SITES = [
    "shutterstock", "alamy", "getty", "adobe",
    "dreamstime", "depositphotos", "123rf",
    "bigstock", "vectorstock", "istock"
]
```

---

## 10. Configuration & Settings

### 10.1 Settings File Location

| Platform | Path |
|----------|------|
| macOS (bundled) | `~/Library/Application Support/Canvex/canva_last_settings.json` |
| Windows (bundled) | `%APPDATA%/Canvex/canva_last_settings.json` |
| Development | `./canva_last_settings.json` |

### 10.2 Settings Schema

```json
{
  "theme": "headshot portrait closeup face",
  "custom_theme": "",
  "resolution": "720p",
  "custom_res": "",
  "format": "PNG",
  "jpg_quality": 90,
  "browser": "Bing Images",
  "last_excel_dir": "/path/to/directory",
  "mappings": [
    ["input_column", "output_column"],
    ["actor", "photo"]
  ],
  "filter_portrait": true,
  "filter_bw": true,
  "filter_cartoon": true,
  "recent_files": [
    "/path/to/file1.xlsx",
    "/path/to/file2.xlsx"
  ]
}
```

### 10.3 Theme Configuration

**Available Themes:**
```python
THEMES = [
    "headshot portrait closeup face",
    "cinematic lighting portrait",
    "studio headshot clean background",
    "dramatic portrait closeup",
    "smiling closeup face",
    "full body portrait",
    "natural daylight portrait",
    "magazine cover portrait",
    "Custom Theme..."
]
```

### 10.4 Resolution Mapping

```python
resolution_table = {
    "240p": 240,
    "360p": 360,
    "480p": 480,
    "720p": 720,
    "1080p": 1080,
    "1440p": 1440,
    "2160p": 2160,
    "3840p": 3840,
}
# Custom: 240-4000 range accepted
```

---

## 11. Error Handling & Logging

### 11.1 Error Handling Strategy

```
┌─────────────────────────────────────────────────────────────────┐
│                    Error Handling Hierarchy                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  Level 1: Image Download Errors                                 │
│  ├─ Retry twice with 0.2s backoff                              │
│  ├─ On failure: Return None, try next URL                      │
│  └─ Logged but processing continues                            │
│                                                                 │
│  Level 2: Search Errors                                         │
│  ├─ Retry page load once                                       │
│  ├─ Google failure: Fallback to Bing                           │
│  └─ All failures: Log warning, skip this cell                  │
│                                                                 │
│  Level 3: Processing Errors                                     │
│  ├─ Caught in try/except/finally                               │
│  ├─ Workbook always closed (checkpoint)                        │
│  ├─ Selenium always quit                                       │
│  ├─ Error log file written                                     │
│  └─ sig_error emitted to UI                                    │
│                                                                 │
│  Level 4: Application Errors                                    │
│  ├─ safe_exit() prevents crash on force quit                   │
│  └─ KeyboardInterrupt handled gracefully                       │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 11.2 Log File Structure

**Normal Log (`_log.txt`):**
```
[START] 2025-01-15 10:30:00
[LOG] Theme: headshot portrait closeup face
[LOG] Search Browser: Bing Images
[LOG] Format: png
[LOG] JPG Quality: 90
[LOG] Resolution: 720px
[LOG] Starting Selenium...
[SEARCH] Tom Hanks
[URLS] (Bing Images) 24 found: ['https://...', ...]
[SEARCH] Brad Pitt
[URLS] (Bing Images) 24 found: ['https://...', ...]
[WARN] No images found: Unknown Person

Time taken: 0h 5m 23s
```

**Error Log (`_ERROR_log.txt`):**
```
[START] 2025-01-15 10:30:00
[LOG] Theme: headshot portrait closeup face
...processing logs...

=== ERROR ===
Traceback (most recent call last):
  File "Canvex.py", line XXX, in run
    ...
Exception: Error message
```

### 11.3 Signal-Based UI Updates

```python
# Worker emits log message
self.sig_log.emit(f"[SEARCH] {term}")

# Main thread receives via signal connection
self.worker.sig_log.connect(lambda m: print(m))

# Progress updates
self.sig_overall.emit(int((ri + 1) / total_rows * 100))
self.sig_step.emit(30)
```

---

## 12. Performance Optimization

### 12.1 Parallel Processing

**Thread Pool Configuration:**
```python
cpus = os.cpu_count() or 4
maxw = min(20, max(6, cpus * 2))  # 6-20 workers
pool = concurrent.futures.ThreadPoolExecutor(max_workers=maxw)
```

**Parallel Download Pattern:**
```python
futures = [pool.submit(dl_resize, u, res) for u in urls[:8]]

for fut in concurrent.futures.as_completed(futures):
    img = fut.result()
    if img and (h > w) and portrait_img is None:
        portrait_img = img
        break  # Early exit on portrait found
```

### 12.2 HTTP Connection Pooling

```python
# Global session with connection pooling
requests_session = requests.Session()
requests_session.headers.update({
    "User-Agent": "Mozilla/5.0..."
})

# Adapter configuration
adapter = HTTPAdapter(
    pool_connections=100,
    pool_maxsize=100,
    max_retries=Retry(total=2, backoff_factor=0.2)
)
requests_session.mount("http://", adapter)
requests_session.mount("https://", adapter)

# Monkey patch for global usage
requests.get = fast_get
```

### 12.3 Image Caching

```python
DOWNLOAD_CACHE = {}

def dl_resize(url, target):
    # Check cache first
    if url in DOWNLOAD_CACHE:
        content = DOWNLOAD_CACHE[url]
    else:
        response = requests.get(url, timeout=7)
        content = response.content
        # Cache images under 1MB
        if len(content) <= 1024 * 1024:
            DOWNLOAD_CACHE[url] = content
```

### 12.4 Selenium Optimization

```python
opts = webdriver.ChromeOptions()

# Performance options
opts.add_argument("--disable-gpu")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-extensions")
opts.add_argument("--incognito")
opts.add_argument("--renderer-process-limit=3")

# Anti-detection
opts.add_argument("--disable-blink-features=AutomationControlled")

# Timeout
driver.set_page_load_timeout(12)
```

### 12.5 Memory Management

```python
# Allow large images
Image.MAX_IMAGE_PIXELS = None

# Cleanup temp files on success
if completed:
    for f in temp_files:
        try:
            os.remove(f)
        except:
            pass
```

---

## 13. Security Considerations

### 13.1 Network Security

| Risk | Mitigation |
|------|------------|
| Malicious image URLs | URL validation before download |
| Timeout attacks | 7-second timeout on requests |
| SSL verification | Default SSL verification enabled |

### 13.2 File System Security

| Risk | Mitigation |
|------|------------|
| Path traversal | Uses os.path.join, no user paths in temp names |
| Temp file exposure | Temp files in system temp dir, cleaned on success |
| Settings file tampering | JSON parsing with exception handling |

### 13.3 Application Security

| Risk | Mitigation |
|------|------------|
| Selenium detection | Anti-detection flags in Chrome options |
| Resource exhaustion | Limited parallel workers, page timeouts |
| Memory overflow | Image pixel limit configurable |

---

## 14. Deployment & Packaging

### 14.1 PyInstaller Configuration

**Spec File (`Canvex.spec`):**
```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Canvex.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app_icon.ico', '.'),
        ('splash.png', '.'),
    ],
    hiddenimports=[
        'PyQt5.QtSvg',
        'qtawesome',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Canvex',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='app_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Canvex',
)

# macOS bundle
app = BUNDLE(
    coll,
    name='Canvex.app',
    icon='app_icon.icns',
    bundle_identifier='com.canvex.app',
)
```

### 14.2 Build Commands

**macOS:**
```bash
pyinstaller Canvex.spec --noconfirm
# Output: dist/Canvex.app
```

**Windows:**
```bash
pyinstaller Canvex.spec --noconfirm
# Output: dist/Canvex/Canvex.exe
```

### 14.3 Required Assets

| File | Purpose | Required |
|------|---------|----------|
| `app_icon.ico` | Windows icon | Yes |
| `app_icon.icns` | macOS icon | Yes (for .app) |
| `splash.png` | Splash screen | Optional |
| `logo.svg` | Alternative logo | Optional |

### 14.4 Distribution Checklist

- [ ] Test on clean macOS installation
- [ ] Test on clean Windows installation
- [ ] Verify Chrome/ChromeDriver compatibility
- [ ] Check code signing (macOS notarization)
- [ ] Verify file associations work
- [ ] Test with various Excel files
- [ ] Verify settings persistence
- [ ] Check temp file cleanup

---

## 15. API Reference

### 15.1 PyQt5 Signals

#### WorkerUltra Signals

| Signal | Signature | Description |
|--------|-----------|-------------|
| `sig_overall` | `pyqtSignal(int)` | Overall progress 0-100 |
| `sig_step` | `pyqtSignal(int)` | Per-item progress 0-100 |
| `sig_log` | `pyqtSignal(str)` | Log message string |
| `sig_done` | `pyqtSignal(str)` | Success with output path |
| `sig_error` | `pyqtSignal(str)` | Error message |

### 15.2 Public Functions

#### Image Search

```python
def fetch_image_urls(
    driver: webdriver.Chrome,
    term: str,
    theme: str,
    browser: str = "Bing Images",
    limit: int = 36
) -> list[str]:
    """
    Fetch image URLs from specified search engine.
    
    Returns list of validated image URLs.
    """
```

#### Image Processing

```python
def dl_resize(
    url: str,
    target: int
) -> Optional[Image.Image]:
    """
    Download and resize image with quality filtering.
    
    Returns PIL Image or None on failure.
    """
```

#### Selenium

```python
def create_driver() -> webdriver.Chrome:
    """
    Create configured Chrome WebDriver.
    
    Returns ready-to-use WebDriver instance.
    """
```

### 15.3 Class Interfaces

#### CanvaImageExcelCreator

```python
class CanvaImageExcelCreator(QWidget):
    def __init__(self):
        """Initialize main window and load settings."""
    
    def load_excel(self) -> None:
        """Open file dialog and load Excel."""
    
    def add_mapping(self) -> None:
        """Add new mapping row to table."""
    
    def start_session(self) -> None:
        """Validate and start processing."""
    
    def cancel_session(self) -> None:
        """Request safe cancellation."""
    
    def load_settings(self) -> None:
        """Load settings from JSON."""
    
    def save_settings(self, mappings: list) -> None:
        """Save settings to JSON."""
```

#### WorkerUltra

```python
class WorkerUltra(QThread):
    def __init__(
        self,
        excel_path: str,
        mappings: list[tuple[str, str]],
        save_path: str,
        theme: str,
        custom_theme: str,
        resolution: str,
        custom_res: str,
        fmt: str,
        jpg_quality: int,
        browser: str,
        selected_sheet: Optional[str] = None
    ):
        """Initialize worker with processing parameters."""
    
    def run(self) -> None:
        """Main processing loop (runs in separate thread)."""
    
    def log(self, msg: str) -> None:
        """Add message to log and emit signal."""
    
    def get_resolution(self) -> int:
        """Parse resolution setting to pixel value."""
```

---

# Appendices

---

## Appendix A: File Structure

```
Canvex/
├── Canvex.py                 # Main application source
├── Canvex.spec               # PyInstaller spec file
├── canva_last_settings.json  # User settings (auto-created)
├── app_icon.ico              # Windows icon
├── app_icon.icns             # macOS icon
├── splash.png                # Splash screen image
├── logo.svg                  # Alternative logo
├── docs/
│   ├── USER_GUIDE.md         # User documentation
│   └── TECHNICAL_DOCS.md     # This document
└── build/
    └── Canvex/               # PyInstaller output
        ├── Canvex            # Executable
        └── ...               # Supporting files
```

---

## Appendix B: Settings Schema

### Full JSON Schema

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "Canvex Settings",
  "type": "object",
  "properties": {
    "theme": {
      "type": "string",
      "description": "Selected image theme preset",
      "default": "headshot portrait closeup face"
    },
    "custom_theme": {
      "type": "string",
      "description": "Custom theme text when 'Custom Theme...' selected",
      "default": ""
    },
    "resolution": {
      "type": "string",
      "enum": ["240p", "360p", "480p", "720p", "1080p", "1440p", "2160p", "3840p", "Custom…"],
      "default": "720p"
    },
    "custom_res": {
      "type": "string",
      "description": "Custom resolution value (240-4000)",
      "default": ""
    },
    "format": {
      "type": "string",
      "enum": ["PNG", "JPG", "WEBP"],
      "default": "PNG"
    },
    "jpg_quality": {
      "type": "integer",
      "minimum": 60,
      "maximum": 100,
      "default": 90
    },
    "browser": {
      "type": "string",
      "enum": ["Bing Images", "Google Images", "DuckDuckGo"],
      "default": "Bing Images"
    },
    "last_excel_dir": {
      "type": "string",
      "description": "Last directory used for file dialogs"
    },
    "mappings": {
      "type": "array",
      "items": {
        "type": "array",
        "items": {"type": "string"},
        "minItems": 2,
        "maxItems": 2
      },
      "description": "Array of [input_column, output_column] pairs"
    },
    "filter_portrait": {
      "type": "boolean",
      "default": true
    },
    "filter_bw": {
      "type": "boolean",
      "default": true
    },
    "filter_cartoon": {
      "type": "boolean",
      "default": true
    },
    "recent_files": {
      "type": "array",
      "items": {"type": "string"},
      "maxItems": 10
    }
  }
}
```

---

## 16. Recent Enhancements (December 2025)

### 16.1 Session Persistence System

#### Last Directory Memory
**Implementation:**
- Stores last directory path in `canva_last_settings.json` under key `"last_excel_dir"`
- Method: `load_basic_settings()` - called at end of `__init__()`
- When user selects file via `QFileDialog`, directory is extracted and saved
- On startup, file browser opens to this directory (if it exists)

**Code:**
```python
def load_basic_settings(self):
    # Restores: theme, resolution, browser, last_dir, filters
    # Called at END of __init__() before any UI interaction
    # Does NOT try to access self.columns (not loaded yet)
```

**File Location:**
- macOS: `~/Library/Application Support/Canvex/canva_last_settings.json`
- Windows: `%APPDATA%/Canvex/canva_last_settings.json`
- Dev: Same folder as `Canvex.py`

#### Recent Files Tracking
**Implementation:**
- Maintains list of last 10 opened files in `"recent_files"` array
- Each entry stores full file path
- Method: `_add_to_recent_files(filepath)` - called when file is opened
- Detects missing files and marks with visual indicator

**Features:**
- Removes duplicates (moves to front of list)
- Validates file existence
- Displays in File menu with ✓/✗ indicators
- Click to open any recent file

**Data Structure:**
```json
{
  "recent_files": [
    "/path/to/file1.xlsx",
    "/path/to/file2.xlsx"
  ]
}
```

### 16.2 Mapping Management System

#### Auto-Save & History
**When Mappings Are Saved:**
- Only when user starts processing and selects output file
- Called from `start_session()` method
- Method: `save_settings(mappings)` saves current mappings

**History Tracking:**
- Last 5 configurations kept in `"mapping_history"` array
- Each entry has:
  - `"timestamp"`: ISO format datetime string
  - `"mappings"`: List of [input_column, output_column] pairs
- Datetime format: `datetime.now().isoformat()`
- Automatic cleanup: keeps only last 5 entries

**Data Structure:**
```json
{
  "mapping_history": [
    {
      "timestamp": "2025-12-30T14:32:15.123456",
      "mappings": [
        ["name", "photo"],
        ["title", "headshot"]
      ]
    }
  ]
}
```

**Merging Settings:**
```python
def save_settings(self, mappings):
    # Preserves existing recent_files
    # Adds new mapping to history
    # Trims history to last 5 entries
    # Auto-creates settings directory if missing
```

#### Previous Mappings Dialog
**Method:** `show_previous_mappings()`
**Dialog Specifications:**
- Window size: 700×600 pixels (resizable, min 600×500)
- List widget: 200px min height
- Preview table: 200-250px height
- Auto-selects first mapping on open

**Components:**
1. **Mapping List (QListWidget):**
   - Shows all configurations with timestamps
   - Format: `"Mapping #N - YYYY-MM-DD HH:MM:SS (X mappings)"`
   - Hover effect: Gray background (#3a3a3c dark, #f0f0f0 light)

2. **Preview Table (QTableWidget):**
   - 2 columns: Input Column, Output Column
   - Shows all mappings for selected configuration
   - Alternating row colors for visibility
   - Updates when selection changes

3. **Load Selection:**
   - Single click: Select, double-click: Load
   - Calls `load_mapping_from_history(mappings)`
   - Dialog closes and mappings appear in main table

4. **Reset Button:**
   - "Reset All Mappings" (orange button)
   - Shows confirmation dialog
   - Clears all rows from mapping table
   - Requires user confirmation

**Implementation:**
```python
def show_previous_mappings(self):
    # Create dialog with proper sizing
    # Populate list from mapping_history
    # Connect signals for live preview
    # Handle load and reset actions
```

### 16.3 Column Mapping Enhancements

#### Smart Column Detection
**Problem Solved:** Old mappings become invalid when columns change

**Solution:**
Method: `load_mapping_from_history(mappings)`
- Checks if each destination column still exists
- If missing: automatically switches to "Create New Column..." mode
- Sets new column name to the original destination
- User can edit before processing

**Code Flow:**
```python
for src, dst in mappings:
    # Add mapping row
    if dst in self.columns:
        # Column exists - use it directly
        dd_out.setCurrentIndex(existing_index)
    else:
        # Column missing - use create new mode
        dd_out.setCurrentIndex(create_new_index)
        txt_new.setText(dst)  # Pre-fill with original name
        txt_new.setVisible(True)
```

#### Dynamic Delete Functionality
**Problem Solved:** Captured row numbers became invalid after deletions

**Solution:**
Methods: `delete_row_by_button(button)` instead of `delete_row(row)`

**Implementation:**
```python
def delete_row_by_button(self, button):
    # Find which row contains this button
    for row in range(self.table.rowCount()):
        if self.table.cellWidget(row, 4) == button:
            self.table.removeRow(row)
            # Renumber remaining rows
            return
    
# Button connection:
btn_del.clicked.connect(lambda _, button=btn_del: 
                        self.delete_row_by_button(button))
```

**Benefits:**
- Works correctly with history-loaded mappings
- Renumbers all rows automatically
- No index conflicts

#### Text Field Visibility Management
**Method:** `toggle_new_col(row)`

**Behavior:**
- Shown: Only when "Create New Column..." is selected
- Hidden: All other cases
- Auto-cleared: When hidden, text is cleared

**Implementation:**
```python
def toggle_new_col(self, row):
    is_create_new = dd_out.currentText() == "Create New Column..."
    
    if is_create_new:
        txt.setVisible(True)
        txt.setEnabled(True)
    else:
        txt.setVisible(False)
        txt.setEnabled(False)
        txt.setText("")  # Clear to prevent confusion
```

### 16.4 Visual & Styling Improvements

#### Native Theme Integration
**Problem Solved:** Custom QComboBox styling with dark background was:
- Making selections invisible
- Not respecting system theme
- Breaking native macOS appearance

**Solution:**
Removed custom background colors from QComboBox CSS:

**Before:**
```css
QComboBox {
    background: #2d2d2d;      /* Dark gray - hides text */
    border: 1px solid #404040;
    color: white;             /* Still invisible on #2d2d2d */
    padding: 6px 10px;
}
QComboBox QAbstractItemView {
    background: #2d2d2d;
    selection-background-color: #0a84ff;
    /* Complex item styling that breaks on macOS */
}
```

**After:**
```css
QComboBox {
    border: 1px solid #404040;    /* Let native background show */
    border-radius: 6px;
    padding: 6px 10px;
    min-height: 22px;
}
QComboBox:hover {
    border: 1px solid #0a84ff;
}
QComboBox:focus {
    border: 2px solid #0a84ff;
}
/* No custom background - uses native rendering */
```

**Impact:**
- Selected text now visible
- Respects Light/Dark theme automatically
- Native macOS dropdown appearance
- Better performance

#### QListWidget Hover Effects
**Location:** Both DARK_STYLE and LIGHT_STYLE constants

**Dark Theme Styling:**
```css
QListWidget {
    background: #252525;
    border: 1px solid #404040;
    border-radius: 6px;
}
QListWidget::item {
    padding: 8px 12px;
    margin: 2px 0px;
    border-radius: 4px;
}
QListWidget::item:hover {
    background: #3a3a3c;      /* Subtle gray */
}
QListWidget::item:selected {
    background: #0a84ff;
    color: white;
}
```

**Light Theme Styling:**
```css
QListWidget::item:hover {
    background: #f0f0f0;      /* Light gray */
}
```

**Features:**
- Hover shows clear visual feedback
- Smooth transitions on macOS 14.2+
- Works in both light and dark modes
- Applied to mappings list in dialog

### 16.5 Bug Fixes

| Issue | Cause | Fix |
|-------|-------|-----|
| Preview not showing mappings | Signal not triggered on auto-select | Call `update_preview()` after `setCurrentRow()` |
| Can't select single mapping | Similar to above | Explicit preview update |
| Delete doesn't work with history | Captured row index changes | Dynamic button lookup via `cellWidget()` |
| Dropdown appears black | Custom dark background | Removed custom background styling |
| Text field always visible | Initial visibility not set | Added `setText("")` and `setVisible(False)` |
| Dialog too cramped | Fixed 500×350 size | Changed to `resize(700, 600)` + min size |
| Hover effects don't work | CSS syntax or macOS incompatibility | Added proper QListWidget::item:hover styling |

---

## Appendix D: Recent Changes Summary

### File Structure Changes
- Settings file automatically created in app directory if missing
- No file structure changes to input Excel files
- Output Excel structure unchanged

### Dependencies Added
- None (used existing PyQt5 infrastructure)

### Code Metrics
- Lines added: ~400 (new methods and styling)
- Methods added: 3 (`reset_all_mappings`, `delete_row_by_button`, improved `toggle_new_col`)
- Classes modified: 1 (`CanvaImageExcelCreator`)

### Testing Recommendations
1. Test on macOS 14.2+ (tested here)
2. Test file opening/recent files with missing files
3. Test delete with 1, 5, 10+ mappings
4. Test load history with changed columns
5. Hover effects on different themes

---

## Appendix E: Future Improvements

**Potential enhancements:**
- Store up to 10 mapping histories instead of 5
- Export/import mapping configurations as JSON
- Batch process multiple files
- Undo/redo functionality
- Drag-and-drop mapping reordering



### Input Formats

| Format | Extension | Support |
|--------|-----------|---------|
| Excel 2007+ | .xlsx | ✓ Full |
| Excel 97-2003 | .xls | ✗ Not supported |
| CSV | .csv | ✗ Not supported |
| OpenDocument | .ods | ✗ Not supported |

### Output Formats

| Format | Extension | Quality | Transparency | Notes |
|--------|-----------|---------|--------------|-------|
| PNG | .png | Lossless | ✓ Yes | Best quality, larger files |
| JPEG | .jpg | 60-100 | ✗ No | Configurable quality |
| WebP | .webp | 95 | ✓ Yes | Modern format, best compression |

### Image Formats (Download)

Accepted extensions in URLs:
- `.jpg`, `.jpeg`
- `.png`
- `.webp`

---

<div align="center">
  <p><em>End of Technical Documentation</em></p>
  <p>Version 1.0 | © 2025 Kunal Pagariya</p>
</div>
