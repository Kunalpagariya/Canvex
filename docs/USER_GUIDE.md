# Canvex User Guide

<div align="center">
  <h2>🖼️ Image Excel Creator</h2>
  <p><em>Automatically search and insert images into Excel files</em></p>
  <p>Version 1.0 | © 2025 Kunal Pagariya</p>
</div>

---

## 📋 Table of Contents

1. [Introduction](#introduction)
2. [Getting Started](#getting-started)
3. [Main Interface Overview](#main-interface-overview)
4. [Step-by-Step Workflow](#step-by-step-workflow)
5. [Configuration Options](#configuration-options)
6. [Column Mappings](#column-mappings)
7. [Settings Panel](#settings-panel)
8. [Output Files](#output-files)
9. [Tips & Best Practices](#tips--best-practices)
10. [Troubleshooting](#troubleshooting)
11. [Keyboard Shortcuts & Navigation](#keyboard-shortcuts--navigation)
12. [FAQ](#faq)

---

## Introduction

### What is Canvex?

Canvex is a powerful desktop application that **automatically searches the web for images** based on text in your Excel spreadsheet and **inserts them directly into a new Excel file**. 

**Perfect for:**
- 📸 Creating employee directories with headshots
- 🎬 Building cast lists with actor photos
- 🏢 Generating product catalogs with images
- 📊 Any data visualization requiring images

### Key Features

| Feature | Description |
|---------|-------------|
| 🔍 **Multi-Engine Search** | Search using Bing, Google, or DuckDuckGo |
| 🎨 **Smart Filtering** | Automatically removes low-quality, B&W, and cartoon images |
| ⚡ **Parallel Processing** | Downloads multiple images simultaneously |
| 💾 **Auto-Save Settings** | Your preferences are remembered between sessions |
| 🌓 **Theme Support** | Light, Dark, or System-following themes |
| 📐 **Flexible Resolution** | From 240p to 4K, or custom values |
| 🎯 **Portrait Priority** | Prefers portrait-oriented images for headshots |

---

## Getting Started

### System Requirements

- **Operating System:** macOS 10.14+ or Windows 10+
- **Internet Connection:** Required for image searches
- **Chrome Browser:** Installed (used for web scraping)
- **RAM:** 4GB minimum, 8GB recommended
- **Storage:** 100MB for app + space for output files

### First Launch

1. **Double-click** the Canvex application icon
2. A **splash screen** will appear briefly
3. The **main window** opens with the drag-and-drop area visible

### Quick Start (5 Minutes)

```
1. Load Excel → 2. Set Theme → 3. Add Mappings → 4. Start → 5. Save
```

---

## Main Interface Overview

### Application Layout

```
┌─────────────────────────────────────────────────────────┐
│  [File] [Settings] [Help] [About]           [Theme]     │  ← Toolbar
├─────────────────────────────────────────────────────────┤
│                                                         │
│        [Select Excel File] or drag & drop               │  ← File Selection
│                                                         │
├─────────────────────────────────────────────────────────┤
│  Image Theme:    [▼ headshot portrait closeup face    ] │
│  Search Browser: [▼ Bing Images                       ] │  ← Configuration
│  Resolution:     [▼ 720p                              ] │
│  Format:         [▼ PNG                               ] │
├─────────────────────────────────────────────────────────┤
│  Column Mappings:                        [+ Add Mapping]│
│  ┌───┬─────────────┬─────────────┬──────────┬───┐      │
│  │ # │ Input Col   │ Output Col  │ New Name │ X │      │  ← Mapping Table
│  ├───┼─────────────┼─────────────┼──────────┼───┤      │
│  │ 1 │ actor_name  │ actor_image │          │ 🗑 │      │
│  └───┴─────────────┴─────────────┴──────────┴───┘      │
├─────────────────────────────────────────────────────────┤
│  Progress: [████████████████░░░░░░░░░░░░░░░] 60%        │  ← Progress Bar
├─────────────────────────────────────────────────────────┤
│              [▶ Start Processing]  [Cancel]             │  ← Action Buttons
└─────────────────────────────────────────────────────────┘
```

### Toolbar Buttons

| Button | Function |
|--------|----------|
| **File** | Open files, view recent files, reveal settings location |
| **Settings** | Configure search engine, filters, resolution, and format |
| **Help** | View this user guide within the app |
| **About** | Application version and contact information |
| **Theme** | Switch between Light, Dark, or System theme |

---

## Step-by-Step Workflow

### Step 1: Load Your Excel File

**Option A: Click to Browse**
1. Click the **"Select Excel File"** button
2. Navigate to your `.xlsx` file
3. Click **Open**

**Option B: Drag and Drop**
1. Open your file explorer/finder
2. Drag the `.xlsx` file onto the Canvex window
3. Release to load

**Sheet Selection:**
- If your Excel has **multiple sheets**, a dialog appears
- Select the sheet you want to process
- Click **Select** or double-click the sheet name

> 💡 **Tip:** Only `.xlsx` files are supported. Convert older `.xls` files first.

### Step 2: Configure Image Settings

#### Image Theme
Choose how images should be searched:

| Theme | Best For |
|-------|----------|
| `headshot portrait closeup face` | Professional headshots, ID photos |
| `cinematic lighting portrait` | Dramatic, artistic portraits |
| `studio headshot clean background` | Corporate/LinkedIn style photos |
| `dramatic portrait closeup` | High-contrast artistic shots |
| `smiling closeup face` | Friendly, approachable photos |
| `full body portrait` | Full-length photos |
| `natural daylight portrait` | Outdoor, natural lighting |
| `magazine cover portrait` | High-fashion style |
| `Custom Theme...` | Enter your own search keywords |

#### Search Browser
Select your preferred search engine:

| Engine | Characteristics |
|--------|-----------------|
| **Bing Images** | ⭐ Recommended. Fastest and most reliable |
| **Google Images** | Alternative results, may be slower |
| **DuckDuckGo** | Privacy-focused, good backup option |

#### Resolution

| Setting | Pixels | Use Case | Speed |
|---------|--------|----------|-------|
| 240p | 240 | Thumbnails | ●●●● Fastest |
| 360p | 360 | Small previews | ●●●○ |
| **480p** | 480 | Standard docs | ●●○○ |
| **720p** | 720 | ⭐ Recommended | ●●○○ |
| 1080p | 1080 | High-quality | ●○○○ |
| 1440p | 1440 | Large displays | ○○○○ Slowest |
| 2160p | 2160 | 4K quality | ○○○○ |
| 3840p | 3840 | Maximum | ○○○○ |
| Custom... | 240-4000 | Your choice | Varies |

#### Image Format

| Format | Quality | File Size | Transparency |
|--------|---------|-----------|--------------|
| **PNG** | ★★★ Best | Large | ✓ Yes |
| **JPG** | ★★ Good | Medium | ✗ No |
| **WEBP** | ★★ Good | Smallest | ✓ Yes |

**JPG Quality Options:**
- `60 (Low)` — Smallest files, visible compression
- `75 (Medium)` — Good balance
- `90 (High)` — High quality, moderate size
- `100 (Ultra)` — Maximum quality, larger files

### Step 3: Set Up Column Mappings

Column mappings tell Canvex which columns contain search terms and where to put the images.

1. Click **"+ Add Mapping"** button
2. Configure the mapping row:

| Field | Description | Example |
|-------|-------------|---------|
| **Input Column** | Column with search text | `actor_name` |
| **Output Column** | Where to insert images | `actor_image` |
| **New Column Name** | For new columns only | `photo` |

**Example Mapping:**
```
Input: "actor_name" (contains "Tom Hanks")
Output: "Create New Column..." → "headshot"
Result: A new "headshot" column with Tom Hanks' photo
```

#### Multiple Mappings
Add as many mappings as needed! Each mapping creates images for one column:

```
Mapping 1: actor1 → image1
Mapping 2: actor2 → image2  
Mapping 3: actor3 → image3
```

### Step 4: Start Processing

1. Click **"▶ Start Processing"** (green button)
2. Choose **where to save** the output file
3. Enter a filename (e.g., `output_with_images.xlsx`)
4. Click **Save**

**During Processing:**
- The **progress bar** shows overall completion
- The **Cancel** button lets you stop safely
- Processing continues in the background

### Step 5: Review Output

When complete, a dialog appears:
- Click **Yes** to open the output file immediately
- Click **No** to close the dialog

**Output Files Created:**
| File | Contents |
|------|----------|
| `your_output.xlsx` | Excel with images inserted |
| `your_output_log.txt` | Processing log (always created) |
| `your_output_ERROR_log.txt` | Error details (only if errors occurred) |

---

## Configuration Options

### Image Theme Details

The theme affects search query construction:
```
Search Query = [Cell Value] + [Theme]
Example: "Tom Hanks" + "headshot portrait closeup face"
```

**Custom Theme Tips:**
- Use descriptive words: `professional`, `corporate`, `natural`
- Add style modifiers: `high quality`, `HD`, `portrait`
- Specify background: `white background`, `studio`

### Search Browser Comparison

| Feature | Bing | Google | DuckDuckGo |
|---------|------|--------|------------|
| Speed | ●●●● | ●●○○ | ●●●○ |
| Reliability | ●●●● | ●●●○ | ●●●○ |
| Image Quality | ●●●● | ●●●● | ●●●○ |
| Rate Limiting | Low | Medium | Low |
| Fallback | — | Bing | — |

> 📝 **Note:** If Google returns no results, Canvex automatically tries Bing as a fallback.

---

## Column Mappings

### Understanding Mappings

```
Excel Input                     Canvex Processing                Excel Output
┌──────────────┐               ┌──────────────────┐           ┌───────────────────┐
│ name         │               │                  │           │ name    │ photo   │
├──────────────┤    Search     │  🔍 Bing Images  │  Insert   ├─────────┼─────────┤
│ Tom Hanks    │ ───────────▶ │  📥 Download     │ ────────▶│ Tom     │ [IMG]   │
│ Brad Pitt    │               │  📐 Resize       │           │ Brad    │ [IMG]   │
└──────────────┘               └──────────────────┘           └─────────┴─────────┘
```

### Creating New Columns

1. In **Output Column**, select **"Create New Column..."**
2. A text field appears
3. Enter the **new column name**
4. The new column is added to the right of existing data

### Multiple Mappings Example

**Input Excel:**
| lead_actor | supporting_actor | director |
|------------|------------------|----------|
| Tom Cruise | Val Kilmer | Tony Scott |
| Keanu Reeves | Laurence Fishburne | The Wachowskis |

**Mappings:**
1. `lead_actor` → `lead_photo` (new)
2. `supporting_actor` → `support_photo` (new)
3. `director` → `director_photo` (new)

**Output Excel:**
| lead_actor | supporting_actor | director | lead_photo | support_photo | director_photo |
|------------|------------------|----------|------------|---------------|----------------|
| Tom Cruise | Val Kilmer | Tony Scott | [IMG] | [IMG] | [IMG] |

---

## Settings Panel

Access via **Settings** button in the toolbar.

### Output Section
- **Resolution:** Quick access to resolution selector
- **Format:** PNG/JPG/WEBP format selector
- **JPG Quality:** Quality slider when JPG is selected

### Search Section
- **Search Engine:** Bing/Google/DuckDuckGo
- **Theme Suffix:** Default theme for new sessions

### Performance Section
- **Download Threads:** Number of parallel downloads (2-20)
- **Request Timeout:** Seconds before download fails (3-30)

### Image Filters Section

| Filter | Effect | Recommended For |
|--------|--------|-----------------|
| **Prioritize portrait images** | Prefers taller-than-wide images | Headshots, portraits |
| **Filter out B&W images** | Excludes grayscale images | Modern, colorful photos |
| **Filter out cartoon images** | Excludes illustrations/graphics | Real photographs only |

**Recommendation Matrix:**

| Use Case | Portrait | B&W Filter | Cartoon Filter |
|----------|----------|------------|----------------|
| Professional headshots | ✓ On | ✓ On | ✓ On |
| Product photos | ✗ Off | ✓ On | ✓ On |
| Artistic portraits | ✓ On | ✗ Off | ✓ On |
| Character illustrations | ✗ Off | ✗ Off | ✗ Off |

---

## Output Files

### Excel Output Structure

The output Excel file contains:

1. **All original data** from the input file
2. **New image columns** based on your mappings
3. **Images embedded** directly in cells

**Image Properties:**
- Scale: 20% of original size
- Position: Anchored to cell
- Row Height: 120 pixels (auto-set)
- Column Width: 22 characters (auto-set)

### Log Files

**Normal Log (`_log.txt`):**
```
[START] 2025-01-15 10:30:00
[LOG] Theme: headshot portrait closeup face
[LOG] Search Browser: Bing Images
[LOG] Resolution: 720px
[SEARCH] Tom Hanks
[URLS] (Bing Images) 24 found: [...]
[SEARCH] Brad Pitt
[URLS] (Bing Images) 24 found: [...]

Time taken: 0h 5m 23s
```

**Error Log (`_ERROR_log.txt`):**
Created only when errors occur. Contains:
- Full processing log
- Error stack trace
- Debug information

---

## 🆕 Recent Features & Enhancements (December 2025)

### Session Persistence

#### Auto-Save Last Directory
- Remembers the last folder where you opened an Excel file
- Next time you open a file, the browser starts in that location
- Stored in: `canva_last_settings.json`

#### Recent Files History
- **Automatic tracking:** Last 10 files you worked with
- **Quick access:** Click File → Recent Files
- **Visual indicators:** See if file exists (✓) or is missing (✗)
- **Click to open:** Opens file directly from the list

Example:
```
Recent Files
├─ ✓ employees_database.xlsx
├─ ✓ actors_filmography.xlsx  
├─ ✗ deleted_file.xlsx (file no longer exists)
└─ ✓ product_catalog.xlsx
```

#### Automatic Settings Backup
- Every setting is automatically saved when you use them
- Includes: Theme, resolution, search engine, filters, browser choice
- Loads automatically on startup
- No manual saving needed!

### Smart Mapping Management

#### Auto-Save Mappings
- Mappings are automatically saved when you start processing
- Last 5 configurations kept with timestamps
- Each entry shows date/time and number of mappings

#### Previous Mappings Dialog
Open via: **File** → **Load Previous Mappings**

**What you can do:**
1. **Browse history:** See all previous mapping configurations
2. **Live preview:** Select any mapping to see which columns were mapped
3. **Quick load:** Double-click or click "Load Selected" to restore
4. **Reset all:** Clear all mappings with confirmation dialog
5. **Select & load:** Choose specific configuration to reuse

**Dialog Features:**
- **Larger window:** 700×600 pixels with resizing (was too small before)
- **Better preview:** Preview table shows all mapped columns clearly
- **Auto-select:** First mapping is selected by default
- **No scrolling issues:** All mappings visible in the list

### Column Mapping Improvements

#### Smart Column Detection
- When you load old mappings, Canvex detects missing columns automatically
- Missing columns automatically switch to "Create New Column..." mode
- You can edit the new column name in the text field below

#### Enhanced Delete Functionality
- **Individual delete:** Click trash button (✕) next to each mapping
- **Works correctly:** Deletes work properly even with loaded history
- **Auto-renumber:** Remaining rows renumber automatically
- **Bulk delete:** Use "Reset All Mappings" to clear everything at once

#### Improved Text Field Visibility
- Text field for new column names only shows when needed
- Completely hidden when not using "Create New Column..."
- Automatically clears when switching away from that option
- No visual clutter!

### Visual & Interface Enhancements

#### Native Theme Integration
- Dropdowns now use system theme colors (not custom dark styling)
- Better visibility of selected values
- Authentic macOS appearance
- Works in both Light and Dark themes

#### List Hover Effects
- Moving mouse over mapping names shows clear feedback
- Subtle background color change indicates interactive element
- **Dark theme:** Gray hover background
- **Light theme:** Light gray hover background
- Smooth and responsive on macOS 14.2+

---

## Tips & Best Practices

### For Best Image Results

| Tip | Why It Helps |
|-----|--------------|
| ✓ Use specific search terms | "John Smith CEO Microsoft" finds better than "John Smith" |
| ✓ Choose appropriate themes | Match theme to content type |
| ✓ Enable all filters for headshots | Removes unwanted image types |
| ✓ Start with 720p | Good balance of quality and speed |
| ✓ Use PNG format | Best quality, no compression artifacts |

### For Faster Processing

| Tip | Impact |
|-----|--------|
| ✓ Use Bing Images | Fastest and most reliable |
| ✓ Lower resolution | Smaller downloads = faster |
| ✓ Stable internet | Avoids timeout retries |
| ✓ Close other browsers | More resources for Canvex |

### For Large Files (1000+ rows)

1. **Process in batches** — Split into smaller files
2. **Use lower resolution** — 480p is sufficient for previews
3. **Choose JPG format** — Smaller file sizes
4. **Monitor progress** — Cancel if stuck

### Column Naming Best Practices

| Good ✓ | Bad ✗ |
|--------|-------|
| `actor_photo` | `photo 1` |
| `employee_headshot` | `image` |
| `product_img` | `column_A` |

---

## Troubleshooting

### Common Issues

#### ❌ "No images found"

**Causes:**
- Search term too vague
- Name misspelled
- Person/item not well-known

**Solutions:**
1. Make search terms more specific
2. Try a different search engine
3. Simplify the theme
4. Check spelling in Excel

#### ❌ "Wrong images appearing"

**Causes:**
- Common name (e.g., "John Smith")
- Theme not matching content

**Solutions:**
1. Add context: "John Smith actor" or "John Smith CEO"
2. Try a different theme
3. Use custom theme with specific keywords

#### ❌ "Processing is very slow"

**Causes:**
- High resolution selected
- Slow internet connection
- Many rows to process

**Solutions:**
1. Lower resolution to 480p or 720p
2. Check internet speed
3. Process in smaller batches

#### ❌ "App appears frozen"

**Causes:**
- Chrome/Selenium starting up
- Large batch processing

**Solutions:**
1. Wait 30 seconds — Selenium needs time to start
2. Check task manager — If CPU is active, processing continues
3. If truly frozen, force quit and check `_ERROR_log.txt`

#### ❌ "Images not appearing in Excel"

**Causes:**
- Excel viewer doesn't show embedded images
- File corrupted during save

**Solutions:**
1. Open in Microsoft Excel (not preview apps)
2. Re-run with different format
3. Check the log file for errors

### Error Messages

| Error | Meaning | Solution |
|-------|---------|----------|
| "Chrome not found" | ChromeDriver issue | Install/update Chrome browser |
| "Connection timeout" | Network issue | Check internet, try again |
| "Permission denied" | File locked | Close the Excel file |
| "Out of memory" | Too many images | Process smaller batches |

---

## Keyboard Shortcuts & Navigation

### General Navigation

| Action | How |
|--------|-----|
| Tab between fields | `Tab` |
| Select dropdown item | `↑` `↓` arrows |
| Confirm selection | `Enter` |
| Cancel dialog | `Esc` |

### During Processing

| Action | How |
|--------|-----|
| Cancel processing | Click **Cancel** button |
| Force quit (emergency) | `Cmd+Q` (Mac) or `Alt+F4` (Windows) |

---

## FAQ

### General Questions

**Q: What file formats are supported?**
> A: Input must be `.xlsx` (Excel 2007+). Output is always `.xlsx`.

**Q: Can I process multiple Excel files at once?**
> A: No, process one file at a time. For batch processing, run Canvex multiple times.

**Q: Are my images saved locally?**
> A: Yes, images are embedded directly in the output Excel file. Temporary files are deleted after processing.

**Q: Does Canvex work offline?**
> A: No, internet connection is required for image searches.

### Technical Questions

**Q: Why does Canvex need Chrome?**
> A: Canvex uses Selenium with Chrome to scrape image search results. This provides more reliable results than API-only approaches.

**Q: Where are settings saved?**
> A: 
> - **macOS:** `~/Library/Application Support/Canvex/canva_last_settings.json`
> - **Windows:** `%APPDATA%/Canvex/canva_last_settings.json`
> - **Development:** Same folder as `Canvex.py`

**Q: Can I customize the image size in Excel?**
> A: Currently, images are inserted at 20% scale with fixed row height (120px). For different sizes, edit the output in Excel.

### Performance Questions

**Q: How long does processing take?**
> A: Depends on rows and resolution. Typical: ~2-5 seconds per row at 720p.

**Q: Can I process 10,000+ rows?**
> A: Technically yes, but recommended to split into batches of 500-1000 for reliability.

**Q: Does resolution affect processing time?**
> A: Yes. Higher resolution = larger downloads = slower processing.

---

## Contact & Support

- **Publisher:** Kunal Pagariya
- **Email:** [kunal.pagariya@outlook.com](mailto:kunal.pagariya@outlook.com)
- **Version:** 1.0
- **© 2025** Kunal Pagariya

---

<div align="center">
  <p><em>Thank you for using Canvex!</em></p>
  <p>If you find this tool helpful, please share it with others who might benefit.</p>
</div>
