# Canvex Changelog — December 2025 Updates

## Version 1.0.1 — Session Management & Mapping Enhancements

### ✨ New Features

#### 1. Session Persistence
- **Last Directory Memory:** App remembers where you last opened an Excel file
- **Recent Files (10):** Quick access to recently opened files with validation
- **Auto-Save Settings:** All preferences automatically saved and restored

#### 2. Smart Mapping Management
- **Mapping History:** Last 5 configurations saved with timestamps
- **Previous Mappings Dialog:** Browse and restore past mapping configurations
- **Live Preview:** See which columns are mapped before loading
- **Reset All:** Clear all mappings with confirmation dialog

#### 3. Column Mapping Improvements
- **Smart Detection:** Auto-detects missing columns and switches to "Create New Column..."
- **Dynamic Delete:** Delete buttons work correctly even with loaded history
- **Text Field Visibility:** New column name field appears only when needed
- **Auto-Clear:** Text field clears automatically when not in use

#### 4. Visual Enhancements
- **Native Theme Integration:** Dropdowns now follow system theme (light/dark)
- **List Hover Effects:** Mapping list items highlight on hover
- **Better Visibility:** Selected dropdown values now clearly visible
- **macOS 14.2+ Support:** Full native styling and animations

---

### 🐛 Bug Fixes

| Issue | Status |
|-------|--------|
| Dropdown selections not visible | ✓ Fixed |
| Previous mappings dialog too small | ✓ Fixed |
| Preview not showing in mappings | ✓ Fixed |
| Can't select single mapping | ✓ Fixed |
| Delete button doesn't work with history | ✓ Fixed |
| Text field always visible | ✓ Fixed |
| Hover effects not working | ✓ Fixed |
| Dark dropdown styling breaking theme | ✓ Fixed |

---

### 📝 Documentation Updates

#### USER_GUIDE.md
- Added comprehensive "Recent Features & Enhancements" section
- Documented all new session persistence features
- Explained mapping management system
- Included visual examples and use cases
- Added tips for using new features

#### TECHNICAL_DOCS.md
- Added Section 16: "Recent Enhancements (December 2025)"
- Documented session persistence implementation
- Detailed mapping management architecture
- Explained visual styling improvements
- Included code examples and data structures
- Added Appendix D with changes summary
- Added Appendix E with future improvements

---

### 🔧 Technical Details

#### Files Modified
- `Canvex.py` - Main application (all enhancements)
- `USER_GUIDE.md` - User documentation
- `TECHNICAL_DOCS.md` - Technical documentation

#### Code Additions
- `load_basic_settings()` - Load basic UI settings on startup
- `load_settings()` - Load mappings after Excel is opened
- `save_settings(mappings)` - Save current mappings with history
- `show_previous_mappings()` - Dialog for viewing/loading previous configurations
- `reset_all_mappings()` - Clear all mappings with confirmation
- `delete_row_by_button(button)` - Dynamic row deletion
- `toggle_new_col(row)` - Smart text field visibility
- `_add_to_recent_files(filepath)` - Track recent files
- QComboBox & QListWidget styling - Native theme support

#### Settings File Structure
```json
{
  "theme": "Dark",
  "resolution": "720",
  "browser": "bing",
  "format": "png",
  "jpg_quality": "90",
  "last_excel_dir": "/path/to/directory",
  "recent_files": ["file1.xlsx", "file2.xlsx"],
  "mapping_history": [
    {
      "timestamp": "2025-12-30T14:32:15.123456",
      "mappings": [["col1", "col2"]]
    }
  ]
}
```

---

### 📊 Statistics

- **Lines Added:** ~400
- **Methods Added:** 3 new, several enhanced
- **Classes Modified:** 1 (CanvaImageExcelCreator)
- **Bugs Fixed:** 8
- **Documentation Added:** 200+ lines

---

### ✅ Testing Completed

- ✓ macOS 14.2 compatibility verified
- ✓ Light and dark theme tested
- ✓ File operations (open, recent, missing files)
- ✓ Mapping creation, deletion, and reset
- ✓ Previous mappings loading and preview
- ✓ Dropdown selection visibility
- ✓ List hover effects
- ✓ Settings persistence across sessions

---

### 🚀 Deployment

**Build Command:**
```bash
pyinstaller Canvex.spec --clean -y
```

**Test Command:**
```bash
open /Users/kunal/Canvex/FinalBuildMac/dist/Canvex.app
```

---

### 📋 Known Limitations

- Mapping history limited to last 5 configurations
- Recent files limited to last 10 files
- Hover effects on QComboBox dropdown items work best on macOS native rendering

---

### 🔮 Future Enhancements

1. Increase mapping history to 10 configurations
2. Export/import mapping configurations as JSON files
3. Batch processing for multiple Excel files
4. Undo/redo functionality for mappings
5. Drag-and-drop mapping reordering
6. Cloud sync for settings

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.1 | Dec 30, 2025 | Session persistence, mapping management, UI fixes |
| 1.0 | 2025 | Initial release |

---

**Author:** Kunal Pagariya  
**Last Updated:** December 30, 2025
