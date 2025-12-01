# Z21 Locomotive Browser GUI

A graphical user interface for browsing and viewing locomotive details from Z21 files.

## Features

- **Locomotive List**: Browse all locomotives in the Z21 file
- **Search**: Filter locomotives by name or address
- **Detailed View**: Click any locomotive to see:
  - Basic information (name, address, speed, direction)
  - Function details with icons, button types, shortcuts
  - CV values (if available)
- **Two-Tab Interface**:
  - **Overview Tab**: Complete text summary
  - **Functions Tab**: Visual function cards with all details

## Usage

### Basic Usage

```bash
# Run with default file (z21_new.z21)
python tools/z21lm_gui.py

# Run with specific file
python tools/z21lm_gui.py z21_new.z21
python tools/z21lm_gui.py rocoData.z21
```

### GUI Layout

**Left Panel:**
- Search box at the top
- Scrollable list of locomotives
- Status bar showing total count

**Right Panel:**
- **Overview Tab**: Complete locomotive information in text format
- **Functions Tab**: Individual function cards showing:
  - Function number
  - Icon/image name
  - Button type (switch/push-button/time button)
  - Position
  - Shortcut key
  - Time delay

### Features

1. **Search**: Type in the search box to filter locomotives
2. **Select**: Click on any locomotive in the list to view details
3. **Navigate**: Use tabs to switch between Overview and Functions views

## Requirements

- Python 3.8+
- tkinter (usually included with Python)
- Z21 parser module (included in project)

## Example

```bash
# Start the GUI
python tools/z21lm_gui.py z21_new.z21

# The window will show:
# - Left: List of 65 locomotives
# - Right: Details panel (empty until you select a locomotive)
# - Search box: Filter by name or address
```

## Tips

- Use the search box to quickly find locomotives by name or address
- Functions are sorted by their position in the UI
- Button types are color-coded in the function cards
- The Overview tab provides a complete text summary
- The Functions tab shows individual cards for each function

