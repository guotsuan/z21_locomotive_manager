# Z21 Locomotive Manager

A Python application to read, parse, and manage `.z21` files used by Roco's Z21 model train control system. This tool allows you to analyze locomotive data, browse function mappings, and export data to JSON format.

## 🔍 Discovery

The `.z21` file format is actually a **ZIP archive** containing:
- **Format 1 (Old)**: `loco_data.xml` - XML file with locomotive and accessory data
- **Format 2 (New)**: `Loco.sqlite` - SQLite database with locomotive data
- Image files (PNG/JPG) for locomotives, wagons, and backgrounds

This project successfully parses both formats and provides tools for inspection and management.

## ✨ Features

- **Dual Format Support**: Reads both XML (old) and SQLite (new) Z21 file formats
- **CLI Tools**: Command-line interface for reading and exporting Z21 files
- **GUI Browser**: Graphical interface for browsing locomotives and their functions
- **Icon Management**: Tools for extracting and matching locomotive function icons
- **JSON Export**: Export locomotive data to JSON for inspection and integration
- **Hex Dump Utility**: Analyze binary file structure
- **SQLite Examination**: Inspect SQLite database contents

## 📋 Requirements

- Python 3.8 or higher
- Virtual environment (recommended)

## 🚀 Installation

1. **Clone the repository** (or navigate to the project directory)

2. **Create a virtual environment**:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
pip install -r requirements.txt
```

## 📖 Usage

### Command-Line Interface

#### Read and Display File Contents

```bash
# Read XML format file
python -m src.cli read rocoData.z21

# Read SQLite format file
python -m src.cli read z21_new.z21
```

#### Export to JSON

```bash
# Export XML format to JSON
python -m src.cli export rocoData.z21 output.json

# Export SQLite format to JSON
python -m src.cli export z21_new.z21 output.json
```

### GUI Browser

Launch the graphical interface to browse locomotives:

```bash
# Run with default file (z21_new.z21)
python tools/z21_gui.py

# Run with specific file
python tools/z21_gui.py z21_new.z21
python tools/z21_gui.py rocoData.z21
```

**GUI Features**:
- Search locomotives by name or address
- View detailed locomotive information
- Browse function mappings with icons
- Two-tab interface: Overview and Functions

### Utility Tools

#### Hex Dump Analysis

```bash
# View first 512 bytes
python tools/hex_dump.py rocoData.z21 -l 512

# View entire file
python tools/hex_dump.py rocoData.z21

# View specific offset
python tools/hex_dump.py rocoData.z21 -o 100 -l 256
```

#### SQLite Database Examination

```bash
python tools/examine_sqlite.py z21_new.z21
```

#### List Locomotives

```bash
python tools/list_locomotives.py z21_new.z21
```

#### Icon Management

```bash
# Extract icons from Z21 file
python tools/extract_icons.py z21_new.z21

# List available icons
python tools/list_icons.py

# Match icons to function names
python tools/match_icons.py
```

## 📁 Project Structure

```
z21_locomitive_manager/
├── README.md              # This file
├── PLAN.md                # Detailed development plan
├── QUICKSTART.md          # Quick start guide
├── requirements.txt       # Python dependencies
├── pytest.ini            # Pytest configuration
├── icon_mapping.json      # Icon name mappings
├── src/                   # Core source code
│   ├── __init__.py
│   ├── binary_reader.py   # Binary file reading utilities
│   ├── cli.py             # Command-line interface
│   ├── data_models.py     # Data structure definitions
│   └── parser.py          # File format parser (XML/SQLite)
├── tools/                 # Utility scripts
│   ├── z21_gui.py         # GUI browser application
│   ├── hex_dump.py        # Hex dump utility
│   ├── examine_sqlite.py  # SQLite database examination
│   ├── list_locomotives.py # List locomotives tool
│   ├── extract_icons.py   # Icon extraction tool
│   ├── list_icons.py      # List icons tool
│   ├── match_icons.py     # Icon matching tool
│   └── GUI_README.md      # GUI documentation
├── icons/                 # Locomotive function icons
├── extracted_icons/       # Extracted icon data
├── tests/                 # Unit tests
│   └── test_reader.py
└── *.z21                  # Sample Z21 files
```

## 🧪 Testing

Run the test suite:

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src

# Run specific test file
python tests/test_reader.py
```

## 📊 Development Status

- [x] **Phase 1: File Analysis** ✅ - Discovered ZIP + XML/SQLite formats
- [x] **Phase 2: Basic Reader** ✅ - Implemented ZIP, XML, and SQLite parsing
- [x] **Phase 3: Data Model** ✅ - Locomotive model working with both formats
- [x] **Phase 4: GUI Browser** ✅ - Graphical interface for browsing locomotives
- [x] **Phase 5: Icon Management** ✅ - Icon extraction and matching tools
- [ ] **Phase 6: Basic Writer** - Write ZIP archives with modified XML/SQLite
- [ ] **Phase 7: Editor Features** - CLI/GUI for editing locomotive data
- [ ] **Phase 8: Advanced Features** - Accessories, layouts, CV editing

## 📝 Supported File Formats

### Format 1: XML (Old Format)
- File: `loco_data.xml` inside ZIP archive
- Example: `rocoData.z21`
- Successfully parsed: 23+ locomotives

### Format 2: SQLite (New Format)
- File: `Loco.sqlite` inside ZIP archive
- Example: `z21_new.z21`
- Successfully parsed: 65+ locomotives

## ⚠️ Important Notes

- **Backup First**: Always backup your `.z21` files before modification
- **Read-Only**: Currently, the tool is read-only (writing support is planned)
- **Compatibility**: Test compatibility with Z21 app after any modifications
- **File Format**: The `.z21` file format is not publicly documented; this project is based on reverse engineering

## 🔧 Data Models

The project uses structured data models:

- **Z21File**: Root container for all Z21 data
- **Locomotive**: Represents a locomotive with address, name, functions, CVs
- **FunctionInfo**: Detailed function information (icon, button type, shortcuts)
- **Accessory**: Turnout/signal/light data
- **Layout**: Track layout configuration
- **Settings**: System settings

## 📚 Documentation

- **QUICKSTART.md**: Quick start guide for analyzing files
- **PLAN.md**: Detailed development plan and architecture
- **tools/GUI_README.md**: GUI browser documentation

## 🤝 Contributing

Contributions are welcome! Areas for improvement:

- Writer functionality (modify and save Z21 files)
- Enhanced GUI features (editing capabilities)
- Additional format support
- Documentation improvements
- Test coverage expansion

## 📄 License

[To be determined]

## 🙏 Acknowledgments

This project was created through reverse engineering of the Z21 file format. Special thanks to the model railway community for their support and feedback.

---

**Note**: This project is not affiliated with Roco or Z21. It is an independent tool for managing Z21 locomotive data files.
