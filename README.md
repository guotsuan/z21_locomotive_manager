# Z21 Locomotive Manager

## Purpose 
For Roco products, it is easy to add your model train to your train library in the Z21 App by simply loading the details and function configuration from the online database. However, for model trains from other manufacturers, the process is less efficient. You must manually enter all details and function mappings one by one in the Z21 App.
 
This Python application allows you to read, parse, and manage `.z21` files used by Roco's Z21 App more conveniently on your computer. With this tool, you can add locomotive data, browse function mappings, and easily export your locomotives back to the Z21 App via AirDrop if you are using a macOS computer.


## ✨ Features

- **Dual Format Support**: Read and display the details and functin mapping of locomotive in Z21 file.
- **GUI Browser**: Graphical interface for browsing locomotives and their functions, import z21loco file. Add or delete locomotive.


## 📋 Requirements

- Python 3.8 or higher
- 

## 🚀 Usage

1. **Clone the repository** (or navigate to the project directory)
2. **Install dependencies**:
```bash
pip install -r requirements.txt
```
3. Launch the graphical interface to browse locomotives:

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

### Format: SQLite (New Format)
- File: `Loco.sqlite` inside ZIP archive
- Example: `z21_new.z21`
- Successfully parsed: 65+ locomotives


## 🤝 Contributing

Contributions are welcome! Areas for improvement:

## 📄 License

This project is licensed under the BSD 3-Clause License.



**Note**: This project is not affiliated with Roco or Z21. It is an independent tool for managing Z21 locomotive data files.
