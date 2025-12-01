# Z21 Locomotive Manager

## Purpose 
For Roco products, it is easy to add your model train to your trains/locomotives library in the Z21 App by simply loading the details and function configuration from the online database. However, for model trains/locomotives from other manufacturers, the process is less efficient. You must manually enter all details and function mappings one by one in the Z21 App.
 
This Python application allows you to read, parse, and manage `.z21` files used by Roco's Z21 App more conveniently on your computer. With this tool, you can add locomotive data, browse function mappings, and easily export your locomotives back to the Z21 App via AirDrop if you are using a macOS computer.

## Future Plans
Use your cell phone camera to take a photo of your instruction sheet with the functions table, then upload it to the application. The app will intelligently extract all the relevant information automatically.


## ✨ Features

- **Dual Format Support**: Read and display the details and functin mapping of locomotive in Z21 file.
- **GUI Browser**: Graphical interface for browsing locomotives and their functions, import z21loco file. Add or delete locomotive.


## 📋 Requirements

- Python 3.8 or higher
- customtkinter>=5.0.0
- Pillow>=9.0.0  
- pytesseract>=0.3.10
- pdf2image>=1.16.3
- pyobjc-framework-AppKit>=9.0.0 

## 🚀 Usage

1. **Clone the repository** (or navigate to the project directory)
2. **Install dependencies**:
```bash
pip install -r requirements.txt
```
3. Launch the graphical interface to browse locomotives:

```bash
# Run with default file (z21_new.z21)
python tools/z21lm_gui.py

# Run with specific file
python tools/z21lm_gui.py z21_new.z21
python tools/z21lm_gui.py rocoData.z21
```

**GUI Features**:
- Search locomotives by name or address
- View detailed locomotive information
- Browse function mappings with icons
- Two-tab interface: Overview and Functions


## 📁 Project Structure

```
z21_locomitive_manager/
├── README.md                    # This file
├── requirements.txt             # Python dependencies
├── icon_mapping.json            # Icon name mappings for function icons
├── src/                         # Core source code
│   ├── __init__.py
│   ├── binary_reader.py         # Binary file reading utilities
│   ├── cli.py                   # Command-line interface
│   ├── data_models.py           # Data structure definitions
│   └── parser.py                # File format parser (XML/SQLite)
├── tools/                       # Utility scripts and GUI
│   ├── __init__.py
│   ├── z21lm_gui.py             # Main GUI browser application (customtkinter)
│   └── z21lm_gui_operations.py  # GUI operations mixin (import, export, etc.)
├── icons/                       # Locomotive function icons (PNG files)
│   └── *.png                    # Function icon images
├── *.z21                        # Z21 database files (ZIP archives)
└── *.z21loco                    # Individual locomotive files
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
