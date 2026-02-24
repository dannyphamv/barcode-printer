# Barcode Printer

A simple, modern Windows application for printing Code128 barcodes to any connected printer. Features a clean GUI with dark/light theme support, real-time barcode preview, and print history tracking.

![Barcode Printer Screenshot](screenshot-dark.avif)

## Features

- **Universal Printer Support** - Works with any Windows printer (thermal, laser, inkjet)
- **Real-time Preview** - See your barcode before printing
- **Print History** - Automatic tracking of all printed barcodes with reprint functionality
- **Dark/Light Theme** - Modern UI with theme switching
- **Batch Printing** - Print multiple copies at once
- **Barcode Caching** - Fast performance with intelligent image caching
- **HiDPI Support** - Crisp display on high-resolution screens
- **Persistent Settings** - Remembers your printer selection and window size

## Installation

### Download the Installer

The easiest way to get started is to download the pre-built installer directly from the [Releases](https://github.com/dannyphamv/barcode-printer/releases) page. Run `LegacyBarcodePrinter_Setup.exe` and follow the prompts — no Python required.

### Run from Source

1. Clone this repository:
```bash
git clone https://github.com/dannyphamv/barcode-printer.git
cd barcode-printer
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python barcode_printer.py
```

### Or create .exe
```bash
pyinstaller --onedir --windowed --noconsole --name="LegacyBarcodePrinter" --add-data="package_color.png:." --icon="favicon.ico" --collect-all=PIL --collect-all=barcode --collect-all=sv_ttk --collect-all=pywinstyles --hidden-import=win32print --hidden-import=win32ui --hidden-import=win32con --hidden-import=win32api barcode_printer.py
```

## Usage

1. **Select a Printer** - Choose your printer from the dropdown menu
2. **Enter Barcode** - Type or scan a barcode value
3. **Preview** - The barcode preview updates automatically
4. **Set Copies** - Choose how many copies to print (default: 1)
5. **Print** - Click the Print button or press Enter
6. **Reprint** - Select items from history and click "Reprint Selected"

## Technical Details

- **Barcode Format**: Code128 (supports alphanumeric characters)
- **Label Size**: 600x300 pixels (automatically scaled to printer)
- **GUI Framework**: Tkinter with sv-ttk theme
- **Printing**: Windows GDI via pywin32
- **Image Processing**: Pillow (PIL)
- **Caching**: LRU cache with 100-item limit for barcode images

## Dependencies

- `sv-ttk` - Modern theme for Tkinter
- `Pillow` - Image processing
- `python-barcode` - Barcode generation
- `pywin32` - Windows printing API
- `pywinstyles` - Windows 10/11 title bar theming
- `pyinstaller` - Packaging the application as EXE

See [requirements.txt](requirements.txt) for specific versions.

## License

This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE) file for details.