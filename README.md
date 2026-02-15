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

## Requirements

- Windows 10 or Windows 11
- Python 3.8 or higher (if running from source)
- A connected printer

## Installation

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

### Or run from .bat file
1. Run this batch file:
```bash
run_barcode_printer.bat
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

See [requirements.txt](requirements.txt) for specific versions.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [sv-ttk](https://github.com/rdbende/Sun-Valley-ttk-theme) - Beautiful modern theme
- [python-barcode](https://github.com/WhyNotHugo/python-barcode) - Barcode generation
- [Pillow](https://python-pillow.org/) - Image processing
- Icon from [Flaticon](https://www.flaticon.com/)
