# barcode_printer.py
# Universal Barcode Printer - Print barcodes to any Windows printer with a simple GUI

import logging
import json
import atexit
from tkinter import messagebox
from tkinter import ttk
from PIL import Image, ImageTk, ImageWin
import barcode as python_barcode
from barcode.writer import ImageWriter
import io
import win32print
import win32ui
import win32con
import tkinter as tk
import sv_ttk
from collections import OrderedDict
import threading
import ctypes
import os
from pathlib import Path


# --- Printer List Cache ---
# _PRINTER_LIST_CACHE stores the list of available printers to avoid repeated lookups
_PRINTER_LIST_CACHE = None


def get_printers(force_refresh=False) -> list[str]:
    """Return a list of available printer names. Refresh if force_refresh is True or cache is empty."""
    global _PRINTER_LIST_CACHE
    if force_refresh or _PRINTER_LIST_CACHE is None:
        _PRINTER_LIST_CACHE = [
            printer[2]
            for printer in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
        ]
    return _PRINTER_LIST_CACHE  # type: ignore


# --- Barcode Image Cache ---
# BARCODE_IMAGE_CACHE stores generated barcode images to avoid regenerating them
# Limit the barcode image cache to 100 items (LRU cache)
BARCODE_IMAGE_CACHE = OrderedDict()
BARCODE_IMAGE_CACHE_MAXSIZE = 100


def generate_label_image(barcode_text: str) -> Image.Image:
    """Generate a label image with a barcode for the given text, using cache for performance."""
    if barcode_text in BARCODE_IMAGE_CACHE:
        # Move to end to mark as recently used
        BARCODE_IMAGE_CACHE.move_to_end(barcode_text)
        return BARCODE_IMAGE_CACHE[
            barcode_text
        ].copy()  # Generate barcode using python-barcode
    code128 = python_barcode.get("code128", barcode_text, writer=ImageWriter())
    with io.BytesIO() as buffer:
        code128.write(buffer)
        buffer.seek(0)
        barcode_img = Image.open(buffer)
        # Create a white label and center the barcode on it
        label_width, label_height = 600, 300
        label_img = Image.new("RGB", (label_width, label_height), 0xFFFFFF)
        barcode_x = (label_width - barcode_img.width) // 2
        barcode_y = (label_height - barcode_img.height) // 2
        label_img.paste(barcode_img, (barcode_x, barcode_y))
        # Add to cache and enforce max size
        BARCODE_IMAGE_CACHE[barcode_text] = label_img.copy()
        BARCODE_IMAGE_CACHE.move_to_end(barcode_text)
        if len(BARCODE_IMAGE_CACHE) > BARCODE_IMAGE_CACHE_MAXSIZE:
            BARCODE_IMAGE_CACHE.popitem(last=False)
        return label_img


def print_image(img: Image.Image, printer_name: str) -> None:
    """Send the given image to the specified printer using Windows APIs."""
    pdc = win32ui.CreateDC()
    if pdc is None:
        logging.error("Failed to create printer device context.")
        raise RuntimeError("Failed to create printer device context.")
    pdc.CreatePrinterDC(printer_name)
    try:
        pdc.StartDoc("Barcode Print")
        pdc.StartPage()
        # Get printable area size
        printable_width = pdc.GetDeviceCaps(win32con.HORZRES)
        printable_height = pdc.GetDeviceCaps(win32con.VERTRES)
        # Resize image to fit printable width (maintain aspect ratio)
        if img.width != printable_width:
            scale = printable_width / img.width
            scaled_width = printable_width
            scaled_height = int(img.height * scale)
            img = img.resize((scaled_width, scaled_height))
        # Center the image on the page
        x1 = 0
        y1 = (
            (printable_height - img.height) // 2 if printable_height > img.height else 0
        )
        x2 = x1 + img.width
        y2 = y1 + img.height
        dib = ImageWin.Dib(img)
        dib.draw(pdc.GetHandleOutput(), (x1, y1, x2, y2))
        pdc.EndPage()
        pdc.EndDoc()
    except Exception as exc:
        logging.error("Print error: %s", exc)
        raise
    finally:
        pdc.DeleteDC()


# Module-level variable to track last previewed barcode value (for pylint compatibility)
_last_preview_value = None


def update_preview(event=None):
    """Update the barcode preview image in the GUI when the barcode entry changes."""
    global _last_preview_value
    _ = event
    barcode_value = entry.get().strip()
    # Only update if value changed
    if _last_preview_value == barcode_value:
        return
    _last_preview_value = barcode_value
    if not barcode_value:
        preview_label.config(image="")
        setattr(preview_label, "image", None)
        return
    try:
        img = generate_label_image(barcode_value)
        # Resize preview for display (use fast filter)
        preview_img = (
            img.resize((400, 200), Image.BOX) if img.size != (400, 200) else img
        )
        tk_img = ImageTk.PhotoImage(preview_img)
        preview_label.config(image=tk_img)
        setattr(preview_label, "image", tk_img)
        # --- Ensure window is large enough for preview and all widgets ---
        root.update_idletasks()
        # Always enforce the minsize, do not shrink below it
        min_width, min_height = 600, 1000  # Enforce your intended minimum
        preview_width = preview_img.width
        preview_height = preview_img.height
        extra_height = 400  # Adjust as needed for your layout
        min_width = max(min_width, preview_width + 100)
        min_height = max(min_height, preview_height + extra_height)
        root.minsize(min_width, min_height)
        # Save the new window size to config for next launch
        config["window_size"] = root.geometry()
        debounced_config_saver.save()
    except (OSError, RuntimeError, ValueError) as exc:
        preview_label.config(image="")
        setattr(preview_label, "image", None)
        print("Preview update error:", exc)


def parse_listbox_entry(item_text: str) -> tuple[str, int]:
    """Parse a listbox entry and return (barcode_text, copies)."""
    if item_text.startswith("Printed: "):
        text = item_text[len("Printed: ") :]
        if " x" in text:
            barcode_text, copies_str = text.rsplit(" x", 1)
            try:
                copies = int(copies_str)
            except ValueError:
                copies = 1
        else:
            barcode_text = text
            copies = 1
    else:
        barcode_text = item_text
        copies = 1
    return barcode_text, copies


def add_tooltip(widget, text):
    """Add a tooltip to a widget that shows in the status bar."""

    def on_enter(_event):
        status_var.set(text)

    def on_leave(_event):
        status_var.set("Ready")

    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)


def threaded_print(
    img: Image.Image, printer_name: str, copies: int, barcode_value: str
) -> None:
    """Threaded print operation to keep UI responsive."""
    global barcode_history
    try:
        set_progress(f"Printing {copies} copies...")
        for i in range(copies):
            set_progress(f"Printing copy {i+1} of {copies}...")
            print_image(img, printer_name)
        # Add to history treeview and persist
        listbox.insert("", "end", values=(barcode_value, copies))
        barcode_history = barcode_history + [
            {"barcode": barcode_value, "copies": copies}
        ]
        save_history(barcode_history)
        entry.delete(0, tk.END)
        update_preview()
        set_progress("Done.")
    except (OSError, RuntimeError) as exc:
        logging.error("Print Error: %s", exc)
        messagebox.showerror("Print Error", str(exc))
        set_progress("")


def handle_print() -> None:
    """Handle the print button click event."""
    barcode_value = entry.get().strip()
    selected_printer = printer_var.get()
    copies = copies_var.get()
    if not barcode_value:
        messagebox.showwarning("Missing Barcode", "Please enter a barcode.")
        return
    if not selected_printer:
        messagebox.showwarning("No Printer Selected", "Please select a printer.")
        return
    try:
        copies_int = int(copies)
        if copies_int < 1:
            raise ValueError("Copies must be at least 1.")
    except ValueError:
        messagebox.showwarning(
            "Invalid Copies", "Please enter a valid number of copies (1 or more)."
        )
        return
    try:
        img = generate_label_image(barcode_value)
        import threading

        # Start print in a background thread
        threading.Thread(
            target=threaded_print,
            args=(img, selected_printer, copies_int, barcode_value),
            daemon=True,
        ).start()
    except (OSError, RuntimeError) as exc:
        logging.error("Print Error: %s", exc)
        messagebox.showerror("Print Error", str(exc))


def threaded_reprint(selected_items, selected_printer):
    """Threaded reprint operation to keep UI responsive."""
    try:
        set_progress(f"Reprinting {len(selected_items)} barcode(s)...")
        for idx, item_id in enumerate(selected_items):
            values = listbox.item(item_id, "values")
            if not values:
                continue
            barcode_text, copies = values[0], int(values[1])
            img = generate_label_image(barcode_text)
            for c in range(copies):
                set_progress(
                    f"Reprinting {idx+1}/{len(selected_items)}: copy {c+1} of {copies}"
                )
                print_image(img, selected_printer)
        set_progress("Done.")
    except (OSError, RuntimeError) as exc:
        logging.error("Reprint Error: %s", exc)
        messagebox.showerror("Print Error", str(exc))
        set_progress("")


def reprint_selected() -> None:
    """Handle the reprint button click event."""
    selected_items = listbox.selection()
    selected_printer = printer_var.get()
    if not selected_printer:
        messagebox.showwarning("No Printer Selected", "Please select a printer.")
        return
    if not selected_items:
        messagebox.showwarning(
            "No Selection", "Please select at least one barcode from the list."
        )
        return
    try:
        import threading

        # Start reprint in a background thread
        threading.Thread(
            target=threaded_reprint,
            args=(selected_items, selected_printer),
            daemon=True,
        ).start()
    except (OSError, RuntimeError) as exc:
        logging.error("Reprint Error: %s", exc)
        messagebox.showerror("Print Error", str(exc))


# --- Configuration ---
# Store config in AppData/UniversalBarcodePrinter/barcode_printer_config.json
APPDATA_DIR = Path(os.getenv("APPDATA", os.path.expanduser("~")))
CONFIG_DIR = APPDATA_DIR / "UniversalBarcodePrinter"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = str(CONFIG_DIR / "barcode_printer_config.json")
DEFAULT_CONFIG = {
    "default_printer": "",
    "window_size": "550x750",
    "language": "en",
}

# --- Internationalization ---
LANGUAGES = {
    "en": {
        "select_printer": "Select Printer:",
        "scan_barcode": "Scan or Enter Barcode:",
        "num_copies": "Number of Copies:",
        "preview": "Preview:",
        "print": "Print",
        "reprint_selected": "Reprint Selected",
        "missing_barcode": "Please enter a barcode.",
        "no_printer": "Please select a printer.",
        "invalid_copies": "Please enter a valid number of copies (1 or more).",
        "no_selection": "Please select at least one barcode from the list.",
        "print_error": "Print Error",
        "reprint_info": "Reprinted {count} barcode(s).",
        "about": "Universal Barcode Printer\nVersion 1.0",
        "help": "Select a printer, scan a barcode, and click Print. Use Reprint to print again.",
    }
    # Add more languages here
}


# --- Utility Functions ---
def load_config():
    """Load configuration from file or return defaults if not found/corrupt."""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        logging.error("Failed to load config: %s", exc)
        return DEFAULT_CONFIG.copy()


def save_config(cfg):
    """Save configuration to file."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except OSError as exc:
        logging.error("Failed to save config: %s", exc)


config = load_config()

# --- Barcode History Persistence ---
HISTORY_FILE = str(CONFIG_DIR / "barcode_history.json")


def load_history():
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return []


def save_history(history):
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f)
    except OSError as exc:
        logging.error("Failed to save history: %s", exc)


barcode_history = load_history()


# --- Internationalization Helper ---
def _(key: str) -> str:
    """Internationalization helper to get translated strings."""
    lang = config.get("language", "en")
    return str(LANGUAGES.get(lang, LANGUAGES["en"]).get(key, key))


# --- Graceful Shutdown ---
def on_exit():
    """Save config and log exit on application close."""
    debounced_config_saver.flush()
    logging.info("Application exited gracefully.")


atexit.register(on_exit)


# --- Debounced Config Save ---
class DebouncedConfigSaver:
    def __init__(self, delay=0.5):
        self.delay = delay
        self._timer = None

    def save(self):
        if self._timer:
            self._timer.cancel()
        self._timer = threading.Timer(self.delay, lambda: save_config(config))
        self._timer.start()

    def flush(self):
        if self._timer:
            self._timer.cancel()
            save_config(config)
            self._timer = None


debounced_config_saver = DebouncedConfigSaver()


# --- HiDPI Scaling for Tkinter (Windows) ---
def set_hidpi_scaling(root):
    try:
        # Only apply on Windows
        if hasattr(ctypes, "windll"):
            user32 = ctypes.windll.user32
            user32.SetProcessDPIAware()
            dpi = (
                user32.GetDpiForSystem()
                if hasattr(user32, "GetDpiForSystem")
                else user32.GetDeviceCaps(user32.GetDC(0), 88)
            )
            scaling = dpi / 72.0
            root.tk.call("tk", "scaling", scaling)
    except Exception as exc:
        print(f"Could not set HiDPI scaling: {exc}")


# --- GUI Setup ---
root = tk.Tk()
set_hidpi_scaling(root)
root.title("Universal Barcode Printer")
root.geometry(config.get("window_size", "550x750"))
root.minsize(600, 1000)


# --- Theme Persistence ---
def get_theme_from_config():
    return config.get("theme", "dark")


def set_theme_in_config(theme):
    config["theme"] = theme
    debounced_config_saver.save()


sv_ttk.set_theme(get_theme_from_config())


# --- Theme Toggle Button ---
def toggle_theme():
    current_theme = sv_ttk.get_theme()
    new_theme = "light" if current_theme == "dark" else "dark"
    sv_ttk.set_theme(new_theme)
    set_theme_in_config(new_theme)
    theme_button.config(
        text=f"Switch to {'Dark' if new_theme == 'light' else 'Light'} Theme"
    )


# Add theme toggle button to the top right
theme_button = ttk.Button(
    root,
    text=f"Switch to {'Light' if get_theme_from_config() == 'dark' else 'Dark'} Theme",
    command=toggle_theme,
)
theme_button.pack(anchor="ne", padx=10, pady=5)


def set_window_size(_event=None):
    """Update config with current window size on resize and debounce save."""
    config["window_size"] = root.geometry()
    debounced_config_saver.save()


root.bind("<Configure>", set_window_size)

# --- Set window icon using a PNG file ---
try:
    icon_img = tk.PhotoImage(
        file="./barcode-scan.png"
    )  # Place your icon.png in the same directory
    root.iconphoto(True, icon_img)
except Exception as exc:
    print(f"Could not set window icon: {exc}")


# --- Widgets ---
ttk.Label(root, text=_("select_printer")).pack(pady=(10, 0))
printer_var = tk.StringVar(value=config.get("default_printer", ""))
printer_dropdown = ttk.Combobox(
    root, textvariable=printer_var, values=get_printers(force_refresh=True), width=50
)
printer_dropdown.pack(pady=(0, 10))


def on_printer_selected(_event=None):
    config["default_printer"] = printer_var.get()
    debounced_config_saver.save()


printer_dropdown.bind("<<ComboboxSelected>>", on_printer_selected)

ttk.Label(root, text=_("scan_barcode")).pack()
entry = ttk.Entry(root, font=("Segoe UI Variable", 12), width=35)
entry.pack(pady=5)
entry.focus()
entry.bind("<KeyRelease>", update_preview)

ttk.Label(root, text=_("num_copies")).pack()
copies_var = tk.StringVar(value="1")
copies_spinbox = ttk.Spinbox(
    root,
    from_=1,
    to=100,
    textvariable=copies_var,
    width=5,
    font=("Segoe UI Variable", 12),
)
copies_spinbox.pack(pady=(0, 10))

ttk.Label(root, text=_("preview")).pack(pady=(10, 0))
preview_label = ttk.Label(root)
preview_label.pack(pady=(0, 10))

# --- Progress Indicator ---
progress_var = tk.StringVar(value="")
progress_label = ttk.Label(
    root, textvariable=progress_var, font=("Segoe UI Variable", 12)
)
progress_label.pack(pady=(0, 5))


def set_progress(msg):
    """Update the progress label in the GUI."""
    progress_var.set(msg)
    root.update_idletasks()


def on_print():
    """Save selected printer and handle print button click."""
    config["default_printer"] = printer_var.get()
    handle_print()


print_button = ttk.Button(
    root, text=_("print"), command=on_print, width=20, style="Accent.TButton"
)
print_button.pack(pady=(5, 10))

# Use ttk.Treeview as a themed replacement for Listbox
tree_columns = ("Barcode", "Copies")
listbox = ttk.Treeview(root, columns=tree_columns, show="headings", height=10)
listbox.heading("Barcode", text="Barcode")
listbox.heading("Copies", text="Copies")
listbox.column("Barcode", width=350)
listbox.column("Copies", width=80, anchor="center")
listbox.pack(pady=10)

# Populate history on startup
for item in barcode_history:
    if isinstance(item, dict):
        barcode, copies = item.get("barcode"), item.get("copies", 1)
    else:
        # Handle legacy format if needed
        barcode, copies = item, 1
    if barcode:
        listbox.insert("", "end", values=(barcode, copies))

reprint_button = ttk.Button(
    root, text=_("reprint_selected"), command=reprint_selected, width=20
)
reprint_button.pack(pady=(5, 10))

entry.bind("<Return>", lambda event: on_print())

# --- Status Bar ---
status_var = tk.StringVar(value="Ready")
status_label = ttk.Label(root, textvariable=status_var, font=("Segoe UI Variable", 12))
status_label.pack(pady=(0, 5))


def set_status(msg):
    """Update the status bar in the GUI."""
    status_var.set(msg)
    root.update_idletasks()


# --- Accessibility: Add tooltips ---
add_tooltip(print_button, "Print the current barcode")
add_tooltip(reprint_button, "Reprint selected barcodes")
add_tooltip(printer_dropdown, "Select a printer")
add_tooltip(entry, "Enter or scan a barcode")
add_tooltip(copies_spinbox, "Set number of copies")

# --- Keyboard Navigation ---
print_button.focus_set()
root.bind("<Alt-p>", lambda e: print_button.invoke())
root.bind("<Alt-r>", lambda e: reprint_button.invoke())


# --- Focus entry when window regains focus ---
def focus_entry_on_window_focus(event=None):
    entry.focus_set()


root.bind("<FocusIn>", focus_entry_on_window_focus)

# Pillow plugin fix for PyInstaller
_img_ref = ImageTk.PhotoImage(Image.new("RGB", (1, 1)))

# --- Start the GUI event loop ---
root.mainloop()
