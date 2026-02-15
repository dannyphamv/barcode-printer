# barcode_printer.py
# Barcode Printer - Print barcodes to any Windows printer with a simple GUI

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
import pywinstyles
import sys

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# === CONSTANTS ===
LABEL_WIDTH = 600
LABEL_HEIGHT = 300
PREVIEW_WIDTH = 400
PREVIEW_HEIGHT = 200
BARCODE_IMAGE_CACHE_MAXSIZE = 100
HISTORY_MAX_SIZE = 100
CONFIG_SAVE_DELAY = 0.5
MIN_WINDOW_WIDTH = 650
MIN_WINDOW_HEIGHT = 1000

# === PRINTER MANAGEMENT ===

_PRINTER_LIST_CACHE = None
_printer_cache_lock = threading.Lock()


def get_printers(force_refresh=False) -> list:
    """Return a list of available printer names.

    Caching prevents repeated expensive system calls to EnumPrinters.
    Thread-safe to prevent race conditions during refresh.
    """
    global _PRINTER_LIST_CACHE

    with _printer_cache_lock:
        if force_refresh or _PRINTER_LIST_CACHE is None:
            try:
                _PRINTER_LIST_CACHE = [
                    printer[2]
                    for printer in win32print.EnumPrinters(
                        win32print.PRINTER_ENUM_LOCAL
                        | win32print.PRINTER_ENUM_CONNECTIONS
                    )
                ]
            except Exception as exc:
                logging.error(f"Failed to enumerate printers: {exc}")
                _PRINTER_LIST_CACHE = []
        return _PRINTER_LIST_CACHE.copy()


# === BARCODE GENERATION & CACHING ===

BARCODE_IMAGE_CACHE = OrderedDict()
_cache_lock = threading.Lock()


def generate_label_image(barcode_text: str) -> Image.Image:
    """Generate a 600x300px label with centered Code128 barcode.

    Thread-safe cache prevents race conditions during concurrent access.
    Returns a copy from cache to prevent modifications affecting cached version.
    """
    if not barcode_text:
        raise ValueError("Barcode text cannot be empty")

    with _cache_lock:
        if barcode_text in BARCODE_IMAGE_CACHE:
            BARCODE_IMAGE_CACHE.move_to_end(barcode_text)
            return BARCODE_IMAGE_CACHE[barcode_text].copy()

    try:
        code128 = python_barcode.get("code128", barcode_text, writer=ImageWriter())
        with io.BytesIO() as buffer:
            code128.write(buffer)
            buffer.seek(0)
            barcode_img = Image.open(buffer).copy()  # Copy to avoid buffer issues

            # Create label with centered barcode
            label_img = Image.new("RGB", (LABEL_WIDTH, LABEL_HEIGHT), 0xFFFFFF)
            barcode_x = (LABEL_WIDTH - barcode_img.width) // 2
            barcode_y = (LABEL_HEIGHT - barcode_img.height) // 2
            label_img.paste(barcode_img, (barcode_x, barcode_y))

            # Update cache
            with _cache_lock:
                BARCODE_IMAGE_CACHE[barcode_text] = label_img.copy()
                BARCODE_IMAGE_CACHE.move_to_end(barcode_text)
                if len(BARCODE_IMAGE_CACHE) > BARCODE_IMAGE_CACHE_MAXSIZE:
                    BARCODE_IMAGE_CACHE.popitem(last=False)

            return label_img
    except Exception as exc:
        logging.error(f"Failed to generate barcode for '{barcode_text}': {exc}")
        raise


def print_image(img: Image.Image, printer_name: str) -> None:
    """Send image to Windows printer using GDI.

    Scales image to fit printer's width while maintaining aspect ratio.
    Centers vertically to avoid cutting off content on smaller paper.
    """
    if not img or not printer_name:
        raise ValueError("Image and printer name are required")

    pdc = None
    try:
        pdc = win32ui.CreateDC()
        if pdc is None:
            raise RuntimeError("Failed to create printer device context")

        pdc.CreatePrinterDC(printer_name)
        pdc.StartDoc("Barcode Print")
        pdc.StartPage()

        printable_width = pdc.GetDeviceCaps(win32con.HORZRES)
        printable_height = pdc.GetDeviceCaps(win32con.VERTRES)

        # Scale to full width to maximize barcode readability
        if img.width != printable_width:
            scale = printable_width / img.width
            scaled_width = printable_width
            scaled_height = int(img.height * scale)
            img = img.resize((scaled_width, scaled_height), Image.LANCZOS)

        # Center vertically
        x1 = 0
        y1 = max(0, (printable_height - img.height) // 2)
        x2 = x1 + img.width
        y2 = y1 + img.height

        dib = ImageWin.Dib(img)
        dib.draw(pdc.GetHandleOutput(), (x1, y1, x2, y2))

        pdc.EndPage()
        pdc.EndDoc()
    except Exception as exc:
        logging.error(f"Print error: {exc}")
        raise
    finally:
        if pdc is not None:
            try:
                pdc.DeleteDC()
            except Exception as exc:
                logging.warning(f"Failed to delete DC: {exc}")


# === UI UPDATE FUNCTIONS ===

_last_preview_value = None
_preview_update_lock = threading.Lock()


def update_preview(event=None):
    """Update barcode preview when entry text changes.

    Debouncing via _last_preview_value prevents excessive image generation.
    Thread-safe to prevent race conditions during rapid typing.
    """
    global _last_preview_value

    barcode_value = entry.get().strip()

    with _preview_update_lock:
        if _last_preview_value == barcode_value:
            return
        _last_preview_value = barcode_value

    if not barcode_value:
        preview_label.config(image="")
        preview_label.image = None
        return

    try:
        img = generate_label_image(barcode_value)

        # Use LANCZOS for better quality downscaling
        preview_img = img.resize((PREVIEW_WIDTH, PREVIEW_HEIGHT), Image.LANCZOS)
        tk_img = ImageTk.PhotoImage(preview_img)

        preview_label.config(image=tk_img)
        preview_label.image = tk_img  # Keep reference

        root.update_idletasks()

        # Update window size constraints
        extra_height = 400
        min_width = max(MIN_WINDOW_WIDTH, PREVIEW_WIDTH + 100)
        min_height = max(MIN_WINDOW_HEIGHT, PREVIEW_HEIGHT + extra_height)
        root.minsize(min_width, min_height)

        config["window_size"] = root.geometry()
        debounced_config_saver.save()
    except Exception as exc:
        logging.error(f"Preview update error: {exc}")
        preview_label.config(image="")
        preview_label.image = None


def parse_listbox_entry(item_text: str) -> tuple:
    """Parse history entry to extract barcode text and copy count.

    Handles legacy format for backward compatibility.
    """
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
    """Display tooltip text in status bar on hover."""

    def on_enter(_event):
        status_var.set(text)

    def on_leave(_event):
        status_var.set("Ready")

    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)


# === PRINTING OPERATIONS ===


def threaded_print(
    img: Image.Image, printer_name: str, copies: int, barcode_value: str
) -> None:
    """Print barcode in background thread to keep UI responsive.

    UI updates queued via root.after() to avoid cross-thread crashes.
    """
    try:

        def set_progress_safe(msg):
            root.after(0, lambda: set_progress(msg))

        set_progress_safe(f"Printing {copies} copies...")

        for i in range(copies):
            set_progress_safe(f"Printing copy {i + 1} of {copies}...")
            print_image(img, printer_name)

        def update_history():
            try:
                update_barcode_history(barcode_value, copies)
                refresh_history_display()
                entry.delete(0, tk.END)
                update_preview()
                set_progress("Done.")
            except Exception as exc:
                logging.error(f"Failed to update history: {exc}")

        root.after(0, update_history)
    except Exception as print_exc:
        error_msg = str(print_exc)

        def show_error():
            logging.error(f"Print Error: {error_msg}")
            messagebox.showerror("Print Error", error_msg)
            set_progress("")

        root.after(0, show_error)


def handle_print() -> None:
    """Validate inputs and initiate print job."""
    barcode_value = entry.get().strip()
    selected_printer = printer_var.get()

    if not barcode_value:
        messagebox.showwarning("Missing Barcode", "Please enter a barcode.")
        return

    if not selected_printer:
        messagebox.showwarning("No Printer Selected", "Please select a printer.")
        return

    try:
        copies_int = int(copies_var.get())
        if copies_int < 1:
            raise ValueError("Copies must be at least 1")
    except ValueError:
        messagebox.showwarning(
            "Invalid Copies", "Please enter a valid number of copies (1 or more)."
        )
        return

    try:
        img = generate_label_image(barcode_value)
        threading.Thread(
            target=threaded_print,
            args=(img, selected_printer, copies_int, barcode_value),
            daemon=True,
        ).start()
    except Exception as exc:
        logging.error(f"Print Error: {exc}")
        messagebox.showerror("Print Error", str(exc))


def threaded_reprint(selected_items, selected_printer):
    """Reprint selected items from history in background thread."""

    def set_progress_safe(msg):
        root.after(0, lambda: set_progress(msg))

    try:
        set_progress_safe(f"Reprinting {len(selected_items)} barcode(s)...")

        for idx, item_id in enumerate(selected_items):
            values = listbox.item(item_id, "values")
            if not values:
                continue

            barcode_text, copies = values[0], int(values[1])
            img = generate_label_image(barcode_text)

            for c in range(copies):
                set_progress_safe(
                    f"Reprinting {idx + 1}/{len(selected_items)}: copy {c + 1} of {copies}"
                )
                print_image(img, selected_printer)

            def update_reprint_count(bt=barcode_text, cp=copies):
                try:
                    update_barcode_history(bt, cp)
                    refresh_history_display()
                except Exception as exc:
                    logging.error(f"Failed to update reprint history: {exc}")

            root.after(0, update_reprint_count)

        root.after(0, lambda: set_progress("Done."))
    except Exception as reprint_exc:
        error_msg = str(reprint_exc)

        def show_error():
            logging.error(f"Reprint Error: {error_msg}")
            messagebox.showerror("Print Error", error_msg)
            set_progress("")

        root.after(0, show_error)


def reprint_selected() -> None:
    """Validate selection and initiate reprint job."""
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
        threading.Thread(
            target=threaded_reprint,
            args=(selected_items, selected_printer),
            daemon=True,
        ).start()
    except Exception as exc:
        logging.error(f"Reprint Error: {exc}")
        messagebox.showerror("Print Error", str(exc))


# === CONFIGURATION & PERSISTENCE ===

APPDATA_DIR = Path(os.getenv("APPDATA", os.path.expanduser("~")))
CONFIG_DIR = APPDATA_DIR / "BarcodePrinter"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "barcode_printer_config.json"
HISTORY_FILE = CONFIG_DIR / "barcode_history.json"

DEFAULT_CONFIG = {
    "default_printer": "",
    "window_size": "550x750",
    "language": "en",
    "theme": "dark",
}

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
        "about": "Barcode Printer\nVersion 1.0",
        "help": "Select a printer, scan a barcode, and click Print.",
    }
}


def load_config():
    """Load config from disk, falling back to defaults on error."""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            loaded = json.load(f)
            # Merge with defaults to handle missing keys
            return {**DEFAULT_CONFIG, **loaded}
    except (OSError, json.JSONDecodeError) as exc:
        logging.warning(f"Failed to load config: {exc}")
        return DEFAULT_CONFIG.copy()


def save_config(cfg):
    """Persist configuration to disk."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except OSError as exc:
        logging.error(f"Failed to save config: {exc}")


config = load_config()

_history_lock = threading.Lock()


def load_history():
    """Load print history from disk."""
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        logging.warning(f"Failed to load history: {exc}")
        return []


def save_history(history):
    """Persist print history to disk."""
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history[:HISTORY_MAX_SIZE], f, indent=2)
    except OSError as exc:
        logging.error(f"Failed to save history: {exc}")


barcode_history = load_history()


def update_barcode_history(barcode_value: str, copies: int):
    """Update history with new print job. Thread-safe."""
    global barcode_history

    with _history_lock:
        found = False
        for idx, item in enumerate(barcode_history):
            if item.get("barcode") == barcode_value:
                barcode_history[idx]["copies"] += copies
                updated_item = barcode_history.pop(idx)
                barcode_history.insert(0, updated_item)
                found = True
                break

        if not found:
            barcode_history.insert(0, {"barcode": barcode_value, "copies": copies})

        barcode_history = barcode_history[:HISTORY_MAX_SIZE]
        save_history(barcode_history)


def refresh_history_display():
    """Refresh the treeview with current history."""
    # Clear existing items
    for item in listbox.get_children():
        listbox.delete(item)

    # Repopulate from history
    with _history_lock:
        for item in barcode_history:
            if isinstance(item, dict):
                barcode = item.get("barcode")
                copies = item.get("copies", 1)
                if barcode:
                    listbox.insert("", tk.END, values=(barcode, copies))


def _(key: str) -> str:
    """Get localized string for current language."""
    lang = config.get("language", "en")
    return str(LANGUAGES.get(lang, LANGUAGES["en"]).get(key, key))


def on_exit():
    """Flush pending config saves before exit."""
    try:
        debounced_config_saver.flush()
        logging.info("Application exited gracefully.")
    except Exception as exc:
        logging.error(f"Error during exit: {exc}")


atexit.register(on_exit)


class DebouncedConfigSaver:
    """Debounce config saves to avoid excessive disk I/O."""

    def __init__(self, delay=CONFIG_SAVE_DELAY):
        self.delay = delay
        self._timer = None
        self._lock = threading.Lock()

    def save(self):
        """Schedule a delayed save."""
        with self._lock:
            if self._timer:
                self._timer.cancel()
            self._timer = threading.Timer(self.delay, self._save_now)
            self._timer.start()

    def _save_now(self):
        """Execute the save operation."""
        save_config(config)

    def flush(self):
        """Force immediate save."""
        with self._lock:
            if self._timer:
                self._timer.cancel()
                self._timer = None
            save_config(config)


debounced_config_saver = DebouncedConfigSaver()


# === HiDPI SUPPORT ===


def set_hidpi_scaling(root):
    """Enable DPI awareness for crisp rendering on high-DPI displays."""
    try:
        if sys.platform == "win32" and hasattr(ctypes, "windll"):
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(
                    2
                )  # PROCESS_PER_MONITOR_DPI_AWARE
            except Exception:
                try:
                    ctypes.windll.user32.SetProcessDPIAware()
                except Exception as exc:
                    logging.warning(f"Failed to set DPI awareness: {exc}")

            try:
                user32 = ctypes.windll.user32
                dpi = (
                    user32.GetDpiForSystem()
                    if hasattr(user32, "GetDpiForSystem")
                    else 96
                )
                scaling = dpi / 96.0
                root.tk.call("tk", "scaling", scaling)
            except Exception as exc:
                logging.warning(f"Failed to set scaling: {exc}")
    except Exception as exc:
        logging.warning(f"Could not set HiDPI scaling: {exc}")


# === GUI INITIALIZATION ===

root = tk.Tk()
set_hidpi_scaling(root)
root.title("Barcode Printer")
root.geometry(config.get("window_size", "550x750"))
root.minsize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)


def apply_theme_to_titlebar(root):
    """Style native title bar to match app theme."""
    try:
        version = sys.getwindowsversion()
        current_theme = sv_ttk.get_theme()

        if version.major == 10 and version.build >= 22000:
            # Windows 11
            header_color = "#1c1c1c" if current_theme == "dark" else "#fafafa"
            pywinstyles.change_header_color(root, header_color)
        elif version.major == 10:
            # Windows 10
            style = "dark" if current_theme == "dark" else "normal"
            pywinstyles.apply_style(root, style)
            # Force title bar update
            root.wm_attributes("-alpha", 0.99)
            root.wm_attributes("-alpha", 1.0)
    except Exception as exc:
        logging.warning(f"Failed to apply theme to titlebar: {exc}")


def get_theme_from_config():
    """Get current theme from config."""
    return config.get("theme", "dark")


def set_theme_in_config(theme):
    """Save theme to config."""
    config["theme"] = theme
    debounced_config_saver.save()


# Set initial theme
sv_ttk.set_theme(get_theme_from_config())


def toggle_theme():
    """Switch between dark and light themes."""
    current_theme = sv_ttk.get_theme()
    new_theme = "light" if current_theme == "dark" else "dark"
    sv_ttk.set_theme(new_theme)
    set_theme_in_config(new_theme)
    theme_button.config(
        text=f"Switch to {'Dark' if new_theme == 'light' else 'Light'} Theme"
    )
    apply_theme_to_titlebar(root)


# Theme toggle button
theme_button = ttk.Button(
    root,
    text=f"Switch to {'Light' if get_theme_from_config() == 'dark' else 'Dark'} Theme",
    command=toggle_theme,
)
theme_button.pack(anchor="ne", padx=10, pady=5)


def save_window_size_on_focus_out(event=None):
    """Save window geometry when focus lost."""
    config["window_size"] = root.geometry()
    debounced_config_saver.save()


root.bind("<FocusOut>", save_window_size_on_focus_out)

# Set window icon
try:
    icon_path = Path("./barcode-scan.png")
    if icon_path.exists():
        icon_img = tk.PhotoImage(file=str(icon_path))
        root.iconphoto(True, icon_img)
except Exception as exc:
    logging.warning(f"Could not set window icon: {exc}")


# === WIDGET CREATION ===

# Printer selection
ttk.Label(root, text=_("select_printer")).pack(pady=(10, 0))
printer_var = tk.StringVar(value=config.get("default_printer", ""))
printer_dropdown = ttk.Combobox(
    root, textvariable=printer_var, values=get_printers(force_refresh=True), width=50
)
printer_dropdown.pack(pady=(0, 10))


def on_printer_selected(_event=None):
    """Save selected printer as default."""
    config["default_printer"] = printer_var.get()
    debounced_config_saver.save()


printer_dropdown.bind("<<ComboboxSelected>>", on_printer_selected)

# Barcode entry
ttk.Label(root, text=_("scan_barcode")).pack()
entry = ttk.Entry(root, font=("Segoe UI Variable", 12), width=35)
entry.pack(pady=5)
entry.focus()
entry.bind("<KeyRelease>", update_preview)

# Copies spinbox
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

# Preview label
ttk.Label(root, text=_("preview")).pack(pady=(10, 0))
preview_label = ttk.Label(root)
preview_label.pack(pady=(0, 10))

# Progress indicator
progress_var = tk.StringVar(value="")
progress_label = ttk.Label(
    root, textvariable=progress_var, font=("Segoe UI Variable", 12)
)
progress_label.pack(pady=(0, 5))


def set_progress(msg):
    """Update progress indicator."""
    progress_var.set(msg)
    root.update_idletasks()


def on_print():
    """Handle print button click."""
    config["default_printer"] = printer_var.get()
    handle_print()


# Print button
print_button = ttk.Button(
    root, text=_("print"), command=on_print, width=20, style="Accent.TButton"
)
print_button.pack(pady=(5, 10))

# History treeview
tree_frame = ttk.Frame(root)
tree_frame.pack(pady=10, padx=10, fill="both", expand=False)

tree_columns = ("Barcode", "Copies")
listbox = ttk.Treeview(tree_frame, columns=tree_columns, show="headings", height=10)
listbox.heading("Barcode", text="Barcode")
listbox.heading("Copies", text="Copies")
listbox.column("Barcode", width=350)
listbox.column("Copies", width=80, anchor="center")

scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=listbox.yview)
listbox.configure(yscrollcommand=scrollbar.set)

listbox.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Populate history
refresh_history_display()

# Reprint button
reprint_button = ttk.Button(
    root, text=_("reprint_selected"), command=reprint_selected, width=20
)
reprint_button.pack(pady=(5, 10))

# Bind Enter key to print
entry.bind("<Return>", lambda event: on_print())

# Status bar
status_var = tk.StringVar(value="Ready")
status_label = ttk.Label(root, textvariable=status_var, font=("Segoe UI Variable", 12))
status_label.pack(pady=(0, 5))


def set_status(msg):
    """Update status bar text."""
    status_var.set(msg)
    root.update_idletasks()


# Add tooltips
add_tooltip(print_button, "Print the current barcode")
add_tooltip(reprint_button, "Reprint selected barcodes from history")
add_tooltip(printer_dropdown, "Select a printer from available devices")
add_tooltip(entry, "Enter or scan a barcode value")
add_tooltip(copies_spinbox, "Set number of copies to print")

# Keyboard shortcuts
root.bind("<Alt-p>", lambda e: print_button.invoke())
root.bind("<Alt-r>", lambda e: reprint_button.invoke())

# Apply theme to titlebar
apply_theme_to_titlebar(root)


def focus_entry_on_window_focus(event=None):
    """Return focus to entry field when window regains focus."""
    entry.focus_set()


root.bind("<FocusIn>", focus_entry_on_window_focus)

# Start the main loop
root.mainloop()
