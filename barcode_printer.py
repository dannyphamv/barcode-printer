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
import pywinstyles, sys


# === PRINTER MANAGEMENT ===

_PRINTER_LIST_CACHE = None


def get_printers(force_refresh=False) -> list[str]:
    """Return a list of available printer names.
    
    Caching prevents repeated expensive system calls to EnumPrinters.
    Refreshing is needed when printers are added/removed during runtime.
    """
    global _PRINTER_LIST_CACHE
    if force_refresh or _PRINTER_LIST_CACHE is None:
        _PRINTER_LIST_CACHE = [
            printer[2]
            for printer in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
        ]
    return _PRINTER_LIST_CACHE  # type: ignore


# === BARCODE GENERATION & CACHING ===

# LRU cache prevents regenerating the same barcode multiple times
# 100-item limit balances memory usage with cache effectiveness
BARCODE_IMAGE_CACHE = OrderedDict()
BARCODE_IMAGE_CACHE_MAXSIZE = 100


def generate_label_image(barcode_text: str) -> Image.Image:
    """Generate a 600x300px label with centered Code128 barcode.
    
    Returns a copy from cache to prevent modifications affecting cached version.
    LRU eviction ensures memory doesn't grow unbounded.
    """
    if barcode_text in BARCODE_IMAGE_CACHE:
        BARCODE_IMAGE_CACHE.move_to_end(barcode_text)
        return BARCODE_IMAGE_CACHE[barcode_text].copy()
    
    code128 = python_barcode.get("code128", barcode_text, writer=ImageWriter())
    with io.BytesIO() as buffer:
        code128.write(buffer)
        buffer.seek(0)
        barcode_img = Image.open(buffer)
        
        # Standard label size chosen for compatibility with common label printers
        label_width, label_height = 600, 300
        label_img = Image.new("RGB", (label_width, label_height), 0xFFFFFF)
        barcode_x = (label_width - barcode_img.width) // 2
        barcode_y = (label_height - barcode_img.height) // 2
        label_img.paste(barcode_img, (barcode_x, barcode_y))
        
        BARCODE_IMAGE_CACHE[barcode_text] = label_img.copy()
        BARCODE_IMAGE_CACHE.move_to_end(barcode_text)
        if len(BARCODE_IMAGE_CACHE) > BARCODE_IMAGE_CACHE_MAXSIZE:
            BARCODE_IMAGE_CACHE.popitem(last=False)
        return label_img


def print_image(img: Image.Image, printer_name: str) -> None:
    """Send image to Windows printer using GDI.
    
    Scales image to fit printer's width while maintaining aspect ratio.
    Centers vertically to avoid cutting off content on smaller paper.
    """
    pdc = win32ui.CreateDC()
    if pdc is None:
        logging.error("Failed to create printer device context.")
        raise RuntimeError("Failed to create printer device context.")
    pdc.CreatePrinterDC(printer_name)
    try:
        pdc.StartDoc("Barcode Print")
        pdc.StartPage()
        
        printable_width = pdc.GetDeviceCaps(win32con.HORZRES)
        printable_height = pdc.GetDeviceCaps(win32con.VERTRES)
        
        # Scale to full width to maximize barcode readability
        if img.width != printable_width:
            scale = printable_width / img.width
            scaled_width = printable_width
            scaled_height = int(img.height * scale)
            img = img.resize((scaled_width, scaled_height))
        
        # Center vertically to prevent partial prints on short paper
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


# === UI UPDATE FUNCTIONS ===

# Track last preview to avoid redundant regeneration on every keystroke
_last_preview_value = None


def update_preview(event=None):
    """Update barcode preview when entry text changes.
    
    Debouncing via _last_preview_value prevents excessive image generation.
    Window resizing ensures all widgets remain visible as preview updates.
    """
    global _last_preview_value
    _ = event
    barcode_value = entry.get().strip()
    
    if _last_preview_value == barcode_value:
        return
    _last_preview_value = barcode_value
    if not barcode_value:
        preview_label.config(image="")
        setattr(preview_label, "image", None)
        return
    try:
        img = generate_label_image(barcode_value)
        # BOX filter is fastest for downsizing preview images
        preview_img = (
            img.resize((400, 200), Image.BOX) if img.size != (400, 200) else img
        )
        tk_img = ImageTk.PhotoImage(preview_img)
        preview_label.config(image=tk_img)
        # Keep reference to prevent garbage collection
        setattr(preview_label, "image", tk_img)
        
        root.update_idletasks()
        # Enforce minimum size to prevent UI elements from being cut off
        min_width, min_height = 650, 1000
        preview_width = preview_img.width
        preview_height = preview_img.height
        extra_height = 400
        min_width = max(min_width, preview_width + 100)
        min_height = max(min_height, preview_height + extra_height)
        root.minsize(min_width, min_height)
        
        config["window_size"] = root.geometry()
        debounced_config_saver.save()
    except (OSError, RuntimeError, ValueError) as exc:
        preview_label.config(image="")
        setattr(preview_label, "image", None)
        print("Preview update error:", exc)


def parse_listbox_entry(item_text: str) -> tuple[str, int]:
    """Parse history entry to extract barcode text and copy count.
    
    Handles legacy format for backward compatibility with older history files.
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
    """Display tooltip text in status bar on hover.
    
    Status bar approach chosen over popup tooltips for cleaner UI.
    """
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
    
    UI updates must be queued via root.after() to avoid cross-thread crashes.
    History updated after successful print to maintain data integrity.
    """
    global barcode_history
    try:
        def set_progress_safe(msg):
            root.after(0, set_progress, msg)
        set_progress_safe(f"Printing {copies} copies...")
        for i in range(copies):
            set_progress_safe(f"Printing copy {i+1} of {copies}...")
            print_image(img, printer_name)
        
        def update_history():
            global barcode_history
            # Update existing entry or create new one
            found = False
            for idx, item in enumerate(barcode_history):
                if item.get("barcode") == barcode_value:
                    barcode_history[idx]["copies"] += copies
                    updated_item = barcode_history.pop(idx)
                    barcode_history = [updated_item] + barcode_history
                    found = True
                    break
            if not found:
                barcode_history = ([{"barcode": barcode_value, "copies": copies}] + barcode_history)[:100]
            else:
                barcode_history = barcode_history[:100]
            save_history(barcode_history)
            
            # Sync treeview with history
            for row in list(listbox.get_children()):
                values = listbox.item(row, "values")
                if values and values[0] == barcode_value:
                    listbox.delete(row)
            listbox.insert("", 0, values=(barcode_value, next((item["copies"] for item in barcode_history if item["barcode"] == barcode_value), copies)))
            entry.delete(0, tk.END)
            update_preview()
            set_progress("Done.")
        root.after(0, update_history)
    except (OSError, RuntimeError) as exc:
        def show_error():
            logging.error("Print Error: %s", exc)
            messagebox.showerror("Print Error", str(exc))
            set_progress("")
        root.after(0, show_error)


def handle_print() -> None:
    """Validate inputs and initiate print job."""
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
        threading.Thread(
            target=threaded_print,
            args=(img, selected_printer, copies_int, barcode_value),
            daemon=True,
        ).start()
    except (OSError, RuntimeError) as exc:
        logging.error("Print Error: %s", exc)
        messagebox.showerror("Print Error", str(exc))


def threaded_reprint(selected_items, selected_printer):
    """Reprint selected items from history in background thread."""
    def set_progress_safe(msg):
        root.after(0, set_progress, msg)
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
                    f"Reprinting {idx+1}/{len(selected_items)}: copy {c+1} of {copies}"
                )
                print_image(img, selected_printer)
            
            def update_reprint_count():
                global barcode_history
                for idx, item in enumerate(barcode_history):
                    if item.get("barcode") == barcode_text:
                        item["copies"] += copies
                        updated_item = barcode_history.pop(idx)
                        barcode_history = [updated_item] + barcode_history
                        break
                else:
                    barcode_history = ([{"barcode": barcode_text, "copies": copies}] + barcode_history)[:100]
                barcode_history = barcode_history[:100]
                save_history(barcode_history)
                
                for row in list(listbox.get_children()):
                    row_values = listbox.item(row, "values")
                    if row_values and row_values[0] == barcode_text:
                        listbox.delete(row)
                listbox.insert("", 0, values=(barcode_text, next((item["copies"] for item in barcode_history if item["barcode"] == barcode_text), copies)))
            root.after(0, update_reprint_count)
        root.after(0, set_progress, "Done.")
    except (OSError, RuntimeError) as exc:
        def show_error():
            logging.error("Reprint Error: %s", exc)
            messagebox.showerror("Print Error", str(exc))
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
    except (OSError, RuntimeError) as exc:
        logging.error("Reprint Error: %s", exc)
        messagebox.showerror("Print Error", str(exc))


# === CONFIGURATION & PERSISTENCE ===

# Store in AppData to persist settings across user sessions
APPDATA_DIR = Path(os.getenv("APPDATA", os.path.expanduser("~")))
CONFIG_DIR = APPDATA_DIR / "BarcodePrinter"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = str(CONFIG_DIR / "barcode_printer_config.json")
DEFAULT_CONFIG = {
    "default_printer": "",
    "window_size": "550x750",
    "language": "en",
}

# Internationalization support - currently English only
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
        "help": "Select a printer, scan a barcode, and click Print. Use Reprint to print again.",
    }
}


def load_config():
    """Load config from disk, falling back to defaults on error.
    
    Graceful degradation ensures app works even with corrupted config files.
    """
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        logging.error("Failed to load config: %s", exc)
        return DEFAULT_CONFIG.copy()


def save_config(cfg):
    """Persist configuration to disk."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except OSError as exc:
        logging.error("Failed to save config: %s", exc)


config = load_config()

# History stored separately to avoid losing it when config is corrupted
HISTORY_FILE = str(CONFIG_DIR / "barcode_history.json")


def load_history():
    """Load print history from disk."""
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return []


def save_history(history):
    """Persist print history, limiting to 100 most recent entries."""
    limited_history = history[-100:]
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(limited_history, f)
    except OSError as exc:
        logging.error("Failed to save history: %s", exc)


barcode_history = load_history()


def _(key: str) -> str:
    """Get localized string for current language."""
    lang = config.get("language", "en")
    return str(LANGUAGES.get(lang, LANGUAGES["en"]).get(key, key))


def on_exit():
    """Flush pending config saves before exit to prevent data loss."""
    debounced_config_saver.flush()
    logging.info("Application exited gracefully.")


atexit.register(on_exit)


class DebouncedConfigSaver:
    """Debounce config saves to avoid excessive disk I/O.
    
    Delays save for 0.5s after last change to batch rapid updates
    (e.g., window resizing generates many geometry changes).
    """
    def __init__(self, delay=0.5):
        self.delay = delay
        self._timer = None

    def save(self):
        if self._timer:
            self._timer.cancel()
        self._timer = threading.Timer(self.delay, lambda: save_config(config))
        self._timer.start()

    def flush(self):
        """Force immediate save, used on app exit."""
        if self._timer:
            self._timer.cancel()
            save_config(config)
            self._timer = None


debounced_config_saver = DebouncedConfigSaver()


# === HiDPI SUPPORT ===

def set_hidpi_scaling(root):
    """Enable DPI awareness for crisp rendering on high-DPI displays.
    
    Without this, app appears blurry on 4K/retina screens.
    Windows-only as other platforms handle DPI automatically.
    """
    try:
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


# === GUI INITIALIZATION ===

root = tk.Tk()
set_hidpi_scaling(root)
root.title("Barcode Printer")
root.geometry(config.get("window_size", "550x750"))
root.minsize(650, 1000)


def apply_theme_to_titlebar(root):
    """Style native title bar to match app theme.
    
    Windows 10/11 require different APIs for title bar customization.
    Alpha workaround forces title bar refresh on Windows 10.
    """
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
        # HACK: Force title bar update by changing transparency
        root.wm_attributes("-alpha", 0.99)
        root.wm_attributes("-alpha", 1)


def get_theme_from_config():
    return config.get("theme", "dark")


def set_theme_in_config(theme):
    config["theme"] = theme
    debounced_config_saver.save()


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


theme_button = ttk.Button(
    root,
    text=f"Switch to {'Light' if get_theme_from_config() == 'dark' else 'Dark'} Theme",
    command=toggle_theme,
)
theme_button.pack(anchor="ne", padx=10, pady=5)


def save_window_size_on_focus_out(event=None):
    """Save window geometry when focus lost to persist user preferences."""
    config["window_size"] = root.geometry()
    debounced_config_saver.save()


root.bind("<FocusOut>", save_window_size_on_focus_out)

try:
    icon_img = tk.PhotoImage(file="./barcode-scan.png")
    root.iconphoto(True, icon_img)
except Exception as exc:
    print(f"Could not set window icon: {exc}")


# === WIDGET CREATION ===

ttk.Label(root, text=_("select_printer")).pack(pady=(10, 0))
printer_var = tk.StringVar(value=config.get("default_printer", ""))
printer_dropdown = ttk.Combobox(
    root, textvariable=printer_var, values=get_printers(force_refresh=True), width=50
)
printer_dropdown.pack(pady=(0, 10))


def on_printer_selected(_event=None):
    """Save selected printer as default for next session."""
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

progress_var = tk.StringVar(value="")
progress_label = ttk.Label(
    root, textvariable=progress_var, font=("Segoe UI Variable", 12)
)
progress_label.pack(pady=(0, 5))


def set_progress(msg):
    """Update progress indicator during print operations."""
    progress_var.set(msg)
    root.update_idletasks()


def on_print():
    """Handle print button click."""
    config["default_printer"] = printer_var.get()
    handle_print()


print_button = ttk.Button(
    root, text=_("print"), command=on_print, width=20, style="Accent.TButton"
)
print_button.pack(pady=(5, 10))

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

# Populate history in reverse to show newest first
for item in reversed(barcode_history):
    if isinstance(item, dict):
        barcode, copies = item.get("barcode"), item.get("copies", 1)
    else:
        # Legacy format compatibility
        barcode, copies = item, 1
    if barcode:
        listbox.insert("", 0, values=(barcode, copies))

reprint_button = ttk.Button(
    root, text=_("reprint_selected"), command=reprint_selected, width=20
)
reprint_button.pack(pady=(5, 10))

entry.bind("<Return>", lambda event: on_print())

status_var = tk.StringVar(value="Ready")
status_label = ttk.Label(root, textvariable=status_var, font=("Segoe UI Variable", 12))
status_label.pack(pady=(0, 5))


def set_status(msg):
    """Update status bar text."""
    status_var.set(msg)
    root.update_idletasks()


# Add tooltips for better UX
add_tooltip(print_button, "Print the current barcode")
add_tooltip(reprint_button, "Reprint selected barcodes")
add_tooltip(printer_dropdown, "Select a printer")
add_tooltip(entry, "Enter or scan a barcode")
add_tooltip(copies_spinbox, "Set number of copies")

# Keyboard shortcuts for accessibility
print_button.focus_set()
root.bind("<Alt-p>", lambda e: print_button.invoke())
root.bind("<Alt-r>", lambda e: reprint_button.invoke())

apply_theme_to_titlebar(root)


def focus_entry_on_window_focus(event=None):
    """Return focus to entry field when window regains focus.
    
    Enables seamless barcode scanner workflow where user can
    scan -> print -> scan without manual clicking.
    """
    entry.focus_set()


root.bind("<FocusIn>", focus_entry_on_window_focus)

# WORKAROUND: Keep reference to prevent PIL cleanup issues in PyInstaller
_img_ref = ImageTk.PhotoImage(Image.new("RGB", (1, 1)))

root.mainloop()
