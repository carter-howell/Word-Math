import os
import pystray
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import messagebox
import keyboard
import win32com.client
import pythoncom
import re
import locale
import threading

# Global variable to store the current hotkey
current_hotkey = 'ctrl+alt+m'
HOTKEY_FILE = "hotkey.txt"  # File to save the hotkey

def evaluate_math_expression(expression):
    """Evaluate a math expression and return the result."""
    try:
        expression = expression.replace("x", "*").replace(",", "")
        result = eval(expression)
        return result
    except Exception as e:
        print(f"Error evaluating expression: {e}")
        return None

def format_number_with_commas(number):
    """Format a number with commas as thousand separators."""
    return locale.format_string("%d", number, grouping=True) if isinstance(number, int) else locale.format_string("%.2f", number, grouping=True)

def process_line_in_word():
    """Process the current line in Microsoft Word."""
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.ActiveDocument
        selection = word.Selection
        selection.HomeKey(Unit=5)
        selection.EndKey(Unit=5, Extend=1)
        line = selection.Text.strip()
        match = re.match(r'^[\d\s.,+\-x/()]+=\s*$', line)
        if match:
            expression = line.rstrip('=').strip()
            result = evaluate_math_expression(expression)
            if result is not None:
                formatted_result = format_number_with_commas(result)
                new_line = f"{expression} = {formatted_result}"
                selection.Text = new_line
                selection.Collapse(0)
                selection.TypeParagraph()
    except Exception as e:
        print(f"Error: {e}")
    finally:
        pythoncom.CoUninitialize()

def create_system_tray():
    def on_quit(icon, item):
        icon.stop()
        os._exit(0)

    def show_window(icon, item):
        # Use a thread-safe approach to interact with the Tkinter GUI
        root.after(0, root.deiconify)
        root.after(100, root.lift)

    # Create a default icon if icon.png is not found
    try:
        icon_image = Image.open("icon.png")  # Ensure you have an icon.png file in the same directory
    except FileNotFoundError:
        # Create a simple default icon
        icon_image = Image.new('RGB', (64, 64), color='white')
        draw = ImageDraw.Draw(icon_image)
        draw.rectangle([16, 16, 48, 48], outline='black', fill='blue')

    menu = pystray.Menu(pystray.MenuItem('Show', show_window), pystray.MenuItem('Quit', on_quit))
    icon = pystray.Icon("MathEvalApp", icon_image, menu=menu)
    icon.run_detached()

def update_hotkey():
    global current_hotkey
    new_hotkey = hotkey_entry.get().strip()
    if new_hotkey:
        try:
            keyboard.remove_hotkey(current_hotkey)
            keyboard.add_hotkey(new_hotkey, process_line_in_word)
            current_hotkey = new_hotkey
            save_hotkey_to_file(new_hotkey)  # Save the new hotkey to a file
            messagebox.showinfo("Success", f"Hotkey updated to {new_hotkey}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update hotkey: {e}")
    else:
        messagebox.showwarning("Warning", "Please enter a valid hotkey")

def save_hotkey_to_file(hotkey):
    """Save the hotkey to a file."""
    with open(HOTKEY_FILE, "w") as file:
        file.write(hotkey)

def load_hotkey_from_file():
    """Load the hotkey from a file. If the file doesn't exist, return the default hotkey."""
    if os.path.exists(HOTKEY_FILE):
        with open(HOTKEY_FILE, "r") as file:
            return file.read().strip()
    return 'ctrl+alt+m'  # Default hotkey

def create_gui():
    global hotkey_entry, root

    root = tk.Tk()
    root.title("MathEvalApp Hotkey Settings")
    root.geometry("300x100")

    tk.Label(root, text="Enter new hotkey:").pack(pady=5)
    hotkey_entry = tk.Entry(root)
    hotkey_entry.pack(pady=5)
    hotkey_entry.insert(0, current_hotkey)

    save_button = tk.Button(root, text="Save", command=update_hotkey)
    save_button.pack(pady=10)

    root.protocol("WM_DELETE_WINDOW", lambda: root.withdraw())  # Minimize to system tray on close

def main():
    global current_hotkey

    locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

    # Load the saved hotkey or use the default
    current_hotkey = load_hotkey_from_file()

    # Create the GUI (starts minimized)
    create_gui()
    root.withdraw()  # Start minimized

    # Start the system tray in a separate thread
    threading.Thread(target=create_system_tray, daemon=True).start()

    print(f"Macro is running. Press '{current_hotkey}' to process the current line in Word.")
    keyboard.add_hotkey(current_hotkey, process_line_in_word)
    root.mainloop()  # Start the Tkinter main loop

if __name__ == "__main__":
    main()