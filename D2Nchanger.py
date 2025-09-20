import os
import sys
import json
import ctypes
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, scrolledtext
from tkinter.ttk import Button, Label, Frame, Style
from PIL import Image, ImageDraw
import pystray
import threading
import logging

# Ensure optional dependencies are handled
try:
    from win32com.client import Dispatch
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("win32com not installed. Startup shortcut creation will be disabled.")

CONFIG_FILE = "wallpaper_config.json"

class WallpaperScheduler:
    def __init__(self):
        self.wallpaper_schedule = {
            "morning": "",
            "afternoon": "",
            "evening": "",
            "night": ""
        }
        self.time_periods = {
            "morning": (6, 12),
            "afternoon": (12, 17),
            "evening": (17, 20),
            "night": (20, 24)
        }
        self.root = None
        self.msg_label = None
        self.tray_icon = None
        self.current_message = ""
        self.log_area = None
        self.time_label = None

        # Configure logging
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

        # Load configuration
        self.load_config()

    def initialize(self):
        """Initialize additional components, including adding to startup."""
        self.add_to_startup()

    def load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as file:
                    self.wallpaper_schedule = json.load(file)
                logging.info("Configuration loaded successfully.")
        except Exception as e:
            logging.error(f"Error loading configuration: {e}")

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w") as file:
                json.dump(self.wallpaper_schedule, file, indent=4)
            logging.info("Configuration saved successfully.")
        except Exception as e:
            logging.error(f"Error saving configuration: {e}")

    def add_to_startup(self):
        if not WIN32_AVAILABLE:
            logging.warning("Cannot create startup shortcut. win32com is not installed.")
            return

        try:
            startup_folder = os.path.join(
                os.getenv("APPDATA"),
                "Microsoft\\Windows\\Start Menu\\Programs\\Startup"
            )

            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
            shortcut_path = os.path.join(startup_folder, "WallpaperScheduler.lnk")

            if not os.path.exists(shortcut_path):
                shell = Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = exe_path
                shortcut.WorkingDirectory = os.path.dirname(exe_path)
                shortcut.Description = "Wallpaper Scheduler"
                shortcut.IconLocation = exe_path
                shortcut.save()

                logging.info(f"Startup shortcut created at: {shortcut_path}")
            else:
                logging.info("Startup shortcut already exists.")

        except Exception as e:
            logging.error(f"Error creating startup shortcut: {e}")

    def create_tray_icon(self):
        def restore():
            self.tray_icon.stop()
            self.root.deiconify()

        def quit_app():
            self.tray_icon.stop()
            self.root.destroy()

        def create_image():
            size = (64, 64)
            image = Image.new("RGBA", size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(image)
            draw.ellipse((8, 8, 56, 56), fill="white", outline="blue", width=2)
            draw.text((20, 20), "WS", fill="blue")
            return image

        menu = pystray.Menu(
            pystray.MenuItem("Restore", restore),
            pystray.MenuItem("Quit", quit_app)
        )
        self.tray_icon = pystray.Icon("Wallpaper Scheduler", create_image(), "Wallpaper Scheduler", menu)

    def update_message(self, message, error=False):
        if self.msg_label:
            color = "red" if error else "green"
            self.msg_label.config(text=message, foreground=color)
            self.root.after(3000, lambda: self.msg_label.config(text=""))

    def save_wallpaper(self, period):
        filepath = filedialog.askopenfilename(filetypes=[("Image Files", "*.jpg;*.png;*.bmp")])
        if filepath and os.path.exists(filepath):
            self.wallpaper_schedule[period] = filepath
            self.save_config()
            self.update_message(f"Wallpaper for {period} updated successfully!")
            logging.info(f"Wallpaper for {period} set to {filepath}")
        else:
            self.update_message("Invalid file selected. Please try again.", error=True)

    def set_wallpaper(self, image_path):
        if os.path.exists(image_path):
            try:
                ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
                logging.info(f"Wallpaper set successfully: {image_path}")
            except Exception as e:
                logging.error(f"Error setting wallpaper: {e}")
        else:
            logging.error(f"Wallpaper file not found: {image_path}")

    def update_time(self):
        if self.time_label:
            current_time = datetime.now().strftime("%I:%M %p")
            self.time_label.config(text=f"Current Time: {current_time}")
            self.root.after(1000, self.update_time)

    def start_changing_wallpapers(self):
        missing_wallpapers = [period for period, path in self.wallpaper_schedule.items() if not path]
        if missing_wallpapers:
            self.update_message("Please set wallpapers for all periods.", error=True)
            logging.warning(f"Missing wallpapers for periods: {missing_wallpapers}")
            return

        self.root.withdraw()
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

        def change_wallpapers():
            current_period = self.get_current_period()
            wallpaper_path = self.wallpaper_schedule.get(current_period, "")
            if wallpaper_path and os.path.exists(wallpaper_path):
                self.set_wallpaper(wallpaper_path)
                logging.info(f"Changed wallpaper for {current_period}")
            else:
                logging.warning(f"No wallpaper found for {current_period}")
            self.root.after(60000, change_wallpapers)

        change_wallpapers()
        self.update_message("Wallpaper Scheduler Started!")

    def get_current_period(self):
        current_hour = datetime.now().hour
        for period, (start, end) in self.time_periods.items():
            if start <= current_hour < end:
                return period
        return "night"

    def create_log_handler(self):
        class TextWidgetHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
                self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

            def emit(self, record):
                msg = self.format(record)
                if self.text_widget:
                    self.text_widget.config(state=tk.NORMAL)
                    self.text_widget.insert(tk.END, msg + '\n')
                    self.text_widget.see(tk.END)
                    self.text_widget.config(state=tk.DISABLED)

        return TextWidgetHandler

    def run(self):
        self.root = tk.Tk()
        self.root.title("Wallpaper Scheduler")
        self.root.geometry("896x705")
        self.root.minsize(896, 705)

        self.create_tray_icon()

        style = Style()
        style.configure("TButton", font=("Helvetica", 10), padding=5)
        style.configure("TLabel", font=("Helvetica", 12))

        header = Label(self.root, text="Wallpaper Scheduler", font=("Helvetica", 16, "bold"), anchor="center")
        header.grid(row=0, column=0, columnspan=2, pady=10, sticky="ew")

        self.time_label = Label(self.root, text="", font=("Helvetica", 12))
        self.time_label.grid(row=1, column=0, columnspan=2, pady=5, sticky="ew")
        self.update_time()

        self.msg_label = Label(self.root, text="", font=("Helvetica", 10), anchor="center")
        self.msg_label.grid(row=2, column=0, columnspan=2, pady=5, sticky="ew")

        for index, period in enumerate(self.wallpaper_schedule.keys(), start=3):
            frame = Frame(self.root)
            frame.grid(row=index, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

            label = Label(frame, text=f"{period.capitalize()} wallpaper:")
            label.pack(side=tk.LEFT, padx=10)

            button = Button(frame, text="Choose File", command=lambda p=period: self.save_wallpaper(p))
            button.pack(side=tk.RIGHT, padx=10)

        start_button = Button(self.root, text="Start Wallpaper Scheduler",
                              command=self.start_changing_wallpapers)
        start_button.grid(row=index + 1, column=0, columnspan=2, pady=20)

        log_label = Label(self.root, text="Logs:", font=("Helvetica", 12))
        log_label.grid(row=index + 2, column=0, columnspan=2, pady=(10, 0))

        self.log_area = scrolledtext.ScrolledText(
            self.root,
            wrap=tk.WORD,
            font=("Courier", 10)
        )
        self.log_area.grid(row=index + 3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.log_area.config(state=tk.DISABLED)

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(index + 3, weight=1)

        log_handler = self.create_log_handler()(self.log_area)
        logging.getLogger().addHandler(log_handler)
        logging.getLogger().setLevel(logging.INFO)

        logging.info("Wallpaper Scheduler UI initialized")
        self.initialize()
        self.root.mainloop()

def main():
    scheduler = WallpaperScheduler()
    scheduler.run()

if __name__ == "__main__":
    main()
