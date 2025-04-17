# gui.py
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Button
from tkinterdnd2 import DND_FILES

MAX_WORKERS_DEFAULT = 3

class CalendarBuddyGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Calendar Buddy")
        self.root.geometry("520x300")
        self.root.configure(bg="#f0f0f0")

        self.scrape_mode = tk.StringVar(value="new")
        self.test_mode = tk.BooleanVar(value=False)
        self.test_limit = tk.IntVar(value=10)
        self.worker_count = tk.IntVar(value=MAX_WORKERS_DEFAULT)
        self.imported_jobs = None

        self.label = Label(root, text="Drag and drop a .txt or .xlsx file here", width=50, height=4, bg="white", relief="ridge")
        self.label.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        self.test_checkbox = tk.Checkbutton(root, text="Test Mode", variable=self.test_mode)
        self.test_checkbox.grid(row=1, column=0, sticky="w", padx=10)

        self.limit_label = Label(root, text="Test Mode Limit:")
        self.limit_label.grid(row=1, column=1, sticky="e")
        self.limit_spinbox = tk.Spinbox(root, from_=1, to=500, textvariable=self.test_limit, width=5)
        self.limit_spinbox.grid(row=1, column=2, sticky="w")

        self.worker_label = Label(root, text="Worker Processes:")
        self.worker_label.grid(row=2, column=0, sticky="w", padx=10)
        self.worker_spinbox = tk.Spinbox(root, from_=1, to=8, textvariable=self.worker_count, width=5)
        self.worker_spinbox.grid(row=2, column=1, sticky="w")

        self.open_button = Button(root, text="Or Click to Browse", command=self.browse_file)
        self.open_button.grid(row=2, column=2, sticky="e", padx=10)

        self.scrape_button = Button(root, text="Run Job Scrape", command=self.close_window)
        self.scrape_button.grid(row=3, column=0, columnspan=3, pady=10)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.handle_drop)

    def handle_drop(self, event):
        self.dropped_file_path = event.data.strip('{}')
        self.process_file()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("Excel files", "*.xlsx")])
        if file_path:
            self.dropped_file_path = file_path
            self.process_file()

    def process_file(self):
        if self.dropped_file_path:
            ext = os.path.splitext(self.dropped_file_path)[1].lower()
            print(f"📁 Loaded file: {self.dropped_file_path}")
            # In a full implementation you would parse and store imported jobs
            self.imported_jobs = []
        else:
            print("⚠️ No file selected.")

    def close_window(self):
        self.root.destroy()

    def init_driver(self):
        from scraper_core import init_driver
        from utils import handle_login
        driver = init_driver(headless=True)
        handle_login(driver)
        return driver
