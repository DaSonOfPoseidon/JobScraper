import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog
from tkinterdnd2 import DND_FILES
from tkcalendar import DateEntry
from datetime import datetime
from utils import parse_imported_jobs
from scrape_runner import run_scrape

class CalendarBuddyGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Calendar Buddy - Job Scraper")
        self.root.geometry("720x600")
        self.start_time = None
        self.jobs_done  = 0

        self.imported_jobs = None
        self.dropped_file_path = None
        self.scrape_mode_choice = tk.StringVar(value="week")
        self.test_mode = tk.BooleanVar(value=False)
        self.test_limit = tk.IntVar(value=10)
        self.worker_count = tk.IntVar(value=6)
        self.export_excel = tk.BooleanVar(value=False)
        self.send_email = tk.BooleanVar(value=False)
        self.run_spreader = tk.BooleanVar(value=True)
        self.base_date = tk.StringVar()

        # === File Input Section ===
        file_frame = ttk.LabelFrame(root, text="Import Job File")
        file_frame.pack(fill="x", padx=10, pady=(10, 5))

        self.label = tk.Label(file_frame, text="Drag and drop a .txt or .xlsx file here", height=2, bg="white", relief="sunken")
        self.label.pack(fill="x", padx=10, pady=5)
        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.handle_drop)

        self.file_label = ttk.Label(file_frame, text="No file loaded.", foreground="gray")
        self.file_label.pack(padx=10, pady=(0, 5))

        # === Settings Section ===
        settings_frame = ttk.LabelFrame(root, text="Run Settings")
        settings_frame.pack(fill="x", padx=10, pady=5)

        ttk.Checkbutton(settings_frame, text="Test Mode", variable=self.test_mode).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Label(settings_frame, text="Test Mode Limit:").grid(row=0, column=1, sticky="e")
        ttk.Spinbox(settings_frame, from_=1, to=500, textvariable=self.test_limit, width=5).grid(row=0, column=2, sticky="w")

        ttk.Label(settings_frame, text="Worker Threads:").grid(row=1, column=0, sticky="w", padx=10)
        ttk.Spinbox(settings_frame, from_=1, to=32, textvariable=self.worker_count, width=5).grid(row=1, column=1, sticky="w")

        ttk.Checkbutton(settings_frame, text="Export Excel", variable=self.export_excel).grid(row=1, column=2, sticky="e", padx=10)
        ttk.Checkbutton(settings_frame, text="Send Email", variable=self.send_email).grid(row=1, column=3, sticky="e", padx=10)
        ttk.Checkbutton(settings_frame, text="Run Spreader (experimental)", variable=self.run_spreader).grid(row=1, column=4, sticky="w", padx=10)

        ttk.Label(settings_frame, text="Calendar Date:").grid(row=2, column=0, sticky="w", padx=10, pady=(5, 0))
        DateEntry(settings_frame, textvariable=self.base_date, width=12).grid(row=2, column=1, sticky="w", pady=(5, 0))

        ttk.Radiobutton(settings_frame, text="Full Week", variable=self.scrape_mode_choice, value="week").grid(row=2, column=2, sticky="w")
        ttk.Radiobutton(settings_frame, text="Single Day", variable=self.scrape_mode_choice, value="day").grid(row=2, column=3, sticky="w")

        # === Action Buttons ===
        button_frame = ttk.Frame(root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Run Job Scrape", command=self.start_scrape_thread).pack(side="left", padx=20)

        # === Log Console ===
        log_frame = ttk.LabelFrame(root, text="Output Log")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, wrap="word", height=12, bg="#f8f8f8", font=("Courier", 10))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

        # === Progress Footer ===
        footer_frame = ttk.Frame(root)
        footer_frame.pack(fill="x", padx=10, pady=5)

        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(footer_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", expand=True, side="left", padx=(0, 10))

        self.throughput_label = ttk.Label(footer_frame, text="0.00 jobs/sec")
        self.throughput_label.pack(side="left", padx=(0,10))

        self.counter_label = ttk.Label(footer_frame, text="0 of 0 completed (0%)")
        self.counter_label.pack(side="right")

    def log(self, message):
        timestamp = time.strftime("[%H:%M:%S]")
        self.log_text.insert(tk.END, f"{timestamp} {message}\n")
        self.log_text.see(tk.END)

    def handle_drop(self, event):
        self.dropped_file_path = event.data.strip('{}')
        self.file_label.config(text=f"Loaded: {os.path.basename(self.dropped_file_path)}")
        self.log(f"📁 Loaded file: {self.dropped_file_path}")
        self.imported_jobs = parse_imported_jobs(self.dropped_file_path)
        self.log(f"✅ Parsed {len(self.imported_jobs)} imported jobs.")

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("Excel files", "*.xlsx")])
        if file_path:
            self.dropped_file_path = file_path
            self.file_label.config(text=f"Loaded: {os.path.basename(file_path)}")
            self.log(f"📁 Loaded file: {file_path}")
            self.imported_jobs = parse_imported_jobs(file_path)
            self.log(f"✅ Parsed {len(self.imported_jobs)} imported jobs.")

    def reset_throughput(self):
        self.start_time = None
        self.jobs_done  = 0
        self.throughput_label.config(text="0.00 jobs/sec (– s/job)")

    def update_throughput(self):
        elapsed = time.perf_counter() - self.start_time
        if not self.jobs_done or elapsed <= 0:
            txt = "0.00 jobs/sec (– s/job)"
        else:
            num_threads = self.worker_count.get()
            jps = self.jobs_done / elapsed
            spj = (elapsed * num_threads) / self.jobs_done
            txt = f"{jps:.2f} jobs/sec ({spj:.2f} s/job)"
        self.throughput_label.config(text=txt)

    def start_scrape_thread(self):
        self.reset_throughput()
        threading.Thread(target=run_scrape, args=(self,), daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = CalendarBuddyGUI(root)
    root.mainloop()
