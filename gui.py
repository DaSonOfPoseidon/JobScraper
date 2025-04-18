# gui.py
import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Button, Text, Scrollbar, END
from tkinterdnd2 import DND_FILES
from concurrent.futures import ThreadPoolExecutor, as_completed
from math import ceil

from utils import parse_imported_jobs, handle_login, export_txt, export_excel, load_cookies, save_cookies
from scraper_core import scrape_jobs, process_job_entries, init_driver

MAX_WORKERS_DEFAULT = 3

class CalendarBuddyGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Calendar Buddy")
        self.root.geometry("600x500")
        self.root.configure(bg="#f0f0f0")

        self.scrape_mode = tk.StringVar(value="new")
        self.test_mode = tk.BooleanVar(value=False)
        self.test_limit = tk.IntVar(value=10)
        self.worker_count = tk.IntVar(value=MAX_WORKERS_DEFAULT)
        self.imported_jobs = None
        self.dropped_file_path = None

        self.label = Label(root, text="Drag and drop a .txt or .xlsx file here", width=60, height=4, bg="white", relief="ridge")
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

        self.scrape_button = Button(root, text="Run Job Scrape", command=self.start_scrape_thread)
        self.scrape_button.grid(row=3, column=0, columnspan=3, pady=10)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.handle_drop)

        self.log_text = Text(root, wrap="word", height=12, width=70, bg="#f8f8f8")
        self.log_text.grid(row=4, column=0, columnspan=3, padx=10, pady=(5, 0))

        self.scrollbar = Scrollbar(root, command=self.log_text.yview)
        self.scrollbar.grid(row=4, column=3, sticky="ns", pady=(5, 0))
        self.log_text.config(yscrollcommand=self.scrollbar.set)

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
            self.log(f"📁 Loaded file: {self.dropped_file_path}")
            self.imported_jobs = parse_imported_jobs(self.dropped_file_path)
            if self.imported_jobs is None:
                self.log("❌ Failed to parse jobs.")
            else:
                self.log(f"✅ Parsed {len(self.imported_jobs)} imported jobs.")
        else:
            self.log("⚠️ No file selected.")

    def log(self, message):
        timestamp = time.strftime("[%H:%M:%S]")
        self.log_text.insert(END, f"{timestamp} {message}\n")
        self.log_text.see(END)

    def start_scrape_thread(self):
        threading.Thread(target=self.run_scrape, daemon=True).start()

    def run_scrape(self):
        self.log("📅 Starting Calendar Job Collection...")
        t0 = time.time()
        self.log("🧪 About to initialize driver...")

        try:
            shared_driver = init_driver(headless=True)
            self.log("✅ Driver initialized.")
        except Exception as e:
            self.log(f"❌ Driver initialization failed: {e}")
            import traceback
            traceback.print_exc()
            return

        raw_jobs = scrape_jobs(
            driver=shared_driver,
            mode=self.scrape_mode.get(),
            imported_jobs=self.imported_jobs,
            selected_day=None,
            test_mode=self.test_mode.get(),
            test_limit=self.test_limit.get(),
            log=self.log
        )

        shared_driver.quit()

        if not raw_jobs:
            self.log("⚠️ No jobs found. Exiting.")
            return

        results = []
        lock = threading.Lock()
        total_jobs = len(raw_jobs)
        completed = [0]

        def thread_task(job_batch):
            local_driver = init_driver(headless=True)
            if not load_cookies(local_driver):
                handle_login(local_driver)
                save_cookies(local_driver)
            try:
                for job in job_batch:
                    result = process_job_entries(local_driver, job, log=self.log)
                    with lock:
                        completed[0] += 1
                        self.log(f"✅ Completed {completed[0]}/{total_jobs}")
                        if result:
                            results.append(result)
            except Exception as e:
                self.log(f"❌ Thread batch failed: {e}")
                import traceback
                traceback.print_exc()
            finally:
                local_driver.quit()

        num_threads = self.worker_count.get()
        batch_size = ceil(len(raw_jobs) / num_threads)
        batches = [raw_jobs[i:i + batch_size] for i in range(0, len(raw_jobs), batch_size)]

        self.log(f"🛠 Processing {total_jobs} jobs with {num_threads} threads (batch size = {batch_size})...")
        with ThreadPoolExecutor(max_workers=num_threads) as executor:
            futures = [executor.submit(thread_task, batch) for batch in batches]
            for _ in as_completed(futures):
                pass

        export_txt(results)
        export_excel(results)

        self.log(f"✅ Done. Total Processed Jobs: {len(results)}")
        self.log(f"⏱️ Total time: {time.time() - t0:.2f} seconds")
