import os
import time
import threading
import asyncio
import tkinter as tk
from tkinter import ttk, filedialog
from tkinterdnd2 import DND_FILES
from tkcalendar import DateEntry
from datetime import datetime
from tkinter import messagebox
from utils import parse_imported_jobs, assign_contractor
from scrape_runner import run_scrape
from spreader import parse_moved_jobs_from_spread
from scraper_core import init_playwright_page
from utils import handle_login, ensure_playwright, __version__, prompt_reassignment

class CalendarBuddyGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Job Scraper v{__version__}")
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
        ttk.Checkbutton(settings_frame, text="Run Spreader (beta)", variable=self.run_spreader).grid(row=1, column=4, sticky="w", padx=10)

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
        elapsed = time.perf_counter() - self.start_time if self.start_time else 0
        if not self.jobs_done or elapsed <= 0:
            txt = "0.00 jobs/sec (– s/job)"
        else:
            num_threads = self.worker_count.get()
            # jobs/sec
            jps = self.jobs_done / elapsed
            # avg sec/job (accounting for concurrency)
            spj = (elapsed * num_threads) / self.jobs_done
            # ETA
            total     = getattr(self, "scrape_total", self.progress_bar["maximum"])
            remaining = total - self.jobs_done
            eta_secs  = spj * remaining
            eta_str   = time.strftime("%M:%S", time.gmtime(eta_secs))

            txt = f"{jps:.2f} jobs/sec ({spj:.2f} s/job)  •  ETA {eta_str}"

        self.throughput_label.config(text=txt)

    def start_scrape_thread(self):
        def target():
            try:
                ensure_playwright()
                # Now Chromium is installed (or an exception was raised)
            except Exception as e:
                # If install failed, abort scraping
                print(f"Playwright setup failed: {e}")
                return

            # Proceed with async scraping
            try:
                asyncio.run(run_scrape(self))
            except Exception as e:
                print(f"Error in run_scrape: {e}")

        self.reset_throughput()
        threading.Thread(target=target, daemon=True).start()

    def show_approve_spread_popup(self, spread_file):
        if not self.run_spreader.get():
            return
        if messagebox.askyesno("Apply Spread Changes?",
                               "Apply contractor reassignments now?"):
            self.apply_spreader(spread_file)

    def apply_spreader(self, spread_file):
        jobs  = parse_moved_jobs_from_spread(spread_file)
        total = len(jobs)
        if not total:
            self.log("No moved jobs to reassign.")
            return

        workers = self.worker_count.get()  # <-- take from GUI
        self.log(f"🔁 Reassigning {total} jobs using {workers} workers…")

        # reset & repurpose the bar
        self.progress_bar["maximum"] = total
        self.progress_var.set(0)
        self.jobs_done   = 0
        self.start_time  = time.perf_counter()

        def _bg():
            asyncio.run(self._apply_spreader_async(jobs, total, workers))
        threading.Thread(target=_bg, daemon=True).start()


    async def _apply_spreader_async(self, jobs, total, workers):
        p, browser, context, _ = await init_playwright_page(headless=True)

        try:
            # 1) log in ONCE on page0
            page0 = await context.new_page()
            await handle_login(page0, log=self.log)

            # 2) create your pool of pages
            pages = [await context.new_page() for _ in range(workers)]
            sem   = asyncio.Semaphore(workers)

            async def do_assign(job):
                async with sem:
                    page = pages.pop()
                    try:
                        url = f"http://inside.sockettelecom.com/workorders/view.php?nCount={job['wo']}"
                        await page.goto(url)
                        await asyncio.sleep(1)  # give UI time
                        await assign_contractor(page, job["wo"], job["contractor"], log=self.log)
                    except Exception as e:
                        self.log(f"❌ WO {job['wo']} failed: {e}")
                    finally:
                        pages.append(page)
                    # ui update back on main thread
                    self.root.after(0, lambda: self._update_spreader_progress(total))

            # 3) dispatch them all
            results = await asyncio.gather(
                *(do_assign(job) for job in jobs),
                return_exceptions=True
            )
            # handle any stray exceptions
            for r in results:
                if isinstance(r, Exception):
                    self.log(f"⚠️ Error in worker: {r}")

        finally:
            await context.close()
            await browser.close()
            await p.stop()
            self.log("✅ Reassignment complete.")


    def _update_spreader_progress(self, total):
        self.jobs_done += 1
        self.progress_var.set(self.jobs_done)

        pct = self.jobs_done / total * 100
        self.counter_label.config(
            text=f"{self.jobs_done} of {total} ({pct:.0f}%)"
        )

        elapsed = time.perf_counter() - self.start_time
        # jobs/sec
        jps = (self.jobs_done / elapsed) if elapsed else 0
        # avg seconds per job (now truly concurrent)
        spj = (elapsed * self.worker_count.get()) / self.jobs_done if self.jobs_done else 0

        remaining = total - self.jobs_done
        eta_secs  = spj * remaining
        eta_str   = time.strftime("%M:%S", time.gmtime(eta_secs))

        self.throughput_label.config(
            text=f"{jps:.2f} jobs/sec ({spj:.2f} s/job)  •  ETA {eta_str}"
        )
    
if __name__ == "__main__":
    root = tk.Tk()
    app = CalendarBuddyGUI(root)
    root.mainloop()
