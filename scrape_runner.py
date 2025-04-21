import os
import re
import time
import threading
from datetime import datetime, timedelta
from math import ceil
from concurrent.futures import ThreadPoolExecutor, as_completed

from scraper_core import scrape_jobs, init_driver, process_job_entries
from utils import export_txt, export_excel, generate_diff_report_and_return,  handle_login

def run_scrape(app):
    app.log("🚀 Starting full scrape...")
    t0 = time.time()

    selected_day = app.base_date.get()
    mode = app.scrape_mode_choice.get()

    try:
        driver = init_driver(headless=True)
        app.log("✅ Driver initialized.")
    except Exception as e:
        app.log(f"❌ Driver init failed: {e}")
        return

    raw_jobs = scrape_jobs(
        driver=driver,
        mode=mode,
        imported_jobs=None,
        selected_day=selected_day,
        test_mode=app.test_mode.get(),
        test_limit=app.test_limit.get(),
        log=app.log
    )
    driver.quit()

    total_jobs = len(raw_jobs)
    results = []
    incomplete = []
    completed = [0]
    app.progress_var.set(0)
    app.progress_bar["maximum"] = total_jobs
    app.counter_label.config(text=f"0 of {total_jobs} completed (0%)")
    lock = threading.Lock()

    def thread_task(job_batch):
        local_driver = init_driver(headless=True)
        handle_login(local_driver)
        try:
            for job in job_batch:
                result = process_job_entries(local_driver, job, log=app.log)
                with lock:
                    completed[0] += 1
                    app.progress_var.set(completed[0])
                    percent = (completed[0] / total_jobs) * 100
                    app.counter_label.config(text=f"{completed[0]} of {total_jobs} completed ({percent:.0f}%)")
                    if result:
                        results.append(result)
                    else:
                        job["error"] = "Failed to parse job details"
                        incomplete.append(job)
        finally:
            local_driver.quit()

    num_threads = app.worker_count.get()
    batch_size = ceil(len(raw_jobs) / num_threads)
    batches = [raw_jobs[i:i + batch_size] for i in range(0, len(raw_jobs), batch_size)]

    app.log("Processing Jobs...")

    with ThreadPoolExecutor(max_workers=num_threads) as executor:
        futures = [executor.submit(thread_task, batch) for batch in batches]
        for _ in as_completed(futures):
            pass

    output_dir = "Outputs"
    os.makedirs(output_dir, exist_ok=True)

    if mode == "day":
        base_date = datetime.strptime(selected_day, "%m/%d/%y")
        date_str = base_date.strftime("%m%d")
        txt_filename = os.path.join(output_dir, f"Jobs{date_str}.txt")
        excel_filename = os.path.join(output_dir, f"Jobs{date_str}.xlsx")
    else:
        base_date = datetime.strptime(selected_day, "%m/%d/%y")
        sunday = base_date - timedelta(days=(base_date.weekday() + 1) % 7)
        saturday = sunday + timedelta(days=6)
        range_str = f"{sunday.strftime('%m%d')}-{saturday.strftime('%m%d')}"
        txt_filename = os.path.join(output_dir, f"Jobs{range_str}.txt")
        excel_filename = os.path.join(output_dir, f"Jobs{range_str}.xlsx")

    export_txt(results, filename=txt_filename)
    if app.export_excel.get():
        export_excel(results, filename=excel_filename)

    start_date = sunday
    end_date = saturday

    output_tag = start_date.strftime("%m%d") if start_date == end_date else f"{start_date.strftime('%m%d')}-{end_date.strftime('%m%d')}"

    if incomplete:
        with open(os.path.join(output_dir, f"UnparsedJobs{output_tag}.txt"), "w") as f:
            for job in incomplete:
                f.write(f"{job.get('time', '?')} - {job.get('name', '?')} - {job.get('cid', '?')} - REASON: {job.get('error', 'Unknown')}\n")

    app.log(f"✅ Scrape complete. {len(results)} jobs saved.")
    app.log(f"⏱️ Duration: {time.time() - t0:.2f}s")

def run_update(app):
    app.log("🔄 Starting Get Updates Mode...")
    t0 = time.time()

    if not app.imported_jobs:
        app.log("⚠️ No baseline jobs to compare. Load a .txt or .xlsx file first.")
        return

    def infer_date_from_filename(path):
        filename = os.path.basename(path)
        match_day = re.search(r"Jobs(\d{4})\.txt", filename)
        match_range = re.search(r"Jobs(\d{4})-(\d{4})\.txt", filename)
        if match_day:
            parsed = datetime.strptime(match_day.group(1), "%m%d").replace(year=datetime.now().year)
            return parsed.date(), "day", match_day.group(1)
        elif match_range:
            start = datetime.strptime(match_range.group(1), "%m%d").replace(year=datetime.now().year)
            end = datetime.strptime(match_range.group(2), "%m%d").replace(year=datetime.now().year)
            return start.date(), "week", f"{match_range.group(1)}-{match_range.group(2)}"
        return None, None, None

    inferred_date, scrape_mode, filename_tag = infer_date_from_filename(app.dropped_file_path)
    if not inferred_date:
        app.log("⚠️ Could not infer date from filename. Use JobsMMDD or JobsMMDD-MMDD.")
        return

    selected_day = inferred_date.strftime("%m/%d/%y")

    try:
        driver = init_driver(headless=True)
        app.log("✅ Driver initialized.")
    except Exception as e:
        app.log(f"❌ Driver initialization failed: {e}")
        return

    scraped_metadata = scrape_jobs(
        driver=driver,
        mode=scrape_mode,
        imported_jobs=None,
        selected_day=selected_day,
        test_mode=app.test_mode.get(),
        test_limit=app.test_limit.get(),
        log=app.log
    )
    driver.quit()

    if not scraped_metadata:
        app.log("⚠️ No metadata jobs found.")
        return

    added, removed, moved = generate_diff_report_and_return(app.imported_jobs, scraped_metadata, filename_tag)
    jobs_to_scrape = added + [new for _, new in moved]

    app.log(f"🔍 Diff results — {len(added)} added, {len(removed)} removed, {len(moved)} moved.")

    if not jobs_to_scrape:
        app.log("✅ No new or moved jobs to scrape. Changes file has been written.")
        return

    results = []
    incomplete = []
    total_jobs = len(jobs_to_scrape)
    completed = [0]
    app.progress_var.set(0)
    app.progress_bar["maximum"] = total_jobs
    app.counter_label.config(text=f"0 of {total_jobs} completed (0%)")
    lock = threading.Lock()

    def thread_task(job_batch):
        local_driver = init_driver(headless=True)
        handle_login(local_driver, log=app.log)
        try:
            for job in job_batch:
                result = process_job_entries(local_driver, job, log=app.log)
                with lock:
                    completed[0] += 1
                    app.progress_var.set(completed[0])
                    percent = (completed[0] / total_jobs) * 100
                    app.counter_label.config(text=f"{completed[0]} of {total_jobs} completed ({percent:.0f}%)")
                    if result:
                        results.append(result)
                    else:
                        job["error"] = "Failed to parse job details"
                        incomplete.append(job)
        finally:
            local_driver.quit()

    num_threads = app.worker_count.get()
    batch_size = ceil(len(jobs_to_scrape) / num_threads)
    batches = [jobs_to_scrape[i:i + batch_size] for i in range(0, len(jobs_to_scrape), batch_size)]

    app.log(f"🛠 Scraping {total_jobs} changed jobs with {num_threads} threads...")
    with ThreadPoolExecutor(max_workers=num_threads) as executor:
        futures = [executor.submit(thread_task, batch) for batch in batches]
        for _ in as_completed(futures):
            pass

    output_dir = "Outputs"
    os.makedirs(output_dir, exist_ok=True)

    txt_filename = os.path.join(output_dir, f"Jobs{filename_tag}_updates.txt")
    excel_filename = os.path.join(output_dir, f"Jobs{filename_tag}_updates.xlsx")
    diff_path = os.path.join(output_dir, f"Changes{filename_tag}.txt")

    export_txt(results, filename=txt_filename)
    if app.export_excel.get():
        export_excel(results, filename=excel_filename)

    if incomplete:
        unparsed_file = os.path.join(output_dir, f"UnparsedJobs{filename_tag}.txt")
        with open(unparsed_file, "w") as f:
            for job in incomplete:
                f.write(f"{job.get('time', '?')} - {job.get('name', '?')} - {job.get('cid', '?')} - REASON: {job.get('error', 'Unknown')}\n")

    app.log(f"📊 Comparison written to: {diff_path}")
    app.log(f"✅ Updates processed. {len(results)} changes scraped.")
    app.log(f"⏱️ Total time: {time.time() - t0:.2f} seconds")