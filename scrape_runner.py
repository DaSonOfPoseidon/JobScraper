import os
import time
import socket
import asyncio
from datetime import datetime, timedelta
from math import ceil

from scraper_core import scrape_jobs, init_playwright_page, process_job_entries
from utils import export_txt, export_excel, generate_changes_file, handle_login, OUTPUT_DIR, PROJECT_ROOT
from emailer import send_job_results
from spreader import run_process as run_spreader

INTERESTING_CODES = [429, 403, 503]

def handle_exports(app, results, txt_filename, excel_filename, unparsed_jobs=None, stats=None):
    files = []

    # 1) TXT export
    export_txt(results, filename=txt_filename)
    files.append(txt_filename)

    # 2) Excel export, if checked
    if app.export_excel.get():
        export_excel(results, filename=excel_filename)
        files.append(excel_filename)
    
    # 3) Add any extra attachments
    if unparsed_jobs:
        files.extend(unparsed_jobs)

    # 3) Email, if checked
    if app.send_email.get():
        date_range = app.base_date.get()
        send_job_results(files, date_range, stats)

async def run_scrape(app):
    is_update = bool(app.imported_jobs)
    app.log("🚀 Starting full scrape...")
    t0 = time.time()

    selected_day = app.base_date.get()
    mode = app.scrape_mode_choice.get()
    send_email = app.send_email.get()

    playwright, browser, context, page = await init_playwright_page(headless=True)

    # 2) explicitly perform login, with its own logging
    app.log("🔐 Attempting to log in…")
    try:
        await handle_login(page, app.log)
        app.log("✅ Login successful.")
    except Exception as e:
        app.log(f"❌ Login failed: {e}")
        browser.close()
        playwright.stop()
        return

    tA = time.time()
    raw_jobs = await scrape_jobs(
        page=page,
        mode=mode,
        imported_jobs=None,
        selected_day=selected_day,
        test_mode=app.test_mode.get(),
        test_limit=app.test_limit.get(),
        log=app.log
    )
    await page.close()
    await context.close()
    tB = time.time()
    print(f"Metadata Scrape took {tB-tA:.2f}s")

    total_jobs = len(raw_jobs)
    results = []
    incomplete = []
    completed = [0]
    app.progress_var.set(0)
    app.progress_bar["maximum"] = total_jobs
    app.counter_label.config(text=f"0 of {total_jobs} completed (0%)")
    lock = asyncio.Lock()

    async def worker(job_batch, idx):
        worker_context, worker_page = await init_playwright_page(browser=browser, playwright=playwright)
        worker_page.on("response", log_response)
        await handle_login(worker_page)
        try:
            for job in job_batch:
                if app.start_time is None:
                    app.start_time = time.perf_counter()

                try:
                    result = await process_job_entries(worker_page, job, log=print)
                except Exception as e:
                    job.setdefault("error", f"Failed: {e}")
                    result = None

                async with lock:
                    completed[0] += 1
                    app.progress_var.set(completed[0])
                    percent = (completed[0] / total_jobs) * 100
                    app.counter_label.config(text=f"{completed[0]} of {total_jobs} completed ({percent:.0f}%)")
                    app.jobs_done += 1
                    app.root.after(0, app.update_throughput)

                    if result:
                        results.append(result)
                    else:
                        job.setdefault("error", "Failed to parse job details")
                        incomplete.append(job)

                        app.log(f"Failed to parse {job.get('cid')}")
        finally:
            await worker_page.close()
            await worker_context.close()


    def log_response(response):
        if response.status in INTERESTING_CODES:
            print(f"\n--- POSSIBLE RATE LIMIT ---")
            print(f"URL: {response.url}")
            print(f"Status: {response.status}")
            print("Headers:")
            for k, v in response.headers.items():
                print(f"    {k}: {v}")
            print("---------------------------\n")

    def attach_logger(page, tag="worker"):
        def on_response(response):
            url = response.url
            status = response.status
            elapsed = response.timing['responseEnd'] - response.timing['requestStart'] \
                if hasattr(response, "timing") else None
            print(f"[{tag}] {url} status={status} elapsed={elapsed}ms")
        page.on("response", on_response)

    num_threads = max(1, app.worker_count.get())
    batch_size = max(1, ceil(len(raw_jobs) / num_threads))
    batches = [raw_jobs[i:i + batch_size] for i in range(0, len(raw_jobs), batch_size)]

    app.log("Processing Jobs...")

    await asyncio.gather(*(worker(batch, i) for i, batch in enumerate(batches)))

    await browser.close()
    await playwright.stop()

    output_dir = OUTPUT_DIR
    os.makedirs(output_dir, exist_ok=True)

    if mode == "day":
        base_date = datetime.strptime(selected_day, "%m/%d/%y")
        date_str = base_date.strftime("%m%d")
        txt_filename = os.path.join(output_dir, f"Jobs{date_str}.txt")
        excel_filename = os.path.join(output_dir, f"Jobs{date_str}.xlsx")
        start_date = end_date = base_date
    else:
        base_date = datetime.strptime(selected_day, "%m/%d/%y")
        sunday = base_date - timedelta(days=(base_date.weekday() + 1) % 7)
        saturday = sunday + timedelta(days=6)
        range_str = f"{sunday.strftime('%m%d')}-{saturday.strftime('%m%d')}"
        txt_filename = os.path.join(output_dir, f"Jobs{range_str}.txt")
        excel_filename = os.path.join(output_dir, f"Jobs{range_str}.xlsx")
        start_date = sunday
        end_date = saturday

    output_tag = start_date.strftime("%m%d") if start_date == end_date else f"{start_date.strftime('%m%d')}-{end_date.strftime('%m%d')}"

    unparsed_file = None
    if incomplete:
        unparsed_file = os.path.join(output_dir, f"UnparsedJobs{output_tag}.txt")
        with open(os.path.join(output_dir, f"UnparsedJobs{output_tag}.txt"), "w") as f:
            for job in incomplete:
                f.write(f"{job.get('time', '?')} - {job.get('name', '?')} - {job.get('cid', '?')} - REASON: {job.get('error', 'Unknown')}\n")

    elapsed = time.time() - t0
    minutes, seconds = divmod(elapsed, 60)
    num_threads = app.worker_count.get()
    mode = app.scrape_mode_choice.get()
    selected_day = app.base_date.get()
    total_jobs = len(results)
    failed_jobs = len(incomplete)
    avg_time = elapsed / total_jobs if total_jobs else 0
    start_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(t0))
    end_time_str = time.strftime("%Y-%m-%d %H:%M:%S")

    hostname = socket.gethostname()

    stats = (
        f"Stats for this run:\n"
        f"---------------------\n"
        f"Scrape Mode:     {mode} ({selected_day})\n"
        f"Threads Used:    {num_threads}\n"
        f"Total Jobs:      {total_jobs}\n"
        f"Failed/Unparsed: {failed_jobs}\n"
        f"Total Time:      {int(minutes)}m {int(seconds)}s\n"
        f"Avg Time/Job:    {avg_time:.2f} sec/job\n"
        f"Start Time:      {start_time_str}\n"
        f"End Time:        {end_time_str}\n"
        f"Host:            {hostname}\n"
    )

    handle_exports(app, results, txt_filename, excel_filename, [unparsed_file] if unparsed_file else None, stats)

    minutes, seconds = divmod(elapsed, 60)
    app.log(f"Scrape complete. {len(results)} jobs saved.")
    if unparsed_file:
        rel_unparsed = os.path.relpath(unparsed_file, PROJECT_ROOT)
        app.log(f"{len(incomplete)} unparsed jobs saved to {rel_unparsed}")

    if is_update:
        def same_day(job):
            try:
                # job["date"] is like "6-16-25"
                dt = datetime.strptime(job["date"], "%m-%d-%y")
                return dt.strftime("%m%d") == output_tag
            except Exception:
                return False

        filtered_old = [j for j in app.imported_jobs if same_day(j)]

        changes_filename = f"{output_tag}Changes.txt"
        changes_path = generate_changes_file(
            filtered_old,
            results,
            changes_filename
        )
        app.log(f"✅ Change summary written to {os.path.relpath(changes_path)}")

    if getattr(app, "run_spreader", None) and app.run_spreader.get():
        try:
            spread_file = run_spreader(txt_filename)
            if os.path.exists(spread_file):
                rel_path = os.path.relpath(spread_file, PROJECT_ROOT)
                app.log(f"(Experimental) Recommended spread saved to {rel_path}")
                if getattr(app, "show_approve_spread_popup", None):
                    app.show_approve_spread_popup(spread_file)  # simplified check since already checked above
            else:
                app.log(f"(Experimental) Spreader failed: {spread_file}")
        except Exception as e:
            app.log(f"(Experimental) Spreader crashed: {e}")
    
    app.log(f"⏱️ Duration: {int(minutes)}:{int(seconds):02d} ({time.time() - t0:.2f}s)")

