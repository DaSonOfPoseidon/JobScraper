from multiprocessing import Manager
from concurrent.futures import ProcessPoolExecutor, as_completed
from scraper_core import scrape_jobs, process_job_entries
from utils import export_txt, export_excel
from gui import CalendarBuddyGUI
import tkinter as tk
from tkinterdnd2 import TkinterDnD
import time

MAX_WORKERS_DEFAULT = 3

if __name__ == "__main__":
    try:
        root = TkinterDnD.Tk()
        gui = CalendarBuddyGUI(root)
        root.mainloop()

        if gui.scrape_mode.get():
            print("📅 Starting Calendar Job Collection...")
            t0 = time.time()

            print("🧪 About to initialize driver...")

            try:
                driver = gui.init_driver()
            except Exception as e:
                print(f"❌ Driver initialization failed: {e}")
                import traceback
                traceback.print_exc()
                input("🔚 Press Enter to close...")
                exit()

            raw_jobs = scrape_jobs(
                driver=driver,
                mode="metadata",
                imported_jobs=gui.imported_jobs,
                selected_day=None,
                test_mode=gui.test_mode.get(),
                test_limit=gui.test_limit.get()
            )
            driver.quit()

            print(f"🛠 Processing {len(raw_jobs)} jobs with {gui.worker_count.get()} workers...")
            completed = 0
            with Manager() as manager:
                results = manager.list()
                with ProcessPoolExecutor(max_workers=gui.worker_count.get()) as executor:
                    futures = {executor.submit(process_job_entries, job): job for job in raw_jobs}
                    for future in as_completed(futures):
                        result = future.result()
                        completed += 1
                        print(f"✅ Completed {completed}/{len(raw_jobs)}")
                        if result:
                            results.append(result)

                export_txt(results)
                export_excel(results)
                print(f"✅ Done. Total Processed Jobs: {len(results)}")
                print(f"⏱️ Total time: {time.time() - t0:.2f} seconds")

    except Exception as e:
        import traceback
        traceback.print_exc()
        input("🔚 Crash occurred. Press Enter to exit...")
