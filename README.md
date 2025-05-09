# Job Scraper

This tool automates scraping of Residential Fiber Install jobs from the internal calendar and exports them as text or Excel reports.

## Overview

* **Full-Week Scrape**: Extracts all jobs for a 7‑day window, any day in that week can be selected.
* **Single-Day Scrape**: Extracts jobs for a specific date.
* **Get Updates**: (Planned) Compare against a provided list to report differences. *Currently non‑functional.*

## Key Features

* **Thread Count**: Adjustable up to 32; However, 8 threads on a mid‑range PC (16 GB RAM, modern i5) will fully utilize resources.
* **Excel Export**: Optionally output to `.xlsx` with auto‑adjusted column widths.
* **Calendar Date Selection**: GUI date picker allows users to select the date they want to parse from a calendar popup for ease of use.
* **Lightning Fast Parsing**: Parses 1000 real world jobs (including errors and load times) in ~6 minutes when configured to 8 threads.

## Typical Workflow

1. **Launch**:

   * (Preferred)
   * Return to the parent Folder (CalendarBuddy)
   * Run `run_scraper.bat`.

   * (Backup)
   * Run `main.py`
   * WARNING: This does not check for code updates or dependancies before running. (s**t may break)
2. **Configure**

   * Select Day or week radial button.
   * Select the date you'd like to parse.
   * Configure your threads (Default is 6)
   * Check the "Export Excel" Box (optional)
   * Press the "Run Job Scrape" Button.
2. **Scrape**:

   * Data is fetched via Selenium, processed by `scraper_core.py`.
   * Multi‑threaded runners in `scrape_runner.py` coordinate parallel browser sessions.
   * Each session navigates to the customer account, waits for work orders to load, identifies the install work order, and scrapes relevant details.
   * Returns data to the `scrape_runner.py` before navigating to the next customer.
3. **Export**:

   * Results saved as `.txt` by default.
   * Toggle in GUI to output `.xlsx`.

## Usage Notes

* **Get Updates** is under development and may be skipped.
* For older PCs, reduce threads to lower resource usage.
* Excel export requires `openpyxl`.
* Other import libraries are currently not mentioned here

## Support & Contribution

Report issues or contribute improvements in the parent repository. Setup and environment details are covered in the top‑level README. Happy scraping!

## Legal

This is a side‑project to streamline internal workflows. There is no warranty—use at your own risk. No parties involved can be held liable for data loss, corruption, or other damage arising from bugs.
Access to this tool requires proper network credentials; links and references are internal only and will not function externally.
