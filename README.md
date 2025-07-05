# Job Scraper

This tool automates scraping, processing, and assignment of Residential Fiber Install jobs from the internal calendar system. It supports exporting reports, emailing results, and reassigning jobs to contractors based on configurable rules.

---

## Overview

- **Full-Week Scrape**: Extracts all Residential Fiber Install jobs for a selected 7-day week.
- **Single-Day Scrape**: Extracts jobs for a specific selected day.
- **Import Job File**: Allows importing `.txt` or `.xlsx` job lists to compare or update assignments.
- **Apply Spreader (Beta)**: Automatically reassigns jobs to contractors based on configurable geographic and capacity rules.
- **Email Integration**: Optionally sends scraped job reports via email with attachments.
- **Excel Export**: Export results as `.xlsx` with automatic column width adjustment.
- **Multi-threaded Scraping**: Supports configurable worker threads for faster scraping and assignment.
- **Test Mode**: Limits jobs processed for testing purposes.

---

## Features

- **GUI Interface** built with tkinter and drag-and-drop support.
- **Playwright Browser Automation** for scraping internal web calendar and customer job details.
- **Configurable Contractor Limits** and geographic priority zones for job assignment.
- **Robust Parsing** of various contractor schedule input formats.
- **Integrated Logging and Progress Bar** for tracking scraping and reassignment status.
- **Automatic Playwright Chromium Installation** and session persistence with saved login state.
- **Update Checking** (planned, minimal stub implemented).
- **CLI Support** for version checking and update via command line flags.

---

## Typical Workflow

1. **Launch**  
   - Run `main.py` or the bundled executable (if available).  
   - The GUI window opens for interaction.

2. **Configure Run**  
   - Drag and drop or browse to import job files (optional).  
   - Select calendar scrape mode: full week or single day.  
   - Pick calendar date via date picker.  
   - Set number of worker threads (default 6, max 32).  
   - Enable optional features: Export Excel, Send Email, Run Spreader (reassignment).

3. **Run Scrape**  
   - Click "Run Job Scrape".  
   - The tool logs in to the internal site automatically if credentials are provided in the .env.  
   - Scrapes job metadata from calendar and fetches detailed customer job info.  
   - Processes jobs concurrently using asyncio and Playwright's "Contexts".

4. **Export & Post-Processing**  
   - Saves results to `.txt` and optionally `.xlsx` files in the Outputs folder.  
   - If enabled, emails results to configured recipients.  
   - If enabled, runs the Spreader algorithm to reassign jobs according to rules and capacity.  
   - Prompts user before applying reassignment changes.

---

## Configuration & Environment

- **.env file** in `Misc` folder holds credentials (`UNITY_USER`, `PASSWORD`) and email SMTP settings.  
- **Spreader Config** is stored in `Misc/spreader_config.json`, with defaults embedded in code.  
- Email sending requires valid SMTP credentials and recipient addresses in `.env`.  
- Playwright Chromium is installed automatically if missing.  
- All output files are saved to the `Outputs` directory.

---

## Requirements

- Python 3.10+  
- Dependencies in `requirements.txt` including:  
  - playwright  
  - tkinterdnd2  
  - tkcalendar  
  - pandas, openpyxl  
  - python-dotenv  
  - tqdm  
  - RapidFuzz  

---

## Limitations & Notes

- The **Get Updates** feature (differencing new vs old job lists) is planned but currently non-functional.  
- The Spreader reassignment algorithm is in beta and requires user confirmation before applying changes.  
- The tool is designed for internal network use; URLs and credentials must have proper access.  
- Excel export requires `openpyxl` and may increase run time.  
- Playwright downloads ~100 MB on first run for Chromium.

---

## Development & Contribution

- The project uses modular Python scripts with asyncio and threading for concurrency.  
- Contributions and bug reports are welcome via the project's repository.  
- To run in development, install dependencies from `requirements.txt` and launch `main.py`.  
- Use `--version` CLI flags for version info

---

## Legal & Disclaimer

This is an internal side-project aimed at streamlining fiber install job workflows. No warranties are provided. Use at your own risk. The tool accesses proprietary internal systems and is not intended for external use.

---

## Contact

For questions or support, please submit an issue on GitHub.