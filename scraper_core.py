# scraper_core.py
import os
import time
import traceback
from datetime import datetime
from datetime import timedelta
import platform
import shutil
from dateutil import parser as dateparser
from dotenv import load_dotenv, set_key
from tkinter import Tk, messagebox, simpledialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

from utils import (
    dismiss_alert, clear_first_time_overlays, NoWOError, NoOpenWOError,
    get_work_order_url, get_job_type_and_address,
    get_contractor_assignments, extract_wo_date,
    handle_login, force_dismiss_any_alert, extract_cid_and_time,
)

CALENDAR_URL = "http://inside.sockettelecom.com/events/calendar.php"
CUSTOMER_URL_TEMPLATE = "http://inside.sockettelecom.com/menu.php?coid=1&tabid=7&parentid=9&customerid={}"

def init_driver(headless: bool = True) -> webdriver.Chrome:
    # 1) Locate the browser binary
    chrome_bin = (
        os.environ.get("CHROME_BIN")
        or shutil.which("google-chrome")
        or shutil.which("chrome")
        or shutil.which("chromium-browser")
        or shutil.which("chromium")
    )
    if not chrome_bin and platform.system() == "Windows":
        # Probe standard Windows install paths
        for p in (
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ):
            if os.path.exists(p):
                chrome_bin = p
                break

    if not chrome_bin or not os.path.exists(chrome_bin):
        raise RuntimeError(
            "Chrome/Chromium binary not found—install it or set CHROME_BIN"
        )

    # 2) Build ChromeOptions
    opts = webdriver.ChromeOptions()
    opts.binary_location = chrome_bin
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")

    # performance tweaks:
    opts.page_load_strategy = "eager"
    opts.add_experimental_option("detach", False)       # or True only when debugging
    opts.add_experimental_option("prefs", {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.fonts": 2,
    })
    opts.add_argument("--blink-settings=imagesEnabled=false")
    opts.set_capability("unhandledPromptBehavior", "dismiss")

    # 3) Pick driver based on platform
    system = platform.system()
    arch = platform.machine().lower()
    if system == "Linux" and ("arm" in arch or "aarch64" in arch):
        # Raspberry Pi / ARM Linux → use distro’s chromedriver
        driver_path = "/usr/bin/chromedriver"
        if not os.path.exists(driver_path):
            raise RuntimeError("ARM chromedriver not found; please `apt install chromium-driver`")
        service = Service(driver_path)
    else:
        # x86 Linux or Windows → download via webdriver-manager
        service = Service(ChromeDriverManager().install())

    # 4) Launch
    return webdriver.Chrome(service=service, options=opts)
    
def scrape_jobs(driver, mode="week", imported_jobs=None, selected_day=None, test_mode=False, test_limit=10, log=print):
    handle_login(driver, log)
    driver.get(CALENDAR_URL)
    
    # Set View (Week or Day)
    try:
        if mode == "week":
            view_button = driver.find_element(By.CSS_SELECTOR, "button.fc-agendaWeek-button")
        else:
            view_button = driver.find_element(By.CSS_SELECTOR, "button.fc-agendaDay-button")
        view_button.click()
        WebDriverWait(driver, 30, poll_frequency=0.1).until(EC.invisibility_of_element_located((By.ID, "spinner")))
        log(f"✅ Switched to {'Week' if mode == 'week' else 'Day'} View.")
    except Exception as e:
        log(f"⚠️ Could not switch view: {e}")

    # Navigate to correct week/day
    if selected_day:
        try:
            raw_selected_date = dateparser.parse(selected_day).date()

            # If mode is week, round down to the Sunday of that week
            if mode == "week":
                weekday = raw_selected_date.weekday()  # 0 = Monday, 6 = Sunday
                days_to_sunday = (raw_selected_date.weekday() + 1) % 7
                target_date = raw_selected_date - timedelta(days=days_to_sunday)
            else:
                target_date = raw_selected_date


            def get_current_calendar_date():
                try:
                    header_text = driver.find_element(By.CSS_SELECTOR, ".fc-center h2").text.strip()
                    # Handle week ranges like "Apr 13 - 19, 2025"
                    if "-" in header_text:
                        parts = header_text.replace("–", "-").split("-")
                        if len(parts) == 2:
                            start_str = parts[0].strip()
                            end_str = parts[1].strip()
                            if "," not in end_str:
                                end_str += ", " + str(datetime.now().year)
                            return dateparser.parse(start_str).date()
                    return dateparser.parse(header_text, fuzzy=True).date()
                except Exception:
                    return None

            current_date = get_current_calendar_date()
            nav_tries = 0

            while current_date and current_date != target_date and nav_tries < 15:
                if current_date < target_date:
                    driver.find_element(By.CSS_SELECTOR, "button.fc-next-button").click()
                else:
                    driver.find_element(By.CSS_SELECTOR, "button.fc-prev-button").click()

                time.sleep(.25)
                current_date = get_current_calendar_date()
                nav_tries += 1

            if current_date == target_date:
                log(f"📆 Calendar now displaying {current_date}")
            else:
                log(f"⚠️ Could not align calendar to {target_date} after {nav_tries} tries.")

        except Exception as e:
            log(f"❌ Failed to navigate to selected day: {e}")

    # Step 3: Wait for jobs to load
    try:
        WebDriverWait(driver, 10).until(lambda d: any("Residential Fiber Install" in el.text for el in d.find_elements(By.CSS_SELECTOR, "a.fc-time-grid-event")))
        log("✅ Jobs loaded and ready to scrape.")
    except:
        log("⚠️ No 'Residential Fiber Install' jobs detected.")

    # Step 4: Extract metadata
    job_links = driver.find_elements(By.CSS_SELECTOR, "a.fc-time-grid-event")
    results = []
    counter = 0

    log("Scraping Calendar...")
    for link in job_links:
        if "Residential Fiber Install" not in link.text:
            continue

        cid, name, time_slot = extract_cid_and_time(link)
        if not cid:
            continue

        counter += 1
        #log(f"{counter} - {cid} queued")

        results.append({
            "cid": cid,
            "name": name,
            "time": time_slot
        })

        if test_mode and len(results) >= test_limit:
            log("🔬 Test mode: Job limit reached. Exiting early.")
            break

    log(f"✅ Queued {len(results)} jobs for processing.")
    return results

def process_job_entries(driver, job, log=print):
    cid = job.get("cid")
    name = job.get("name")
    time_slot = job.get("time")
    customer_url = CUSTOMER_URL_TEMPLATE.format(cid)

    try:
        driver.get(customer_url)
        clear_first_time_overlays(driver)

        try:
            driver.switch_to.default_content()
            WebDriverWait(driver, 5, poll_frequency=0.1).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView")))
        except Exception:
            job["error"] = f"Failed to load customer page for {cid}"
            return None

        def _wait_for_work_order(driver):
            try:
                return get_work_order_url(driver, log=None)
            except (NoWOError, NoOpenWOError):
                raise
            except Exception:
                return False

        try:
            workorder_url, wo_number = WebDriverWait(driver, 5, poll_frequency=0.1).until(
                _wait_for_work_order
            )
        except (NoWOError, NoOpenWOError) as e:
            job["error"] = str(e)
            return None
        except TimeoutException:
            workorder_url, wo_number = None, None
        
        if not workorder_url:
            job["error"] = job.get("error", " No WO found")
            return None

        driver.get(workorder_url)
        dismiss_alert(driver)

        job_type, address = get_job_type_and_address(driver)
        contractor_info = get_contractor_assignments(driver)
        job_date = extract_wo_date(driver)

        return {
            "company": contractor_info,
            "date": job_date,
            "time": time_slot,
            "name": name,
            "cid": cid,
            "type": job_type,
            "address": address,
            "wo": wo_number
        }

    except Exception as e:
        log(f"Couldn't parse {cid}: {e}")
        traceback.print_exc()
        return None