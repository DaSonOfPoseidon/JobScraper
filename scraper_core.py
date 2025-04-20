# UPDATED scraper_core.py
import os
import time
import traceback
from datetime import datetime
from datetime import timedelta
from dateutil import parser as dateparser
from dotenv import load_dotenv, set_key
from tkinter import Tk, messagebox, simpledialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils import (
    dismiss_alert, clear_first_time_overlays,
    get_work_order_url, get_job_type_and_address,
    get_contractor_assignments, extract_wo_date,
    handle_login, force_dismiss_any_alert, extract_cid_and_time,
    perform_login
)

CALENDAR_URL = "http://inside.sockettelecom.com/events/calendar.php"
CUSTOMER_URL_TEMPLATE = "http://inside.sockettelecom.com/menu.php?coid=1&tabid=7&parentid=9&customerid={}"

def init_driver(headless=True):
    options = Options()
    options.page_load_strategy = 'eager'
    options.add_experimental_option("detach", True)

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

    try:
        driver_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "../chromedriver.exe"))
        service = Service(driver_path)
        service.creationflags = 0x08000000 
        return webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print("❌ ChromeDriver failed to initialize.")
        print("🔧 Possible issues: wrong ChromeDriver version, not executable, path mismatch.")
        raise e
    
def scrape_jobs(driver, mode="week", imported_jobs=None, selected_day=None, test_mode=False, test_limit=10, log=print):
    CALENDAR_URL = "http://inside.sockettelecom.com/events/calendar.php"
    driver.get(CALENDAR_URL)
    wait = WebDriverWait(driver, 20)

    force_dismiss_any_alert(driver)

    if "login.php" in driver.current_url or "Username" in driver.page_source:
        log("🔐 Login screen detected. Logging in...")
        handle_login(driver)
        driver.get(CALENDAR_URL)

    # Set View (Week or Day)
    try:
        if mode == "week":
            view_button = driver.find_element(By.CSS_SELECTOR, "button.fc-agendaWeek-button")
        else:
            view_button = driver.find_element(By.CSS_SELECTOR, "button.fc-agendaDay-button")
        view_button.click()
        time.sleep(1)
        wait.until(EC.invisibility_of_element_located((By.ID, "spinner")))
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

                time.sleep(1)
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
        wait.until(lambda d: any("Residential Fiber Install" in el.text for el in d.find_elements(By.CSS_SELECTOR, "a.fc-time-grid-event")))
        log("✅ Jobs loaded and ready to scrape.")
    except:
        log("⚠️ No 'Residential Fiber Install' jobs detected.")

    # Step 4: Extract metadata
    job_links = driver.find_elements(By.CSS_SELECTOR, "a.fc-time-grid-event")
    results = []
    counter = 0

    for link in job_links:
        if "Residential Fiber Install" not in link.text:
            continue

        cid, name, time_slot = extract_cid_and_time(link)
        if not cid:
            continue

        counter += 1
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        log(f"{timestamp} {counter} - {cid} queued")

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
        #log(f"🌐 Loading customer page for {cid}")

        clear_first_time_overlays(driver) #10 second wait?

        try:
            driver.switch_to.default_content()
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView")))
        except Exception:
            log(f"❌ Could not switch to MainView iframe for CID {cid}")
            return None

        workorder_url, wo_number = None, None
        for attempt in range(3):
            workorder_url, wo_number = get_work_order_url(driver)
            if workorder_url:
                break
            #log(f"🔁 Retry {attempt + 1}/3: No WO yet for CID {cid}")
            time.sleep(1)

        if not workorder_url:
            log(f"⚠️ No WO found for {cid}.")
            return None

        driver.get(workorder_url)
        dismiss_alert(driver)
        time.sleep(1)

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
        log(f"❌ Failed to process job for CID {cid}: {e}")
        traceback.print_exc()
        return None