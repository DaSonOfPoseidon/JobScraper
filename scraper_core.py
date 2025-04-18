# UPDATED scraper_core.py
import os
import time
import traceback
from datetime import datetime
from dotenv import load_dotenv
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
    handle_login, force_dismiss_any_alert, extract_cid_and_time
)

load_dotenv(dotenv_path=".env")

CALENDAR_URL = "http://inside.sockettelecom.com/events/calendar.php"
CUSTOMER_URL_TEMPLATE = "http://inside.sockettelecom.com/menu.php?coid=1&tabid=7&parentid=9&customerid={}"

USERNAME = os.getenv("UNITY_USER")
PASSWORD = os.getenv("PASSWORD")

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
        service = Service("./chromedriver.exe")
        service.creationflags = 0x08000000  # CREATE_NO_WINDOW
        return webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print("❌ ChromeDriver failed to initialize.")
        print("🔧 Possible issues: wrong ChromeDriver version, not executable, path mismatch.")
        raise e
    
def scrape_jobs(driver, mode="metadata", imported_jobs=None, selected_day=None, test_mode=False, test_limit=10, log=print):
    driver.get(CALENDAR_URL)
    wait = WebDriverWait(driver, 30)
    force_dismiss_any_alert(driver)

    if "login.php" in driver.current_url or "Username" in driver.page_source:
        log("🔐 Login screen detected before clicking Week button.")
        handle_login(driver)
        driver.get(CALENDAR_URL)

    try:
        week_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'week')]")))
        week_button.click()
        time.sleep(1)
        wait.until(EC.invisibility_of_element_located((By.ID, "spinner")))
        log("✅ Switched to Week View.")
    except Exception as e:
        log(f"⚠️ Could not switch to Week View: {e}")

    try:
        next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.fc-next-button")))
        next_button.click()
        time.sleep(1)
        wait.until(EC.invisibility_of_element_located((By.ID, "spinner")))
        log("✅ Navigated to Next Week.")
    except Exception as e:
        log(f"⚠️ Could not advance to next week: {e}")

    try:
        wait.until(lambda d: any("Residential Fiber Install" in el.text for el in d.find_elements(By.CSS_SELECTOR, "a.fc-time-grid-event")))
        log("✅ Jobs loaded and ready to scrape.")
    except:
        log("⚠️ No 'Residential Fiber Install' jobs detected.")

    job_links = driver.find_elements(By.CSS_SELECTOR, "a.fc-time-grid-event")
    results = []
    counter = 0

    for link in job_links:
        if "Residential Fiber Install" not in link.text:
            continue

        cid, name, time_slot = extract_cid_and_time(link)
        if not cid:
            continue

        # Allow duplicates in scraping - address matching is unavailable at this stage
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

    log(f"✅ Queued {len(results)} metadata jobs for processing.")
    return results

def process_job_entries(driver, job, log=print):
    cid = job.get("cid")
    name = job.get("name")
    time_slot = job.get("time")
    customer_url = CUSTOMER_URL_TEMPLATE.format(cid)

    try:
        driver.get(customer_url)
        log(f"🌐 Loading customer page for {cid}")

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
            log(f"🔁 Retry {attempt + 1}/3: No WO yet for CID {cid}")
            time.sleep(1)

        if not workorder_url:
            log(f"⚠️ Still no WO found for {cid} after retries.")
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
