# scraper_core.py
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
    extract_cid_and_time, force_dismiss_any_alert,
    handle_login
)

load_dotenv(dotenv_path=".env")

CALENDAR_URL = "http://inside.sockettelecom.com/events/calendar.php"
CUSTOMER_URL_TEMPLATE = "http://inside.sockettelecom.com/menu.php?coid=1&tabid=7&parentid=9&customerid={}"

USERNAME = os.getenv("UNITY_USER")
PASSWORD = os.getenv("PASSWORD")


def init_driver(headless=True):
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium import webdriver
    import os

    options = Options()
    options.page_load_strategy = 'eager'
    options.add_experimental_option("detach", True)

    # ✅ Suppress Chrome logging
    options.add_argument("--log-level=3")  # Only fatal errors
    options.add_argument("--silent")

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

    # 🪟 Hide command window on Windows
    service = Service("./chromedriver.exe")
    if os.name == "nt":
        service.creationflags = 0x08000000  # CREATE_NO_WINDOW

    return webdriver.Chrome(service=service, options=options)

def scrape_jobs(driver, mode="metadata", imported_jobs=None, selected_day=None, test_mode=False, test_limit=10):
    driver.get(CALENDAR_URL)
    wait = WebDriverWait(driver, 30)
    force_dismiss_any_alert(driver)

    try:
        week_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'week')]")))
        week_button.click()
        time.sleep(1)
        wait.until(EC.invisibility_of_element_located((By.ID, "spinner")))
        print("✅ Switched to Week View.")
    except Exception as e:
        print(f"⚠️ Could not switch to Week View: {e}")

    try:
        next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.fc-next-button")))
        next_button.click()
        time.sleep(1)
        wait.until(EC.invisibility_of_element_located((By.ID, "spinner")))
        print("✅ Navigated to Next Week.")
    except Exception as e:
        print(f"⚠️ Could not advance to next week: {e}")

    try:
        wait.until(lambda d: any("Residential Fiber Install" in el.text for el in d.find_elements(By.CSS_SELECTOR, "a.fc-time-grid-event")))
        print("✅ Jobs loaded and ready to scrape.")
    except:
        print("⚠️ No 'Residential Fiber Install' jobs detected.")

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
        print(f"{timestamp} {counter} - {cid} queued")

        results.append({
            "cid": cid,
            "name": name,
            "time": time_slot
        })

        if test_mode and len(results) >= test_limit:
            print("🔬 Test mode: Job limit reached. Exiting early.")
            break

    print(f"✅ Queued {len(results)} metadata jobs for multiprocessing.")
    return results


def process_job_entries(job):
    driver = init_driver()
    cid = job.get("cid")
    name = job.get("name")
    time_slot = job.get("time")
    customer_url = CUSTOMER_URL_TEMPLATE.format(cid)

    try:
        driver.get(customer_url)
        print(f"🌐 Loaded customer page for CID {cid}: {driver.current_url}")

        handle_login(driver)
        clear_first_time_overlays(driver)

        driver.switch_to.default_content()
        try:
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView")))
        except Exception:
            print(f"❌ Could not switch to MainView iframe for CID {cid}")
            print("🔎 Page source snippet:\n", driver.page_source[:1000])
            return None

        workorder_url, wo_number = None, None
        for attempt in range(3):
            workorder_url, wo_number = get_work_order_url(driver)
            if workorder_url:
                break
            print(f"🔁 Retry {attempt + 1}/3: No WO yet for CID {cid}")
            time.sleep(1)

        if not workorder_url:
            print(f"⚠️ Still no WO found for {cid} after retries.")
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
        print(f"❌ Failed to process job for CID {cid}: {e}")
        traceback.print_exc()
        return None

    finally:
        driver.quit()