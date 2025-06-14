# utils.py
import os
import re
import time
import pickle
from tkinter import Tk, messagebox, simpledialog
import pandas as pd
import threading
from datetime import datetime
from collections import defaultdict
from dotenv import load_dotenv, set_key
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

HERE = os.path.abspath(os.path.dirname(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(HERE, os.pardir))

BASE_URL = "http://inside.sockettelecom.com/"
COOKIE_FILE = os.path.join(PROJECT_ROOT, "cookies.pkl")

OUTPUT_DIR = os.path.join(PROJECT_ROOT, "Outputs")
ENV_PATH = os.path.join(PROJECT_ROOT, ".env")

cookie_lock = threading.Lock()
class NoWOError(Exception):
    pass
class NoOpenWOError(Exception):
    pass

# Setup + Creds
def prompt_for_credentials():
    login_window = Tk()
    login_window.withdraw()

    USERNAME = simpledialog.askstring("Login", "Enter your USERNAME:", parent=login_window)
    PASSWORD = simpledialog.askstring("Login", "Enter your PASSWORD:", parent=login_window, show="*")

    login_window.destroy()
    return USERNAME, PASSWORD

def save_env_credentials(USERNAME, PASSWORD):
    dotenv_path = ENV_PATH
    if not os.path.exists(dotenv_path):
        with open(dotenv_path, "w") as f:
            f.write("")
    set_key(dotenv_path, "UNITY_USER", USERNAME)
    set_key(dotenv_path, "PASSWORD", PASSWORD)

def check_env_or_prompt_login(log=print):
    load_dotenv()
    username = os.getenv("UNITY_USER")
    password = os.getenv("PASSWORD")

    if username and password:
        log("✅ Loaded stored credentials.")
        return username, password

    while True:
        username, password = prompt_for_credentials()
        if not username or not password:
            messagebox.showerror("Login Cancelled", "Login is required to continue.")
            return None, None

        save_env_credentials(username, password)
        log("✅ Credentials captured and saved to .env.")
        return username, password


# Login + Session
def save_cookies(driver, filename=COOKIE_FILE):
    with cookie_lock:
        raw = driver.get_cookies()
    safe = []
    for c in raw:
        entry = {k: c[k] for k in ("name","value","domain","path","expiry","secure","httpOnly") if k in c}
        safe.append(entry)
    with open(filename, "wb") as f:
        pickle.dump(safe, f)

def load_cookies(driver, base_url=BASE_URL, filename=COOKIE_FILE):
    if not os.path.exists(filename):
        return False

    try:
        with cookie_lock:
            saved = pickle.load(open(filename, "rb"))
    except Exception:
        # Bad file → delete and force a fresh login next time
        os.remove(filename)
        return False

    # Navigate to host so cookies can be applied
    driver.get(base_url)
    host = base_url.split("://",1)[1].split("/",1)[0]

    injected = False
    for c in saved:
        c["domain"] = host
        try:
            driver.add_cookie(c)
            injected = True
        except Exception as e:
            # skip malformed cookies
            print(f"⚠️ Skipping cookie {c.get('name')}: {e}")

    if not injected:
        # No valid cookies → delete file to avoid repeated failures
        os.remove(filename)
        return False

    driver.refresh()
    return True


def handle_login(driver, log=print):
    driver.get(BASE_URL)

    if load_cookies(driver):
        if not login_failed(driver):
            log("✅ Session restored via cookies.")
            clear_first_time_overlays(driver)
            return
        else:
            log("⚠️ Cookie session invalid—deleting and retrying with credentials.")
            try: os.remove(COOKIE_FILE)
            except OSError: pass

    # === Fallback to credential login ===
    while "login.php" in driver.current_url or "Username" in driver.page_source:
        username, password = check_env_or_prompt_login(log)
        if not (username and password):
            log("❌ Login cancelled.")
            return

        perform_login(driver, username, password)
        WebDriverWait(driver, 10).until(
            lambda d: "menu.php" in d.current_url or "calendar" in d.page_source
        )

        if not login_failed(driver):
            save_cookies(driver)   # <-- write fresh cookies
            log("✅ Logged in with username/password.")
            return
        else:
            log("❌ Login failed. Re-prompting...")

def perform_login(driver, USERNAME, PASSWORD):
    driver.get("http://inside.sockettelecom.com/system/login.php")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))
    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.ID, "login").click()
    clear_first_time_overlays(driver)

def login_failed(driver):
    try:
        return (
            "login.php" in driver.current_url
            or "Username" in driver.page_source
            or "Invalid username or password" in driver.page_source
        )
    except Exception:
        return True  # if we can't read the page, assume failure


# Popups
def dismiss_alert(driver):
    try:
        for _ in range(3):
            WebDriverWait(driver, 1).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.dismiss()
    except:
        pass

def force_dismiss_any_alert(driver): 
    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.dismiss()
    except:
        pass

def clear_first_time_overlays(driver):
    # Dismiss alert if present
    try:
        WebDriverWait(driver, 0.5).until(EC.alert_is_present())
        driver.switch_to.alert.dismiss()
    except:
        pass

    # Known popup buttons
    buttons = [
        "//form[@id='valueForm']//input[@type='button']",
        "//form[@id='f']//input[@type='button']"
    ]
    for xpath in buttons:
        try:
            WebDriverWait(driver, 0.5).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        except:
            pass

    # Iframe switch loop
    for _ in range(3):
        try:
            WebDriverWait(driver, 0.5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView")))
            return
        except:
            time.sleep(0.25)
    print("❌ Could not switch to MainView iframe.")


# Scraping
def extract_cid_and_time(job_elem):
    try:
        text = job_elem.find_element(By.CLASS_NAME, "fc-title").text
        time_text = job_elem.find_element(By.CLASS_NAME, "fc-time").get_attribute("data-start")
        name, cid = re.findall(r"(.*?) - (\d{4}-\d{4}-\d{4})", text)[0]
        return cid.strip(), name.strip(), time_text.strip()
    except Exception as e:
        print(f"❌ Error extracting CID/time: {e}")
        return None, None, None

def get_contractor_assignments(driver):
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "contractorsection")))
        section_text = driver.find_element(By.CLASS_NAME, "contractorsection").text.strip()
        for line in section_text.splitlines():
            if " - (Primary" in line:
                return line.split(" - ")[0].strip()
        for line in section_text.splitlines():
            if "assigned to this work order" not in line and "Contractors" not in line:
                return line.split(" - ")[0].strip()
        return "Unknown"
    except Exception as e:
        print(f"❌ Could not find contractor name: {e}")
        return "Unknown"

def get_work_order_url(driver, log=None):
    try:
        dismiss_alert(driver)
        WebDriverWait(driver, 5, poll_frequency=0.05).until(
            lambda d: d.find_element(By.ID, "workShow").is_displayed()
        )

        # gather all install-fiber WOs
        rows = driver.find_elements(
            By.XPATH, "//div[@id='workShow']//table//tr[position()>1]"
        )
        install_wos = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 5:
                continue
            wo_num = cols[0].text.strip()
            desc   = cols[1].text.strip().lower()
            job_type = cols[2].text.strip().lower()
            status = cols[3].text.strip().lower()
            url    = cols[4].find_element(By.TAG_NAME, "a").get_attribute("href")
            if "install" in job_type and "fiber" in job_type:
                install_wos.append((int(wo_num), status, url))

        # 1) no install WOs at all
        if not install_wos:
            if log:
                log("⚠️ No install WOs found on page")
            raise NoWOError("No install work orders found")

        # 2) any “in process” WOs?
        in_process = [wo for wo in install_wos if "in process" in wo[1]]
        if in_process:
            best = max(in_process, key=lambda x: x[0])
            return best[2], best[0]

        # 3) any “open” or “scheduled” WOs?
        open_wos = [wo for wo in install_wos if "open" in wo[1] or "scheduled" in wo[1]]
        if open_wos:
            best = max(open_wos, key=lambda x: x[0])
            return best[2], best[0]

        # 4) install WOs exist but none are open/scheduled
        if log:
            log("⚠️ Found install WOs, but none are open or scheduled")
        raise NoOpenWOError("Install work orders exist but none are open")

    except (NoWOError, NoOpenWOError):
        # propagate our custom errors so caller can set job["error"] appropriately
        raise

    except Exception as e:
        # all other errors get your original fallback
        if log:
            log(f"❌ Error finding work order URL: {e}")
        else:
            print(f"❌ Error finding work order URL: {e}")
        return None, None

def get_job_type_and_address(driver):
    wait = WebDriverWait(driver, 10)

    # 1) Grab the address for city-inspection
    address = "Unknown"
    try:
        address = driver.find_element(
            By.XPATH,
            "//a[contains(@href, 'viewServiceMap')]"
        ).text.strip()
    except:
        pass

    # 2) Read the packageName <b> text for actual package info
    package_info = ""
    try:
        pkg_elem = driver.find_element(
            By.CSS_SELECTOR,
            ".packageName.text-indent b"
        )
        package_info = pkg_elem.text.strip()
    except:
        pass
    pkg_lower = package_info.lower()

    # 3) Grab the description (for connectorized check)
    desc_text = ""
    try:
        desc_text = driver.find_element(
            By.XPATH,
            "//td[contains(text(), 'Description:')]/following-sibling::td"
        ).text.strip().lower()
    except:
        pass

    # 4) OFFICIAL 5 Gig check 
    if "5 gig" in pkg_lower and "2.5" not in pkg_lower:
        is_connectorized = "connectorized" in desc_text
        has_phone        = "bundle" in pkg_lower or "phone" in pkg_lower

        if is_connectorized:
            job_type = "Connectorized 5 Gig Bundle" if has_phone else "Connectorized 5 Gig"
        else:
            job_type = "5 Gig Fiber Bundle"       if has_phone else "5 Gig Naked Fiber"

        return job_type, address

    # 5) Conversion Fallback: no package + Jefferson City → 5 Gig Conversion (until IT fixes the WO)
    if not package_info and "jefferson city" in address.lower():
        return "5 Gig Conversion", address

    # 6) 2.5G branch
    if "2.5" in pkg_lower:
        is_connectorized = "connectorized" in desc_text
        has_phone        = "bundle" in pkg_lower or "phone" in pkg_lower

        if is_connectorized:
            job_type = "Connectorized 2.5G Bundle" if has_phone else "Connectorized 2.5G"
        else:
            job_type = "2.5G Fiber Bundle"       if has_phone else "2.5G Naked Fiber"

        return job_type, address

    # 7) OTHERWISE: your existing Connectorized / Bundle / Naked logic
    is_connectorized = "connectorized" in desc_text
    has_phone        = "bundle" in pkg_lower or "phone" in pkg_lower

    if is_connectorized:
        job_type = "Connectorized Bundle" if has_phone else "Connectorized"
    else:
        job_type = "Fiber Bundle"       if has_phone else "Naked Fiber"

    return job_type, address


def extract_wo_date(driver):
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "scheduledEventList")))
        text = driver.find_element(By.ID, "scheduledEventList").text.strip()
        match = re.search(r"(\d{4}-\d{2}-\d{2})", next((line for line in text.splitlines() if "Fiber Install" in line), ""))
        if match:
            parsed_date = datetime.strptime(match.group(1), "%Y-%m-%d")
            return parsed_date.strftime("%#m-%#d-%y") if os.name == "nt" else parsed_date.strftime("%-m-%-d-%y")
    except Exception as e:
        print(f"⚠️ Could not extract WO date: {e}")
    return "Unknown"


# Time + Data
def get_output_tag(start, end): #File date stamp
    if start == end:
        return start.strftime("%m%d")
    return f"{start.strftime('%m%d')}-{end.strftime('%m%d')}"

def parse_time(t):
    try:
        return datetime.strptime(t.strip(), "%I:%M")
    except:
        return datetime.min

def get_sort_key(time_str):
    time_str = time_str.strip()
    hour = int(time_str.split(":")[0])
    # Treat hours before 6 as PM
    if hour in [1, 2, 3, 4, 5]:
        hour += 12
    return hour


# I/O
def generate_diff_report_and_return(imported_jobs, scraped_jobs, output_tag=""):

    def job_key(job):
        return f"{job.get('cid')}|{job.get('time')}"

    old_jobs = {job_key(job): job for job in imported_jobs}
    new_jobs = {job_key(job): job for job in scraped_jobs}

    old_cids = {job.get("cid"): job for job in imported_jobs}
    new_cids = {job.get("cid"): job for job in scraped_jobs}

    added = [job for k, job in new_jobs.items() if k not in old_jobs]
    removed = [job for k, job in old_jobs.items() if k not in new_jobs]
    moved = [(old_cids[cid], new_cids[cid]) for cid in old_cids if cid in new_cids and job_key(old_cids[cid]) != job_key(new_cids[cid])]

    # Write report
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename_tag = datetime.now().strftime("%m%d%H%M")
    report_path = os.path.join(OUTPUT_DIR, f"Changes{filename_tag}.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("=== Added ===\n")
        for job in added:
            f.write(f"{job.get('time', '?')} - {job.get('name', '?')} - {job.get('cid', '?')} - {job.get('type', '?')} - {job.get('address', '?')} - WO {job.get('wo', '?')}\n")
        f.write("\n=== Removed ===\n")
        for job in removed:
            f.write(f"{job.get('time', '?')} - {job.get('name', '?')} - {job.get('cid', '?')} - {job.get('type', '?')} - {job.get('address', '?')} - WO {job.get('wo', '?')}\n")
        f.write("\n=== Moved ===\n")
        for old, new in moved:
            f.write(f"CID {old['cid']}: {old['time']} → {new['time']} | {old['address']} → {new['address']}\n")

    return added, removed, moved

def export_txt(jobs, filename=None):
    jobs_by_company = defaultdict(lambda: defaultdict(list))
    for job in jobs:
        jobs_by_company[job["company"]][job["date"]].append(job)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_name = os.path.basename(filename) if filename else "Jobs.txt"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    with open(out_path, "w", encoding="utf-8") as f:
        for company, days in jobs_by_company.items():
            f.write(f"{company}\n\n")
            for date, entries in sorted(days.items()):
                f.write(f"{date}\n")
                for job in sorted(entries, key=lambda j: (get_sort_key(j['time']), j['name'].lower())):
                    f.write(f"{job['time']} - {job['name']} - {job['cid']} - {job['type']} - {job['address']} - WO {job['wo']}\n")
                f.write("\n")
            f.write("\n")

def export_excel(jobs, filename=None):
    jobs_by_company = defaultdict(lambda: defaultdict(list))
    for job in jobs:
        jobs_by_company[job["company"]][job["date"]].append(job)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_name = os.path.basename(filename) if filename else "Jobs.txt"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    rows = []
    for company, days in jobs_by_company.items():
        rows.append([company])
        for date, entries in sorted(days.items()):
            rows.append([date])
            for job in sorted(entries, key=lambda j: (get_sort_key(j['time']), j['name'].lower())):
                rows.append([
                    job['time'],
                    job['name'],
                    job['cid'],
                    job['type'],
                    job['address'],
                    f"WO {job['wo']}"
                ])
            rows.append([])
        rows.append([])

    # Write to Excel
    df = pd.DataFrame(rows)
    
    df.to_excel(out_path, index=False, header=False)

    # Autofit column widths
    wb = load_workbook(out_path)
    ws = wb.active
    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_length + 2
    wb.save(out_path)

def parse_imported_jobs(file_path):
    jobs = []
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".txt":
            with open(file_path, "r") as f:
                for line in f:
                    if " - " in line and "WO" in line:
                        parts = line.strip().split(" - ")
                        if len(parts) >= 6:
                            jobs.append({
                                "time": parts[0],
                                "name": parts[1],
                                "cid": parts[2],
                                "type": parts[3],
                                "address": parts[4],
                                "wo": parts[5].replace("WO ", "").strip()
                            })
        elif ext == ".xlsx":
            df = pd.read_excel(file_path, header=None)
            for row in df.itertuples(index=False):
                if len(row) >= 6 and isinstance(row[5], str) and "WO" in row[5]:
                    jobs.append({
                        "time": str(row[0]),
                        "name": str(row[1]),
                        "cid": str(row[2]),
                        "type": str(row[3]),
                        "address": str(row[4]),
                        "wo": row[5].replace("WO ", "").strip()
                    })
    except Exception as e:
        print(f"❌ Failed to parse imported file: {e}")
        return None
    return jobs
