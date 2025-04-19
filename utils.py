# utils.py
import os
import re
import time
import pickle
import pandas as pd
from datetime import datetime
from collections import defaultdict
from dotenv import load_dotenv
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

load_dotenv(dotenv_path=".env")
USERNAME = os.getenv("UNITY_USER")
PASSWORD = os.getenv("PASSWORD")

def dismiss_alert(driver):
    try:
        for _ in range(3):
            WebDriverWait(driver, 1).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.dismiss()
            time.sleep(0.5)
    except:
        pass

def force_dismiss_any_alert(driver): 
    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.dismiss()
        time.sleep(0.25)
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

def get_sort_key(time_str):
    time_str = time_str.strip()
    hour = int(time_str.split(":")[0])
    # Treat hours 1, 3, 5 as PM
    if hour in [1, 2, 3, 4, 5]:
        hour += 12
    return hour

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

def get_work_order_url(driver):
    try:
        dismiss_alert(driver)
        WebDriverWait(driver, 5, poll_frequency=0.05).until(lambda d: d.find_element(By.ID, "workShow").is_displayed())
        rows = driver.find_elements(By.XPATH, "//div[@id='workShow']//table//tr[position()>1]")
        install_wos = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 5:
                continue
            wo_num = cols[0].text.strip()
            desc = cols[1].text.strip().lower()
            status = cols[3].text.strip().lower()
            url = cols[4].find_element(By.TAG_NAME, "a").get_attribute("href")
            if "install" in desc and "fiber" in desc:
                install_wos.append((int(wo_num), status, url))
        for wo in install_wos:
            if wo[1] == "in process":
                return wo[2], wo[0]
        if install_wos:
            return max(install_wos, key=lambda x: x[0])[2], max(install_wos, key=lambda x: x[0])[0]
        return None, None
    except Exception as e:
        print(f"❌ Could not find work order URL: {e}")
        return None, None

def get_job_type_and_address(driver):
    wait = WebDriverWait(driver, 10)
    job_type = "Unknown"
    address = "Unknown"
    is_connectorized = False
    has_phone = False
    try:
        desc_text = driver.find_element(By.XPATH, "//td[contains(text(), 'Description:')]/following-sibling::td").text.strip().lower()
        is_connectorized = "connectorized" in desc_text
    except: pass
    try:
        service_div = driver.find_element(By.CLASS_NAME, "servicesDiv").text.strip().lower()
        has_phone = "phone" in service_div
    except: pass
    job_type = ("Connectorized Bundle" if has_phone else "Connectorized") if is_connectorized else ("Fiber Bundle" if has_phone else "Naked Fiber")
    try:
        address = driver.find_element(By.XPATH, "//a[contains(@href, 'viewServiceMap')]").text.strip()
    except: pass
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

def parse_time(t):
    try:
        return datetime.strptime(t.strip(), "%I:%M")
    except:
        return datetime.min

def export_txt(jobs, filename="Outputs/Jobs.txt"):
    jobs_by_company = defaultdict(lambda: defaultdict(list))
    for job in jobs:
        jobs_by_company[job["company"]][job["date"]].append(job)

    with open(filename, "w") as f:
        for company, days in jobs_by_company.items():
            f.write(f"{company}\n\n")
            for date, entries in sorted(days.items()):
                f.write(f"{date}\n")
                for job in sorted(entries, key=lambda j: get_sort_key(j['time'])):
                    f.write(f"{job['time']} - {job['name']} - {job['cid']} - {job['type']} - {job['address']} - WO {job['wo']}\n")
                f.write("\n")
            f.write("\n")

def export_excel(jobs, filename="Outputs/Jobs.txt"):
    jobs_by_company = defaultdict(lambda: defaultdict(list))
    for job in jobs:
        jobs_by_company[job["company"]][job["date"]].append(job)

    rows = []
    for company, days in jobs_by_company.items():
        rows.append([company])
        for date, entries in sorted(days.items()):
            rows.append([date])
            for job in sorted(entries, key=lambda j: get_sort_key(j['time'])):
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
    df.to_excel(filename, index=False, header=False)

    # Autofit column widths
    wb = load_workbook(filename)
    ws = wb.active
    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_length + 2
    wb.save(filename)

def save_cookies(driver, filename="cookies.pkl"):
    with open(filename, "wb") as f:
        pickle.dump(driver.get_cookies(), f)

def load_cookies(driver, filename="cookies.pkl"):
    if not os.path.exists(filename): return False
    try:
        with open(filename, "rb") as f:
            cookies = pickle.load(f)
        driver.get("http://inside.sockettelecom.com/")
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.refresh()
        clear_first_time_overlays(driver)
        return True
    except Exception:
        if os.path.exists(filename): os.remove(filename)
        return False

def handle_login(driver):
    driver.get("http://inside.sockettelecom.com/")
    if load_cookies(driver):
        if "login.php" in driver.current_url or "USERNAME" in driver.page_source:
            perform_login(driver)
            time.sleep(2)
            save_cookies(driver)
        else:
            print("✅ Session restored via cookies.")
            clear_first_time_overlays(driver)
    else:
        perform_login(driver)
        time.sleep(2)
        save_cookies(driver)

def perform_login(driver, username=None, password=None):
    username = username or USERNAME
    password = password or PASSWORD

    driver.get("http://inside.sockettelecom.com/system/login.php")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))
    driver.find_element(By.NAME, "username").clear()
    driver.find_element(By.NAME, "username").send_keys(username)
    driver.find_element(By.NAME, "password").clear()
    driver.find_element(By.NAME, "password").send_keys(password)
    driver.find_element(By.ID, "login").click()

    clear_first_time_overlays(driver)

    try:
        # Wait up to 3s for the error message to appear
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Incorrect username or password')]"))
        )
        print("❌ Login failed (error message detected).")
        return False
    except TimeoutException:
        pass  # No error message found, continue

    if "login.php" in driver.current_url:
        print("❌ Login failed (still on login.php).")
        return False

    print("✅ Login complete and overlays cleared.")
    return True

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

def generate_diff_report_and_return(imported_jobs, scraped_jobs):
    from collections import defaultdict
    from datetime import datetime

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
    os.makedirs("Outputs", exist_ok=True)
    filename_tag = datetime.now().strftime("%m%d%H%M")
    with open(f"Outputs/Changes{filename_tag}.txt", "w") as f:
        f.write("=== Added ===\n")
        for job in added:
            f.write(f"{job['time']} - {job['name']} - {job['cid']} - {job['type']} - {job['address']} - WO {job['wo']}\n")
        f.write("\n=== Removed ===\n")
        for job in removed:
            f.write(f"{job['time']} - {job['name']} - {job['cid']} - {job['type']} - {job['address']} - WO {job['wo']}\n")
        f.write("\n=== Moved ===\n")
        for old, new in moved:
            f.write(f"CID {old['cid']}: {old['time']} → {new['time']} | {old['address']} → {new['address']}\n")

    return added, removed, moved
