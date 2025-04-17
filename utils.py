# utils.py
import time
import re
import os
import pandas as pd
from collections import defaultdict
from dotenv import load_dotenv
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

load_dotenv(dotenv_path=".env")

USERNAME = os.getenv("UNITY_USER")
PASSWORD = os.getenv("PASSWORD")

def dismiss_alert(driver):
    try:
        for _ in range(3):  # Try up to 3 alerts
            WebDriverWait(driver, 1).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            #print(f"⚠️ Dismissing alert: {alert.text}")
            alert.dismiss()
            time.sleep(0.5)
    except:
        pass

def force_dismiss_any_alert(driver):
    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        #print(f"⚠️ Force dismissing alert: {alert.text}")
        alert.dismiss()
        time.sleep(0.25)
    except:
        pass

def clear_first_time_overlays(driver):
    try:
        WebDriverWait(driver, 2).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        #print(f"⚠️ Dismissing alert: {alert.text}")
        alert.dismiss()
        time.sleep(0.25)
    except:
        pass

    # Try to close Vision modal fast
    try:
        close_btn = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//form[@id='valueForm']//input[@type='button']"))
        )
        #print("⚠️ Closing Vision modal...")
        close_btn.click()
        time.sleep(0.25)
    except:
        pass

    # Try to close soulkiller popup fast
    try:
        soulkiller_close = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//form[@id='f']//input[@type='button']"))
        )
        print("Closing soulkiller popup...")
        soulkiller_close.click()
        time.sleep(0.25)
    except:
        pass

    # Try switching to iframe quickly with fallback loop
    for _ in range(4):  # Retry up to ~4x within ~2s total
        try:
            WebDriverWait(driver, 0.5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView")))
            #print("✅ Switched into MainView iframe.")
            return
        except:
            time.sleep(0.5)
    print("❌ Could not switch to MainView iframe.")

def extract_cid_and_time(job_elem):
    try:
        text = job_elem.find_element(By.CLASS_NAME, "fc-title").text
        time_text = job_elem.find_element(By.CLASS_NAME, "fc-time").get_attribute("data-start")
        name, cid = re.findall(r"(.*?) - (\d{4}-\d{4}-\d{4})", text)[0]
        return cid.strip(), name.strip(), time_text.strip()
    except Exception as e:
        print(f"\u274c Error extracting CID/time: {e}")
        return None, None, None

def get_contractor_assignments(driver):
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "contractorsection"))
        )
        section_elem = driver.find_element(By.CLASS_NAME, "contractorsection")
        section_text = section_elem.text.strip()
        # print(f"[DEBUG] contractorsection raw text:\n{section_text}")

        # Extract the contractor name (line containing 'Primary')
        lines = section_text.splitlines()
        for line in lines:
            if " - (Primary" in line:
                contractor_name = line.split(" - ")[0].strip()
                #print(f"[DEBUG] contractor_name parsed: {contractor_name}")
                return contractor_name

        # Fallback if 'Primary' not found, just use first non-header line
        for line in lines:
            if "assigned to this work order" not in line and "Contractors" not in line:
                contractor_name = line.split(" - ")[0].strip()
                #print(f"[DEBUG] contractor_name fallback: {contractor_name}")
                return contractor_name

        #print("[DEBUG] No contractor lines matched.")
        return "Unknown"

    except Exception as e:
        #print("[DEBUG] contractorsection not found or failed to load.")
        print(f"❌ Could not find contractor name: {e}")
        return "Unknown"

def get_work_order_url(driver):
    try:
        dismiss_alert(driver)

        # Faster polling
        wait = WebDriverWait(driver, 5, poll_frequency=0.05)
        wait.until(
            lambda d: d.find_element(By.ID, "workShow").is_displayed()
        )

        workorder_rows = driver.find_elements(By.XPATH, "//div[@id='workShow']//table//tr[position()>1]")

        if not workorder_rows:
            print("⚠️ Customer has no work orders.")
            return None, None

        install_wos = []
        for row in workorder_rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            wo_num = cols[0].text.strip()
            desc = cols[1].text.strip()
            status = cols[3].text.strip()
            url = cols[4].find_element(By.TAG_NAME, "a").get_attribute("href")

            desc_lower = desc.lower()
            if "install" in desc_lower and "fiber" in desc_lower:
                install_wos.append((int(wo_num), status, url))

        for wo in install_wos:
            if wo[1].lower() == "in process":
                return wo[2], wo[0]

        if install_wos:
            max_wo = max(install_wos, key=lambda x: x[0])
            return max_wo[2], max_wo[0]

        print("⚠️ No Fiber Install WOs found, even though customer has work orders.")
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

    # === Step 1: Check Work Order Description for Legacy vs Connectorized ===
    try:
        desc_elem = driver.find_element(By.XPATH, "//td[contains(text(), 'Description:')]/following-sibling::td")
        description_text = desc_elem.text.strip().lower()
        #print(f"[DEBUG] WO Description: {description_text}")
        is_connectorized = "connectorized" in description_text
    except Exception as e:
        print(f"❌ Failed to get WO description for connectorized check: {e}")

    # === Step 2: Check Services Section for Phone/Bundles ===
    try:
        wait.until(lambda d: d.find_element(By.CLASS_NAME, "servicesDiv").text.strip() != "")
        service_div = driver.find_element(By.CLASS_NAME, "servicesDiv").text.strip().lower()
        #print(f"[DEBUG] servicesDiv: {service_div}")
        has_phone = "phone" in service_div # add any other phone indicators here
    except Exception as e:
        print(f"❌ Failed to check service info for phone bundle: {e}")

    # === Final Job Type Decision ===
    if is_connectorized:
        job_type = "Connectorized Bundle" if has_phone else "Connectorized"
    else:
        job_type = "Fiber Bundle" if has_phone else "Naked Fiber"

    # === Address Detection ===
    try:
        address_elem = driver.find_element(By.XPATH, "//a[contains(@href, 'viewServiceMap')]")
        address = address_elem.text.strip()
        #print(f"[DEBUG] address found: {address}")
    except Exception as e:
        print(f"❌ get_job_type_and_address() - unexpected error getting address: {e}")

    #print(f"[DEBUG] Final job type: {job_type}")
    return job_type, address

def extract_wo_date(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "scheduledEventList"))
        )
        container = driver.find_element(By.ID, "scheduledEventList")
        text = container.text.strip()
        #print(f"[DEBUG] scheduledEventList text:\n{text}")

        # Split and scan only relevant install entries
        lines = text.splitlines()
        install_line = next((line for line in lines if "Fiber Install" in line), None)

        if not install_line:
            print("⚠️ No Residential Fiber Install event found in scheduledEventList.")
            return "Unknown"

        match = re.search(r"(\d{4}-\d{2}-\d{2})", install_line)
        if match:
            date_str = match.group(1)
            parsed_date = datetime.strptime(date_str, "%Y-%m-%d")
            formatted = parsed_date.strftime("%#m-%#d-%y") if os.name == "nt" else parsed_date.strftime("%-m-%-d-%y")
            #print(f"[DEBUG] parsed WO date: {formatted}")
            return formatted

        print("⚠️ Date not found in install line.")
        return "Unknown"

    except Exception as e:
        print(f"⚠️ Could not extract WO date: {e}")
        return "Unknown"

def save_cookies(driver, filename="cookies.pkl"):
    import pickle
    with open(filename, "wb") as f:
        pickle.dump(driver.get_cookies(), f)

def load_cookies(driver, filename="cookies.pkl"):
    import pickle

    if not os.path.exists(filename):
        return False

    try:
        with open(filename, "rb") as f:
            cookies = pickle.load(f)

        driver.get("http://inside.sockettelecom.com/")
        for cookie in cookies:
            driver.add_cookie(cookie)

        driver.refresh()

        # 🧹 Immediately clear overlays after refresh
        clear_first_time_overlays(driver)
        return True
    except (EOFError, pickle.UnpicklingError) as e:
        print(f"⚠️ Deleting corrupted cookie file: {filename}")
        os.remove(filename)
        return False

def handle_login(driver):
    driver.get("http://inside.sockettelecom.com/")

    if load_cookies(driver):
        print("🍪 Cookies loaded. Verifying session...")

        # Check if login is still required after cookies load
        if "login.php" in driver.current_url or "Username" in driver.page_source:
            print("⚠️ Cookie session invalid — login required again.")
            perform_login(driver)
            time.sleep(2)
            save_cookies(driver)
        else:
            print("✅ Session restored via cookies.")
            clear_first_time_overlays(driver)
    else:
        print("🔐 No cookies found. Performing login...")
        perform_login(driver)
        time.sleep(2)
        save_cookies(driver)

def perform_login(driver):
    driver.get("http://inside.sockettelecom.com/system/login.php")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))

    dismiss_alert(driver)  # <-- Just in case alert is up already

    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.ID, "login").click()

    dismiss_alert(driver)  # <-- In case alert pops right after clicking login

    # Wait until redirected
    WebDriverWait(driver, 10).until(lambda d: "menu.php" in d.current_url or "Dashboard" in d.page_source)

    clear_first_time_overlays(driver)  # All modals/panels/etc

    print("✅ Login complete and overlays cleared.")

def export_txt(jobs, filename="calendar_output.txt"):
    jobs_by_company = defaultdict(lambda: defaultdict(list))
    for job in jobs:
        jobs_by_company[job["company"]][job["date"]].append(job)
    with open(filename, "w") as f:
        for company, days in jobs_by_company.items():
            f.write(f"{company}\n\n")
            for date, entries in sorted(days.items()):
                f.write(f"{date}\n")
                for job in entries:
                    line = f"{job['time']} - {job['name']} - {job['cid']} - {job['type']} - {job['address']} - WO {job['wo']}"
                    f.write(line + "\n")
                f.write("\n")
            f.write("\n")

def export_excel(jobs, filename="calendar_output.xlsx"):
    from openpyxl.utils import get_column_letter
    from openpyxl import load_workbook

    jobs_by_company = defaultdict(lambda: defaultdict(list))
    for job in jobs:
        jobs_by_company[job["company"]][job["date"]].append(job)

    rows = []
    for company, days in jobs_by_company.items():
        rows.append([company])
        for date, entries in sorted(days.items()):
            rows.append([date])
            for job in entries:
                rows.append([
                    job['time'], job['name'], job['cid'],
                    job['type'], job['address'], f"WO {job['wo']}"
                ])
            rows.append([])
        rows.append([])

    # Write to Excel
    df = pd.DataFrame(rows)
    df.to_excel(filename, index=False, header=False)

    # Auto-adjust column widths
    wb = load_workbook(filename)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2  # Add padding

    wb.save(filename)
