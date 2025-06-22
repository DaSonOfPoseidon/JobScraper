# utils.py
import os
import re
from tkinter import Tk, messagebox, simpledialog
import pandas as pd
from datetime import datetime
from collections import defaultdict
from dotenv import load_dotenv, set_key
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from playwright.async_api import TimeoutError as PlaywrightTimeout


HERE = os.path.abspath(os.path.dirname(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(HERE, os.pardir))

BASE_URL = "http://inside.sockettelecom.com/"

OUTPUT_DIR = os.path.join(PROJECT_ROOT, "Outputs")
ENV_PATH = os.path.join(PROJECT_ROOT, ".env")

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
async def handle_login(page, log=print):
    await page.goto("http://inside.sockettelecom.com/")
    # If already logged in:
    if "login.php" not in page.url:
        log("✅ Session restored with stored state.")
        await clear_first_time_overlays(page)
        return

    # Otherwise, login with creds
    user, pw = check_env_or_prompt_login(log)
    await page.goto("http://inside.sockettelecom.com/system/login.php")
    await page.fill("input[name='username']", user)
    await page.fill("input[name='password']", pw)
    await page.click("#login")
    await page.wait_for_selector("iframe#MainView", timeout=10_000)
    await clear_first_time_overlays(page)
    # Save state for next run
    await page.context.storage_state(path="state.json")
    log("✅ Logged in via credentials.")

# Browser Interaction
async def clear_first_time_overlays(page):
    selectors = [
        'xpath=//input[@id="valueForm1" and @type="button"]',
        'xpath=//input[@value="Close This" and @type="button"]',
        'xpath=//form[starts-with(@id,"valueForm")]//input[@type="button"]',
        'xpath=//form[@id="f"]//input[@type="button"]'
    ]
    for sel in selectors:
        while True:
            try:
                btn = await page.wait_for_selector(sel, timeout=500)
                await btn.click()
                await page.wait_for_timeout(200)
            except PlaywrightTimeout:
                break

async def extract_cid_and_time(link, text):
    try:
        # Split lines
        lines = text.strip().split("\n")
        if len(lines) < 3:
            return None, None, None
        time_slot = lines[0].strip()
        # lines[1] = job type/area, usually not needed here
        third_line = lines[2].strip()
        # Now split the third line by ' - '
        parts = third_line.split(" - ")
        if len(parts) < 3:
            return None, None, None
        name = parts[0].strip()
        cid = parts[1].strip()
        # Optional: parse order #
        order_match = re.search(r'Order\s*#:(\d+)', parts[2])
        order_num = order_match.group(1) if order_match else None
        return cid, name, time_slot
    except Exception as e:
        print(f"❌ Error extracting CID/time: {e}")
        return None, None, None

async def get_contractor_assignments(page):
    try:
        # Wait for contractor section to show up
        await page.wait_for_selector(".contractorsection", timeout=5000)
        section_elem = await page.query_selector(".contractorsection")
        if not section_elem:
            return "Unknown"
        section_text = (await section_elem.inner_text()).strip()
        lines = section_text.splitlines()
        for line in lines:
            if " - (Primary" in line:
                return line.split(" - ")[0].strip()
        for line in lines:
            if "assigned to this work order" not in line and "Contractors" not in line:
                return line.split(" - ")[0].strip()
        return "Unknown"
    except Exception as e:
        print(f"❌ Could not find contractor name: {e}")
        return "Unknown"

async def get_work_order_url(frame, log=print):
    """
    Find newest (highest-numbered) in-process Fiber Install WO on customer page.
    Raises NoWOError or NoOpenWOError if not found.
    Returns (absolute_url, wo_number).
    """
    try:
        # Wait for the work order table to load (more general selector)
        try:
            await frame.wait_for_selector("#custWork #workShow table", timeout=10_000)
        except Exception:
            log("❌ Work Orders table not found inside frame!")
            raise NoWOError("Work Orders table not found!")
        rows = await frame.query_selector_all("#custWork #workShow table tr")
        if not rows:
            log("❌ No work order rows found inside table.")
            raise NoWOError("No work order rows found!")

        matches = []
        for row in rows:
            tds = await row.query_selector_all("td")
            if len(tds) < 5:
                continue

            first = (await tds[0].inner_text()).strip()
            # skip header or invalid rows
            if first == "#" or not first.isdigit():
                continue

            wo_number = int(first)
            job_type = (await tds[2].inner_text()).strip()
            status   = (await tds[3].inner_text()).strip()
            # Only match "Fiber Install" AND "In Process"
            if all(s in job_type for s in ["Fiber", "Install"]) and "In Process" in status:
                link_td = await tds[4].query_selector("a")
                href = await link_td.get_attribute('href') if link_td else None
                matches.append((wo_number, href))

        if not matches:
            # Fallback: any "Fiber Install", even if not "In Process"
            for row in rows:
                tds = await row.query_selector_all("td")
                if len(tds) < 5:
                    continue
                job_type = (await tds[2].inner_text()).strip()
                if all(s in job_type for s in ["Fiber", "Install"]):
                    raise NoOpenWOError("No open (In Process) Fiber Install WO found.")
            raise NoWOError("No Fiber Install WO found at all.")

        # Return WO with the highest number (most recent)
        wo_number, url = max(matches, key=lambda x: x[0])
        # Make sure URL is absolute
        if url and url.startswith("/"):
            url = "http://inside.sockettelecom.com" + url
        return url, wo_number

    except NoWOError:
        raise
    except NoOpenWOError:
        raise
    except Exception as e:
        log(f"❌ Error in get_work_order_url: {e}")
        raise NoWOError(str(e))

async def get_job_type_and_address(page):
    # 1) Grab the address for city-inspection
    address = "Unknown"
    try:
        addr_elem = await page.query_selector("a[href*='viewServiceMap']")
        if addr_elem:
            address = (await addr_elem.inner_text()).strip()
    except Exception:
        pass

    # 2) Read the packageName <b> text for actual package info
    package_info = ""
    try:
        pkg_elem = await page.query_selector(".packageName.text-indent b")
        if pkg_elem:
            package_info = (await pkg_elem.inner_text()).strip()
    except Exception:
        pass
    pkg_lower = package_info.lower()

    # 3) Grab the description (for connectorized check)
    desc_text = ""
    try:
        desc_elem = await page.query_selector("td.detailHeader:has-text('Description:') + td.detailData")
        if desc_elem:
            desc_text = (await desc_elem.inner_text()).strip().lower()
        else:
            # fallback for XPath
            desc_elem = await page.query_selector("//td[contains(text(), 'Description:')]/following-sibling::td[1]")
            if desc_elem:
                desc_text = (await desc_elem.inner_text()).strip().lower()
    except Exception:
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

async def extract_wo_date(page):
    try:
        # Wait for the scheduled event section to load
        await page.wait_for_selector("#scheduledEventList", timeout=10_000)
        text = (await page.locator("#scheduledEventList").inner_text()).strip()
        # Find line with "Fiber Install"
        fiber_line = next((line for line in text.splitlines() if "Fiber Install" in line), "")
        match = re.search(r"(\d{4}-\d{2}-\d{2})", fiber_line)
        if match:
            parsed_date = datetime.strptime(match.group(1), "%Y-%m-%d")
            # Windows: %#m, Linux/mac: %-m
            fmt = "%#m-%#d-%y" if os.name == "nt" else "%-m-%-d-%y"
            return parsed_date.strftime(fmt)
    except Exception as e:
        print(f"⚠️ Could not extract WO date: {e}")
    return "Unknown"

async def assign_contractor(page, wo_number, desired_contractor_full, log=print):  # COMPANY ASSIGNMENT
    try:
        # 🧠 Trigger the assignment UI via JavaScript
        await page.evaluate(f"assignContractor('{wo_number}');")
        await page.wait_for_selector("#ContractorID", timeout=10_000)
        await page.wait_for_timeout(500)  # let modal settle

        # ✅ Get current contractor assignment from the page
        contractor_texts = [
            text.strip()
            for text in await page.locator("b").all_inner_texts()
        ]
        current_contractor = None
        for text in contractor_texts:
            if " - (Primary" in text:
                current_contractor = text.split(" - ")[0].strip()
                break

        if current_contractor == desired_contractor_full:
            log(f"✅ Contractor '{current_contractor}' already assigned to WO #{wo_number}")
            return

        log(f"🧹 Reassigning from '{current_contractor}' → '{desired_contractor_full}'")

        # 🧽 Remove currently assigned contractors
        remove_links = page.locator("a", has_text="Remove")
        count = await remove_links.count()
        for i in range(count):
            try:
                link = remove_links.nth(i)
                await link.scroll_into_view_if_needed()
                await link.click()
                await page.wait_for_timeout(500)
            except Exception as e:
                log(f"❌ Could not remove contractor: {e}")

        # 🏷️ Assign the new contractor
        contractor_dropdown = page.locator("#ContractorID")
        await contractor_dropdown.select_option(label=desired_contractor_full)
        role_dropdown = page.locator("#ContractorType")
        await role_dropdown.select_option(label="Primary")

        # Hide any modal overlays if needed (FileList)
        try:
            file_list_elem = page.locator("#FileList")
            if await file_list_elem.is_visible():
                await page.evaluate("(el) => el.style.display = 'none'", file_list_elem)
        except Exception:
            pass  # It's okay if FileList isn't present

        assign_button = page.locator("input[type='button'][value='Assign']")
        await assign_button.click()

        log(f"🏷️ Assigned contractor '{desired_contractor_full}' to WO #{wo_number}")

    except Exception as e:
        log(f"❌ Contractor assignment process failed for WO #{wo_number}: {e}")

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
def generate_changes_file(old_list, new_list, changes_filename):
    def stringify(j):
        return f"{j['time']} - {j['name']} - {j['cid']} - {j['type']} - {j['address']} - WO {j['wo']}"

    # build sets of lines per company
    old_by_co = defaultdict(set)
    new_by_co = defaultdict(set)
    for j in old_list:
        old_by_co[j['company']].add(stringify(j))
    for j in new_list:
        new_by_co[j['company']].add(stringify(j))

    # all companies seen
    companies = sorted(set(old_by_co) | set(new_by_co))

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, changes_filename)

    with open(path, 'w', encoding='utf-8') as f:
        for co in companies:
            f.write(f"{co}\n")
            f.write("Added:\n")
            for line in sorted(new_by_co[co] - old_by_co.get(co, set())):
                f.write(f"  {line}\n")
            f.write("\nRemoved:\n")
            for line in sorted(old_by_co[co] - new_by_co.get(co, set())):
                f.write(f"  {line}\n")
            f.write("\n")

    return path

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
    ext = os.path.splitext(file_path)[1].lower()
    jobs = []

    if ext == ".txt":
        with open(file_path, "r", encoding="utf-8") as f:
            current_company = None
            current_date    = None

            for raw in f:
                line = raw.strip()
                if not line:
                    continue

                # 1) Company header: any line without " - " and no digits
                if " - " not in line and not any(c.isdigit() for c in line):
                    current_company = line
                    continue

                # 2) Date header: matches M-D-YY or M-D-YYYY
                if re.match(r'^\d{1,2}-\d{1,2}-\d{2,4}$', line):
                    current_date = line
                    continue

                # 3) Job line
                if " - " in line and "WO" in line:
                    parts = [p.strip() for p in line.split(" - ")]
                    if len(parts) >= 6:
                        time, name, cid, typ, addr, wo = parts[:6]
                        jobs.append({
                            "company": current_company,
                            "date":    current_date,
                            "time":    time,
                            "name":    name,
                            "cid":     cid,
                            "type":    typ,
                            "address": addr,
                            "wo":      wo.replace("WO ", "")
                        })

    elif ext == ".xlsx":
        df = pd.read_excel(file_path, header=None)
        current_company = None
        current_date    = None

        for row in df.itertuples(index=False):
            # flatten row to single string list
            cells = [str(c).strip() for c in row if c and str(c).strip()]
            if not cells:
                continue

            line = " - ".join(cells)
            # same detection logic as above
            if " - " not in line and not any(c.isdigit() for c in line):
                current_company = line
                continue
            if re.match(r'^\d{1,2}-\d{1,2}-\d{2,4}$', line):
                current_date = line
                continue
            if " - " in line and "WO" in line:
                parts = [p.strip() for p in line.split(" - ")]
                if len(parts) >= 6:
                    time, name, cid, typ, addr, wo = parts[:6]
                    jobs.append({
                        "company": current_company,
                        "date":    current_date,
                        "time":    time,
                        "name":    name,
                        "cid":     cid,
                        "type":    typ,
                        "address": addr,
                        "wo":      wo.replace("WO ", "")
                    })

    return jobs
