# scraper_core.py
import traceback
import time
from datetime import datetime
from datetime import timedelta
from dateutil import parser as dateparser
import asyncio
from playwright.async_api import async_playwright, Page, TimeoutError as PlaywrightTimeout

from utils import (
    clear_first_time_overlays, NoWOError, NoOpenWOError,
    get_work_order_url, get_job_type_and_address,
    get_contractor_assignments, extract_wo_date, extract_cid_and_time
)

CALENDAR_URL = "http://inside.sockettelecom.com/events/calendar.php"
CUSTOMER_URL_TEMPLATE = "http://inside.sockettelecom.com/menu.php?coid=1&tabid=7&parentid=9&customerid={}"

async def init_playwright_page(headless: bool = True, browser=None, playwright=None):
    # If browser is not passed, start everything fresh
    if browser is None or playwright is None:
        playwright = await async_playwright().start()
        browser = await playwright.chromium.launch(
            headless=headless,
            args=[
                "--disable-gpu",
                "--no-sandbox",
                "--disable-usb-keyboard-detect",
                "--disable-hid-detection",
                "--log-level=3",
                "--blink-settings=imagesEnabled=false",
                "--headless=new"
            ],
        )
        owns_browser = True  # So caller can close these later
    else:
        owns_browser = False

    context = await browser.new_context(
        java_script_enabled=True,
        bypass_csp=True,
        viewport={"width": 1920, "height": 1080},
        storage_state="state.json"
    )
    async def block_resources(route):
        if route.request.resource_type in ["image"]:
            await route.abort()
        else:
            await route.continue_()
    await context.route("**/*", block_resources)
    page = await context.new_page()

    if owns_browser:
        return playwright, browser, context, page
    else:
        return context, page
   
async def scrape_jobs(page: Page, mode="week", imported_jobs=None, selected_day=None, test_mode=False, test_limit=10, log=print):
    # Assumes: you are already logged in and on the right context/page
    await page.goto(CALENDAR_URL)
    
    # Step 1: Set View (Week or Day)
    try:
        if mode == "week":
            btn = await page.wait_for_selector("button.fc-agendaWeek-button", timeout=8000)
        else:
            btn = await page.wait_for_selector("button.fc-agendaDay-button", timeout=8000)
        await btn.click()
        await page.wait_for_selector("#spinner", state="hidden", timeout=10_000)
        log(f"✅ Switched to {'Week' if mode == 'week' else 'Day'} View.")
    except Exception as e:
        log(f"⚠️ Could not switch view: {e}")

    # Step 2: Navigate to correct week/day
    if selected_day:
        try:
            raw_selected_date = dateparser.parse(selected_day).date()
            if mode == "week":
                weekday = raw_selected_date.weekday()  # 0=Monday
                days_to_sunday = (raw_selected_date.weekday() + 1) % 7
                target_date = raw_selected_date - timedelta(days=days_to_sunday)
            else:
                target_date = raw_selected_date

            async def get_current_calendar_date():
                try:
                    header_text = await page.locator(".fc-center h2").text_content()
                    header_text = header_text.strip()
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

            current_date = await get_current_calendar_date()
            nav_tries = 0

            while current_date and current_date != target_date and nav_tries < 15:
                if current_date < target_date:
                    await page.click("button.fc-next-button")
                else:
                    await page.click("button.fc-prev-button")
                await asyncio.sleep(.25)
                current_date = await get_current_calendar_date()
                nav_tries += 1

            if current_date == target_date:
                log(f"📆 Calendar now displaying {current_date}")
            else:
                log(f"⚠️ Could not align calendar to {target_date} after {nav_tries} tries.")
        except Exception as e:
            log(f"❌ Failed to navigate to selected day: {e}")

    # Step 3: Wait for jobs to load
    try:
        await page.wait_for_selector(
            'a.fc-time-grid-event:has-text("Residential Fiber Install")', timeout=10_000
        )
        log("✅ Jobs loaded and ready to scrape.")
    except PlaywrightTimeout:
        log("⚠️ No 'Residential Fiber Install' jobs detected.")

    # Step 4: Extract metadata
    job_links = await page.query_selector_all('a.fc-time-grid-event')
    results = []
    counter = 0

    log("Scraping Calendar...")
    for link in job_links:
        text = await link.inner_text()
        if "Residential Fiber Install" not in text:
            continue

        # You must pass a compatible async extract_cid_and_time or adapt below:
        cid, name, time_slot = await extract_cid_and_time(link, text) if extract_cid_and_time else (None, None, None)
        if not cid:
            continue

        counter += 1
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

async def process_job_entries(page: Page, job: dict, log=print):
    cid = job.get("cid")
    name = job.get("name")
    time_slot = job.get("time")
    customer_url = CUSTOMER_URL_TEMPLATE.format(cid)

    try:
        t0 = time.perf_counter()
        await page.goto(customer_url)
        #print(f"[{cid}] Customer page took {time.perf_counter() - t0:.2f}s to load")
        if clear_first_time_overlays:
            t0 = time.perf_counter()
            await clear_first_time_overlays(page)
            #print(f"[{cid}] Clear overlay took {time.perf_counter() - t0:.2f}s")

        # Switch to MainView iframe
        try:
            t0 = time.perf_counter()
            await page.wait_for_selector('iframe[name="MainView"]', timeout=10_000)
            frame = page.frame(name="MainView")
            #print(f"[{cid}] Find and assign MainView took {time.perf_counter() - t0:.2f}s")
        except PlaywrightTimeout:
            frame = page.main_frame()

        # Wait for WO
        async def _wait_for_work_order():
            try:
                return await get_work_order_url(frame, log=log)
            except (NoWOError, NoOpenWOError):
                raise
            except Exception:
                return False

        try:
            t0 = time.perf_counter()
            workorder_url, wo_number = await asyncio.wait_for(_wait_for_work_order(), timeout=5)
            #print(f"[{cid}] WO URL Scrape took {time.perf_counter() - t0:.2f}s")
        except (NoWOError, NoOpenWOError) as e:
            job["error"] = str(e)
            log(f"[{cid}] No work order found: {e}")
            return None
        except asyncio.TimeoutError:
            workorder_url, wo_number = None, None

        if not workorder_url:
            log(f"[{cid}] No workorder_url found, skipping job")
            return None

        t0 = time.perf_counter()
        await page.goto(workorder_url)
        #print(f"[{cid}] WO took {time.perf_counter() - t0:.2f}s to load")

        t0 = time.perf_counter()
        job_type, address = (await get_job_type_and_address(page)) if get_job_type_and_address else (None, None)
        #print(f"[{cid}] job type took {time.perf_counter() - t0:.2f}s")
        t0 = time.perf_counter()
        contractor_info = (await get_contractor_assignments(page)) if get_contractor_assignments else None
        print(f"[{cid}] Contractor extract took {time.perf_counter() - t0:.2f}s")
        t0 = time.perf_counter()
        job_date = (await extract_wo_date(page))
        #print(f"[{cid}] Date extract took {time.perf_counter() - t0:.2f}s")

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