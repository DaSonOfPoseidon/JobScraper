import re
import sys
from collections import defaultdict, deque
import tkinter as tk
from datetime import datetime
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
import os
import json
from utils import prompt_reassignment, MISC_DIR, __version__

# ----------- CONFIGURABLE RULES -----------
CONTRACTORS = [
    "TGS Fiber",
    "Tex-Star Communications",
    "Subterraneus Installs",
    "Pifer Quality Communications",
    "Advanced Electric",
    "All Clear",
    "Maverick",
    "Socket",
    "North Sky",
    "None Assigned",
    "Unassigned",
]

LIMITS = {
        "Tex-Star Communications": 7,
        "North Sky": 3,
        "Maverick": 2,
        "Subterraneus Installs": 9,
        "TGS Fiber": 8,
        "Pifer Quality Communications": 3,
        "All Clear": 1,
        "Advanced Electric": 0,  # Never gets assigned
        "Unassigned": float("inf"),
        "Socket": float("inf"),
}

DEFAULT_LIMITS = {
    "Tex-Star Communications": 7,
    "North Sky": 3,
    "Maverick": 2,
    "Subterraneus Installs": 9,
    "TGS Fiber": 8,
    "Pifer Quality Communications": 3,
    "All Clear": 1,
    "Advanced Electric": 0,
    "Unassigned": 9999,  # use 9999 instead of infinity for JSON
    "Socket": 9999,
}

    # Helper: Area to assignment order map
AREA_PRIORITY = {
    "greater_boone": [
        "Tex-Star Communications",
        "North Sky",
        "overflow_subt_tgs",  # special handling for round robin
        "Maverick",
        "Unassigned",
    ],
        "jc": [
            "Maverick",
            "All Clear",
            "overflow_subt_tgs",
            "Unassigned",
        ],
        "west": [
            "TGS Fiber",
            "Unassigned"
        ],
        "rolla": [
            "Pifer Quality Communications",
            "Subterraneus Installs",
            "Unassigned"
        ],
        "kirksville": [
            "Subterraneus Installs",
            "Unassigned"
        ],
        "ofallon": [
            "Subterraneus Installs",
            "Unassigned"
        ],
        "unknown": [
            "Unassigned"
        ]
}
CITY_AREA = {
        # Greater Boone/Columbia & neighbors (customize as needed)
        "columbia": "greater_boone",
        "hallsville": "greater_boone",
        "sturgeon": "greater_boone",
        "centralia": "greater_boone",
        "ashland": "greater_boone",
        "boonville": "greater_boone",
        "rocheport": "greater_boone",
        "harrisburg": "greater_boone",
        "moberly": "greater_boone",
        "fulton": "greater_boone",
        "clark": "greater_boone",
        "fayette": "greater_boone",
        "new franklin": "greater_boone",

        # JC area
        "jefferson city": "jc",
        "lohman": "jc",

        # West
        "clinton": "west",
        "oak grove": "west",
        "warrensburg": "west",
        "sedalia": "west",

        # Rolla/Lebanon
        "rolla": "rolla",
        "lebanon": "rolla",

        # Kirksville
        "kirksville": "kirksville",

        # O'Fallon
        "o'fallon": "ofallon",
        "o fallon": "ofallon"
}

SOCKET_FORCED_STREETS = [
    "endeavor",
    "kentsfield",
    "briarmont",
    "bearfield",
    "discovery ridge"
]

CITY_LIST = list(CITY_AREA.keys())

CONFIG_PATH = os.path.join(MISC_DIR, "spreader_config.json")

DEFAULT_CONFIG = {
    "limits": DEFAULT_LIMITS,
    "forced_streets": SOCKET_FORCED_STREETS,
    "area_priority": AREA_PRIORITY,
    "city_area": CITY_AREA
}

AREA_COVERAGE = {
    "greater_boone": {"Tex-Star Communications", "North Sky", "Maverick", "Subterraneus Installs", "TGS Fiber"},
    "jc": {"Maverick", "All Clear", "Subterraneus Installs", "TGS Fiber"},
    "west": {"TGS Fiber"},
    "rolla": {"Pifer Quality Communications", "Subterraneus Installs"},
    "kirksville": {"Subterraneus Installs"},
    "ofallon": {"Subterraneus Installs"},
    "unknown": set()
}

TIMESLOT_ORDER = ['8:00', '10:00', '12:00', '1:00', '3:00']

# --- Utility functions ---
def parse_city(address):
    # Remove state and zip at the end (e.g. ', MO 65109' or ', 65101')
    address = address.strip()
    state_zip = re.search(r',?\s*[A-Z]{2}\s*\d{5}$', address)
    if state_zip:
        address = address[:state_zip.start()]
    # Now find the longest city at the end of the remaining address
    address_lower = address.lower()
    for city in sorted(CITY_LIST, key=lambda c: -len(c)):
        # Remove apostrophes for matching, ignore case, allow city to be just before any end
        city_clean = city.replace("'", "").lower()
        if address_lower.replace("'", "").endswith(city_clean):
            return city
        # Try matching with city as a word near the end (with or without 'apt', etc.)
        if re.search(rf"\b{re.escape(city_clean)}\b", address_lower.replace("'", "")):
            return city
    return ""

def extract_timeslot(job_line):
    m = re.match(r'^(\d{1,2}:\d{2})', job_line)
    return m.group(1) if m else "Unknown"

def extract_address(job_line):
    m = re.search(r'- ([^-]+) - WO ', job_line)
    return m.group(1) if m else ""

def address_triggers_socket(address):
    addr_lower = address.lower()
    return any(street in addr_lower for street in SOCKET_FORCED_STREETS)

def extract_customer_name(job_line):
    # job_line example: "8:00 - Janet Brown - 0991-7706-6125 - Connectorized - 1829 ... - WO 490915"
    parts = job_line.split(" - ")
    if len(parts) >= 2:
        return parts[1].strip().lower()  # lowercase for consistent sorting
    return ""

def detect_date(line):
    return bool(re.match(r'^\d{1,2}-\d{1,2}-\d{2,4}$', line.strip()))

def is_job_line(line):
    return bool(re.match(r'^\d{1,2}:\d{2}', line)) and 'WO ' in line

def parse_moved_jobs_from_spread(spread_file):
    moved_jobs = []
    with open(spread_file, encoding="utf-8") as f:
        current_contractor = None
        for line in f:
            line = line.strip()
            if not line:
                continue
            print(f"Line read: '{line}'")
            if line in CONTRACTORS:
                current_contractor = line
                print(f"Found contractor header: '{line}'")
                continue
            m = re.match(r".*WO (\d+).*(# MOVED.*)", line, re.IGNORECASE)
            if m and current_contractor:
                print(f"Found moved job under '{current_contractor}': {line}")
                moved_jobs.append({
                    "contractor": current_contractor,
                    "wo": m.group(1),
                    "line": line
                })
    print(f"Total moved jobs found: {len(moved_jobs)}")
    return moved_jobs

def parse_input(filename):
    # Returns {section: [{'date':..., 'jobs': [...]}]}
    sections = defaultdict(list)
    with open(filename, encoding='utf-8') as f:
        current_section = None
        current_date = None
        for line in f:
            line = line.rstrip('\n')
            section = detect_section(line)
            if section:
                current_section = section
                current_date = None
                continue
            if detect_date(line):
                current_date = line.strip()
                sections[current_section].append({'date': current_date, 'jobs': []})
                continue
            if is_job_line(line):
                if current_section and current_date:
                    sections[current_section][-1]['jobs'].append(line)
                continue
    return sections

def slot_key(slot):
    try:
        idx = TIMESLOT_ORDER.index(slot)
        return (idx, 0)
    except ValueError:
        h, m = map(int, slot.split(":"))
        return (len(TIMESLOT_ORDER), h*60+m)

def log_job_changes(jobs, output_sections, contractors):
    added = defaultdict(list)
    removed = defaultdict(list)

    # Build reverse map: for each job, where did it end up?
    final_assignment = {}
    for contractor, slotdict in output_sections.items():
        for jobs_by_slot in slotdict.values():
            for job in jobs_by_slot:
                final_assignment[job] = contractor

    for job in jobs:
        orig = job['original_contractor']
        final = final_assignment.get(job['line'], None)
        if orig != final:
            if final is not None:
                added[final].append(job['line'])
            removed[orig].append(job['line'])

    return added, removed

def write_change_log(added, removed, filename="job_changes.log"):
    with open(filename, "w", encoding="utf-8") as f:
        for contractor in CONTRACTORS:
            added_jobs = added.get(contractor, [])
            removed_jobs = removed.get(contractor, [])
            if not added_jobs and not removed_jobs:
                continue
            f.write(f"{contractor}\n\n")
            if added_jobs:
                f.write("Added:\n")
                for job in added_jobs:
                    f.write(f"{job}\n")
            if removed_jobs:
                f.write("Removed:\n")
                for job in removed_jobs:
                    f.write(f"{job}\n")
            f.write("\n")
    print(f"[DONE] Change log written to {filename}")

def detect_section(line):
    """
    Only recognize lines matching a contractor name as a new section.
    """
    section = line.strip()
    return section if section in CONTRACTORS else None

def reassign_jobs(sections):
    # Flatten jobs list
    jobs = []
    for contractor, days in sections.items():
        for day in days:
            for job in day['jobs']:
                parts = [p.strip() for p in job.split(" - ")]
                slot = extract_timeslot(job)
                date = day['date']
                city = parse_city(extract_address(job))
                jobs.append({
                    'original_contractor': contractor,
                    'contractor': contractor,
                    'date': date,
                    'time': slot,
                    'city': city,
                    'line': job,
                    'job_type': parts[3] if len(parts) > 3 else ""
                })
    apply_forced_assignments(jobs)

    # Sort by date, timeslot, then customer
    jobs.sort(key=lambda j: (j['date'], slot_key(j['time']), extract_customer_name(j['line'])))

    slot_counts = defaultdict(lambda: defaultdict(int))  # contractor -> (date,time) -> count
    move_comments = {}
    output_sections = {c: defaultdict(list) for c in CONTRACTORS}

    # Group jobs by area
    jobs_by_area = defaultdict(lambda: defaultdict(list))
    for job in jobs:
        area = CITY_AREA.get(job['city'].lower(), 'unknown')
        key = (job['date'], job['time'])
        jobs_by_area[area][key].append(job)

    # Collect unassigned JC jobs for overflow
    jc_unassigned = []

    # 1) Assign all non-Boone-area jobs first (skip 'greater_boone')
    for area, slots in jobs_by_area.items():
        if area == 'greater_boone':
            continue
        priority = []
        for c in AREA_PRIORITY.get(area, []):
            if c == 'overflow_subt_tgs':
                priority.extend(['TGS Fiber', 'Subterraneus Installs'])
            elif not c.startswith('overflow_'):
                priority.append(c)
        for key, slot_jobs in slots.items():
            if area == 'jc':
                unassigned = assign_strict_priority(
                    slot_jobs, key, priority, output_sections, slot_counts, move_comments,
                    return_unassigned=True
                )
                jc_unassigned.extend(unassigned)
            else:
                assign_strict_priority(
                    slot_jobs, key, priority, output_sections, slot_counts, move_comments
                )

    # 2) Assign all Boone-area jobs ('greater_boone')
    gb_priority = []
    if 'greater_boone' in jobs_by_area:
        for c in AREA_PRIORITY.get('greater_boone', []):
            if c == 'overflow_subt_tgs':
                gb_priority.extend(['TGS Fiber', 'Subterraneus Installs'])
            elif not c.startswith('overflow_'):
                gb_priority.append(c)
        for key, slot_jobs in jobs_by_area['greater_boone'].items():
            assign_strict_priority(
                slot_jobs, key, gb_priority, output_sections, slot_counts, move_comments
            )

    # 3) Roll any JC unassigned jobs into Greater Boone overflow
    for job in jc_unassigned:
        key = (job['date'], job['time'])
        placed = False
        for contractor in gb_priority:
            if slot_counts[contractor][key] < LIMITS.get(contractor, 0):
                output_sections[contractor][key].append(job['line'])
                slot_counts[contractor][key] += 1
                move_comments[job['line']] = (
                    f"MOVED from {job['original_contractor']} (jc overflow to greater_boone)"
                )
                placed = True
                break
        if not placed:
            output_sections['Unassigned'][key].append(job['line'])
            slot_counts['Unassigned'][key] += 1
            move_comments[job['line']] = (
                f"MOVED from {job['original_contractor']} (jc overflow unassigned fallback)"
            )

    return output_sections, move_comments, jobs

def assign_greater_boone_jobs(jobs, key, output_sections, slot_counts, move_comments):
    # This function is preserved for backward compatibility but
    # your logic is now centralized in assign_strict_priority below,
    # so here just call assign_strict_priority for the Greater Boone overflow contractors.

    subt = "Subterraneus Installs"
    tgs = "TGS Fiber"
    area = "greater_boone"

    priority_list = []
    for c in AREA_PRIORITY.get(area, []):
        if c.startswith("overflow_"):
            if c == "overflow_subt_tgs":
                priority_list.extend([tgs, subt])
        else:
            priority_list.append(c)

    assign_strict_priority(jobs, key, priority_list, output_sections, slot_counts, move_comments)

def assign_strict_priority(jobs, key, priority_list, output_sections, slot_counts, move_comments, return_unassigned=False):
    unassigned_jobs = []
    for job in jobs:
        orig = job['contractor']
        forced = job.get('forced_contractor')
        assigned = False

        # If forced contractor present, try assign there first (ignore priority)
        if forced:
            if slot_counts[forced][key] < LIMITS.get(forced, 0):
                output_sections[forced][key].append(job['line'])
                slot_counts[forced][key] += 1
                assigned = True
                if orig != forced:
                    move_comments[job['line']] = f"MOVED from {orig} (forced assignment to {forced})"
                continue
            else:
                # Forced contractor is full, consider fallback
                pass

        # Try original contractor if valid and has capacity
        if orig in priority_list and slot_counts[orig][key] < LIMITS.get(orig, 0):
            output_sections[orig][key].append(job['line'])
            slot_counts[orig][key] += 1
            assigned = True
            continue

        # Strict priority fallback
        for contractor in priority_list:
            if slot_counts[contractor][key] < LIMITS.get(contractor, 0):
                output_sections[contractor][key].append(job['line'])
                slot_counts[contractor][key] += 1
                assigned = True
                if orig != contractor:
                    move_comments[job['line']] = f"MOVED from {orig} (area overflow, assigned to {contractor})"
                break

        if not assigned:
            if return_unassigned:
                unassigned_jobs.append(job)
            else:
                output_sections["Unassigned"][key].append(job['line'])
                slot_counts["Unassigned"][key] += 1
                if orig != "Unassigned":
                    move_comments[job['line']] = f"MOVED from {orig} (area unassigned fallback)"

    if return_unassigned:
        return unassigned_jobs

def parse_date_str(date_str):
    try:
        return datetime.strptime(date_str, "%m-%d-%y")
    except ValueError:
        return datetime.strptime(date_str, "%m-%d-%Y")

def write_output(sections, move_comments, filename="output.txt"):
    with open(filename, "w", encoding="utf-8") as f:
        for contractor in CONTRACTORS:
            dayslots = sections[contractor]
            if not dayslots:
                continue
            f.write(f"{contractor}\n\n")
            
            by_date = defaultdict(list)
            for (date, slot), jobs in dayslots.items():
                by_date[date].append((slot, jobs))
            
            for date in sorted(by_date, key=parse_date_str):
                f.write(f"{date}\n")
                slotjobs = sorted(by_date[date], key=lambda x: slot_key(x[0]))
                for slot, jobs in slotjobs:
                    jobs_sorted = sorted(jobs, key=lambda job: extract_customer_name(job).lower())
                    for job in jobs_sorted:
                        note = f"  # {move_comments[job]}" if job in move_comments else ""
                        f.write(f"{job}{note}\n")
                f.write("\n")
    print(f"[DONE] Output written to {filename}")

def run_process(file_path):
    global LIMITS, SOCKET_FORCED_STREETS, AREA_PRIORITY
    try:
        ensure_config_file_exists()
        config = load_spreader_config()

        LIMITS = config.get("limits", DEFAULT_LIMITS)
        SOCKET_FORCED_STREETS = config.get("forced_streets", SOCKET_FORCED_STREETS)
        AREA_PRIORITY = config.get("area_priority", AREA_PRIORITY)
        
        sections = parse_input(file_path)
        final_sections, move_comments, jobs = reassign_jobs(sections)
        base, ext = os.path.splitext(file_path)
        out_file = base + "_spread.txt"
        write_output(final_sections, move_comments, filename=out_file)
        added, removed = log_job_changes(jobs, final_sections, CONTRACTORS)
        write_change_log(added, removed, filename=base + "_changelog.txt")
        return out_file
    except Exception as e:
        return str(e)

def ensure_config_file_exists():
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)
        print(f"[INFO] Created default config file at {CONFIG_PATH}")

def load_spreader_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)

    # Replace sentinel 9999 back to float('inf') in limits
    def replace_9999_with_inf(obj):
        if isinstance(obj, dict):
            return {k: replace_9999_with_inf(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [replace_9999_with_inf(i) for i in obj]
        elif obj == 9999:
            return float('inf')
        else:
            return obj

    return replace_9999_with_inf(config)

def save_spreader_config(config):
    # Replace float('inf') with sentinel 9999 before saving
    def replace_inf(obj):
        if isinstance(obj, dict):
            return {k: replace_inf(v) for k, v in obj.items()}
        elif isinstance(obj, float) and (obj == float('inf')):
            return 9999
        elif isinstance(obj, list):
            return [replace_inf(i) for i in obj]
        else:
            return obj

    safe_config = replace_inf(config)

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(safe_config, f, indent=2)

def apply_forced_assignments(jobs):
    for job in jobs:
        address = extract_address(job['line'])
        job_type = job.get('job_type', '').lower()

        if address_triggers_socket(address):
            job['forced_contractor'] = "Socket"
            continue

        if "conversion" in job_type:
            job['forced_contractor'] = "Socket"
            continue

        # Otherwise no forced contractor
        job['forced_contractor'] = None

def open_settings_gui(root):
    config = load_spreader_config()
    limits = config.get("limits", {})
    forced_streets = config.get("forced_streets", [])

    settings_win = tk.Toplevel(root)
    settings_win.title("Spreader Configuration")

    entries = {}
    # Contractor Limits Section
    tk.Label(settings_win, text="Contractor Limits", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=2, pady=(10, 0))

    for i, contractor in enumerate(CONTRACTORS, start=1):
        tk.Label(settings_win, text=contractor, width=30, anchor="w").grid(row=i, column=0, padx=10, pady=2)
        limit_val = limits.get(contractor, 0)
        if limit_val == float('inf'):
            limit_val = 9999  # show as large number
        entry = tk.Entry(settings_win, width=10)
        entry.insert(0, str(limit_val))
        entry.grid(row=i, column=1, padx=10, pady=2)
        entries[contractor] = entry

    # Forced Streets Section (simple multi-line Text widget)
    row_forced_streets = len(CONTRACTORS) + 2
    tk.Label(settings_win, text="Forced Streets (one per line):", font=("Segoe UI", 11, "bold")).grid(row=row_forced_streets, column=0, columnspan=2, pady=(20, 5))
    text_forced_streets = tk.Text(settings_win, height=8, width=40)
    text_forced_streets.grid(row=row_forced_streets + 1, column=0, columnspan=2, padx=10)
    text_forced_streets.insert("1.0", "\n".join(forced_streets))

    def save_and_close():
        new_limits = {}
        try:
            for c, e in entries.items():
                val = int(e.get())
                if val < 0:
                    raise ValueError("Limit cannot be negative")
                new_limits[c] = float('inf') if val == 9999 else val
        except ValueError as ve:
            tk.messagebox.showerror("Invalid input", f"Please enter valid non-negative integers.\n{ve}")
            return

        new_forced_streets = [line.strip() for line in text_forced_streets.get("1.0", "end").splitlines() if line.strip()]

        # Update the config dict and save
        config['limits'] = new_limits
        config['forced_streets'] = new_forced_streets

        save_spreader_config(config)

        # Update global vars if used elsewhere
        global LIMITS, SOCKET_FORCED_STREETS
        LIMITS = new_limits
        SOCKET_FORCED_STREETS = new_forced_streets

        tk.messagebox.showinfo("Saved", "Configuration updated successfully.")
        settings_win.destroy()

    btn_save = tk.Button(settings_win, text="Save", command=save_and_close)
    btn_save.grid(row=row_forced_streets + 2, column=0, pady=10)

    btn_cancel = tk.Button(settings_win, text="Cancel", command=settings_win.destroy)
    btn_cancel.grid(row=row_forced_streets + 2, column=1, pady=10)

    settings_win.transient(root)
    settings_win.grab_set()
    root.wait_window(settings_win)

def start_gui():
    def on_drop(event):
        file_path = event.data.strip('{}')  # Handles filenames with spaces
        if not file_path.lower().endswith(".txt"):
            messagebox.showerror("Error", "Please drop a .txt file!")
            return
        status.set("Processing...")
        root.update()
        out_file = run_process(file_path)
        if os.path.exists(out_file):
            status.set(f"Done! Output: {out_file}")
            messagebox.showinfo("Finished", f"Output written to:\n{out_file}")
        else:
            status.set("Error!")
            messagebox.showerror("Error", f"Processing failed:\n{out_file}")
        prompt_reassignment(root, out_file)

    root = TkinterDnD.Tk()
    root.title(f"ButterKnife v{__version__}")
    root.geometry("480x180")

    label = tk.Label(root, text="\nDrop your jobs .txt file here", font=("Segoe UI", 15), pady=16)
    label.pack(expand=True)

    btn_settings = tk.Button(root, text="Edit Spreader Config", command=lambda: open_settings_gui(root))
    btn_settings.pack(pady=(0, 10))

    status = tk.StringVar(value="Waiting for file...")
    status_label = tk.Label(root, textvariable=status, fg="blue", font=("Segoe UI", 11))
    status_label.pack(pady=(0, 16))

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)

    root.mainloop()

if __name__ == "__main__":
    if "--version" in sys.argv:
        print(__version__)
        sys.exit(0)

    ensure_config_file_exists()
    config = load_spreader_config()

    LIMITS = config.get("limits", DEFAULT_LIMITS)
    SOCKET_FORCED_STREETS = config.get("forced_streets", SOCKET_FORCED_STREETS)
    AREA_PRIORITY = config.get("area_priority", AREA_PRIORITY)

    start_gui()
