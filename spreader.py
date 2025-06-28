import re
import sys
from collections import defaultdict, deque
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
import os
import json

__version__ = "0.1.0"

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
    "Unassigned"
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

CITY_LIST = list(CITY_AREA.keys())
LIMITS_FILE = "contractor_limits.json"

AREA_COVERAGE = {
    "greater_boone": {"Tex-Star Communications", "North Sky", "Maverick", "Subterraneus Installs", "TGS Fiber"},
    "jc": {"Maverick", "All Clear", "Subterraneus Installs", "TGS Fiber"},
    "west": {"TGS Fiber"},
    "rolla": {"Pifer Quality Communications", "Subterraneus Installs"},
    "kirksville": {"Subterraneus Installs"},
    "ofallon": {"Subterraneus Installs"},
    "unknown": set()
}

# Slot limits
TEXSTAR_LIMIT = 7
PIFER_LIMIT = 3
ALLCLEAR_LIMIT = 1
ADVANCED_LIMIT = 0
NORTHSKY_LIMIT = 2
SUBT_LIMIT = 9
TGS_LIMIT = 8
MAVERICK_LIMIT = 2

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

def detect_section(line):
    if ' - ' in line or not line.strip():
        return None
    if re.match(r'^\d{1,2}-\d{1,2}-\d{2,4}$', line.strip()):
        return None
    return line.strip()

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
            if line in CONTRACTORS:
                current_contractor = line
                continue
            # Match job line with WO and "# MOVED"
            m = re.match(r".*WO (\d+).*(# MOVED.*)", line, re.IGNORECASE)
            if m and current_contractor:
                moved_jobs.append({
                    "contractor": current_contractor,
                    "wo": m.group(1),
                    "line": line
                })
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

def reassign_jobs(sections):
    # Parse all jobs into a flat list with fields
    jobs = []
    for contractor, days in sections.items():
        for day in days:
            for job in day['jobs']:
                slot = extract_timeslot(job)
                addr = extract_address(job)
                city = parse_city(addr)
                parts = [p.strip() for p in job.split(" - ")]
                job_type = parts[3] if len(parts) > 3 else ""
                jobs.append({
                    'original_contractor': contractor,
                    'contractor': contractor,
                    'date': day['date'],
                    'time': slot,
                    'city': city,
                    'line': job,
                    'job_type': job_type
                })

    # Track per-slot assignments per contractor
    slot_counts = defaultdict(lambda: defaultdict(int))  # company -> (date, time) -> count

    # Track round-robin overflow state for each slot for SubT/TGS
    rr_state = defaultdict(lambda: deque(["Subterraneus Installs", "TGS Fiber"]))

    move_comments = {}
    output_sections = {c: defaultdict(list) for c in CONTRACTORS}

    # For minimal churn, pre-compute which companies can cover which areas
    AREA_COVERAGE = {}
    for area, priorities in AREA_PRIORITY.items():
        # flatten round robin marker for coverage
        AREA_COVERAGE[area] = set([p for p in priorities if not p.startswith('overflow_')])

    for job in jobs:
        key = (job['date'], job['time'])
        orig = job['contractor']
        city = job['city']
        job_line = job['line']

        # 1. Special job type override (from your rules)
        if job['job_type'].lower() == "5 gig conversion":
            if orig != 'Socket':
                move_comments[job_line] = f"FORCED to Socket by job type rule (was {orig})"
            output_sections['Socket'][key].append(job_line)
            slot_counts['Socket'][key] += 1
            continue

        # 2. Figure out area by city
        area = CITY_AREA.get(city.lower(), "unknown")
        priorities = AREA_PRIORITY.get(area, ["Unassigned"])

        # 3. Minimal churn: if original contractor can cover this area and is under slot limit, keep it
        if orig in AREA_COVERAGE.get(area, set()) and slot_counts[orig][key] < LIMITS.get(orig, 0):
            output_sections[orig][key].append(job_line)
            slot_counts[orig][key] += 1
            continue

        # 4. Assign according to area priority
        assigned = False
        for p in priorities:
            if p.startswith("overflow_"):
                # Round robin overflow for SubT/TGS
                subt_full = slot_counts["Subterraneus Installs"][key] >= LIMITS.get("Subterraneus Installs", 0)
                tgs_full = slot_counts["TGS Fiber"][key] >= LIMITS.get("TGS Fiber", 0)
                if subt_full and tgs_full:
                    continue
                for _ in range(2):  # Try each once
                    target = rr_state[key][0]
                    if slot_counts[target][key] < LIMITS.get(target, 0):
                        if orig != target:
                            move_comments[job_line] = f"MOVED from {orig} (overflow, balanced to {target})"
                        output_sections[target][key].append(job_line)
                        slot_counts[target][key] += 1
                        assigned = True
                        rr_state[key].rotate(-1)
                        break
                    else:
                        rr_state[key].rotate(-1)
                if assigned:
                    break
            else:
                if slot_counts[p][key] < LIMITS.get(p, 0):
                    if orig != p:
                        move_comments[job_line] = f"MOVED from {orig}"
                    output_sections[p][key].append(job_line)
                    slot_counts[p][key] += 1
                    assigned = True
                    break

        # 5. If not assigned, put in Unassigned
        if not assigned:
            if orig != "Unassigned":
                move_comments[job_line] = f"MOVED from {orig} (overflow, unassigned)"
            output_sections["Unassigned"][key].append(job_line)
            slot_counts["Unassigned"][key] += 1

    return output_sections, move_comments, jobs

def write_output(sections, move_comments, filename="output.txt"):
    # slot ordering
    with open(filename, "w", encoding="utf-8") as f:
        for contractor in CONTRACTORS:
            dayslots = sections[contractor]
            if not dayslots:
                continue
            f.write(f"{contractor}\n\n")
            # Gather all dates in order
            by_date = defaultdict(list)
            for (date, slot), jobs in dayslots.items():
                by_date[date].append((slot, jobs))
            for date in sorted(by_date):
                f.write(f"{date}\n")
                # sort by slot order
                slotjobs = sorted(by_date[date], key=lambda x: slot_key(x[0]))
                for slot, jobs in slotjobs:
                    for job in jobs:
                        note = f"  # {move_comments[job]}" if job in move_comments else ""
                        f.write(f"{job}{note}\n")
                f.write("\n")
    print(f"[DONE] Output written to {filename}")

def run_process(file_path):
    try:
        LIMITS = load_limits()
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

def ensure_limits_file_exists(filename=LIMITS_FILE):
    if not os.path.exists(filename):
        with open(filename, "w") as f:
            json.dump(DEFAULT_LIMITS, f, indent=2)
        print(f"[INFO] Created default limits file at {filename}")

def load_limits(filename=LIMITS_FILE):
    try:
        with open(filename, "r") as f:
            data = json.load(f)
            for k in ["Unassigned", "Socket"]:
                if k in data and data[k] == 9999:
                    data[k] = float("inf")
            return data
    except Exception as e:
        print(f"Warning: Could not load limits from {filename}, using defaults. {e}")
        return {
            "Tex-Star Communications": 7,
            "North Sky": 3,
            "Maverick": 2,
            "Subterraneus Installs": 9,
            "TGS Fiber": 8,
            "Pifer Quality Communications": 3,
            "All Clear": 1,
            "Advanced Electric": 0,
            "Unassigned": float("inf"),
            "Socket": float("inf"),
        }

def save_limits(limits, filename=LIMITS_FILE):
    # Convert float('inf') back to 9999 or a large int before saving
    to_save = {}
    for k, v in limits.items():
        if v == float('inf'):
            to_save[k] = 9999
        else:
            to_save[k] = v
    with open(filename, "w") as f:
        json.dump(to_save, f, indent=2)

def open_settings_gui(root):
    limits = load_limits()

    settings_win = tk.Toplevel(root)
    settings_win.title("Contractor Limits Settings")

    entries = {}

    tk.Label(settings_win, text="Contractor", font=("Segoe UI", 11, "bold"), width=30).grid(row=0, column=0, padx=10, pady=5)
    tk.Label(settings_win, text="Limit", font=("Segoe UI", 11, "bold"), width=10).grid(row=0, column=1, padx=10, pady=5)

    for i, contractor in enumerate(CONTRACTORS, start=1):
        tk.Label(settings_win, text=contractor, width=30, anchor="w").grid(row=i, column=0, padx=10, pady=2)
        limit_val = limits.get(contractor, 0)
        if limit_val == float('inf'):
            limit_val = 9999  # show as large number

        entry = tk.Entry(settings_win, width=10)
        entry.insert(0, str(limit_val))
        entry.grid(row=i, column=1, padx=10, pady=2)
        entries[contractor] = entry

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

        save_limits(new_limits)
        global LIMITS
        LIMITS = new_limits
        tk.messagebox.showinfo("Saved", "Contractor limits updated successfully.")
        settings_win.destroy()

    btn_save = tk.Button(settings_win, text="Save", command=save_and_close)
    btn_save.grid(row=len(CONTRACTORS)+1, column=0, pady=10)

    btn_cancel = tk.Button(settings_win, text="Cancel", command=settings_win.destroy)
    btn_cancel.grid(row=len(CONTRACTORS)+1, column=1, pady=10)

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

    root = TkinterDnD.Tk()
    root.title("Job Assignment Reorganizer")
    root.geometry("480x180")

    label = tk.Label(root, text="\nDrop your jobs .txt file here", font=("Segoe UI", 15), pady=16)
    label.pack(expand=True)

    btn_settings = tk.Button(root, text="Edit Contractor Limits", command=lambda: open_settings_gui(root))
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

    LIMITS = load_limits()        
    ensure_limits_file_exists()
    start_gui()
