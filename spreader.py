import re
from collections import defaultdict, deque
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
import requests
import os
import time

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
    "Uncategorized"
]

# City assignment rules
CITY_MAP = {
    'clinton': 'tgs',
    'oak grove': 'tgs',
    'warrensburg': 'tgs',
    'sedalia': 'tgs',
    'kirksville': 'subt',
    "o fallon": 'subt',
    "o'fallon": 'subt',
    'rolla': 'pifer',
    'lebanon': 'pifer',
    'columbia': 'texstar',
    'fayette': 'maverick',
    'jefferson city': 'jeffcity',
    'lohman': 'jeffcity',
    'sturgeon': 'subt',
    'harrisburg': 'subt',
    'ashland': 'subt',
    'boonville': 'subt',
    'centralia': 'subt',
    'rocheport': 'subt',
    'moberly': 'subt',
    'hallsville': 'subt',
    'fulton': 'subt',
    'clark': 'subt',
}

CITY_LIST = [
    'clinton', 'oak grove', 'warrensburg', 'sedalia', 'kirksville',
    "o'fallon", "o fallon", 'rolla', 'lebanon', 'columbia', 'fayette', 'jefferson city', 
    'lohman', 'sturgeon', 'harrisburg', 'ashland', 'boonville', 'centralia',
    'rocheport', 'moberly', 'hallsville', 'fulton', 'clark'
    # Add more as needed!
]

# Slot limits
TEXSTAR_LIMIT = 8
PIFER_LIMIT = 3
ALLCLEAR_LIMIT = 1
ADVANCED_LIMIT = 1

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

def get_city_group(city):
    return CITY_MAP.get(city.lower(), 'uncategorized')

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
    # Parse all jobs into flat list
    jobs = []
    for contractor, days in sections.items():
        for day in days:
            for job in day['jobs']:
                slot = extract_timeslot(job)
                addr = extract_address(job)
                city = parse_city(addr)
                # ----- Extract job type as the 4th split item (strip spaces) -----
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

    # Track slot counts
    texstar_ct = defaultdict(int)
    pifer_ct = defaultdict(int)
    advanced_ct = defaultdict(int)
    allclear_ct = defaultdict(int)
    maverick_ct = defaultdict(int)
    subt_rr = deque()
    tgs_rr = deque()
    move_comments = {}

    output_sections = {c: defaultdict(list) for c in CONTRACTORS}

    # ROUND ROBIN for overflow
    rr_state = deque(['Subterraneus Installs', 'TGS Fiber'])

    for job in jobs:
        key = (job['date'], job['time'])
        cty = job['city']
        group = get_city_group(cty)
        orig = job['contractor']

        # ----- 5 Gig Conversion override -----
        if job['job_type'].lower() == "5 gig conversion":
            if orig != 'Socket':
                move_comments[job['line']] = f"FORCED to Socket by job type rule (was {orig})"
            output_sections['Socket'][key].append(job['line'])
            continue

        # FAYETTE -> Maverick
        if group == 'maverick':
            if orig != 'Maverick':
                move_comments[job['line']] = f"MOVED from {orig}"
            output_sections['Maverick'][key].append(job['line'])
            continue

        # Tex-Star: Columbia (limit 8)
        if group == 'texstar':
            if orig == 'Tex-Star Communications' and texstar_ct[key] < TEXSTAR_LIMIT:
                output_sections['Tex-Star Communications'][key].append(job['line'])
                texstar_ct[key] += 1
            elif texstar_ct[key] < TEXSTAR_LIMIT:
                if orig != 'Tex-Star Communications':
                    move_comments[job['line']] = f"MOVED from {orig}"
                output_sections['Tex-Star Communications'][key].append(job['line'])
                texstar_ct[key] += 1
            else:
                # --- Begin new overflow logic ---
                subt_count = len(output_sections['Subterraneus Installs'][key])
                tgs_count = len(output_sections['TGS Fiber'][key])
                # First fill either company up to 8, whichever has fewer (ties go to SubT for consistency)
                if subt_count < 8 or tgs_count < 8:
                    if subt_count <= tgs_count and subt_count < 8:
                        target = 'Subterraneus Installs'
                    elif tgs_count < 8:
                        target = 'TGS Fiber'
                    else:
                        target = 'Subterraneus Installs'  # If both are 8 (shouldn't happen here)
                else:
                    # Both have at least 8: alternate to keep them even
                    if subt_count <= tgs_count:
                        target = 'Subterraneus Installs'
                    else:
                        target = 'TGS Fiber'
                move_comments[job['line']] = f"MOVED from {orig} (TexStar overflow, balanced to {target})"
                output_sections[target][key].append(job['line'])
            continue

        # Pifer: Rolla/Lebanon (limit 3)
        if group == 'pifer':
            if orig == 'Pifer Quality Communications' and pifer_ct[key] < PIFER_LIMIT:
                output_sections['Pifer Quality Communications'][key].append(job['line'])
                pifer_ct[key] += 1
            elif pifer_ct[key] < PIFER_LIMIT:
                if orig != 'Pifer Quality Communications':
                    move_comments[job['line']] = f"MOVED from {orig}"
                output_sections['Pifer Quality Communications'][key].append(job['line'])
                pifer_ct[key] += 1
            else:
                target = rr_state[0]
                if orig != target:
                    move_comments[job['line']] = f"MOVED from {orig}"
                    output_sections[target][key].append(job['line'])
                rr_state.rotate(-1)
            continue

        # Jefferson City: Advanced (1), then All Clear (1), then round robin
        if group == 'jeffcity':
            if advanced_ct[key] < ADVANCED_LIMIT:
                if orig != 'Advanced Electric':
                    move_comments[job['line']] = f"MOVED from {orig}"
                output_sections['Advanced Electric'][key].append(job['line'])
                advanced_ct[key] += 1
            elif allclear_ct[key] < ALLCLEAR_LIMIT:
                if orig != 'All Clear':
                    move_comments[job['line']] = f"MOVED from {orig}"
                output_sections['All Clear'][key].append(job['line'])
                allclear_ct[key] += 1
            else:
                target = rr_state[0]
                if orig != target:
                    move_comments[job['line']] = f"MOVED from {orig}"
                output_sections[target][key].append(job['line'])
                rr_state.rotate(-1)
            continue

        # SubT: All O'Fallon and Kirksville
        if group == 'subt':
            if orig != 'Subterraneus Installs':
                move_comments[job['line']] = f"MOVED from {orig}"
            output_sections['Subterraneus Installs'][key].append(job['line'])
            continue

        # TGS: ALL Clinton, Oak Grove, Warrensburg, Sedalia
        if group == 'tgs':
            if orig != 'TGS Fiber':
                move_comments[job['line']] = f"MOVED from {orig}"
            output_sections['TGS Fiber'][key].append(job['line'])
            continue

        # Anything else: Uncategorized
        if orig != 'Uncategorized':
            move_comments[job['line']] = f"MOVED from {orig}"
        output_sections['Uncategorized'][key].append(job['line'])

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

    status = tk.StringVar(value="Waiting for file...")
    status_label = tk.Label(root, textvariable=status, fg="blue", font=("Segoe UI", 11))
    status_label.pack(pady=(0, 16))

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
