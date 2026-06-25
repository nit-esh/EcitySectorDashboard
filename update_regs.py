#!/usr/bin/env python3
"""
update_regs.py — Reads Karnataka All Events Count.xlsx (Numbers format),
extracts registration counts, and injects them into the dashboard HTML.

Run this script from Terminal whenever the Numbers file is updated:
    python3 update_regs.py

Requires: macOS (uses osascript to export Numbers → CSV)
"""

import subprocess, csv, os, re, json, sys, tempfile, glob, datetime

# ── Paths ────────────────────────────────────────────────────────────────────
BASE   = os.path.dirname(os.path.abspath(__file__))
HTML   = os.path.join(BASE, "isha_multi_centre_dashboard_v2.html")
TMPCSV = tempfile.mktemp(suffix=".csv")

# Auto-detect the Numbers/Excel file in Santhosha Data folder
# You can rename the file to anything — the script picks it up automatically
SANTHOSHA_DIR = os.path.join(BASE, "Santhosha Data")
_candidates = (
    glob.glob(os.path.join(SANTHOSHA_DIR, "*.xlsx")) +
    glob.glob(os.path.join(SANTHOSHA_DIR, "*.numbers")) +
    glob.glob(os.path.join(SANTHOSHA_DIR, "*.xls"))
)
# Filter to files that look like event count sheets (contains "Count" or "Event" or "Registration")
_match = [f for f in _candidates if re.search(r'count|event|registration', os.path.basename(f), re.I)]
if _match:
    XLSX = _match[0]
elif _candidates:
    XLSX = _candidates[0]   # fallback: just take the first file found
else:
    # If a file path is passed as argument, use that
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        XLSX = sys.argv[1]
    else:
        print(f"✗ No Numbers/Excel file found in: {SANTHOSHA_DIR}")
        print("  Either put the file there, or run: python3 update_regs.py /path/to/file.numbers")
        sys.exit(1)

print(f"  Using file: {os.path.basename(XLSX)}")

# ── Step 1: Export Numbers file → CSV via AppleScript ────────────────────────
def export_to_csv():
    script = f'''
tell application "Numbers"
    open POSIX file "{XLSX}"
    delay 2
    set theDoc to front document
    export theDoc to POSIX file "{TMPCSV}" as CSV
    close theDoc saving no
end tell
'''
    print(f"Exporting {os.path.basename(XLSX)} → CSV …")
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if r.returncode != 0:
        print(f"  ✗ AppleScript error: {r.stderr.strip()}")
        sys.exit(1)
    if not os.path.exists(TMPCSV):
        print("  ✗ CSV not created — is Numbers installed?")
        sys.exit(1)
    print(f"  ✓ Exported to {TMPCSV}")

# ── Step 2: Normalise helpers ─────────────────────────────────────────────────
CENTRE_PATTERNS = [
    (r'sadhguru sannidhi',         'sadhguru sannidhi'),
    (r'electronic city',           'electronic'),
    (r'kanakapura',                'kanakapura'),
    (r'sarjapur|sargapur',         'sarjapur'),
    (r'hsr layout|hsr',            'hsr'),
    (r'marathahalli|marathali',    'marathahalli'),
    (r'malleswaram|malleshwaram',  'malleswaram'),
    (r'vijayanagar|vijayanagara',  'vijayanagar'),
    (r'jayanagar|jaynagar',        'jayanagar'),
    (r'banaswadi|banasawadi',      'banaswadi'),
    (r'hebbal|hebbala',            'hebbal'),
    (r'indiranagar|indira nagar',  'indiranagar'),
    (r'mysuru|mysore',             'mysuru'),
    (r'hubballi|hubli',            'hubballi'),
    (r'koramangala',               'koramangala'),
    (r'chikkaballapur',            'chikkaballapur'),
]

PROG_PATTERNS = [
    (r'inner engineering|isha yoga|\bie\b',  'inner engineering'),
    (r'bhava spandana',                       'bhava spandana'),
    (r'shoonya',                              'shoonya intensive'),
    (r'samyama',                              'samyama'),
    (r'vairagya',                             'vairagya'),
    (r'eye care.*(upayoga|upa yoga)',           'eye care upa yoga'),
    (r'eye care.*(shanmukhi)',                 'eye care shanmukhi'),
    (r'eye care',                             'eye care'),
    (r'shanmukhi',                            'shanmukhi'),
    (r'angamardana',                          'angamardana'),
    (r'surya kriya.*surya shakti|surya shakti.*surya kriya', 'surya kriya surya shakti'),
    (r'surya kriya',                          'surya kriya'),
    (r'surya shakti',                         'surya shakti'),
    (r'yogasanas|yoga asana',                 'yogasanas'),
    (r'bhuta shuddhi',                        'bhuta shuddhi'),
    (r'bhastrika',                            'bhastrika'),
    (r'hatha yoga for children|hatha children','hatha children'),
    (r'hatha yoga',                           'hatha yoga'),
    (r'thoppukarnam',                         'thoppukarnam'),
    (r'jala neti',                            'jala neti'),
    (r'sutra neti',                           'sutra neti'),
    (r'nauli',                                'nauli'),
    (r'kapalbhati',                           'kapalbhati'),
    (r'trataka',                              'trataka'),
    (r'guru pooja|guru puja',                 'guru pooja'),
    (r'upa yoga|upayoga',                     'upa yoga'),
    (r'isha kriya',                           'isha kriya'),
    (r'chit shakti',                          'chit shakti'),
    (r'nada aradhana',                        'nada aradhana'),
    (r'satsang|sathsang',                     'satsang'),
    (r'isha janani',                          'isha janani'),
]

_MONTH_MAP = {
    'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06',
    'jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12'
}

def extract_start_date(text):
    """Extract ISO start date from names like '...Jul 2 - 5, 2026' → '2026-07-02'"""
    m = re.search(r'\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+(\d{1,2})\b', text.lower())
    ym = re.search(r'\b(20\d{2})\b', text)
    if m and ym:
        return f"{ym.group(1)}-{_MONTH_MAP[m.group(1)]}-{m.group(2).zfill(2)}"
    return None

def norm_centre(text):
    t = text.lower()
    for pat, key in CENTRE_PATTERNS:
        if re.search(pat, t):
            return key
    return None

def norm_prog(text):
    t = text.lower()
    for pat, key in PROG_PATTERNS:
        if re.search(pat, t):
            return key
    return None

# ── Step 3: Parse CSV ─────────────────────────────────────────────────────────
def parse_csv():
    regs = {}  # {centre: {prog: count}}
    skipped = 0
    matched = 0

    with open(TMPCSV, newline='', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        for row in reader:
            if not row:
                continue
            name = row[0].strip() if len(row) > 0 else ''
            # Count column — find first numeric column after col 0
            count_str = ''
            for col in row[1:]:
                col = col.strip()
                if col and re.match(r'^-?\d+(\.\d+)?$', col):
                    count_str = col
                    break
            if not name or not count_str:
                continue
            try:
                count = int(float(count_str))
            except:
                continue
            if count < 0:
                continue

            centre = norm_centre(name)
            prog   = norm_prog(name)

            if not centre or not prog:
                skipped += 1
                continue

            # Use date-qualified key when a date is found in the row name
            # (prevents summing multiple batches of the same programme)
            date = extract_start_date(name)
            prog_key = f"{prog}|{date}" if date else prog

            if centre not in regs:
                regs[centre] = {}
            # For dated keys, overwrite (last row wins); for generic keys, sum
            if date:
                regs[centre][prog_key] = count
            else:
                regs[centre][prog_key] = regs[centre].get(prog_key, 0) + count
            matched += 1

    print(f"  ✓ Parsed: {matched} entries matched, {skipped} rows skipped")
    return regs

# ── Step 4: Inject into dashboard HTML ───────────────────────────────────────
MARKER_START = "/* __LIVE_REGS_START__ */"
MARKER_END   = "/* __LIVE_REGS_END__ */"

def inject_html(regs):
    with open(HTML, 'r', encoding='utf-8') as f:
        html = f.read()

    regs_json = json.dumps(regs, indent=2, ensure_ascii=False)
    block = f"{MARKER_START}\nconst LIVE_REGS = {regs_json};\n{MARKER_END}"

    if MARKER_START in html:
        # Replace existing block
        pattern = re.escape(MARKER_START) + r'.*?' + re.escape(MARKER_END)
        html = re.sub(pattern, block, html, flags=re.DOTALL)
        print(f"  ✓ Updated LIVE_REGS in dashboard HTML")
    else:
        print(f"  ✗ Markers not found in HTML — has the dashboard been set up?")
        sys.exit(1)

    with open(HTML, 'w', encoding='utf-8') as f:
        f.write(html)

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print("\n── Karnataka Registration Counts Updater ──")
    export_to_csv()
    print("Parsing registration data …")
    regs = parse_csv()

    # Print summary
    total = sum(sum(p.values()) for p in regs.values())
    print(f"\n  Centres: {len(regs)}")
    for centre, progs in sorted(regs.items()):
        print(f"    {centre}: {progs}")
    print(f"  Total registrations captured: {total}")

    print(f"\nInjecting into dashboard …")
    inject_html(regs)

    # Cleanup
    try: os.remove(TMPCSV)
    except: pass

    print("\n✓ Done — open isha_multi_centre_dashboard_v2.html to see updated counts.\n")
