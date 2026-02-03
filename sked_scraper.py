import sys
import os
import re
import hashlib
import requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

def parse_iso_time(time_str):
    if not time_str: return None
    t = time_str.strip().upper()
    if t in ["ALL DAY", "TBC", "EVENING", "TBD", "TBA"]: return t
    try:
        if ':' not in t: t = re.sub(r'(\d+)', r'\1:00', t)
        return datetime.strptime(t, "%I:%M%p").strftime("%H:%M:%S")
    except ValueError:
        try: return datetime.strptime(t, "%H:%M").strftime("%H:%M:%S")
        except: return t

def get_uid(owner, date_str, note_str):
    seed = f"{owner}-{date_str}-{note_str}".lower().strip()
    return hashlib.md5(seed.encode()).hexdigest() + "@script.gen"

def run_conversion(source):
    try:
        if source.startswith('http'):
            response = requests.get(source, timeout=10)
            response.raise_for_status()
            html_content = response.text
            # Use a temporary name until we find the real one in the HTML
            base_name = "temp_schedule"
        else:
            with open(source, 'r', encoding='utf-8') as f:
                html_content = f.read()
            base_name = os.path.splitext(source)[0]

        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Get Calendar Name from the first significant text found
        text_elements = [t.get_text(strip=True) for t in soup.find_all(['p', 'span', 'h1', 'h2']) if t.get_text(strip=True)]
        cal_name = text_elements[0] if text_elements else "Schedule"
        
        # Find the year
        year_match = re.search(r'\d{4}', " ".join(text_elements[:10]))
        year = year_match.group(0) if year_match else str(datetime.now().year)

        cal = Calendar()
        cal.add('prodid', f'-//Generic// {cal_name} //EN')
        cal.add('version', '2.0')
        cal.add('x-wr-calname', cal_name)

        tables = soup.find_all('table')
        for table in tables:
            rows = table.find_all('tr')
            matrix = {}
            current_section = "General"

            for r_idx, row in enumerate(rows):
                if row.get('colspan') or len(row.find_all('td')) == 1:
                    current_section = row.get_text(strip=True)
                    continue

                cols = row.find_all(['td', 'th'])
                c_idx = 0
                for cell in cols:
                    while (r_idx, c_idx) in matrix: c_idx += 1
                    content = cell.get_text(strip=True)
                    rowspan, colspan = int(cell.get('rowspan', 1)), int(cell.get('colspan', 1))
                    for r in range(r_idx, r_idx + rowspan):
                        for c in range(c_idx, c_idx + colspan):
                            matrix[(r, c)] = content
                    c_idx += colspan

                raw_date = matrix.get((r_idx, 1))
                raw_time = matrix.get((r_idx, 2))
                note = matrix.get((r_idx, 4)) or "Scheduled Event"
                
                if not raw_date or "DATE" in raw_date.upper(): continue

                try:
                    d_parts = raw_date.split()
                    day, mon = d_parts[1].zfill(2), d_parts[2][:3]
                    iso_date = datetime.strptime(f"{day} {mon} {year}", "%d %b %Y").strftime("%Y-%m-%d")
                except: continue

                event = Event()
                event.add('summary', f"{cal_name}: {note}")
                event.add('location', matrix.get((r_idx, 3), ""))
                event.add('description', f"Section: {current_section}\nRef: {matrix.get((r_idx, 0))}")
                event.add('uid', get_uid(cal_name, iso_date, note))

                times = re.split(r'\s*-\s*', str(raw_time))
                start_iso = parse_iso_time(times[0])
                end_iso = parse_iso_time(times[1]) if len(times) > 1 else None

                try:
                    if start_iso and ":" in start_iso:
                        start_dt = datetime.strptime(f"{iso_date} {start_iso}", "%Y-%m-%d %H:%M:%S")
                        event.add('dtstart', start_dt)
                        if end_iso and ":" in end_iso:
                            event.add('dtend', datetime.strptime(f"{iso_date} {end_iso}", "%Y-%m-%d %H:%M:%S"))
                        else:
                            event.add('dtend', start_dt + timedelta(hours=1))
                    else:
                        d = datetime.strptime(iso_date, "%Y-%m-%d").date()
                        event.add('dtstart', d)
                        event.add('dtend', d + timedelta(days=1))
                    cal.add_component(event)
                except: continue

        # Save file with a clean name based on the calendar title
        clean_name = re.sub(r'[\\/*?:"<>|]', "", cal_name).replace(" ", "_")
        output_file = f"{clean_name}.ics"
        with open(output_file, 'wb') as f:
            f.write(cal.to_ical())
        print(f"Success: Created {output_file}")

    except Exception as e:
        print(f"Error processing {source}: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_conversion(sys.argv[1])
