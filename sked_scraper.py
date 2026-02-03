import sys, os, re, hashlib, requests
from datetime import datetime, timedelta
import pytz
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

NZ_TZ = pytz.timezone('Pacific/Auckland')

def parse_iso_time(time_str):
    if not time_str: return None
    t = str(time_str).strip().upper()
    # Robust regex to find 12:25pm, 1:15pm, 1pm, 6:00pm, etc.
    match = re.search(r'(\d{1,2})(?::(\d{2}))?\s*([AP]M)', t)
    if match:
        hr, mn, ampm = match.groups()
        mn = mn if mn else "00"
        # Convert to 24hr format for processing
        return datetime.strptime(f"{hr}:{mn}{ampm}", "%I:%M%p")
    return None

def clean_text(text):
    if not text: return ""
    return " ".join(text.replace('\xa0', ' ').split()).strip()

def run_conversion(source, output_folder="calendars"):
    if not os.path.exists(output_folder): os.makedirs(output_folder)
    master_path = os.path.join(output_folder, "Master_Schedule.ics")

    try:
        if source.startswith('http'):
            response = requests.get(source, timeout=15)
            html_content = response.text
        else:
            with open(source, 'r', encoding='utf-8') as f: html_content = f.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        year = "2026"
        cal_name = "Taharoto Strings"

        if os.path.exists(master_path):
            with open(master_path, 'rb') as f: master_cal = Calendar.from_ical(f.read())
        else:
            master_cal = Calendar()
            master_cal.add('x-wr-calname', 'Master Schedule')
            master_cal.add('x-wr-timezone', 'Pacific/Auckland')
            master_cal.add('version', '2.0')

        for row in soup.find_all('tr'):
            cells = [clean_text(c.get_text(" ", strip=True)) for c in row.find_all(['td', 'th'])]
            if not cells: continue
            row_text = " | ".join(cells)
            
            # 1. DATE EXTRACTION
            date_matches = re.findall(r'(\d{1,2})(?:st|nd|rd|th)?\s+([a-zA-Z]{3,9})', row_text, re.IGNORECASE)
            if not date_matches: continue

            found_dates = []
            for d, m in date_matches:
                try: found_dates.append(datetime.strptime(f"{d.zfill(2)} {m[:3]} {year}", "%d %b %Y"))
                except: continue
            if not found_dates: continue

            start_date = found_dates[0]
            end_date = found_dates[-1]

            # 2. TIME EXTRACTION - The "Harvest All" Method
            # This finds all instances of times like 3:30pm and 6:00pm in the row
            all_times = re.findall(r'\d{1,2}(?::\d{2})?\s*[AP]M', row_text, re.IGNORECASE)
            
            # Map location and note based on column positions
            loc_val = cells[3] if len(cells) > 3 else "WGHS EC"
            note_val = cells[-1]

            if "NO REHEARSAL" in row_text.upper() and "KBB" not in row_text.upper(): continue

            event = Event()
            event.add('summary', f"{cal_name}: {note_val}")
            event.add('location', loc_val)
            event.add('uid', hashlib.md5(f"{cal_name}{start_date}{row_text}".encode()).hexdigest() + "@bot")

            if all_times:
                # TIMED EVENT
                t_start = parse_iso_time(all_times[0])
                dt_start = NZ_TZ.localize(datetime.combine(start_date.date(), t_start.time()))
                event.add('dtstart', dt_start)
                
                # If a second time (end time) was found in the row, use it!
                if len(all_times) > 1:
                    t_end = parse_iso_time(all_times[-1]) # Take the last time found as end time
                    dt_end = NZ_TZ.localize(datetime.combine(start_date.date(), t_end.time()))
                    event.add('dtend', dt_end)
                else:
                    # Default 50 mins if no end time is written
                    event.add('dtend', dt_start + timedelta(minutes=50))
            else:
                # ALL-DAY EVENT
                event.add('dtstart', start_date.date())
                event.add('dtend', end_date.date() + timedelta(days=1))

            master_cal.add_component(event)
            print(f"Synced: {start_date.date()} | Time: {all_times} | {note_val}")

        with open(master_path, 'wb') as f: f.write(master_cal.to_ical())

    except Exception as e: print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1: run_conversion(sys.argv[1])
