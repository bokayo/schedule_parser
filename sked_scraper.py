import sys, os, re, hashlib, requests
from datetime import datetime, timedelta
import pytz
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

NZ_TZ = pytz.timezone('Pacific/Auckland')

def parse_iso_time(time_str):
    if not time_str: return None
    t = str(time_str).strip().upper()
    match = re.search(r'(\d{1,2})(?::(\d{2}))?\s*([AP]M)', t)
    if match:
        hr, mn, ampm = match.groups()
        mn = mn if mn else "00"
        return datetime.strptime(f"{hr}:{mn}{ampm}", "%I:%M%p")
    return None

def clean_text(text):
    if not text: return ""
    # Remove non-breaking spaces and collapse all whitespace
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

        # This logic mimics the JSON extractor: it processes every table row individually
        for row in soup.find_all('tr'):
            # Combine all text in the row into one searchable string
            cells = [clean_text(c.get_text(" ", strip=True)) for c in row.find_all(['td', 'th'])]
            if not cells: continue
            
            row_text = " | ".join(cells)
            
            # 1. FIND ALL DATES IN THE ROW
            # This regex captures "15th May", "16th May", "May 15", etc.
            date_matches = re.findall(r'(\d{1,2})(?:st|nd|rd|th)?\s+([a-zA-Z]{3,9})', row_text, re.IGNORECASE)
            
            if not date_matches:
                continue

            # Convert matches to datetime objects
            found_dates = []
            for d, m in date_matches:
                try:
                    found_dates.append(datetime.strptime(f"{d.zfill(2)} {m[:3]} {year}", "%d %b %Y"))
                except: continue
            
            if not found_dates: continue

            # Define Start and End Date
            start_date = found_dates[0]
            end_date = found_dates[-1]

            # 2. IDENTIFY TIME, LOCATION, AND NOTE
            # We look for a clock time (e.g. 12:25pm)
            time_match = re.search(r'(\d{1,2})(?::(\d{2}))?\s*([AP]M)', row_text, re.IGNORECASE)
            
            # Heuristic for Note/Location based on common table structure
            # Summary is usually the last thing in the row
            note_val = cells[-1]
            # Location is usually the second to last thing in the row
            loc_val = cells[-2] if len(cells) > 3 else "WGHS"

            # Skip if it's an explicit "No Rehearsal" row
            if "NO REHEARSAL" in row_text.upper() and "KBB" not in row_text.upper():
                continue

            event = Event()
            event.add('summary', f"{cal_name}: {note_val}")
            event.add('location', loc_val)
            event.add('uid', hashlib.md5(f"{cal_name}{start_date}{row_text}".encode()).hexdigest() + "@bot")

            if time_match:
                # TIMED EVENT
                start_obj = parse_iso_time(time_match.group())
                dt_start = NZ_TZ.localize(datetime.combine(start_date.date(), start_obj.time()))
                event.add('dtstart', dt_start)
                # Default to 1 hour if no end time found
                event.add('dtend', dt_start + timedelta(hours=1))
            else:
                # ALL-DAY EVENT (May 15-16)
                event.add('dtstart', start_date.date())
                # End date is the morning of the day AFTER the event ends
                event.add('dtend', end_date.date() + timedelta(days=1))
                # Add the 'All Day' or 'Workshop' info to the summary if it was in the time column
                if "ALL DAY" in row_text.upper():
                    event['summary'] = f"{event['summary']} (All Day)"

            master_cal.add_component(event)
            print(f"Captured: {start_date.date()} to {end_date.date()} - {note_val}")

        with open(master_path, 'wb') as f:
            f.write(master_cal.to_ical())

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_conversion(sys.argv[1])
