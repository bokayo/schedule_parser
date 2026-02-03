import sys, os, re, hashlib, requests
from datetime import datetime, timedelta
import pytz
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

# Set the correct New Zealand Timezone
NZ_TZ = pytz.timezone('Pacific/Auckland')

def parse_iso_time(time_str):
    """Helper to turn '12:25pm' into a datetime object."""
    if not time_str: return None
    t = str(time_str).strip().upper()
    if any(word in t for word in ["ALL DAY", "TBC", "TBA", "TBD", "NO REHEARSAL", "EXAMS"]): return None
    try:
        # Standardize formats like '1pm' to '1:00pm'
        if ':' not in t: t = re.sub(r'(\d+)', r'\1:00', t)
        return datetime.strptime(t, "%I:%M%p")
    except: return None

def run_conversion(source, output_folder="calendars"):
    if not os.path.exists(output_folder): os.makedirs(output_folder)
    master_path = os.path.join(output_folder, "Ensemble_Schedule.ics")

    try:
        if source.startswith('http'):
            response = requests.get(source, timeout=15)
            html_content = response.text
        else:
            with open(source, 'r', encoding='utf-8') as f: html_content = f.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        all_text = [t.get_text(" ", strip=True) for t in soup.find_all(['p', 'span', 'h1', 'td']) if t.get_text(strip=True)]
        
        cal_name = next((x for x in all_text if "SCHEDULE" in x.upper() or len(x) > 10), "Schedule")
        cal_name = cal_name.split('\n')[0].strip()
        year = next((re.search(r'\d{4}', x).group(0) for x in all_text[:20] if re.search(r'\d{4}', x)), "2026")

        if os.path.exists(master_path):
            with open(master_path, 'rb') as f: master_cal = Calendar.from_ical(f.read())
        else:
            master_cal = Calendar()
            master_cal.add('x-wr-calname', 'Ensemble Schedule')
            master_cal.add('x-wr-timezone', 'Pacific/Auckland')
            master_cal.add('version', '2.0')

        for table in soup.find_all('table'):
            rows = table.find_all('tr')
            matrix = {}
            col_map = {}

            for r_idx, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                curr_col = 0
                for cell in cells:
                    while (r_idx, curr_col) in matrix: curr_col += 1
                    content = cell.get_text(" ", strip=True)
                    if "DATE" in content.upper(): col_map['date'] = curr_col
                    if "TIME" in content.upper(): col_map['time'] = curr_col
                    if "LOCATION" in content.upper(): col_map['loc'] = curr_col
                    if "NOTES" in content.upper(): col_map['note'] = curr_col
                    
                    rowspan, colspan = int(cell.get('rowspan', 1)), int(cell.get('colspan', 1))
                    for r in range(r_idx, r_idx + rowspan):
                        for c in range(curr_col, curr_col + colspan): matrix[(r, c)] = content
                    curr_col += colspan

                if 'date' in col_map and 'time' in col_map:
                    date_val = matrix.get((r_idx, col_map['date']), "")
                    m = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+([a-zA-Z]{3,9})', date_val)
                    
                    if m:
                        day, mon = m.group(1).zfill(2), m.group(2)[:3]
                        try:
                            iso_date = datetime.strptime(f"{day} {mon} {year}", "%d %b %Y")
                            time_val = matrix.get((r_idx, col_map['time']), "")
                            loc_val = matrix.get((r_idx, col_map.get('loc', 99)), "TBA")
                            note_val = matrix.get((r_idx, col_map.get('note', 99)), "Rehearsal")

                            if any(x in time_val.upper() for x in ["NO REHEARSAL", "EXAMS"]): continue

                            # --- UPDATED TIME PARSING ---
                            # Split "12:25pm - 1:15pm" into ["12:25pm", "1:15pm"]
                            parts = re.split(r'\s*-\s*|\s*TO\s*', str(time_val).upper())
                            start_obj = parse_iso_time(parts[0])
                            end_obj = parse_iso_time(parts[1]) if len(parts) > 1 else None

                            if start_obj:
                                event = Event()
                                event.add('summary', f"{cal_name}: {note_val}")
                                event.add('location', loc_val)
                                event.add('uid', hashlib.md5(f"{cal_name}{iso_date}{note_val}".encode()).hexdigest() + "@bot")
                                
                                start_dt = NZ_TZ.localize(datetime.combine(iso_date.date(), start_obj.time()))
                                event.add('dtstart', start_dt)
                                
                                if end_obj:
                                    end_dt = NZ_TZ.localize(datetime.combine(iso_date.date(), end_obj.time()))
                                    event.add('dtend', end_dt)
                                else:
                                    # Fallback if only start time is provided
                                    event.add('dtend', start_dt + timedelta(minutes=50))
                                
                                master_cal.add_component(event)
                        except: continue

        with open(master_path, 'wb') as f: f.write(master_cal.to_ical())
        print(f"Sync complete. Locations and precise end-times included.")

    except Exception as e: print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1: run_conversion(sys.argv[1])
