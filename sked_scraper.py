import sys, os, re, hashlib, requests
from datetime import datetime, timedelta
import pytz
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

# Set the correct New Zealand Timezone
NZ_TZ = pytz.timezone('Pacific/Auckland')

def parse_iso_time(time_str):
    if not time_str: return None
    t = str(time_str).strip().upper()
    if any(word in t for word in ["ALL DAY", "TBC", "EVENING", "TBA", "TBD", "NO REHEARSAL", "EXAMS"]): return None
    try:
        if ':' not in t: t = re.sub(r'(\d+)', r'\1:00', t)
        return datetime.strptime(t, "%I:%M%p")
    except: return None

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
        all_text = [t.get_text(" ", strip=True) for t in soup.find_all(['p', 'span', 'h1', 'td']) if t.get_text(strip=True)]
        
        cal_name = next((x for x in all_text if "SCHEDULE" in x.upper() or len(x) > 10), "Schedule")
        cal_name = cal_name.split('\n')[0].strip()
        year = next((re.search(r'\d{4}', x).group(0) for x in all_text[:20] if re.search(r'\d{4}', x)), "2026")

        if os.path.exists(master_path):
            with open(master_path, 'rb') as f: master_cal = Calendar.from_ical(f.read())
        else:
            master_cal = Calendar()
            master_cal.add('x-wr-calname', 'Master Rehearsal Schedule')
            master_cal.add('x-wr-timezone', 'Pacific/Auckland')
            master_cal.add('version', '2.0')

        for table in soup.find_all('table'):
            rows = table.find_all('tr')
            matrix = {}
            # Step 1: Map the table headers to find exactly where 'LOCATION' and 'TIME' are
            col_map = {"DATE": 1, "TIME": 2, "LOCATION": 3, "NOTES": 4} 
            
            for r_idx, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                c_idx = 0
                for cell in cells:
                    while (r_idx, c_idx) in matrix: c_idx += 1
                    content = cell.get_text(" ", strip=True)
                    # Update mapping if we see headers
                    if "DATE" in content.upper(): col_map["DATE"] = c_idx
                    if "TIME" in content.upper(): col_map["TIME"] = c_idx
                    if "LOCATION" in content.upper(): col_map["LOCATION"] = c_idx
                    if "NOTES" in content.upper(): col_map["NOTES"] = c_idx
                    
                    rowspan, colspan = int(cell.get('rowspan', 1)), int(cell.get('colspan', 1))
                    for r in range(r_idx, r_idx + rowspan):
                        for c in range(c_idx, c_idx + colspan): matrix[(r, c)] = content
                    c_idx += colspan

                # Step 2: Extract data based on the discovered map
                date_val = matrix.get((r_idx, col_map["DATE"]), "")
                m = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+([a-zA-Z]{3,9})', date_val)
                
                if m:
                    day, mon = m.group(1).zfill(2), m.group(2)[:3]
                    try:
                        iso_date = datetime.strptime(f"{day} {mon} {year}", "%d %b %Y")
                        time_val = matrix.get((r_idx, col_map["TIME"]), "")
                        loc_val = matrix.get((r_idx, col_map["LOCATION"]), "TBA")
                        note_val = matrix.get((r_idx, col_map["NOTES"]), "Rehearsal")

                        if any(x in time_val.upper() for x in ["NO REHEARSAL", "EXAMS"]): continue

                        event = Event()
                        event.add('summary', f"{cal_name}: {note_val}")
                        event.add('location', loc_val) # This is now pulling from the specific LOCATION column
                        event.add('uid', hashlib.md5(f"{cal_name}{iso_date}{note_val}".encode()).hexdigest() + "@bot")
                        
                        start_time_obj = parse_iso_time(re.split(r'\s*-\s*', str(time_val))[0])
                        if start_time_obj:
                            start_dt = NZ_TZ.localize(datetime.combine(iso_date.date(), start_time_obj.time()))
                            event.add('dtstart', start_dt)
                            event.add('dtend', start_dt + timedelta(minutes=50))
                        else:
                            event.add('dtstart', iso_date.date())
                            event.add('dtend', iso_date.date() + timedelta(days=1))
                        
                        master_cal.add_component(event)
                    except: continue

        with open(master_path, 'wb') as f: f.write(master_cal.to_ical())
        print(f"Successfully merged {cal_name} with specific location mapping.")

    except Exception as e: print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1: run_conversion(sys.argv[1])
