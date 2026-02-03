import sys, os, re, hashlib, requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

def parse_iso_time(time_str):
    if not time_str: return None
    t = str(time_str).strip().upper()
    if any(word in t for word in ["ALL DAY", "TBC", "EVENING", "TBA"]): return t
    try:
        if ':' not in t: t = re.sub(r'(\d+)', r'\1:00', t)
        return datetime.strptime(t, "%I:%M%p").strftime("%H:%M:%S")
    except:
        try: return datetime.strptime(t, "%H:%M").strftime("%H:%M:%S")
        except: return t

def run_conversion(source):
    try:
        response = requests.get(source, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # IMPROVED: Smarter Metadata Extraction
        all_text = [t.get_text(" ", strip=True) for t in soup.find_all(['p', 'span', 'h1', 'td']) if t.get_text(strip=True)]
        cal_name = next((x for x in all_text if len(x) > 3), "General_Schedule")
        year = next((re.search(r'\d{4}', x).group(0) for x in all_text[:20] if re.search(r'\d{4}', x)), str(datetime.now().year))

        cal = Calendar()
        cal.add('x-wr-calname', cal_name)
        cal.add('prodid', f'-//Global// {cal_name} //EN')
        cal.add('version', '2.0')

        found_events = 0
        for table in soup.find_all('table'):
            rows = table.find_all('tr')
            matrix = {}
            for r_idx, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                c_idx = 0
                for cell in cells:
                    while (r_idx, c_idx) in matrix: c_idx += 1
                    content = cell.get_text(" ", strip=True)
                    for r in range(r_idx, r_idx + int(cell.get('rowspan', 1))):
                        for c in range(c_idx, c_idx + int(cell.get('colspan', 1))):
                            matrix[(r, c)] = content
                    c_idx += int(cell.get('colspan', 1))

                # IMPROVED: Scan columns 0, 1, and 2 for a valid date
                iso_date = None
                date_col = -1
                for col_to_check in [0, 1, 2]:
                    val = matrix.get((r_idx, col_to_check), "")
                    match = re.search(r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', val, re.I)
                    if match:
                        day, mon = match.group(2).zfill(2), match.group(3)[:3]
                        iso_date = datetime.strptime(f"{day} {mon} {year}", "%d %b %Y").strftime("%Y-%m-%d")
                        date_col = col_to_check
                        break
                
                if iso_date:
                    note = matrix.get((r_idx, date_col + 3), matrix.get((r_idx, date_col + 2), "Event"))
                    time_val = matrix.get((r_idx, date_col + 1), "")
                    
                    event = Event()
                    event.add('summary', f"{cal_name}: {note}")
                    event.add('uid', hashlib.md5(f"{cal_name}{iso_date}{note}".encode()).hexdigest() + "@bot")
                    
                    times = re.split(r'\s*-\s*', str(time_val))
                    start_iso = parse_iso_time(times[0])
                    
                    if start_iso and ":" in str(start_iso):
                        event.add('dtstart', datetime.strptime(f"{iso_date} {start_iso}", "%Y-%m-%d %H:%M:%S"))
                        event.add('dtend', event['dtstart'].dt + timedelta(hours=1))
                    else:
                        event.add('dtstart', datetime.strptime(iso_date, "%Y-%m-%d").date())
                        event.add('dtend', datetime.strptime(iso_date, "%Y-%m-%d").date() + timedelta(days=1))
                    
                    cal.add_component(event)
                    found_events += 1

        output_file = re.sub(r'\W+', '', cal_name) + ".ics"
        with open(output_file, 'wb') as f: f.write(cal.to_ical())
        print(f"Finished: {found_events} events saved to {output_file}")

    except Exception as e: print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1: run_conversion(sys.argv[1])
