import sys, os, re, hashlib, requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

def parse_iso_time(time_str):
    if not time_str: return None
    t = str(time_str).strip().upper()
    if any(word in t for word in ["ALL DAY", "TBC", "EVENING", "TBA", "TBD"]): return t
    try:
        if ':' not in t: t = re.sub(r'(\d+)', r'\1:00', t)
        return datetime.strptime(t, "%I:%M%p").strftime("%H:%M:%S")
    except: return None

def run_conversion(source, output_folder="calendars"):
    # 1. Create directory explicitly
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        if source.startswith('http'):
            response = requests.get(source, timeout=15)
            html_content = response.text
        else:
            with open(source, 'r', encoding='utf-8') as f:
                html_content = f.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        all_text = [t.get_text(" ", strip=True) for t in soup.find_all(['p', 'span', 'h1', 'td']) if t.get_text(strip=True)]
        
        # Determine Calendar Name and Year
        cal_name = next((x for x in all_text if "SCHEDULE" in x.upper() or len(x) > 10), "Schedule")
        cal_name = cal_name.split('\n')[0].strip()
        year = next((re.search(r'\d{4}', x).group(0) for x in all_text[:20] if re.search(r'\d{4}', x)), str(datetime.now().year))

        cal = Calendar()
        cal.add('x-wr-calname', cal_name)
        cal.add('x-wr-timezone', 'Pacific/Auckland')
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
                    rowspan, colspan = int(cell.get('rowspan', 1)), int(cell.get('colspan', 1))
                    for r in range(r_idx, r_idx + rowspan):
                        for c in range(c_idx, c_idx + colspan):
                            matrix[(r, c)] = content
                    c_idx += colspan

                # Scan columns 0, 1, 2 for the date
                for col_idx in range(3):
                    val = matrix.get((r_idx, col_idx), "")
                    m = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+([a-zA-Z]{3,9})', val)
                    if m:
                        day, mon = m.group(1).zfill(2), m.group(2)[:3]
                        try:
                            iso_date = datetime.strptime(f"{day} {mon} {year}", "%d %b %Y").strftime("%Y-%m-%d")
                            time_val = matrix.get((r_idx, col_idx + 1), "")
                            loc = matrix.get((r_idx, col_idx + 2), "")
                            note = matrix.get((r_idx, col_idx + 3), "Scheduled Event")

                            event = Event()
                            event.add('summary', f"{cal_name}: {note}")
                            event.add('location', loc)
                            event.add('uid', hashlib.md5(f"{cal_name}{iso_date}{note}".encode()).hexdigest() + "@bot")
                            
                            start_iso = parse_iso_time(re.split(r'\s*-\s*', str(time_val))[0])
                            if start_iso and ":" in str(start_iso):
                                start_dt = datetime.strptime(f"{iso_date} {start_iso}", "%Y-%m-%d %H:%M:%S")
                                event.add('dtstart', start_dt)
                                event.add('dtend', start_dt + timedelta(hours=1))
                            else:
                                d = datetime.strptime(iso_date, "%Y-%m-%d").date()
                                event.add('dtstart', d)
                                event.add('dtend', d + timedelta(days=1))
                            
                            cal.add_component(event)
                            found_events += 1
                            break 
                        except: continue

        # Save to the specific calendars/ folder
        clean_name = re.sub(r'\W+', '_', cal_name).strip("_")
        output_path = os.path.join(output_folder, f"{clean_name}.ics")
        with open(output_path, 'wb') as f:
            f.write(cal.to_ical())
        print(f"Exported {found_events} events to {output_path}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_conversion(sys.argv[1])
