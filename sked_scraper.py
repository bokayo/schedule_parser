import sys, os, re, hashlib, requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from icalendar import Calendar, Event

def parse_iso_time(time_str):
    if not time_str: return None
    t = str(time_str).strip().upper()
    if any(word in t for word in ["ALL DAY", "TBC", "EVENING", "TBA", "TBD"]): return t
    try:
        # Handles 12:25pm or 12pm
        if ':' not in t: t = re.sub(r'(\d+)', r'\1:00', t)
        return datetime.strptime(t, "%I:%M%p").strftime("%H:%M:%S")
    except: return None

def run_conversion(source, output_folder="calendars"):
    if not os.path.exists(output_folder): os.makedirs(output_folder)

    try:
        # Load from URL or Local File
        if source.startswith('http'):
            response = requests.get(source, timeout=15)
            html_content = response.text
        else:
            with open(source, 'r', encoding='utf-8') as f: html_content = f.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Identify the Calendar Name (e.g., TAHAROTO STRINGS 2026)
        title_element = soup.find(class_="c38") or soup.find('p')
        cal_name = title_element.get_text(strip=True) if title_element else "Schedule"
        year = "2026" # Hardcoded based on your specific file content

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
                    rowspan = int(cell.get('rowspan', 1))
                    colspan = int(cell.get('colspan', 1))
                    for r in range(r_idx, r_idx + rowspan):
                        for c in range(c_idx, c_idx + colspan):
                            matrix[(r, c)] = content
                    c_idx += colspan

                # TARGET: The Date Column (Column 1 in this file)
                date_val = matrix.get((r_idx, 1), "")
                # Regex to handle ordinals: matches "Wed 28th Jan" or "Wed 4th Feb"
                date_match = re.search(r'([a-zA-Z]{3})\s+(\d{1,2})(?:st|nd|rd|th)?\s+([a-zA-Z]{3,9})', date_val)
                
                if date_match:
                    day_num = date_match.group(2).zfill(2)
                    month_str = date_match.group(3)[:3]
                    iso_date = datetime.strptime(f"{day_num} {month_str} {year}", "%d %b %Y").strftime("%Y-%m-%d")
                    
                    # Extract Notes, Time, and Location based on columns 2, 3, 4
                    time_raw = matrix.get((r_idx, 2), "")
                    location = matrix.get((r_idx, 3), "")
                    notes = matrix.get((r_idx, 4), "Rehearsal")

                    # Handle "No Rehearsal" or merged status rows
                    if "No Rehearsal" in time_raw or "No Rehearsal" in date_val:
                        summary = f"{cal_name}: NO REHEARSAL"
                    else:
                        summary = f"{cal_name}: {notes}"

                    event = Event()
                    event.add('summary', summary)
                    event.add('location', location)
                    event.add('uid', hashlib.md5(f"{cal_name}{iso_date}{summary}".encode()).hexdigest() + "@bot")
                    
                    time_parts = re.split(r'\s*-\s*', time_raw)
                    start_iso = parse_iso_time(time_parts[0])
                    
                    if start_iso:
                        start_dt = datetime.strptime(f"{iso_date} {start_iso}", "%Y-%m-%d %H:%M:%S")
                        event.add('dtstart', start_dt)
                        event.add('dtend', start_dt + timedelta(minutes=50)) # Default 50m rehearsal
                    else:
                        d = datetime.strptime(iso_date, "%Y-%m-%d").date()
                        event.add('dtstart', d)
                        event.add('dtend', d + timedelta(days=1))
                    
                    cal.add_component(event)
                    found_events += 1

        # Overwrite file in the calendars/ folder
        clean_filename = re.sub(r'\W+', '_', cal_name).strip("_") + ".ics"
        save_path = os.path.join(output_folder, clean_filename)
        with open(save_path, 'wb') as f: f.write(cal.to_ical())
        print(f"Success: {found_events} events -> {save_path}")

    except Exception as e: print(f"Processing Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1: run_conversion(sys.argv[1])
