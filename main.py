import requests
from bs4 import BeautifulSoup
import os
from dotenv import load_dotenv
import pandas as pd

# 1. Load credentials from .env
load_dotenv()

def scrape_to_excel(id, session, is_lecturer=False):
    """Scrapes data and returns a list of dictionaries for Excel"""
    type_char = 'N' if is_lecturer else 'G'
    url = f"https://planzajec.uek.krakow.pl/index.php?typ={type_char}&id={id}&okres=2"

    response = session.get(url)
    response.raise_for_status()
    response.encoding = "utf-8"
    
    soup = BeautifulSoup(response.text, 'html.parser')
    excel_data = []
    
    table_rows = soup.find_all('tr')

    for row in table_rows[1:]:
        columns = row.find_all('td')
        if len(columns) >= 6:
            date_str = columns[0].text.strip()
            if not date_str: continue # Skip empty rows
            
            day_time_str = columns[1].text.strip()
            subject = columns[2].text.strip()
            class_type = columns[3].text.strip()
            
            # Logic for Group vs Lecturer tables
            if not is_lecturer:
                teacher = columns[4].text.strip()
                location_td = columns[5]
            else:
                location_td = columns[4]
                teacher = "N/A"

            # Extract location text or link
            link = location_td.find("a")
            location = link["href"] if link and link.get("href") else location_td.text.strip()

            # Clean up the time string (e.g., "08:00-09:30")
            time_only = day_time_str.split('(')[0].strip() if '(' in day_time_str else day_time_str

            excel_data.append({
                "Date": date_str,
                "Time": time_only,
                "Subject": subject,
                "Type": class_type,
                "Location": location,
                "Teacher": teacher
            })

    return excel_data

def run_update():
    username = os.getenv("UEK_LOGIN")
    password = os.getenv("UEK_PASSWORD")
    
    if not username or not password:
        print("Error: UEK_LOGIN or UEK_PASSWORD not set in .env")
        return

    session = requests.Session()
    session.auth = (username, password)
    
    # PUT YOUR GROUP ID HERE
    my_group_id = "252671" 
    
    print(f"Fetching data for {my_group_id}...")
    data = scrape_to_excel(my_group_id, session)

    if data:
        # 1. Convert your list of dictionaries into a Pandas Table (DataFrame)
        df = pd.DataFrame(data)
        
        # 2. Define the filename
        filename = f"uek_schedule_{my_group_id}.xlsx"
        
        # 3. Use the ExcelWriter to create a real .xlsx file directly on your computer
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Schedule')
            
            # Auto-adjust columns width for perfect formatting
            for column in df:
                column_width = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                # chr(65) is 'A', so this maps 0->A, 1->B, 2->C, etc.
                writer.sheets['Schedule'].column_dimensions[chr(65 + col_idx)].width = column_width + 2
                
        print(f"Success! Saved beautifully formatted Excel file: {filename}")
    else:
        print("No data found.")

if __name__ == "__main__":
    run_update()