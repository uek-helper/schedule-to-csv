import streamlit as st
import requests
from bs4 import BeautifulSoup
import io
import pandas as pd

def scrape_data(id, username, password, is_lecturer):
    session = requests.Session()
    session.auth = (username, password)
    type_char = 'N' if is_lecturer else 'G'
    url = f"https://planzajec.uek.krakow.pl/index.php?typ={type_char}&id={id}&okres=2"
    
    response = session.get(url)
    if response.status_code == 401:
        return None, "Login Failed: Check Credentials"
    
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')
    
    excel_data = []
    table_rows = soup.find_all('tr')
    
    for row in table_rows[1:]:
        columns = row.find_all('td')
        
        if len(columns) >= 5: 
            date_str = columns[0].text.strip()
            if not date_str: continue 
            
            teacher = ""
            group = ""
            location = ""
            day_of_week = ""
            start = ""
            end = ""
            
            if not is_lecturer and len(columns) >= 6:
                teacher = columns[4].text.strip()
                location = columns[5].text.strip()
            elif is_lecturer and len(columns) >= 6:
                location = columns[4].text.strip()
                group = columns[5].text.strip()
            elif is_lecturer and len(columns) == 5:
                location = columns[4].text.strip()
            
            raw_time = columns[1].text.strip() 
            clean_time = raw_time.split('(')[0].strip() 
            
            if " " in clean_time:
                day_of_week, time_range = clean_time.split(' ', 1)
            else:
                time_range = clean_time
                
            if "-" in time_range:
                start, end = time_range.split('-', 1)
            else:
                start = time_range

            # Build the base dictionary shared by both
            entry = {
                "Date": date_str,
                "Day": day_of_week,
                "Starting": start.strip(),
                "Ending": end.strip(),
                "Subject": columns[2].text.strip(),
                "Type": columns[3].text.strip(),
                "Location": location
            }
            
            # Strictly assign the unique column based on user type
            if is_lecturer:
                entry["Group"] = group
            else:
                entry["Teacher"] = teacher
                
            excel_data.append(entry)
            
    return excel_data, None

# -- STREAMLIT UI --
st.set_page_config(page_title="UEK Schedule Exporter", page_icon="🎓")



st.title("🎓 UEK Schedule to Excel")
st.markdown("Enter your ID to generate a downloadable Excel file (.xlsx).")

user_type = st.radio("I am a:", ["Student/Group", "Lecturer"])
group_id = st.text_input("Enter ID (e.g., 188141)", placeholder="Look at your plan URL for 'id=...'")

if st.button("Generate Schedule"):
    username = st.secrets["UEK_LOGIN"]
    password = st.secrets["UEK_PASSWORD"]
    is_lecturer = (user_type == "Lecturer")
    
    with st.spinner("Scraping UEK Portal..."):
        data, error = scrape_data(group_id, username, password, is_lecturer)
        
        if error:
            st.error(error)
        elif not data:
            st.warning("No schedule found for this ID.")
        else:
            df = pd.DataFrame(data)
            buffer = io.BytesIO()
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Schedule')
                
                workbook = writer.book
                worksheet = writer.sheets['Schedule']
                
                for column in df:
                    column_width = max(df[column].astype(str).map(len).max(), len(column))
                    col_idx = df.columns.get_loc(column)
                    worksheet.column_dimensions[chr(65 + col_idx)].width = column_width + 2
            
            st.success("Schedule successfully formatted!")
            
            st.download_button(
                label="📥 Download Excel File (.xlsx)",
                data=buffer.getvalue(),
                file_name=f"uek_schedule_{group_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.image("picture.png", caption="UEK Schedule Exporter", use_container_width=True)