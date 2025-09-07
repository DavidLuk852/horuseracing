import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

# Function to scrape horse names from HKJC race card
def get_horse_names(race_date, racecourse, race_no):
    url = f"https://racing.hkjc.com/racing/information/Chinese/Racing/RaceCard.aspx?RaceDate={race_date}&Racecourse={racecourse}&RaceNo={race_no}"
    response = requests.get(url)
    
    if response.status_code != 200:
        return None
    
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table', class_='starter f_tac f_fs13 draggable hiddenable')
    
    if not table:
        return None
    
    horse_names = []
    tbody = table.find('tbody')
    if tbody:
        for tr in tbody.find_all('tr'):
            tds = tr.find_all('td')
            if len(tds) >= 4:
                horse_name = tds[3].get_text(strip=True)
                horse_names.append(horse_name)
    
    return horse_names

# Function to get last three race records from Excel
def get_last_three_records(horse_name, df):
    horse_records = df[df['馬名'] == horse_name].sort_values(by='日期', ascending=False)
    return horse_records.head(3)

# Main program
def main():
    race_date = input("Enter Race Date (YYYY/MM/DD): ")
    racecourse = input("Enter Racecourse (ST for Sha Tin, HV for Happy Valley): ").upper()

    excel_file = "HK_Racing_Auto_Data.xlsx"
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        print(f"Error: {excel_file} not found!")
        return

    output_file = f"Race_Records_{race_date.replace('/', '_')}_{racecourse}.xlsx"
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    race_no = 1
    while True:
        print(f"Processing Race {race_no}...")
        horse_names = get_horse_names(race_date, racecourse, race_no)
        
        if not horse_names:
            print(f"No data found for Race {race_no}. Stopping.")
            break
        
        # Prepare data for this race in a tall structure
        race_data = []
        for horse_name in horse_names:
            last_three = get_last_three_records(horse_name, df)
            if not last_three.empty:
                for idx, record in enumerate(last_three.itertuples(), 1):  # Start counting from 1
                    race_data.append({
                        '第幾仗': idx,  # 1 for last race, 2 for second last, 3 for third last
                        'Horse Name': horse_name,
                        '日期': record.日期,
                        '賽事名稱': record.賽事名稱,
                        '名次': record.名次,
                        '馬號': record.馬號,
                        '馬齡': record.馬齡,
                        '騎師': record.騎師,
                        '練馬師': record.練馬師,
                        '實際負磅': record.實際負磅,
                        '排位體重': record.排位體重,
                        '檔位': record.檔位,
                        '頭馬距離': record.頭馬距離,
                        '沿途走位': record.沿途走位,
                        '完成時間': record.完成時間,
                        '配備': record.配備,
                        '獨贏賠率': record.獨贏賠率,
                        '賽事班次': record.賽事班次,
                        '途程': record.途程,
                        '賽道': record.賽道,
                        '場地狀況': record.場地狀況,
                        '分段1': getattr(record, '分段1', None),
                        '分段2': getattr(record, '分段2', None),
                        '分段3': getattr(record, '分段3', None),
                        '分段4': getattr(record, '分段4', None),
                        '分段5': getattr(record, '分段5', None),
                        '分段6': getattr(record, '分段6', None),
                        '頭段': getattr(record, '頭段', None),
                        '末段': getattr(record, '末段', None),
                        '頭段指數': getattr(record, '頭段指數', None),
                        '末段指數': getattr(record, '末段指數', None),
                        '時間指數': getattr(record, '時間指數', None)
                    })
        
        # Convert to DataFrame and save to Excel sheet
        if race_data:
            race_df = pd.DataFrame(race_data)
            # Define column order with 第幾仗 as the first column
            columns = [
                '第幾仗', 'Horse Name', '日期', '賽事名稱', '名次', '馬號', '馬齡', '騎師', '練馬師', 
                '實際負磅', '排位體重', '檔位', '頭馬距離', '沿途走位', '完成時間', 
                '配備', '獨贏賠率', '賽事班次', '途程', '賽道', '場地狀況', 
                '分段1', '分段2', '分段3', '分段4', '分段5', '分段6', '頭段', '末段',
                '頭段指數', '末段指數', '時間指數'
            ]
            # Filter to only include columns that exist in race_df
            available_columns = [col for col in columns if col in race_df.columns]
            race_df = race_df[available_columns]
            race_df.to_excel(writer, sheet_name=f'Race_{race_no}', index=False)
        
        race_no += 1
    
    writer.close()
    print(f"Data saved to {output_file}")

if __name__ == "__main__":
    main()