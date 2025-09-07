import requests
from bs4 import BeautifulSoup
import openpyxl
from datetime import datetime
import re  # 如果之前沒有導入
import logging  # 如果之前有使用
import pandas as pd

# Excel 文件路徑
excel_file = "HK_Racing_Auto_Data.xlsx"

# 內建標準時間字典，區分沙田 (ST) 和跑馬地 (HV)｛
standard_times = {
    ("ST", 1000.0, "一級賽"): {"pace": 33.65, "final": 22.25, "total": "55.9"},
    ("ST", 1000.0, "二級賽"): {"pace": 33.65, "final": 22.25, "total": "55.9"},
    ("ST", 1000.0, "三級賽"): {"pace": 33.65, "final": 22.25, "total": "55.9"},
    ("ST", 1000.0, "第一班"): {"pace": 33.7, "final": 22.35, "total": "56.05"},
    ("ST", 1000.0, "第二班"): {"pace": 33.7, "final": 22.35, "total": "56.05"},
    ("ST", 1000.0, "第三班"): {"pace": 33.7, "final": 22.75, "total": "56.45"},
    ("ST", 1000.0, "第四班"): {"pace": 33.75, "final": 22.9, "total": "56.65"},
    ("ST", 1000.0, "第五班"): {"pace": 34.1, "final": 22.9, "total": "57"},
    ("ST", 1000.0, "新馬賽"): {"pace": 34.05, "final": 22.6, "total": "56.65"},
    ("ST", 1200.0, "一級賽"): {"pace": 45.75, "final": 22.4, "total": "1:08.15"},
    ("ST", 1200.0, "二級賽"): {"pace": 45.75, "final": 22.4, "total": "1:08.15"},
    ("ST", 1200.0, "三級賽"): {"pace": 45.75, "final": 22.4, "total": "1:08.15"},
    ("ST", 1200.0, "第一班"): {"pace": 45.85, "final": 22.6, "total": "1:08.45"},
    ("ST", 1200.0, "第二班"): {"pace": 46, "final": 22.65, "total": "1:08.65"},
    ("ST", 1200.0, "第三班"): {"pace": 46.05, "final": 22.95, "total": "1:09.00"},
    ("ST", 1200.0, "第四班"): {"pace": 46.2, "final": 23.15, "total": "1:09.35"},
    ("ST", 1200.0, "第四班（條件限制）"): {"pace": 46.2, "final": 23.15, "total": "1:09.35"},
    ("ST", 1200.0, "第五班"): {"pace": 46.25, "final": 23.3, "total": "1:09.55"},
    ("ST", 1200.0, "新馬賽"): {"pace": 46.9, "final": 23, "total": "1:09.90"},
    ("ST", 1400.0, "一級賽"): {"pace": 58.7, "final": 22.4, "total": "1:21.10"},
    ("ST", 1400.0, "二級賽"): {"pace": 58.7, "final": 22.4, "total": "1:21.10"},
    ("ST", 1400.0, "三級賽"): {"pace": 58.7, "final": 22.4, "total": "1:21.10"},
    ("ST", 1400.0, "第一班"): {"pace": 58.55, "final": 22.7, "total": "1:21.25"},
    ("ST", 1400.0, "第二班"): {"pace": 58.45, "final": 23, "total": "1:21.45"},
    ("ST", 1400.0, "第三班"): {"pace": 58.4, "final": 23.25, "total": "1:21.65"},
    ("ST", 1400.0, "第四班（條件限制）"): {"pace": 58.6, "final": 23.4, "total": "1:22.00"},
    ("ST", 1400.0, "第四班"): {"pace": 58.6, "final": 23.4, "total": "1:22.00"},
    ("ST", 1400.0, "第五班"): {"pace": 58.65, "final": 23.65, "total": "1:22.30"},
    ("ST", 1600.0, "一級賽"): {"pace": 71.15, "final": 22.75, "total": "1:33.90"},
    ("ST", 1600.0, "二級賽"): {"pace": 71.15, "final": 22.75, "total": "1:33.90"},
    ("ST", 1600.0, "三級賽"): {"pace": 71.15, "final": 22.75, "total": "1:33.90"},
    ("ST", 1600.0, "第一班"): {"pace": 71.05, "final": 23, "total": "1:34.05"},
    ("ST", 1600.0, "四歲"): {"pace": 71.15, "final": 23.1, "total": "1:34.25"},
    ("ST", 1600.0, "第二班"): {"pace": 71.15, "final": 23.1, "total": "1:34.25"},
    ("ST", 1600.0, "第三班（條件限制）"): {"pace": 71.2, "final": 23.5, "total": "1:34.70"},
    ("ST", 1600.0, "第三班"): {"pace": 71.2, "final": 23.5, "total": "1:34.70"},
    ("ST", 1600.0, "第四班"): {"pace": 71.2, "final": 23.7, "total": "1:34.90"},
    ("ST", 1600.0, "第五班"): {"pace": 71.55, "final": 23.9, "total": "1:35.45"},
    ("ST", 1800.0, "一級賽"): {"pace": 84.35, "final": 22.75, "total": "1:47.1"},
    ("ST", 1800.0, "二級賽"): {"pace": 84.35, "final": 22.75, "total": "1:47.1"},
    ("ST", 1800.0, "三級賽"): {"pace": 84.35, "final": 22.75, "total": "1:47.1"},
    ("ST", 1800.0, "第一班"): {"pace": 83.9, "final": 23.4, "total": "1:47.30"},
    ("ST", 1800.0, "四歲"): {"pace": 83.9, "final": 23.4, "total": "1:47.30"},
    ("ST", 1800.0, "第二班"): {"pace": 83.9, "final": 23.4, "total": "1:47.30"},
    ("ST", 1800.0, "第三班"): {"pace": 83.95, "final": 23.55, "total": "1:47.5"},
    ("ST", 1800.0, "第四班"): {"pace": 84.1, "final": 23.76, "total": "1:47.85"},
    ("ST", 1800.0, "第五班"): {"pace": 84.25, "final": 24.2, "total": "1:48.45"},
    ("ST", 2000.0, "一級賽"): {"pace": 97.3, "final": 23.2, "total": "2:00.50"},
    ("ST", 2000.0, "二級賽"): {"pace": 97.3, "final": 23.2, "total": "2:00.50"},
    ("ST", 2000.0, "三級賽"): {"pace": 97.3, "final": 23.2, "total": "2:00.50"},
    ("ST", 2000.0, "第一班"): {"pace": 98, "final": 23.2, "total": "2:01.20"},
    ("ST", 2000.0, "四歲"): {"pace": 98.3, "final": 23.4, "total": "2:01.70"},
    ("ST", 2000.0, "第二班"): {"pace": 98.3, "final": 23.4, "total": "2:01.70"},
    ("ST", 2000.0, "第三班"): {"pace": 98.35, "final": 23.55, "total": "2:01.90"},
    ("ST", 2000.0, "第四班"): {"pace": 98.6, "final": 23.75, "total": "2:02.35"},
    ("ST", 2000.0, "第五班"): {"pace": 98.45, "final": 24.2, "total": "2:02.65"},
    ("ST", 2200.0, "第五班"): {"pace": 111.75, "final": 24, "total": "2:15.75"},
    ("ST", 2400.0, "一級賽"): {"pace": 123.05, "final": 23.95, "total": "2:27.00"},
    ("HV", 1000.0, "第一班"): {"pace": 33.45, "final": 22.95, "total": "56.40"},
    ("HV", 1000.0, "第二班"): {"pace": 33.45, "final": 22.95, "total": "56.40"},
    ("HV", 1000.0, "第三班"): {"pace": 33.5, "final": 23.15, "total": "56.65"},
    ("HV", 1000.0, "第四班"): {"pace": 33.85, "final": 23.35, "total": "57.2"},
    ("HV", 1000.0, "第五班"): {"pace": 34, "final": 23.35, "total": "57.35"},
    ("HV", 1200.0, "第一班"): {"pace": 45.85, "final": 23.25, "total": "1:09.10"},
    ("HV", 1200.0, "第二班"): {"pace": 45.8, "final": 23.5, "total": "1:09.30"},
    ("HV", 1200.0, "第三班"): {"pace": 46.05, "final": 23.55, "total": "1:09.60"},
    ("HV", 1200.0, "第四班"): {"pace": 46.35, "final": 23.55, "total": "1:09.90"},
    ("HV", 1200.0, "第五班"): {"pace": 46.45, "final": 23.65, "total": "1:10.10"},
    ("HV", 1650.0, "第一班"): {"pace": 75.7, "final": 23.4, "total": "1:39.10"},
    ("HV", 1650.0, "第二班"): {"pace": 75.65, "final": 23.65, "total": "1:39.30"},
    ("HV", 1650.0, "第三班"): {"pace": 76.05, "final": 23.85, "total": "1:39.90"},
    ("HV", 1650.0, "第四班"): {"pace": 76.05, "final": 24.05, "total": "1:40.10"},
    ("HV", 1650.0, "第五班"): {"pace": 76.1, "final": 24.2, "total": "1:40.30"},
    ("HV", 1800.0, "三級賽"): {"pace": 85.05, "final": 23.9, "total": "1:48.95"},
    ("HV", 1800.0, "第二班"): {"pace": 85.2, "final": 23.95, "total": "1:49.15"},
    ("HV", 1800.0, "第三班"): {"pace": 85.5, "final": 23.95, "total": "1:49.45"},
    ("HV", 1800.0, "第四班"): {"pace": 85.35, "final": 24.3, "total": "1:49.65"},
    ("HV", 1800.0, "第五班"): {"pace": 85.4, "final": 24.55, "total": "1:49.95"},
    ("HV", 2200.0, "第三班"): {"pace": 112.3, "final": 24.3, "total": "2:16.60"},
    ("HV", 2200.0, "第四班"): {"pace": 112.7, "final": 24.35, "total": "2:17.05"},
    ("HV", 2200.0, "第五班"): {"pace": 112.75, "final": 24.6, "total": "2:17.35"},
    ("ST_AWT", 1200.0, "第二班"): {"pace": 45.3, "final": 23.05, "total": "1:08.35"},
    ("ST_AWT", 1200.0, "第三班"): {"pace": 45.35, "final": 23.2, "total": "1:08.55"},
    ("ST_AWT", 1200.0, "第四班"): {"pace": 45.4, "final": 23.55, "total": "1:08.95"},
    ("ST_AWT", 1200.0, "第五班"): {"pace": 45.7, "final": 23.65, "total": "1:09.35"},
    ("ST_AWT", 1650.0, "第一班"): {"pace": 73.9, "final": 23.9, "total": "1:37.80"},
    ("ST_AWT", 1650.0, "第二班"): {"pace": 74.5, "final": 23.9, "total": "1:38.40"},
    ("ST_AWT", 1650.0, "第三班"): {"pace": 74.65, "final": 23.95, "total": "1:38.6"},
    ("ST_AWT", 1650.0, "第四班"): {"pace": 74.9, "final": 24.15, "total": "1:39.05"},
    ("ST_AWT", 1650.0, "第五班"): {"pace": 75.15, "final": 24.3, "total": "1:39.45"},
    ("ST_AWT", 1800.0, "第三班"): {"pace": 84.1, "final": 23.95, "total": "1:48.05"},
    ("ST_AWT", 1800.0, "第四班"): {"pace": 84.4, "final": 24.15, "total": "1:48.55"},
    ("ST_AWT", 1800.0, "第五班"): {"pace": 85.15, "final": 24.3, "total": "1:49.45"},
}

# 將時間格式轉換為秒
def convert_time_to_seconds(time_str):
    if not time_str or time_str in ["N/A", "-", None]:
        return None
    time_str = str(time_str).strip()
    if time_str == "N/A":
        return None
    
    parts = time_str.split(":")
    if len(parts) == 2:  # MM:SS 或 MM:SS.TT 格式
        minutes = float(parts[0])
        seconds_parts = parts[1].split(".")
        seconds = float(seconds_parts[0])
        fractions = float(seconds_parts[1]) / 100 if len(seconds_parts) > 1 else 0
        return minutes * 60 + seconds + fractions
    elif len(parts) == 1:  # SS.TT 格式
        seconds_parts = parts[0].split(".")
        seconds = float(seconds_parts[0])
        fractions = float(seconds_parts[1]) / 100 if len(seconds_parts) > 1 else 0
        return seconds + fractions
    return None

# 初始化 Excel 文件
def initialize_excel():
    try:
        wb = openpyxl.load_workbook(excel_file)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "賽馬自動記錄"
        headers = [
            "日期", "賽事名稱", "賽事編號", "名次", "馬號", "馬名", "馬齡", "騎師", "練馬師", "實際負磅", "排位體重", 
            "檔位", "頭馬距離", "沿途走位", "完成時間", "配備", "獨贏賠率", 
            "賽事班次", "途程", "賽道", "場地狀況",
            "分段1", "分段2", "分段3", "分段4", "分段5", "分段6",
            "頭段", "末段",
            "頭段指數", "末段指數", "時間指數"
        ]
        ws.append(headers)
        wb.save(excel_file)
    return wb

# 從賽果頁面抓取資料
def fetch_race_data(url, racecourse):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")
        
        # 提取賽事編號
        race_tab = soup.find("div", class_="race_tab")
        race_number = "N/A"
        if race_tab:
            race_header = race_tab.find("td", colspan="16")
            if race_header and "第" in race_header.text and "場" in race_header.text:
                match = re.search(r'\((\d+)\)', race_header.text)
                if match:
                    race_number = match.group(1)  # 提取括號中的數字，例如 "425"
                    race_number = str(int(race_number)).zfill(3)  # 將數字轉換為三位數，例如 "001"

        race_info = extract_race_info(soup)
        race_table = soup.find("table", class_="f_tac table_bd draggable")
        if not race_table:
            print("未找到賽果表格")
            return []

        tbody = race_table.find("tbody")
        if not tbody:
            print("賽果表格中未找到 tbody")
            return []

        date_str = url.split("RaceDate=")[-1].split("&")[0]
        race_no = url.split("RaceNo=")[-1].split("&")[0]
        race_name = f"賽事 {date_str} 第{race_no}場"

        sectional_data = fetch_sectional_times(date_str, race_no)
       
        race_data = []
        for row in tbody.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) < 12:
                print(f"賽果表格欄位不足，跳過該行: {cols}")
                continue

            place = ''.join(filter(str.isdigit, cols[0].text.strip())) or cols[0].text.strip()
            horse_number = cols[1].text.strip()
            horse_link = cols[2].find("a")
            if not horse_link:
                print(f"未找到馬名連結，跳過該行: {row}")
                continue
            horse_name = horse_link.text.strip()
            horse_url = "https://racing.hkjc.com" + horse_link["href"]
            # 傳遞 race_number 給 fetch_horse_age，確保 race_number 為有效值
            horse_data = fetch_horse_age(horse_url, race_number if race_number != "N/A" else None)
            horse_age = horse_data["age"]
            equipment = horse_data["equipment"]  # 獲取配備
            jockey_link = cols[3].find("a")
            jockey = jockey_link.text.strip() if jockey_link else cols[3].text.strip() or "N/A"
            trainer_link = cols[4].find("a")
            trainer = trainer_link.text.strip() if trainer_link else cols[4].text.strip() or "N/A"
            weight = cols[5].text.strip()
            body_weight = cols[6].text.strip()
            gate = cols[7].text.strip()
            distance = cols[8].text.strip()  # 頭馬距離
            position_divs = cols[9].find_all("div")
            positions = " ".join([div.text.strip() for div in position_divs if div.text.strip().isdigit()])
            finish_time = cols[10].text.strip()
            odds = cols[11].text.strip()

            # 添加調試輸出，檢查 finish_time 和 odds
            print(f"完成時間: {finish_time}")
            print(f"獨贏賠率: {odds}")

            # 處理 head_time 和 final_time
            sectional_times = sectional_data.get(horse_number, ["N/A"] * 6)
            head_time = "N/A"
            final_time = "N/A"
            if sectional_times != ["N/A"] * 6:
                try:
                    valid_times = [float(time) for time in sectional_times if time not in ["N/A", "---"]]
                    if valid_times:
                        if len(valid_times) >= 3:
                            head_time = sum(valid_times[:-1])
                            final_time = valid_times[-1]
                        elif len(valid_times) == 2:
                            head_time = valid_times[0]
                            final_time = valid_times[1]
                        else:
                            final_time = valid_times[0]
                except ValueError as e:
                    print(f"無法解析分段時間: {sectional_times}, 錯誤: {e}")

            # 處理 place 和 positions，確保不包含 "---"
            if place in ["---", "N/A"]:
                place = "N/A"
            print(f"名次: {place}")
            if positions in ["---", "N/A"]:
                positions = "N/A"
            print(f"沿途走位: {positions}")

            # 處理完成時間，確保不包含 "---"
            if finish_time in ["---", "N/A"]:
                finish_time = "N/A"
                total_time_seconds = None
            else:
                total_time_seconds = convert_time_to_seconds(finish_time)

            # 處理獨贏賠率，確保不包含 "---"
            if odds in ["---", "N/A"]:
                odds = "N/A"

            # 定義 track_with_racecourse
            track_with_racecourse = f"{racecourse}{race_info['track']}" if racecourse and race_info['track'] else race_info['track'] or "N/A"

            # 提取途程和班次，處理 "N/A" 或 "---" 的情況
            distance_value = None
            if race_info["distance"] and race_info["distance"].strip() not in ["N/A", "---"]:
                try:
                    distance_value = float(race_info["distance"].replace("米", "").strip())
                    print(f"賽事途程 (浮點數): {distance_value}")
                except ValueError as e:
                    print(f"無法將途程 {race_info['distance']} 轉換為浮點數: {e}")
                    distance_value = None
            else:
                print(f"賽事途程: {race_info['distance']} (無效或 N/A)")

            class_value = race_info["class"] if race_info["class"] not in ["N/A", "---"] else "N/A"

            # 處理其他數值欄位，確保不包含 "---" 或 "N/A"，並添加調試輸出
            if weight in ["---", "N/A"]:
                weight = "N/A"
            print(f"權重: {weight}")
            if body_weight in ["---", "N/A"]:
                body_weight = "N/A"
            print(f"體重: {body_weight}")
            if gate in ["---", "N/A"]:
                gate = "N/A"
            print(f"檔位: {gate}")

            # 將賽馬場轉換為代碼，並考慮全天候跑道
            if racecourse == "沙田" and "全天候" in race_info["track"]:
                racecourse_code = "ST_AWT"  # 全天候跑道
            else:
                racecourse_code = "ST" if racecourse == "沙田" else "HV" if racecourse == "跑馬地" else None

            # 查找標準時間
            standard_key = (racecourse_code, distance_value, class_value) if racecourse_code and distance_value and class_value else None
            print(f"查找標準時間 - 賽馬場: {racecourse_code}, 途程: {distance_value}, 班次: {class_value}, 標準鍵: {standard_key}")
            standard_raw = standard_times.get(standard_key, {"pace": None, "final": None, "total": None})

            # 將標準時間轉換為秒（如果需要）
            standard = {
                "pace": convert_time_to_seconds(standard_raw["pace"]) if isinstance(standard_raw["pace"], str) else standard_raw["pace"],
                "final": convert_time_to_seconds(standard_raw["final"]) if isinstance(standard_raw["final"], str) else standard_raw["final"],
                "total": convert_time_to_seconds(standard_raw["total"]) if isinstance(standard_raw["total"], str) else standard_raw["total"]
            }

            # 計算指數（標準時間 - 實際時間）
            head_index = "N/A"
            final_index = "N/A"
            time_index = "N/A"

            if standard["pace"] is not None and head_time != "N/A":
                head_index = round(standard["pace"] - head_time, 2)
            if standard["final"] is not None and final_time != "N/A":
                final_index = round(standard["final"] - final_time, 2)
            if standard["total"] is not None and total_time_seconds is not None:
                time_index = round(standard["total"] - total_time_seconds, 2)

            # 處理頭馬距離，允許 "N/A" 或 "---"
            if distance in ["---", "N/A"]:
                distance = "N/A"

            # 將配備添加到 race_data，位於 "完成時間" 後
            race_data.append([
                date_str, race_name, race_number, place, horse_number, horse_name, horse_age, jockey, trainer, weight, 
                body_weight, gate, distance, positions, finish_time, equipment, odds,
                race_info["class"], race_info["distance"], track_with_racecourse, race_info["going"]
            ] + sectional_times + [head_time, final_time, head_index, final_index, time_index])
        return race_data
    except Exception as e:
        print(f"抓取賽果失敗: {e}")
        return []

def extract_race_info(soup):
    race_info = {
        "class": "N/A",
        "distance": "N/A",
        "track": "N/A",
        "going": "N/A",
    }

    race_tab = soup.find("div", class_="race_tab")
    if not race_tab:
        print("未找到賽事資訊表格")
        return race_info

    class_distance_row = race_tab.find("td", style="width: 385px;")
    if class_distance_row:
        class_distance_text = class_distance_row.text.strip()
        if " - " in class_distance_text:
            race_info["class"] = class_distance_text.split(" - ")[0].strip()
            race_info["distance"] = class_distance_text.split(" - ")[1].strip()
        else:
            distance_match = re.search(r"\d+米", class_distance_text)
            if distance_match:
                race_info["distance"] = distance_match.group()
            race_info["class"] = class_distance_text.replace(race_info["distance"], "").strip()

    going_row = race_tab.find("td", string=lambda x: x and "場地狀況 :" in x)
    if going_row:
        going_value = going_row.find_next("td")
        race_info["going"] = going_value.text.strip() if going_value else "N/A"

    track_row = race_tab.find("td", string=lambda x: x and "賽道 :" in x)
    if track_row:
        track_value = track_row.find_next("td")
        race_info["track"] = track_value.text.strip() if track_value else "N/A"

    print(f"提取賽事資訊 - 班次: {race_info['class']}, 途程: {race_info['distance']}")  # 診斷輸出
    return race_info

# 從分段時間頁面抓取資料
def fetch_sectional_times(date, race_no):
    # 將日期從 YYYY/MM/DD 轉換為 DD/MM/YYYY
    date_obj = datetime.strptime(date, "%Y/%m/%d")
    formatted_date = date_obj.strftime("%d/%m/%Y")
    
    # 生成正確的分段時間 URL，只包含 RaceDate 和 RaceNo
    sectional_url = f"https://racing.hkjc.com/racing/information/chinese/Racing/DisplaySectionalTime.aspx?RaceDate={formatted_date}&RaceNo={race_no}"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        print(f"正在訪問: {sectional_url}")
        response = requests.get(sectional_url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")
        
        sectional_table = soup.find("table", class_="table_bd f_tac race_table")
        if not sectional_table:
            print("未找到分段時間表格")
            return {}

        print("找到表格，開始解析...")
        sectional_data = {}
        tbody = sectional_table.find("tbody")
        if not tbody:
            print("分段時間表格無數據")
            return {}

        for row in tbody.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) < 10:  # 確保有足夠的欄位
                print(f"跳過短於 10 欄的行: {row}")
                continue

            try:
                horse_number = cols[1].text.strip()
                sectional_times = []
                for i in range(3, 9):  # 第 1-6 段
                    time = "N/A"
                    p_tags = cols[i].find_all("p")
                    if p_tags:
                        for p in p_tags:
                            if "f_clear" not in p.get("class", []):
                                time_text = p.text.strip()
                                if time_text and p.find("span", class_="color_blue2"):
                                    time = time_text.split()[0]  # 只取 "22.48"，排除 200m 細分
                                elif time_text:
                                    time = time_text
                    sectional_times.append(time)
                print(f"馬匹 {horse_number} 的分段時間: {sectional_times}")
                sectional_data[horse_number] = sectional_times
            except IndexError as e:
                print(f"解析分段時間行失敗（索引錯誤）: {row}, 錯誤: {e}")
                continue
            except Exception as e:
                print(f"解析分段時間行失敗（其他錯誤）: {row}, 錯誤: {e}")
                continue

        return sectional_data
    except Exception as e:
        logging.error(f"抓取分段時間失敗: {e}")
        return {}

def fetch_horse_age(horse_url, race_number=None):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(horse_url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")

        # 查找馬齡，添加錯誤處理
        age_row = soup.find("td", string=lambda x: x and "出生地 / 馬齡" in x) if soup else None
        age = "N/A"
        if age_row:
            age_td = age_row.find_next_sibling("td")
            if age_td:
                next_age_td = age_td.find_next_sibling("td")
                if next_age_td:
                    age_text = next_age_td.text.strip()
                    age = age_text.split("/")[-1].strip() if "/" in age_text else "N/A"  # 提取馬齡

        # 查找配備（從馬匹歷史記錄表格中提取，根據賽事編號）
        equip_table = soup.find("table", class_="bigborder")
        equipment = "N/A"
        if equip_table and race_number:
            # 查找所有可能的賽事記錄行
            race_rows = equip_table.find_all("tr")
            if race_rows:
                # 遍歷每一行，查找與賽事編號匹配的行
                for row in race_rows:
                    # 檢查第一個 <td> 是否包含有效的賽事編號連結
                    race_link_td = row.find("td", align="center")
                    if race_link_td:
                        race_link = race_link_td.find("a", class_="htable_eng_text")
                        if race_link and race_link.text.strip() == str(race_number):
                            tds = row.find_all("td")
                            if len(tds) >= 18:  # 確保有足夠的 <td>（根據你的 HTML 結構）
                                equipment = tds[-2].text.strip() if tds[-2].text.strip() else "N/A"  # 倒數第二個 <td> 是配備
                            break  # 找到匹配的賽事後退出循環

        print(f"馬匹 {horse_url} 的馬齡: {age}, 配備: {equipment}, 賽事編號: {race_number}")
        return {"age": age, "equipment": equipment}
    except Exception as e:
        print(f"提取馬匹資料失敗: {e}")
        return {"age": "N/A", "equipment": "N/A"}

# 修改 fetch_race_urls 函數，返回場地名稱和URL列表
def fetch_race_urls(date):
    base_url = "https://racing.hkjc.com/racing/information/Chinese/Racing/LocalResults.aspx"
    url = f"{base_url}?RaceDate={date}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")

        race_table = soup.find("table", class_="f_fs12 js_racecard")
        if not race_table:
            print("未找到賽事表格")
            return None, []

        race_urls = []
        racecourse = None

        racecourse_span = race_table.find("span", style="color:#666666;font-size:12px;font-weight:700;font-family:Arial,Verdana,Helvetica,sans-serif;")
        if racecourse_span:
            racecourse_name = racecourse_span.text.strip().replace(":", "").strip()
            if racecourse_name == "跑馬地":
                racecourse = "跑馬地"
            elif racecourse_name == "沙田":
                racecourse = "沙田"
            else:
                print(f"未知的賽馬場名稱: {racecourse_name}")
                return None, []

        for td in race_table.find_all("td"):
            link = td.find("a", href=True)
            if link:
                race_url = "https://racing.hkjc.com" + link["href"]
                if race_url not in race_urls:
                    race_urls.append(race_url)
            else:
                img = td.find("img")
                if img and "racecard_rt_1_o.gif" in img["src"]:
                    if racecourse:
                        racecourse_code = "HV" if racecourse == "跑馬地" else "ST"
                        race_url = f"https://racing.hkjc.com/racing/information/Chinese/Racing/LocalResults.aspx?RaceDate={date}&Racecourse={racecourse_code}&RaceNo=1"
                        race_urls.append(race_url)

        return racecourse, race_urls
    except Exception as e:
        print(f"獲取賽事 URL 失敗: {e}")
        return None, []
           
# 儲存到 Excel
def save_to_excel(wb, data):
    ws = wb.active
    for row in data:
        ws.append(row)
    wb.save(excel_file)
    print(f"資料已儲存至 {excel_file}")

# 主程式
def main():
    wb = initialize_excel()
    date = input("請輸入日期 (格式: YYYY/MM/DD): ")

    racecourse, race_urls = fetch_race_urls(date)
    if not race_urls:
        print(f"未找到 {date} 的賽事資料")
        return

    all_race_data = []
    for url in race_urls:
        print(f"正在處理賽事: {url}")
        race_data = fetch_race_data(url, racecourse)
        if race_data:
            all_race_data.extend(race_data)

    if all_race_data:
        save_to_excel(wb, all_race_data)
    else:
        print("無資料可儲存")

if __name__ == "__main__":
    main()