# main.py
import os
import time
import traceback
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo  # Python 3.9+ 可用 zoneinfo 取得台灣時區

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

import gspread
from oauth2client.service_account import ServiceAccountCredentials


def main():
    try:
        taiwan_tz = ZoneInfo("Asia/Taipei")
        today = datetime.now(tz=taiwan_tz)
        print(f"[INFO] 今天是 {today.strftime('%Y/%m/%d')}，接下來要抓取往前 5 天的資料。")

        # 連線到 Google Sheet
        json_keyfile_path = "service_account.json"
        worksheet = connect_google_sheet(
            json_keyfile_path=json_keyfile_path,
            sheet_name="Yung資料庫",           # 你可以改
            worksheet_name="原價屋網路PC組裝數RD"       # 你可以改
        )

        # 用一個 list 先把各天算出來 (day_str, year_folder_str)
        # offset=1 => 昨天, offset=2 => 前天, ..., offset=5 => 前5天
        days_info = []
        for offset in range(1, 6):
            target_date = today - timedelta(days=offset)
            minguo_year = target_date.year - 1911
            day_str = f"{minguo_year:03}{target_date.month:02}{target_date.day:02}"
            year_folder_str = f"{minguo_year}年"
            days_info.append((day_str, year_folder_str))

        # 1) 先抓昨天 (index=0) 與 前天 (index=1)，各嘗試一次
        results = [0, 0, 0, 0, 0]  # 預設五天的組裝數都給 0
        for i in range(2):
            day_str, year_folder_str = days_info[i]
            assemble_count = single_attempt_coolpc(year_folder_str, day_str)
            results[i] = assemble_count
            print(f"[INFO] 第{i+1}天(偏移: {i+1}) => {day_str} => assemble_count={assemble_count}")

        # 2) 如果昨天和前天都為 0，再額外重開最多三次
        if results[0] == 0 and results[1] == 0:
            print("[WARNING] 昨天與前天都抓不到資料，再重開三次資料夾嘗試看看～")
            for retry in range(3):
                for i in range(2):
                    # 如果已經抓到了 (>0) 就不用再抓
                    if results[i] != 0:
                        continue
                    day_str, year_folder_str = days_info[i]
                    assemble_count = single_attempt_coolpc(year_folder_str, day_str)
                    if assemble_count > 0:
                        results[i] = assemble_count
                        print(f"[INFO] 重開第{retry+1}次 => 抓到 {day_str}={assemble_count}")
                # 如果昨天或前天在這次重開中抓到了，也繼續再試其他天
            print(f"[INFO] 重開結束，昨天({days_info[0][0]})={results[0]}，前天({days_info[1][0]})={results[1]}")

        # 3) 再處理剩下的天數 offset=3,4,5
        for i in range(2, 5):
            day_str, year_folder_str = days_info[i]
            assemble_count = single_attempt_coolpc(year_folder_str, day_str)
            results[i] = assemble_count
            print(f"[INFO] 第{i+1}天(偏移: {i+1}) => {day_str} => assemble_count={assemble_count}")

        # 4) 全部抓完以後，寫回試算表
        #    注意：要用 update_or_append() 一天一天處理，以免有人只抓部分
        for i in range(5):
            day_str, _ = days_info[i]
            assemble_count = results[i]
            update_or_append(worksheet, (day_str, assemble_count))

    except Exception as e:
        print("[ERROR] 程式出現例外:")
        traceback.print_exc()


def connect_google_sheet(json_keyfile_path, sheet_name, worksheet_name):
    """
    連線到Google Sheet (sheet_name)，並打開指定worksheet(worksheet_name)。
    回傳gspread的worksheet物件。
    """
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)
    client = gspread.authorize(creds)
    sh = client.open(sheet_name)
    worksheet = sh.worksheet(worksheet_name)
    return worksheet


def single_attempt_coolpc(year_folder, day_str):
    """
    嘗試一次開 Selenium、進入「year_folder / day_str」資料夾。
    若能抓到組裝數，就回傳；否則回傳 0。
    """
    assemble_count = 0
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=options
        )
        wait = WebDriverWait(driver, 20)

        driver.get("https://www.coolpc.com.tw/photo/#/shared_space/folder/156?_k=tr98b7")
        time.sleep(10)

        # 點「每日組裝分享 (僅網路部)」
        share_folder_xpath = "//div[@class='css-106gz8u' and text()='每日組裝分享 (僅網路部)']"
        wait.until(EC.element_to_be_clickable((By.XPATH, share_folder_xpath))).click()
        time.sleep(10)

        # 點「xxx年」資料夾
        year_xpath = f"//div[@class='css-106gz8u' and text()='{year_folder}']"
        wait.until(EC.element_to_be_clickable((By.XPATH, year_xpath))).click()
        time.sleep(10)

        # 點「day_str」資料夾, e.g. 1140206
        day_xpath = f"//div[@class='css-106gz8u' and text()='{day_str}']"
        wait.until(EC.element_to_be_clickable((By.XPATH, day_xpath))).click()
        time.sleep(10)

        # 等footer出現，抓組裝數
        footer_elem = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='synofoto-folder-wall-footer']"))
        )
        footer_text = footer_elem.text.strip()  # e.g. "48 個項目"
        count_str = footer_text.split(" ")[0]
        assemble_count = int(count_str)
        print(f"[INFO] 成功抓到『{day_str}』的組裝數 = {assemble_count}")

    except Exception as e:
        print(f"[WARNING] 嘗試抓取 {year_folder}/{day_str} 時失敗，回傳 0：{e}")

    finally:
        try:
            driver.quit()
        except:
            pass

    return assemble_count


def update_or_append(worksheet, row_data):
    """
    只讀試算表最後 5 行(純文字)，若第一欄有同樣 day_str 就覆蓋，否則插到最後。
    """
    day_str, assemble_count = row_data
    current_data = worksheet.get_all_values()
    row_count = len(current_data)

    # 如果整張表都還沒資料，就第一行塞入
    if row_count == 0:
        worksheet.update("A1:B1", [[day_str, assemble_count]])
        print(f"[INFO] 試算表是空的，已新增第一行: {row_data}")
        return

    # 只看最後 5 行 (不足就整張表)
    start_row = max(1, row_count - 4)
    end_row = row_count
    range_cells = f"A{start_row}:B{end_row}"
    last_data = worksheet.get(range_cells)

    matched_row = None
    for i, row in enumerate(last_data):
        actual_row = start_row + i  # 轉回整張表的行號
        if len(row) > 0 and row[0] == day_str:
            matched_row = actual_row
            break

    if matched_row:
        cell_range = f"A{matched_row}:B{matched_row}"
        worksheet.update(cell_range, [[day_str, assemble_count]])
        print(f"[INFO] 找到同日期 '{day_str}'，已覆蓋到第 {matched_row} 行，組裝數={assemble_count}")
    else:
        new_row = row_count + 1
        cell_range = f"A{new_row}:B{new_row}"
        worksheet.update(cell_range, [[day_str, assemble_count]])
        print(f"[INFO] 沒找到 '{day_str}'，已新增到第 {new_row} 行，組裝數={assemble_count}")


if __name__ == "__main__":
    main()
