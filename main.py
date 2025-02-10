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
        # 1) 取得今天 (台灣時區) 以便往前推 5 天
        taiwan_tz = ZoneInfo("Asia/Taipei")
        today = datetime.now(tz=taiwan_tz)

        print(f"[INFO] 今天是 {today.strftime('%Y/%m/%d')}，接下來要抓取往前 5 天的資料。")

        # 2) 連線 Google Sheet (假設 service_account.json 已由 workflow 放在檔案系統)
        json_keyfile_path = "service_account.json"
        worksheet = connect_google_sheet(
            json_keyfile_path=json_keyfile_path,
            sheet_name="Yung資料庫",          # 你可以改
            worksheet_name="原價屋網路PC組裝數RD"      # 你可以改
        )

        # 3) 依序處理最近 5 天
        for offset in range(1, 6):
            # offset=1 => 昨天；offset=2 => 前天；... => offset=5 => 前 5 天
            target_date = today - timedelta(days=offset)
            # 轉民國年 (西元 - 1911)
            minguo_year = target_date.year - 1911
            day_str = f"{minguo_year:03}{target_date.month:02}{target_date.day:02}"
            year_folder_str = f"{minguo_year}年"

            print(f"\n[INFO] 處理 {target_date.strftime('%Y/%m/%d')} => 民國{year_folder_str} / {day_str}")

            # 4) 用 Selenium 爬取組裝數
            assemble_count = crawl_coolpc(year_folder_str, day_str)

            # 5) 把結果更新到 Google Sheet
            update_or_append(worksheet, (day_str, assemble_count))

    except Exception as e:
        print("[ERROR] 程式出現例外:")
        traceback.print_exc()


def connect_google_sheet(json_keyfile_path, sheet_name, worksheet_name):
    """
    連線到Google Sheet (sheet_name)，並且打開指定worksheet(worksheet_name)
    回傳gspread的worksheet物件
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


def update_or_append(worksheet, row_data):
    """
    只讀試算表的最後 5 行資料(純文字)，在第一欄(索引0)找是否已有 row_data[0] (day_str)。
    - 若找到相同字串，就更新該行
    - 若沒找到，就 append 到最後一行
    """
    day_str, assemble_count = row_data
    current_data = worksheet.get_all_values()
    row_count = len(current_data)
    if row_count == 0:
        # 試算表目前是空的，直接放到第一行
        worksheet.update("A1:B1", [[day_str, assemble_count]])
        print(f"[INFO] 試算表是空的，已新增第一行: {row_data}")
        return

    # 只看最後 5 行範圍
    # （如果不足 5 行，就從第1行開始）
    start_row = max(1, row_count - 4)
    end_row = row_count

    # 讀出這範圍的資料(2欄)
    range_cells = f"A{start_row}:B{end_row}"
    last_data = worksheet.get(range_cells)  # 這是一個 2D list

    matched_row = None
    for i, row in enumerate(last_data):
        # row可能像 ['1140206', '63']
        # i 是 0-based 的索引，但真正的試算表行號 = start_row + i
        actual_row = start_row + i
        if len(row) > 0 and row[0] == day_str:
            matched_row = actual_row
            break

    if matched_row:
        # 有找到就更新
        cell_range = f"A{matched_row}:B{matched_row}"
        worksheet.update(cell_range, [[day_str, assemble_count]])
        print(f"[INFO] 找到同日期 '{day_str}'，已覆蓋到第 {matched_row} 行，組裝數={assemble_count}")
    else:
        # 沒找到就插到最後
        new_row = row_count + 1
        cell_range = f"A{new_row}:B{new_row}"
        worksheet.update(cell_range, [[day_str, assemble_count]])
        print(f"[INFO] 沒找到 '{day_str}'，已新增到第 {new_row} 行，組裝數={assemble_count}")


def crawl_coolpc(year_folder, day_str):
    """
    用 Selenium 抓取「每日組裝分享 -> year_folder -> day_str」的 footer 組裝數
    回傳組裝數 (int)
    """
    print("[INFO] 開啟 Selenium 瀏覽器(Headless)")
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=options
    )
    wait = WebDriverWait(driver, 20)

    try:
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

        assemble_count = 0
        try:
            footer_elem = wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='synofoto-folder-wall-footer']"))
            )
            footer_text = footer_elem.text.strip()  # e.g. "48 個項目"
            count_str = footer_text.split(" ")[0]
            assemble_count = int(count_str)
            print(f"[INFO] 抓到『{day_str}』的組裝數 = {assemble_count}")
        except:
            print("[WARNING] 無法取得 footer 或轉成 int 失敗，預設當作 0")

    finally:
        driver.quit()

    return assemble_count


if __name__ == "__main__":
    main()
