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
        # 1) 計算「昨日」的日期 (台灣時區)
        taiwan_tz = ZoneInfo("Asia/Taipei")
        today = datetime.now(tz=taiwan_tz)
        yesterday = today - timedelta(days=1)
        
        # 轉民國年 (西元 - 1911)
        minguo_year = yesterday.year - 1911
        # ex: 2024/02/07(今天) -> 昨天2024/02/06 -> 民國113, 月02, 日06 => "1130206"
        day_str = f"{minguo_year:03}{yesterday.month:02}{yesterday.day:02}"
        year_folder_str = f"{minguo_year}年"

        print(f"[INFO] 今天是 {today.strftime('%Y/%m/%d')}，要抓取『{year_folder_str}』資料夾底下的『{day_str}』。")

        # 2) 從 GitHub Secrets拿到 base64編碼過的 service_account.json 內容
        #    在 GitHub Actions 執行時，會事先在 workflow 用 echo/base64 decode生成檔案 "service_account.json"
        #    這裡我們假設該檔案已經在目前路徑了
        json_keyfile_path = "service_account.json"

        # 3) 連線 Google Sheet
        worksheet = connect_google_sheet(
            json_keyfile_path=json_keyfile_path,
            sheet_name="Yung資料庫",           # 你可以改
            worksheet_name="原價屋網路PC組裝數RD"         # 你可以改
        )

        # 4) 啟動 Selenium (Headless Chrome) 來爬
        assemble_count = crawl_coolpc(year_folder_str, day_str)
        
        # 5) 把結果 append 到Sheet最下面
        #    這裡只append 一筆 (day_str, assemble_count)
        append_row_to_sheet(worksheet, (day_str, assemble_count))

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


def append_row_to_sheet(worksheet, row_data):
    """
    將 row_data (例如 ("1130206", 30)) append到worksheet最下面
    """
    # 取得目前所有資料
    current_data = worksheet.get_all_values()
    last_row = len(current_data)
    start_row = last_row + 1  # 下一空行
    cell_range = f"A{start_row}:B{start_row}"

    # 轉成worksheet.update需要的 2D陣列
    update_values = [ [row_data[0], row_data[1]] ]
    worksheet.update(cell_range, update_values)
    print(f"[INFO] 已將資料 {row_data} 追加到試算表 (第{start_row}行)")


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

    # 點「day_str」資料夾, e.g. 1130206
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
        print("[WARNING] 抓不到footer文字，或轉成int失敗，預設0")

    driver.quit()
    return assemble_count


if __name__ == "__main__":
    main()
