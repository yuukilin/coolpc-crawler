# main.py
import json
import time
import traceback
from dataclasses import dataclass
from datetime import datetime, timedelta
from enum import Enum
from pathlib import Path
from zoneinfo import ZoneInfo  # Python 3.9+ 可用 zoneinfo 取得台灣時區

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

import gspread
from oauth2client.service_account import ServiceAccountCredentials


PENDING_FILE = Path("pending_dates.json")


class FetchStatus(Enum):
    SUCCESS = "success"
    PENDING = "pending"
    ERROR = "error"


@dataclass(frozen=True)
class FetchResult:
    status: FetchStatus
    count: int = None


def parse_day(day_str):
    if not isinstance(day_str, str) or len(day_str) != 7 or not day_str.isdigit():
        raise ValueError(f"無效的民國日期：{day_str!r}")
    return datetime(int(day_str[:3]) + 1911, int(day_str[3:5]), int(day_str[5:]))


def load_pending_dates(path=PENDING_FILE):
    if not path.exists():
        return set()
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("version") != 1 or not isinstance(payload.get("dates"), list):
        raise ValueError("待補抓清單格式錯誤")
    for day_str in payload["dates"]:
        parse_day(day_str)
    return set(payload["dates"])


def save_pending_dates(dates, path=PENDING_FILE):
    temporary_path = path.with_suffix(".tmp")
    temporary_path.write_text(
        json.dumps({"version": 1, "dates": sorted(dates)}, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    temporary_path.replace(path)


def target_days(today, pending_dates):
    dates = set(pending_dates)
    for offset in range(1, 6):
        day = today - timedelta(days=offset)
        dates.add(f"{day.year - 1911:03}{day.month:02}{day.day:02}")
    for day_str in dates:
        if parse_day(day_str).date() >= today.date():
            raise ValueError(f"待補抓日期必須早於今天：{day_str}")
    return sorted(dates, reverse=True)


def main():
    try:
        taiwan_tz = ZoneInfo("Asia/Taipei")
        today = datetime.now(tz=taiwan_tz)
        days = target_days(today, load_pending_dates())
        # 先保留所有目標；中斷或寫入失敗時，下次仍會補抓。
        remaining = set(days)
        save_pending_dates(remaining)
        print(f"[INFO] 今天是 {today.strftime('%Y/%m/%d')}，抓取近 5 天及待補抓日期，共 {len(days)} 天。")

        # 連線到 Google Sheet
        json_keyfile_path = "service_account.json"
        worksheet = connect_google_sheet(
            json_keyfile_path=json_keyfile_path,
            sheet_name="Yung資料庫",           # 你可以改
            worksheet_name="原價屋網路PC組裝數RD"       # 你可以改
        )

        results = {}
        year_dates_cache = {}
        for day_str in days:
            year_folder = f"{int(day_str[:3])}年"
            result = single_attempt_coolpc(year_folder, day_str, year_dates_cache)
            # 尚未上架不在同次執行反覆重試；真正讀取失敗才重開瀏覽器。
            for retry in range(3):
                if result.status != FetchStatus.ERROR:
                    break
                print(f"[WARNING] {day_str} 讀取失敗，重試第 {retry + 1} 次。")
                result = single_attempt_coolpc(year_folder, day_str, year_dates_cache)
            results[day_str] = result
            count = result.count if result.status == FetchStatus.SUCCESS else 0
            update_or_append(worksheet, (day_str, count))
            if result.status == FetchStatus.SUCCESS:
                remaining.remove(day_str)
                save_pending_dates(remaining)

        pending_days = [day for day, result in results.items() if result.status == FetchStatus.PENDING]
        if pending_days:
            print(f"::notice::日期相簿尚未上架，已寫入 0 並列入下次補抓：{', '.join(pending_days)}")

        failed_days = [day for day, result in results.items() if result.status == FetchStatus.ERROR]
        if failed_days:
            print(f"::notice::頁面讀取失敗，已寫入 0 並列入下次補抓：{', '.join(failed_days)}")

    except Exception as e:
        print("[ERROR] 程式出現例外:")
        traceback.print_exc()
        raise


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


def single_attempt_coolpc(year_folder, day_str, year_dates_cache=None):
    """
    嘗試一次開 Selenium、進入「year_folder / day_str」資料夾。
    回傳成功、尚未上架或讀取失敗；主程式將未取得的數值寫為 0。
    """
    driver = None
    stage = "開啟瀏覽器"
    try:
        # 同一次執行已確認缺席的日期，不必逐日重開瀏覽器。
        if year_dates_cache is not None and year_folder in year_dates_cache:
            if day_str not in year_dates_cache[year_folder]:
                print(f"[PENDING] {year_folder}/{day_str} 尚未上架；本次寫入 0，下次補抓。")
                return FetchResult(FetchStatus.PENDING)
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=options
        )
        wait = WebDriverWait(driver, 20)

        stage = "開啟原價屋相簿"
        driver.get("https://www.coolpc.com.tw/photo/#/shared_space/folder/156?_k=tr98b7")
        time.sleep(10)

        # 點「每日組裝分享 (僅網路部)」
        stage = "開啟每日組裝分享"
        share_folder_xpath = "//div[@class='css-106gz8u' and text()='每日組裝分享 (僅網路部)']"
        wait.until(EC.element_to_be_clickable((By.XPATH, share_folder_xpath))).click()
        time.sleep(10)

        # 點「xxx年」資料夾
        stage = "載入年份目錄"
        year_xpath = f"//div[@class='css-106gz8u' and text()='{year_folder}']"
        wait.until(EC.element_to_be_clickable((By.XPATH, year_xpath))).click()
        time.sleep(10)

        # 年份側欄目錄由相簿一次載入，先確認它已展開且有日期清單。
        # 目錄載入逾時仍是 ERROR；只有已載入清單中沒有日期才是 PENDING。
        available_dates = wait.until(lambda browser: loaded_year_dates(browser, year_folder))
        if year_dates_cache is not None:
            year_dates_cache[year_folder] = available_dates
        if day_str not in available_dates:
            print(f"[PENDING] {year_folder}/{day_str} 尚未上架；本次寫入 0，下次補抓。")
            return FetchResult(FetchStatus.PENDING)

        # 點「day_str」資料夾, e.g. 1140206
        stage = "開啟日期相簿"
        day_xpath = f"//div[@class='css-106gz8u' and text()='{day_str}']"
        wait.until(EC.element_to_be_clickable((By.XPATH, day_xpath))).click()
        time.sleep(10)

        # 等footer出現，抓組裝數
        stage = "讀取組裝數"
        footer_elem = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='synofoto-folder-wall-footer']"))
        )
        footer_text = footer_elem.text.strip()  # e.g. "48 個項目"
        count_str = footer_text.split(" ")[0]
        parsed_count = int(count_str)
        if parsed_count < 0:
            raise ValueError("組裝數不可為負數")
        print(f"[INFO] 成功抓到『{day_str}』的組裝數 = {parsed_count}")
        return FetchResult(FetchStatus.SUCCESS, parsed_count)

    except Exception as e:
        print(f"[WARNING] {year_folder}/{day_str} 在「{stage}」失敗，本次寫入 0：{type(e).__name__}: {e}")
        return FetchResult(FetchStatus.ERROR)

    finally:
        try:
            driver.quit()
        except:
            pass


def loaded_year_dates(driver, year_folder):
    year_xpath = f"//div[@class='css-106gz8u' and text()='{year_folder}']/ancestor::li[1]"
    year_items = driver.find_elements(By.XPATH, year_xpath)
    if not year_items:
        return False
    year_item = year_items[0]
    lists = year_item.find_elements(By.XPATH, "./div/ul")
    if not lists or not lists[0].is_displayed():
        return False
    names = lists[0].find_elements(
        By.XPATH, "./li/div[contains(@class, 'synofoto-treebeard-container')]/div[@class='css-106gz8u']"
    )
    dates = {name.text.strip() for name in names}
    if not dates:
        return False
    # 結構或命名異常不能當成「尚未上架」，否則可能吞掉網站改版錯誤。
    for date in dates:
        parse_day(date)
        if int(date[:3]) != int(year_folder[:-1]):
            raise ValueError("年份目錄出現其他年度日期")
    return dates


def update_or_append(worksheet, row_data):
    """
    搜尋整張試算表，避免補抓超過五天的日期時新增重複資料。
    """
    day_str, assemble_count = row_data
    if type(assemble_count) is not int or assemble_count < 0:
        raise ValueError("拒絕將未取得或無效組裝數寫入試算表")
    current_data = worksheet.get_all_values()
    row_count = len(current_data)

    # 如果整張表都還沒資料，就第一行塞入
    if row_count == 0:
        worksheet.update(range_name="A1:B1", values=[[day_str, assemble_count]])
        print(f"[INFO] 試算表是空的，已新增第一行: {row_data}")
        return

    matched_row = None
    for actual_row, row in enumerate(current_data, start=1):
        if len(row) > 0 and row[0] == day_str:
            matched_row = actual_row
            break

    if matched_row:
        cell_range = f"A{matched_row}:B{matched_row}"
        worksheet.update(range_name=cell_range, values=[[day_str, assemble_count]])
        print(f"[INFO] 找到同日期 '{day_str}'，已覆蓋到第 {matched_row} 行，組裝數={assemble_count}")
    else:
        new_row = row_count + 1
        cell_range = f"A{new_row}:B{new_row}"
        worksheet.update(range_name=cell_range, values=[[day_str, assemble_count]])
        print(f"[INFO] 沒找到 '{day_str}'，已新增到第 {new_row} 行，組裝數={assemble_count}")


if __name__ == "__main__":
    main()
