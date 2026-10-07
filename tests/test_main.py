import contextlib
import io
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import Mock, patch

from selenium.common.exceptions import TimeoutException

import main as crawler


class CrawlerTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.previous_directory = os.getcwd()
        os.chdir(self.temporary.name)
        self.output = io.StringIO()
        self.redirect = contextlib.redirect_stdout(self.output)
        self.redirect.__enter__()
        self.redirect_errors = contextlib.redirect_stderr(self.output)
        self.redirect_errors.__enter__()

    def tearDown(self):
        self.redirect_errors.__exit__(None, None, None)
        self.redirect.__exit__(None, None, None)
        os.chdir(self.previous_directory)
        self.temporary.cleanup()

    def run_main(self, fetch, update=None):
        def fetch_with_cache(year, day, cache):
            return fetch(year, day)

        with patch.object(crawler, "datetime", wraps=datetime) as clock, \
                patch.object(crawler, "connect_google_sheet", return_value=Mock()), \
                patch.object(crawler, "single_attempt_coolpc", side_effect=fetch_with_cache) as attempts, \
                patch.object(crawler, "update_or_append", side_effect=update) as writes:
            clock.now.return_value = datetime(2026, 10, 7)
            crawler.main()
        return attempts, writes

    def test_pending_dates_do_not_fail_or_write_zero(self):
        def fetch(year, day):
            if day in {"1151005", "1151006"}:
                return crawler.FetchResult(crawler.FetchStatus.PENDING)
            return crawler.FetchResult(crawler.FetchStatus.SUCCESS, 29)

        attempts, writes = self.run_main(fetch)
        self.assertEqual(attempts.call_count, 5)
        self.assertEqual(writes.call_count, 3)
        self.assertEqual(crawler.load_pending_dates(), {"1151005", "1151006"})
        self.assertIn("::notice::", self.output.getvalue())

    def test_actual_errors_still_fail_after_retries(self):
        def fetch(year, day):
            status = crawler.FetchStatus.ERROR if day == "1151006" else crawler.FetchStatus.SUCCESS
            return crawler.FetchResult(status, 0 if status == crawler.FetchStatus.SUCCESS else None)

        with self.assertRaisesRegex(RuntimeError, "頁面讀取失敗.*1151006"):
            self.run_main(fetch)
        self.assertEqual(crawler.load_pending_dates(), {"1151006"})
        self.assertEqual(self.output.getvalue().count("重試第"), 3)

    def test_zero_is_success_and_not_retried(self):
        attempts, writes = self.run_main(lambda year, day: crawler.FetchResult(crawler.FetchStatus.SUCCESS, 0))
        self.assertEqual(attempts.call_count, 5)
        self.assertEqual(writes.call_count, 5)
        self.assertTrue(all(call.args[1][1] == 0 for call in writes.call_args_list))
        self.assertEqual(crawler.load_pending_dates(), set())

    def test_older_pending_date_is_recovered_after_five_days(self):
        crawler.save_pending_dates({"1150920"})
        attempts, writes = self.run_main(lambda year, day: crawler.FetchResult(crawler.FetchStatus.SUCCESS, 12))
        self.assertEqual(attempts.call_count, 6)
        self.assertIn(("1150920", 12), [call.args[1] for call in writes.call_args_list])
        self.assertEqual(crawler.load_pending_dates(), set())

    def test_sheet_failure_keeps_date_for_next_run(self):
        crawler.save_pending_dates({"1150920"})
        with self.assertRaisesRegex(RuntimeError, "sheet unavailable"):
            self.run_main(
                lambda year, day: crawler.FetchResult(crawler.FetchStatus.SUCCESS, 12),
                update=RuntimeError("sheet unavailable"),
            )
        self.assertEqual(crawler.load_pending_dates(), {"1150920", "1151002", "1151003", "1151004", "1151005", "1151006"})

    def test_old_pending_survives_when_still_absent(self):
        crawler.save_pending_dates({"1150920"})
        attempts, _ = self.run_main(lambda year, day: crawler.FetchResult(crawler.FetchStatus.PENDING))
        self.assertEqual(attempts.call_count, 6)
        self.assertIn("1150920", crawler.load_pending_dates())

    def test_corrupt_pending_state_is_not_silently_discarded(self):
        Path("pending_dates.json").write_text('{"version":1,"dates":["1150230"]}')
        with self.assertRaises(ValueError):
            crawler.load_pending_dates()

    def test_cross_year_pending_dates_keep_their_year(self):
        days = crawler.target_days(datetime(2027, 1, 2), {"1151220"})
        self.assertIn("1160101", days)
        self.assertIn("1151231", days)
        self.assertIn("1151220", days)

    def test_future_pending_date_is_rejected(self):
        with self.assertRaises(ValueError):
            crawler.target_days(datetime(2026, 10, 7), {"1151007"})

    def test_old_sheet_row_is_updated_instead_of_duplicated(self):
        sheet = Mock()
        sheet.get_all_values.return_value = [["1150920", "9"]] + [[f"115100{i}", "1"] for i in range(1, 7)]
        crawler.update_or_append(sheet, ("1150920", 12))
        sheet.update.assert_called_once_with(range_name="A1:B1", values=[["1150920", 12]])

    def test_unfetched_or_negative_values_are_never_written(self):
        sheet = Mock()
        for value in (None, -1, True, "0"):
            with self.assertRaises(ValueError):
                crawler.update_or_append(sheet, ("1151006", value))
        sheet.update.assert_not_called()

    def attempt(self, responses):
        driver = Mock()
        wait = Mock()
        wait.until.side_effect = responses
        with patch.object(crawler.webdriver, "Chrome", return_value=driver), \
                patch.object(crawler, "ChromeDriverManager") as manager, \
                patch.object(crawler, "ChromeService"), \
                patch.object(crawler, "WebDriverWait", return_value=wait), \
                patch.object(crawler.time, "sleep"):
            manager.return_value.install.return_value = "/unused-driver"
            result = crawler.single_attempt_coolpc("115年", "1151006")
        driver.quit.assert_called_once()
        return result

    def test_loaded_directory_missing_date_is_pending(self):
        result = self.attempt([Mock(), Mock(), {"1151004"}])
        self.assertEqual(result.status, crawler.FetchStatus.PENDING)
        self.assertIsNone(result.count)

    def test_confirmed_missing_dates_reuse_the_year_listing(self):
        with patch.object(crawler.webdriver, "Chrome") as browser:
            result = crawler.single_attempt_coolpc("115年", "1151006", {"115年": {"1151004"}})
        self.assertEqual(result.status, crawler.FetchStatus.PENDING)
        browser.assert_not_called()

    def test_directory_load_timeout_is_error(self):
        result = self.attempt([Mock(), Mock(), TimeoutException("year list unavailable")])
        self.assertEqual(result.status, crawler.FetchStatus.ERROR)
        self.assertIn("載入年份目錄", self.output.getvalue())
        self.assertIn("TimeoutException", self.output.getvalue())

    def test_existing_date_click_timeout_is_error(self):
        result = self.attempt([Mock(), Mock(), {"1151006"}, TimeoutException("date unclickable")])
        self.assertEqual(result.status, crawler.FetchStatus.ERROR)
        self.assertIn("開啟日期相簿", self.output.getvalue())

    def test_count_parse_error_is_error(self):
        result = self.attempt([Mock(), Mock(), {"1151006"}, Mock(), Mock(text="讀取中")])
        self.assertEqual(result.status, crawler.FetchStatus.ERROR)

    def test_zero_count_in_source_is_valid(self):
        result = self.attempt([Mock(), Mock(), {"1151006"}, Mock(), Mock(text="0 個項目")])
        self.assertEqual(result, crawler.FetchResult(crawler.FetchStatus.SUCCESS, 0))

    def test_directory_has_not_loaded_until_expanded_and_populated(self):
        driver = Mock()
        year = Mock()
        listing = Mock()
        driver.find_elements.return_value = [year]
        year.find_elements.return_value = [listing]
        listing.is_displayed.return_value = False
        self.assertFalse(crawler.loaded_year_dates(driver, "115年"))
        listing.is_displayed.return_value = True
        listing.find_elements.return_value = []
        self.assertFalse(crawler.loaded_year_dates(driver, "115年"))
        listing.find_elements.return_value = [Mock(text="1151004")]
        self.assertEqual(crawler.loaded_year_dates(driver, "115年"), {"1151004"})

    def test_unexpected_directory_structure_is_error(self):
        driver = Mock()
        year = Mock()
        listing = Mock()
        driver.find_elements.return_value = [year]
        year.find_elements.return_value = [listing]
        listing.is_displayed.return_value = True
        for name in ("讀取中", "1141004"):
            listing.find_elements.return_value = [Mock(text=name)]
            with self.assertRaises(ValueError):
                crawler.loaded_year_dates(driver, "115年")


if __name__ == "__main__":
    unittest.main()
