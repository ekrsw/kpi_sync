import datetime
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from typing import List
import pandas as pd

import settings


logger = logging.getLogger(__name__)

class Scraper:
    def __init__(self,
                 url: str = settings.REPORTER_URL,
                 id: str = settings.REPORTER_ID) -> None:
        self.url = url
        self.id = id
        self.df = pd.DataFrame()
        self.driver = None

    def create_driver(self) -> None:
        try:
            if self.driver:
                self.close_driver()
            
            logger.info("driverを作成します。")
            options = Options()

            # ブラウザを表示させるか
            if settings.HEADLESS_MODE:
                options.add_argument('--headless')
            
            # コマンドプロンプトのログを表示させない。
            options.add_argument('--disable-logging')
            options.add_argument('--disable-extensions')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--log-level=3')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])

            # driverの作成
            self.driver = webdriver.Chrome(options=options)
            self.driver.implicitly_wait(5)
        except Exception as e:
            logger.error(f"driverの作成に失敗しました。: {e}")
    
    def login(self):
        """レポーターにログイン"""
        try:
            self.driver.get(self.url)
            logger.debug(f"URL {self.url}にアクセスしました。")

            # ID入力
            self.driver.find_element(By.ID, 'logon-operator-id').send_keys(self.id)
            logger.debug("IDを入力しました。")

            # ログインボタンをクリック
            self.driver.find_element(By.ID, 'logon-btn').click()
            logger.debug("ログインボタンをクリックしました。")

        except Exception as e:
            logger.error(f"ログインに失敗しました。: {e}")
            raise

    def call_template(self, template: List[str]):
        """テンプレート呼び出し、指定の集計期間を表示"""
        try:
            # テンプレート呼び出し
            self.driver.find_element(By.ID, 'template-title-span').click()
            el1 = self.driver.find_element(By.ID, 'download-open-range-select')
            s1 = Select(el1)
            s1.select_by_visible_text(template[0])
            el2 = self.driver.find_element(By.ID, 'template-download-select')
            s2 = Select(el2)
            s2.select_by_value(template[1])
            self.driver.find_element(By.ID, 'template-creation-btn').click()
        except Exception as e:
            logger.error(f"テンプレートの呼び出しに失敗しました。: {e}")

    def filter_by_date(self, start_date, end_date, input_id):
        pass

    def select_tabs(self, tab_id):
        pass

    def close_driver(self):
        """driverを閉じる"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("driverを正常に閉じました。")
            except Exception as e:
                logger.error("driverの閉鎖に失敗しました。: {e}")
            finally:
                self.driver = None