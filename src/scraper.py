from bs4 import BeautifulSoup
import datetime
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
from typing import List
import pandas as pd

import settings


logger = logging.getLogger(__name__)

class Base:
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
            
            logger.info("driverを作成しています。")
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
    
    def login(self) -> None:
        """レポーターにログイン"""
        try:
            logger.info("ログインを試みています。")
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

    def call_template(self, template: str) -> None:
        """
        テンプレート呼び出し

        Parameters
        ----------
        template : List[str]
            表示したいテンプレート。
        """
        try:
            # テンプレート呼び出し
            logger.info(f"テンプレートを呼び出しています。: [{template}]")
            self.driver.find_element(By.ID, 'template-title-span').click()
            el2 = self.driver.find_element(By.ID, 'template-download-select')
            s2 = Select(el2)
            s2.select_by_value(template)
            self.driver.find_element(By.ID, 'template-creation-btn').click()
        except Exception as e:
            logger.error(f"テンプレートの呼び出しに失敗しました。: {e}")

    def create_report(self, element_id: str = "0"):
        """
        レポート作成
        
        Parameters
        ----------
        element_id : str
            選択するタブによって、"0" or "1"
        """
        try:
            self.driver.find_element(By.ID, f'panel-td-create-report-{element_id}').click()
        except Exception as e:
            logger.error(f"レポートの作成に失敗しました。: {e}")

    def select_tabs(self, tab_element_id: str = "1"):
        """
        レポートのタブ切り替え

        Parameters
        ----------
        tab_element_id : str
            選択するタブによって、"1" or "2"
        """
        self.driver.find_element(By.ID, f'normal-title{tab_element_id}').click()
    
    def create_dateframe(self, list_name: str) -> pd.DataFrame:
        """
        
        'normal-list1-dummy-0'
        'normal-list2-dummy-1'
        """
        try:
            logger.debug("ページソースをエンコードしています。")
            html = self.driver.page_source.encode('utf-8')
        except Exception as e:
            logger.error(f"ページソースをUTF-8でエンコード中にエラーが発生しました。: {e}")
        try:
            logger.debug("HTMLをパースしています。")
            soup = BeautifulSoup(html, 'lxml')
        except Exception as e:
            logger.error(f"HTMLのパース中にエラーが発生しました。: {e}")
        
        data_table = []

        # headerのリストを作成
        header_table = soup.find(id=f'{list_name}-table-head-table')
        xmp = header_table.thead.tr.find_all('xmp')
        header_list = [i.string for i in xmp]
        data_table.append(header_list)

        # bodyのリストを作成
        body_table = soup.find(id=f'{list_name}-table-body-table')
        tr = body_table.tbody.find_all('tr')
        for td in tr:
            xmp = td.find_all('xmp')
            row = [i.string for i in xmp]
            data_table.append(row)
        
        # テーブルをDataFrameに変換
        df = pd.DataFrame(data_table[1:], columns=data_table[0])
        df.set_index(df.columns[0], inplace=True)
        return df

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


class Scraper(Base):
    def scrape_ctstage_report(self, templates: List[str], stop_event):
        results = {}
        if stop_event.is_set():
            logger.info(f"スクレイピング処理が停止されました。")
        try:
            self.create_driver()
            self.login()
            for template in templates:
                retries = 0
                while retries < settings.REPORTER_MAX_RETRIES and not stop_event.is_set():
                    try:
                        self.call_template(template)
                        self.create_report(element_id="0")
                        time.sleep(1)

                        df1 = self.create_dateframe('normal-list1-dummy-0')

                        self.select_tabs(tab_element_id="2")
                        self.create_report(element_id="1")
                        time.sleep(1)

                        df2 = self.create_dateframe('normal-list2-dummy-1')

                        # 必要なデータを辞書に保存
                        template_result = {
                            "total_calls": int(df1.iloc[0, 0]), # 総着信数
                            "IVR_interruptions_before_response": int(df1.iloc[0, 1]), # IVR応答前放棄呼数
                            "ivr_interruptions": int(df1.iloc[0, 2]), # IVR切断数
                            "time_out": int(df2.iloc[0, 0]), # タイムアウト数
                            "abandoned_during_operator": int(df2.iloc[0, 1]) # ACD放棄呼数
                        }
                        results[template] = template_result
                        break

                    except Exception as e:
                        retries += 1
                        logger.error(f"{template}のスクレイピング中にエラーが発生しました({retries}回目)。: {e}")
                        self.close_driver()
                        self.create_driver()
                        self.login()
            return results
                        
        except Exception as e:
            logger.error(f"スクレイピング中に予期しないエラーが発生しました。")
            return results
        finally:
            self.close_driver()
            logger.debug("Web Driverを終了しました。")
