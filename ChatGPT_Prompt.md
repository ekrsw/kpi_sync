# プロジェクト構成
kpi_sync/
├── data/
│   ├── TS_todays_activity.xlsx
│   ├── TS_todays_close.xlsx
│   └── TS_todays_support.xlsx
├── src/
│   ├── controller.py
│   ├── excel_sync.py
│   └── scraper.py
├── .env
├── main.py
├── requirements.txt
└── settings.py

# コード詳細
## src/controller.py
```controller.py
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import threading

from src.excel_sync import SynchronizedExcelProcessor
from src.scraper import Scraper
import settings


logger = logging.getLogger(__name__)

def sync_process_controller():
    stop_event = threading.Event()

    excel_processor = SynchronizedExcelProcessor(
        file_paths=settings.EXCEL_FILES,
        max_retries=settings.SYNC_MAX_RETRIES,
        retry_delay=settings.SYNC_RETRY_DELAY,
        refresh_interval=settings.REFRESH_INTERVAL
    )

    # scraper処理をここに書く
    scraper = Scraper()

    # Excelが開いているかを確認して開いている場合はExcelを強制終了する。
    SynchronizedExcelProcessor.check_and_close(settings.EXCEL_FILES)

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []

        # Excelファイルの処理をタスクとして追加
        logger.debug("Excelファイルの処理をタスクとして追加しています。")
        for file_path in excel_processor.file_paths:
            futures.append(
                executor.submit(excel_processor.process_file, file_path, stop_event)
            )
        
        # scraping処理をタスクとして追加
        logger.debug("スクレイピングの処理をタスクとして追加しています。")
        futures.append(
            executor.submit(scraper.test, settings.TEMPLATES, stop_event)
        )

        try:
            for future in as_completed(futures):
                result = future.result()
        except KeyboardInterrupt:
            logger.info("停止信号を受け取りました。全てのタスクを停止します。")
            stop_event.set()
        
        except Exception as e:
            logger.error(f"エラーが発生しました。: {e}")
            stop_event.set()
```
## src/excel_sync.py
```excel_sync.py
import logging
import openpyxl
import os
import pythoncom  # COM初期化に必要
import time
from typing import List
import win32com.client

import settings

# ロガーの設定
logger = logging.getLogger(__name__)

class SynchronizedExcelProcessor:
    def __init__(self, file_paths: List[str],
                 max_retries: int = settings.SYNC_MAX_RETRIES,
                 retry_delay: int = settings.SYNC_RETRY_DELAY,
                 refresh_interval: int = settings.REFRESH_INTERVAL
                 ) -> None:
        """
        Excelファイルの同期処理を管理するクラス。

        Parameters
        ----------
        file_paths : List[str]
            同期するExcelファイルのパスのリスト。
        max_retries : int, optional
            同期失敗時の最大リトライ回数（デフォルトは設定ファイルから）。
        retry_delay : float, optional
            リトライ間の待機時間（秒、デフォルトは設定ファイルから）。
        refresh_interval : int, optional
            CalculationState を確認する際の待機時間（秒、デフォルトは設定ファイルから）。。
        """
        self.file_paths = file_paths
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.refresh_interval = refresh_interval

    def process_file(self, file_path, stop_event):
        """
        個別のExcelファイルを処理します。

        Parameters
        ----------
        file_path : str
            処理するExcelファイルのパス。
        """
        logger.debug(f"{file_path}の処理を開始します。")
        if stop_event.is_set():
            logger.info(f"{file_path}の処理が停止されました。")
            return
        if not os.path.exists(file_path):
            logger.warning(f"ファイルが存在しません。: {file_path}")
            return
        
        try:
            # COMライブラリを初期化
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            logger.debug("COMライブラリを初期化しました。")

            # Excelアプリケーション作成
            excel = self._create_excel_app()
            logger.debug('Excelアプリケーションを作成しました。')

            logger.info(f"{file_path}の同期を開始します。")
            retries = 0

            while retries < self.max_retries and not stop_event.is_set():
                try:
                    workbook = excel.Workbooks.Open(file_path)
                    logger.debug("ワークブックを開きました。")
                    workbook.RefreshAll()
                    logger.debug("ワークブックを更新しています。")
                    time.sleep(self.refresh_interval)
                    workbook.Save()
                    logger.debug("ワークブックを保存しました。")
                    workbook.Close()
                    logger.info(f"{file_path}の同期が完了しました。")
                    break
                except Exception as e:
                    retries += 1
                    logger.info(f"{file_path}の同期中にエラーが発生しました。（{retries}回目）: {e}")
                    if retries >= self.max_retries:
                        logger.error(f"{file_path}の同期に失敗しました。最大リトライ回数に達しました。: {e}")
                        # リトライ回数を超えたのでwhileループを抜けます。
                        break

                    else:
                        logger.info(f"{file_path}の同期を再試行します。")
                        time.sleep(self.retry_delay)
                    
                    # Excelのクローズ処理を行います。
                    self._close_app(excel)

                    # 次の処理に備えてExcelアプリケーションを作成します。
                    excel = self._create_excel_app()
                    logger.info('Excelアプリケーションを作成しました。')

        except Exception as e:
            logger.error(f"{file_path}の同期処理中に予期しないエラーが発生しました。{e}")
        finally:
            # COMライブラリを終了
            pythoncom.CoUninitialize()
            logger.debug("COMライブラリを終了しました。")
    
    def _create_excel_app(self) -> win32com.client.CDispatch:
        """
        Excelアプリケーションを起動し、設定を行うヘルパーメソッド。

        Returns
        -------
        excel_app : COMObject
            起動したExcelアプリケーションオブジェクト。
        """
        logger.info("Excelアプリケーションを起動します。")
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        return excel_app
    
    @staticmethod
    def check_and_close(file_names):
        for file_name in file_names:
            try:
                with open(file_name, "r+b"):
                    logger.debug(f"{file_name}は閉じられています。")
            except PermissionError:
                logger.info(f"{file_name}は開かれています。")
                try:
                    os.system("taskkill /F /IM excel.exe")  # Excelプロセスを強制終了
                    logger.info("Excelアプリケーションを強制終了しました。")
                except Exception as e:
                    logger.error(f"Excelを強制終了する際にエラーが発生しました: {e}")
            except Exception as e:
                logger.error(f"{file_name}の確認中にエラーが発生しました。Excelアプリケーションを強制終了します。: {e}")
                try:
                    os.system("taskkill /F /IM excel.exe")  # Excelプロセスを強制終了
                    logger.info("Excelアプリケーションを強制終了しました。")
                except Exception as e:
                    logger.error(f"Excelを強制終了する際にエラーが発生しました: {e}")

    def _close_app(self, excel) -> None:
        """
        Excelアプリケーションを終了するヘルパー関数。
        """
        try:
            excel.Quit()
            logger.info("Excelアプリケーションを終了しました。")
        except Exception as quit_e:
            logger.warning(f"Excelの終了中にエラーが発生しました。: {quit_e}")
```
## src/scraper.py
```scraper.py
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
    def test(self, templates: List[str], stop_event):
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

                        print("総着信数: ", df1.iloc[0, 0]) # 総着信数
                        print("IVR応答前放棄呼数: ", df1.iloc[0, 1]) # IVR応答前放棄呼数
                        print("IVR切断数: ", df1.iloc[0, 2]) # IVR切断数
                        print("タイムアウト数: ", df2.iloc[0, 0]) # タイムアウト数
                        print("ACD放棄呼数: ", df2.iloc[0, 1]) # ACD放棄呼数
                        break

                    except Exception as e:
                        retries += 1
                        logger.error(f"{template}のスクレイピング中にエラーが発生しました({retries}回目)。: {e}")
                        self.close_driver()
                        self.create_driver()
                        self.login()
                        
        except Exception as e:
            logger.error(f"スクレイピング中に予期しないエラーが発生しました。")
        finally:
            self.close_driver()
            logger.debug("Web Driverを終了しました。")
```
## .env
```.env
REPORTER_URL = http://example.com
REPORTER_ID = my_id
```
## main.py
```main.py
import logging
from src.controller import sync_process_controller
import time

import settings

LOG_FILE = settings.LOG_FILE

def setup_logging(log_file):
    # ロギングの設定
    logging.basicConfig(
        level=logging.INFO, # ログレベル
        format='%(asctime)s:%(levelname)s:%(name)s:%(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='a', encoding='utf-8'),
            logging.StreamHandler() # コンソールへ出力
        ]
    )

setup_logging(LOG_FILE)
logger = logging.getLogger(__name__)

if __name__ == '__main__':
    start = time.time()
    sync_process_controller()
    end = time.time()
    time_diff = end - start
    logger.info(f"処理が正常に終了しました。（処理時間: {time_diff} 秒）")
    print(f"処理時間: {time_diff}")
```
## requirements.txt
```requirements.txt
pandas>=1.0.0
requests>=2.0.0
beautifulsoup4>=4.0.0
lxml==5.3.0
openpyxl>=3.0.0
pywin32>=300
python-dotenv>=1.0.0
asyncio
selenium==4.27.1
```
## settings.py
```settings.py
import os
from dotenv import load_dotenv

# .envファイルの読込み
load_dotenv()

# ファイルパスのBaseの設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ログファイルのパス
LOG_FILE = "kpi_sync.log"

# Excelファイルの名前とパス
ACTIVITY_FILE_NAME = 'TS_todays_activity.xlsx'
CLOSE_FILE_NAME = 'TS_todays_close.xlsx'
SUPPORT_FILE_NAME = 'TS_todays_support.xlsx'

ACTIVITY_FILE = os.path.join(BASE_DIR, 'data', ACTIVITY_FILE_NAME)
CLOSE_FILE = os.path.join(BASE_DIR, 'data', CLOSE_FILE_NAME)
SUPPORT_FILE = os.path.join(BASE_DIR, 'data', SUPPORT_FILE_NAME)

EXCEL_FILES = [ACTIVITY_FILE, CLOSE_FILE, SUPPORT_FILE]

# Excel同期処理の設定
SYNC_MAX_RETRIES = 5  # 同期失敗時の最大リトライ回数
SYNC_RETRY_DELAY = 2  # リトライ間の待機時間（秒）
REFRESH_INTERVAL = 5  # 更新が完了するまで待機する時間（秒）

# CTStageレポーター関係設定
REPORTER_URL = os.getenv('REPORTER_URL')
REPORTER_ID = os.getenv('REPORTER_ID')
HEADLESS_MODE = False
REPORTER_MAX_RETRIES = 5
TEMPLATE_SS = 'TEMPLATE_SS'
TEMPLATE_TVS = 'TEMPLATE_TVS'
TEMPLATE_KMN = 'TEMPLATE_KMN'
TEMPLATE_HHD = 'TEMPLATE_HHD'
TEMPLATES = [TEMPLATE_SS, TEMPLATE_TVS, TEMPLATE_KMN, TEMPLATE_HHD]
```