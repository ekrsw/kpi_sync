import concurrent.futures
import logging
import os
import pythoncom  # COM初期化に必要
import threading
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
                    else:
                        logger.info(f"{file_path}の同期を再試行します。")
                        time.sleep(self.retry_delay)
                try:
                    excel.Quit()
                    logger.info("Excelアプリケーションを終了しました。")
                except Exception as quit_e:
                    logger.warning(f"Excelの終了中にエラーが発生しました。: {quit_e}")

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
    