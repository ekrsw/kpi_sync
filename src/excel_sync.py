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
        retry_delay : int, optional
            リトライ間の待機時間（秒、デフォルトは設定ファイルから）。
        refresh_interval : int, optional
            CalculationState を確認する際の待機時間（秒、デフォルトは設定ファイルから）。。
        """
        self.file_paths = file_paths
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.refresh_interval = refresh_interval

    def process_file(self, file_path, stop_event) -> dict:
        """
        このクラスのメインのメソッド。
        これを実行するだけ。

        Parameters
        ----------
        file_path : str
            処理するExcelファイルのパス。
        stop_event : threading.Event
            処理を停止するためのイベント。
        
        Returns
        -------
        dict
            ファイルの処理結果。
        """
        self._sync_file(file_path, stop_event)

        if settings.ACTIVITY_FILE in file_path:
            from .processors.activity_processor import ActivityProcessor
            activity = ActivityProcessor(file_path)
            activity.load_data()
            result = activity.process()
            return result
        elif settings.CLOSE_FILE in file_path:
            result = {}
            return result
        elif settings.SUPPORT_FILE in file_path:
            from .processors.support_processor import SupportProcessor
            support = SupportProcessor(file_path)
            support.load_data()
            result = support.process()
            return result
        else:
            logger.error(f"ファイル名がPathに含まれていません。{file_path}")

    
    def _sync_file(self, file_path, stop_event) -> None:
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
    def check_and_close(file_names: List[str]) -> None:
        """
        指定されたExcelファイルが開かれているかを確認し、開かれている場合はExcelアプリケーションを強制終了します。
        
        Parameters
        ----------
        file_names : List[str]
            確認するExcelファイルのパスのリスト。
        """
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
    