from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import threading

from src.excel_processors.excel_sync import SynchronizedExcelProcessor
from src.calculator.kpi_calculator import KpiCalculator
from src.scraper import Scraper
import settings


logger = logging.getLogger(__name__)

def sync_processor() -> dict:
    """
    Excelファイルの処理とスクレイピング処理を同期的に実行する。
    
    Returns
    -------
    dict
        処理結果を格納した辞書型オブジェクト
    """
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
        results = {}

        # Excelファイルの処理をタスクとして追加
        logger.debug("Excelファイルの処理をタスクとして追加しています。")
        for file_path in excel_processor.file_paths:
            futures.append(
                executor.submit(excel_processor.process_file, file_path, stop_event)
            )
        
        # scraping処理をタスクとして追加
        logger.debug("スクレイピングの処理をタスクとして追加しています。")
        futures.append(
            executor.submit(scraper.scrape_ctstage_report, settings.TEMPLATES, stop_event)
        )

        try:
            for future in as_completed(futures):
                result = future.result()
                if isinstance(result, dict):
                    results.update(result)
                else:
                    logger.error(f"処理結果が辞書型ではありません。: {result}")
            return results
        except KeyboardInterrupt:
            logger.info("停止信号を受け取りました。全てのタスクを停止します。")
            stop_event.set()
        
        except Exception as e:
            logger.error(f"エラーが発生しました。: {e}")
            stop_event.set()

def kpi_calculator(results: dict) -> dict:
    """
    KPIを計算する。
    """
    return results

def operator_calculator(results: dict) -> dict:
    """
    オペレーター別のKPIを計算する。
    """
    return results