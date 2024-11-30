from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import threading

from src.excel_sync import SynchronizedExcelProcessor
import settings


logger = logging.getLogger(__name__)

def process_controll():
    stop_event = threading.Event()

    excel_processor = SynchronizedExcelProcessor(
        file_paths=settings.EXCEL_FILES,
        max_retries=settings.SYNC_MAX_RETRIES,
        retry_delay=settings.SYNC_RETRY_DELAY,
        refresh_interval=settings.REFRESH_INTERVAL
    )

    # scraper処理をここに書く

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []

        # Excelファイルの処理をタスクとして提出
        for file_path in excel_processor.file_paths:
            futures.append(
                executor.submit(excel_processor.process_file, file_path, stop_event)
            )
        
        # scraping処理をタスクとして提出

        try:
            for future in as_completed(futures):
                result = future.result()
        except KeyboardInterrupt:
            logger.info("停止信号を受け取りました。全てのタスクを停止します。")
            stop_event.set()
        
        except Exception as e:
            logger.error(f"エラーが発生しました。: {e}")
            stop_event.set()
        
