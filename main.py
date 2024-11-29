import logging
from src.excel_sync import SynchronizedExcelProcessor

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
    s = SynchronizedExcelProcessor(settings.EXCEL_FILES)
    s.start()
    s.thread.join()
