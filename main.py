import logging
from src.controller import sync_process_controller
import time

import settings

LOG_FILE = settings.LOG_FILE
LOG_LEVEL = getattr(logging, settings.LOG_LEVEL.upper(), logging.INFO)

def setup_logging(log_file, log_level):
    # ロギングの設定
    logging.basicConfig(
        level=log_level, # ログレベル
        format='%(asctime)s:%(levelname)s:%(name)s:%(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='a', encoding='utf-8'),
            logging.StreamHandler() # コンソールへ出力
        ]
    )

setup_logging(LOG_FILE, LOG_LEVEL)
logger = logging.getLogger(__name__)

if __name__ == '__main__':
    start = time.time()
    sync_process_controller()
    end = time.time()
    time_diff = end - start
    logger.info(f"処理が正常に終了しました。（処理時間: {time_diff} 秒）")