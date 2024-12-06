import logging
from src.controller import sync_processor
from src.calculator.kpi_calculator import KpiCalculator
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
    result = sync_processor()
    kpi_calculator = KpiCalculator(result)
    
    for k, v in result.items():
        print(f"{k}: {v}")
    
    print("総着信数: ", kpi_calculator.total_calls('TVS'))
    print("自動音声ガイダンス途中切断数: ", kpi_calculator.ivr_interruptions('TVS'))
    print("オペレーター呼出途中放棄数: ", kpi_calculator.abandoned_during_operator('TVS'))
    print("留守電数: ", kpi_calculator.voicemails('TVS'))
    print("留守電放棄件数: ", kpi_calculator.abandoned_in_ivr('TVS'))
    print("放棄呼数: ", kpi_calculator.abandoned_calls('TVS'))
    print("応答件数: ", kpi_calculator.responses('TVS'))
    print("応答率: ", kpi_calculator.response_rate('TVS'))
    print("電話問い合わせ件数: ", kpi_calculator.phone_inquiries('TVS'))
    print("直受け対応件数: ", kpi_calculator.direct_handling('TVS'))
    

    end = time.time()
    time_diff = end - start
    logger.info(f"処理が正常に終了しました。（処理時間: {time_diff} 秒）")

    