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
    
    print("TVS_総着信数: ", kpi_calculator.total_calls('TVS'))
    print("TVS_オペレーター呼出途中放棄数: ", kpi_calculator.abandoned_during_operator('TVS'))
    print("TVS_留守電数: ", kpi_calculator.voicemails('TVS'))
    print("TVS_留守電放棄件数: ", kpi_calculator.abandoned_in_ivr('TVS'))
    print("TVS_放棄呼数: ", kpi_calculator.abandoned_calls('TVS'))
    print("TVS_応答件数: ", kpi_calculator.responses('TVS'))
    print("TVS_応答率: ", kpi_calculator.response_rate('TVS'))
    print("TVS_電話問い合わせ件数: ", kpi_calculator.phone_inquiries('TVS'))
    print("TVS_直受け対応件数: ", kpi_calculator.direct_handling('TVS'))
    print("TVS_電話問い合わせ件数: ", kpi_calculator.phone_inquiries('TVS'))
    print("TVS_直受け対応件数: ", kpi_calculator.direct_handling('TVS'))
    print("TVS_直受け率: ", kpi_calculator.direct_handling_rate('TVS'))
    print("お待たせ20分以上: ", kpi_calculator.waiting_for_callback_count_over_20min('TVS'))
    
    print("SS_総着信数: ", kpi_calculator.total_calls('SS'))
    print("SS_オペレーター呼出途中放棄数: ", kpi_calculator.abandoned_during_operator('SS'))
    print("SS_留守電数: ", kpi_calculator.voicemails('SS'))
    print("SS_留守電放棄件数: ", kpi_calculator.abandoned_in_ivr('SS'))
    print("SS_放棄呼数: ", kpi_calculator.abandoned_calls('SS'))
    print("SS_応答件数: ", kpi_calculator.responses('SS'))
    print("SS_応答率: ", kpi_calculator.response_rate('SS'))
    print("SS_電話問い合わせ件数: ", kpi_calculator.phone_inquiries('SS'))
    print("SS_直受け対応件数: ", kpi_calculator.direct_handling('SS'))
    print("SS_電話問い合わせ件数: ", kpi_calculator.phone_inquiries('SS'))
    print("SS_直受け対応件数: ", kpi_calculator.direct_handling('SS'))
    print("SS_直受け率: ", kpi_calculator.direct_handling_rate('SS'))
    print("お待たせ20分以上: ", kpi_calculator.waiting_for_callback_count_over_20min('SS'))

    end = time.time()
    time_diff = end - start
    logger.info(f"処理が正常に終了しました。（処理時間: {time_diff} 秒）")

    