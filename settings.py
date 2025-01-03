import datetime
import os
from dotenv import load_dotenv

# .envファイルの読込み
load_dotenv()

# ファイルパスのBaseの設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ロギングの設定
LOG_FILE = "kpi_sync.log"
LOG_LEVEL = "INFO"

# 各種ファイル名とパスの設定
ACTIVITY_FILE_NAME = 'TS_todays_activity.xlsx'
CLOSE_FILE_NAME = 'TS_todays_close.xlsx'
SUPPORT_FILE_NAME = 'TS_todays_support.xlsx'
OPERATORS_FILE_NAME = 'operators.xlsx'
SHIFT_SCHEDULE_NAME = f'{datetime.datetime.now().strftime('%Y%m')}_Campaign_ScheduleList.csv'

ACTIVITY_FILE = os.path.join(BASE_DIR, 'data', ACTIVITY_FILE_NAME)
CLOSE_FILE = os.path.join(BASE_DIR, 'data', CLOSE_FILE_NAME)
SUPPORT_FILE = os.path.join(BASE_DIR, 'data', SUPPORT_FILE_NAME)
OPERATORS_FILE = os.path.join(BASE_DIR, 'data', OPERATORS_FILE_NAME)
SHIFT_SCHEDULE = os.path.join(BASE_DIR, 'data', 'shift_schedule', SHIFT_SCHEDULE_NAME)

EXCEL_FILES = [ACTIVITY_FILE, CLOSE_FILE, SUPPORT_FILE]

# Excel同期処理の設定
SYNC_MAX_RETRIES = 5  # 同期失敗時の最大リトライ回数
SYNC_RETRY_DELAY = 2  # リトライ間の待機時間（秒）
REFRESH_INTERVAL = 5  # 更新が完了するまで待機する時間（秒）

# シリアル値
SERIAL_20_MINUTES = 0.0138888888888889
SERIAL_30_MINUTES = 0.0208333333333333
SERIAL_40_MINUTES = 0.0277777777777778
SERIAL_60_MINUTES = 0.0416666666666667

# CTStageレポーター関係設定
REPORTER_URL = os.getenv('REPORTER_URL')
REPORTER_ID = os.getenv('REPORTER_ID')
HEADLESS_MODE = True
REPORTER_MAX_RETRIES = 5
TEMPLATE_SS = 'TEMPLATE_SS'
TEMPLATE_TVS = 'TEMPLATE_TVS'
TEMPLATE_KMN = 'TEMPLATE_KMN'
TEMPLATE_HHD = 'TEMPLATE_HHD'
TEMPLATE_OP = 'TEMPLATE_OP'
TEMPLATES = [TEMPLATE_SS, TEMPLATE_TVS, TEMPLATE_KMN, TEMPLATE_HHD, TEMPLATE_OP]

USE_ADDITION = True