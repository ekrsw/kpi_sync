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