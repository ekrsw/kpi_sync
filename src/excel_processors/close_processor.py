from src.excel_processors.base import BaseProcessor
import pandas as pd
import datetime
import settings
import logging

logger = logging.getLogger(__name__)


class CloseProcessor(BaseProcessor):
    def process(self):
        df = self.df.copy()
        result = {}
        try:
            logger.debug("クローズデータのフィルタリング、整形を開始します。")
            # 最初の3列をスキップし、5列目をインデックスとして設定します
            try:
                df = df.iloc[:, 3:].set_index(df.columns[5])
                df.reset_index(inplace=True)
            except Exception as e:
                logger.error(f"クローズデータの不要な列の削除と、所有者をインデックスに設定している間にエラーが発生しました。: {e}")
                raise
            
            logger.debug("完了日時を日付型に変換します。")
            try:
                df['完了日時'] = pd.to_datetime(df['完了日時'])
            except Exception as e:
                logger.error(f"完了日時を日付型に変換中にエラーが発生しました。: {e}")
                raise

            # DataFrameのインデックスを日付でソートする
            df.sort_values(by=['完了日時'], inplace=True)
            df.reset_index(drop=True, inplace=True)

            # 日付範囲でフィルタリング
            start_date = datetime.datetime.combine(datetime.date.today(), datetime.time.min)
            end_date = datetime.datetime.combine(datetime.date.today() + datetime.timedelta(days=1), datetime.time.min)
            df = df[(df['完了日時'] >= start_date) & (df['完了日時'] < end_date)]
            df.set_index(['所有者'], inplace=True)

            counts = df.index.value_counts()

            df = pd.DataFrame(counts).reset_index()
            df.columns = ['氏名', 'クローズ']
            df = df.set_index(df.columns[0])

            return df

        except Exception as e:
            logger.error(f"クローズデータのフィルタリング、整形中にエラーが発生しました。: {e}")
            raise