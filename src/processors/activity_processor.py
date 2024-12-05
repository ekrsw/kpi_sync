from .base import BaseProcessor
import pandas as pd
import datetime
import settings
import logging

logger = logging.getLogger(__name__)

class ActivityProcessor(BaseProcessor):
    def process(self):
        """
        Activityファイルのデータを指定された条件でフィルタリングおよび整形します。
        """
        df = self.df.copy()
        result = {}
        try:
            # '件名'列を文字列型に変換
            df['件名'] = df['件名'].astype(str)

            # 件名に「【受付】」が含まれていないもののみ残す
            df = df[~df['件名'].str.contains('【受付】', na=False)]
            logger.info(f"Filtered out rows containing '【受付】'. Remaining rows: {df.shape[0]}")

            # 日付範囲でフィルタリング
            start_date = datetime.date.today()
            end_date = datetime.date.today()
            df = self.filtered_by_date_range(df, '登録日時 (関連) (サポート案件)', start_date, end_date)
            logger.info(f"Filtered DataFrame by date range {start_date} to {end_date}. Remaining rows: {df.shape[0]}")

            # 案件番号でソートし、最も早い日時を残して重複を削除
            df = df.sort_values(by=['案件番号 (関連) (サポート案件)', '登録日時'])
            df = df.drop_duplicates(subset='案件番号 (関連) (サポート案件)', keep='first')
            logger.info(f"Sorted and dropped duplicates. Remaining rows: {df.shape[0]}")

            # '登録日時'と'登録日時 (関連) (サポート案件)'をdatetime型に変換
            # self.df['登録日時'] = pd.to_datetime(self.df['登録日時'], unit='d', origin='1899-12-30')
            # self.df['登録日時 (関連) (サポート案件)'] = pd.to_datetime(self.df['登録日時 (関連) (サポート案件)'], unit='d', origin='1899-12-30')

            # 時間差を計算
            df['時間差'] = df['登録日時'] - df['登録日時 (関連) (サポート案件)']
            df['時間差'] = df['時間差'].fillna(0.0)
            logger.info("Calculated '時間差' column.")

            # 受付タイプが折返しor留守電のものを抽出
            _df_ss_tvs_kmn = df[(df['受付タイプ (関連) (サポート案件)'] == '折返し') | (df['受付タイプ (関連) (サポート案件)'] == '留守電')]
            _df_hhd = df[(df['受付タイプ (関連) (サポート案件)'] == 'HHD入電（折返し）') | (df['受付タイプ (関連) (サポート案件)'] == '留守電')]

            df_ss = _df_ss_tvs_kmn[(_df_ss_tvs_kmn['サポート区分 (関連) (サポート案件)'] == 'SS')]
            df_tvs = _df_ss_tvs_kmn[(_df_ss_tvs_kmn['サポート区分 (関連) (サポート案件)'] == 'TVS')]
            df_kmn = _df_ss_tvs_kmn[(_df_ss_tvs_kmn['サポート区分 (関連) (サポート案件)'] == '顧問先')]
            df_hhd = _df_hhd[(_df_hhd['サポート区分 (関連) (サポート案件)'] == 'HHD')]

            # グループ別に分割
            result['cb_0_20_ss'], result['cb_20_30_ss'], result['cb_30_40_ss'], result['cb_40_60_ss'], result['cb_60over_ss'], result['cb_not_include_ss'] = self.group_activities_by_callback_duration(df_ss)
            result['cb_0_20_tvs'], result['cb_20_30_tvs'], result['cb_30_40_tvs'], result['cb_40_60_tvs'], result['cb_60over_tvs'], result['cb_not_include_tvs'] = self.group_activities_by_callback_duration(df_tvs)
            result['cb_0_20_kmn'], result['cb_20_30_kmn'], result['cb_30_40_kmn'], result['cb_40_60_kmn'], result['cb_60over_kmn'], result['cb_not_include_kmn'] = self.group_activities_by_callback_duration(df_kmn)
            result['cb_0_20_hhd'], result['cb_20_30_hhd'], result['cb_30_40_hhd'], result['cb_40_60_hhd'], result['cb_60over_hhd'], result['cb_not_include_hhd'] = self.group_activities_by_callback_duration(df_hhd)

        except Exception as e:
            logger.error(f"活動データのフィルタリング、整形中にエラーが発生しました。: {e}")
            raise
        
        logger.debug(f"活動処理結果： {result}")
        return result
    
    def group_activities_by_callback_duration(self, df):
        towenty_minutes = settings.SERIAL_20_MINUTES
        thirty_minutes = settings.SERIAL_30_MINUTES
        forty_minutes = settings.SERIAL_40_MINUTES
        sixty_minutes = settings.SERIAL_60_MINUTES

        cb_0_20 = df[(df['時間差'] <= towenty_minutes)].shape[0]
        cb_20_30 = df[(df['時間差'] > towenty_minutes) & (df['時間差'] <= thirty_minutes)].shape[0]
        cb_30_40 = df[(df['時間差'] > thirty_minutes) & (df['時間差'] <= forty_minutes)].shape[0]
        cb_40_60 = df[(df['時間差'] > forty_minutes) & (df['時間差'] <= sixty_minutes) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_60over = df[(df['時間差'] > sixty_minutes) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_not_include =df[(df['時間差'] > sixty_minutes) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')].shape[0]

        return cb_0_20, cb_20_30, cb_30_40, cb_40_60, cb_60over, cb_not_include
    
    def convert_to_pending_num(self, df):
        towenty_minutes = 0.0138888888888889
        thirty_minutes = 0.0208333333333333
        forty_minutes = 0.0277777777777778
        sixty_minutes = 0.0416666666666667

        wfc_over_20 = self.create_wfc_list(df[df['お待たせ時間'] >= towenty_minutes])
        wfc_over30 = self.create_wfc_list(df[df['お待たせ時間'] >= thirty_minutes])
        wfc_over40 = self.create_wfc_list(df[df['お待たせ時間'] >= forty_minutes])
        wfc_over60 = self.create_wfc_list(df[df['お待たせ時間'] >= sixty_minutes])

        return wfc_over_20, wfc_over30, wfc_over40, wfc_over60
    
    def create_wfc_list(self, df) -> str:
        _ = list(df.loc[:, '案件番号 (関連) (サポート案件)'])
        l = map(lambda x: str(x), _)
        return ','.join(l)

    @staticmethod
    def datetime_to_serial(dt: datetime.datetime, base_date=datetime.datetime(1899, 12, 30)) -> float:
        """
        datetimeオブジェクトをシリアル値に変換する。

        :param dt: 変換するdatetimeオブジェクト
        :param base_date: シリアル値の基準日（デフォルトは1899年12月30日）
        :return: シリアル値
        """
        return (dt - base_date).total_seconds() / (24 * 60 * 60)
