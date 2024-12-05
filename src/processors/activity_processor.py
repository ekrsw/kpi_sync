from .base import BaseProcessor
import pandas as pd
import datetime
import settings
import logging

logger = logging.getLogger(__name__)

TOWENTY_MINUTES = settings.SERIAL_20_MINUTES
THIRTY_MINUTES = settings.SERIAL_30_MINUTES
FORTY_MINUTES = settings.SERIAL_40_MINUTES
SIXTY_MINUTES = settings.SERIAL_60_MINUTES

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
        
        wfc_result = self.waiting_for_callback(start_date, end_date)
        result.update(wfc_result)
        logger.debug(f"活動処理結果： {result}")
        return result
    
    def waiting_for_callback(self, start_date, end_date):
        df = self.df.copy()
        # 受付けタイプ「直受け」「折返し」「留守電」のみ残す
        df = df[(df['受付タイプ (関連) (サポート案件)'] == '折返し') | (df['受付タイプ (関連) (サポート案件)'] == '留守電')]

        # 指標に含めないが「いいえ」のもののみ残す
        df = df[df['指標に含めない (関連) (サポート案件)'] == 'いいえ']

        df = df[(df['顛末コード (関連) (サポート案件)'] == '対応中') | (df['顛末コード (関連) (サポート案件)'] == '対応待ち')]

        # 件名に「【受付】」が含まれているもののみ残す。
        df['件名'] = df['件名'].astype(str) 
        contains_df = df[df['件名'] == '【受付】']
        uncontains_df = df[df['件名'] != '【受付】']

        only_contains_df = pd.merge(contains_df, uncontains_df, on='案件番号 (関連) (サポート案件)', how='outer', indicator=True)
        result = only_contains_df[only_contains_df['_merge'] == 'left_only']
        s = result['案件番号 (関連) (サポート案件)'].unique()
        df = df[df['案件番号 (関連) (サポート案件)'].isin(s)]

        # 案件番号、登録日時でソート
        df.sort_values(by=['案件番号 (関連) (サポート案件)', '登録日時'], inplace=True)

        # 同一案件番号の最初の活動のみ残して他は削除  
        df.drop_duplicates(subset='案件番号 (関連) (サポート案件)', keep='first', inplace=True)
        
        # サポート案件の登録日時と、活動の登録日時をPandas Datetime型に変換して、差分を'お待たせ時間'カラムに格納、NaNは０変換
        current_serial = self.current_time_to_serial()
        df['お待たせ時間'] = (current_serial - df['登録日時 (関連) (サポート案件)'])

        df = self.filtered_by_date_range(df, '登録日時 (関連) (サポート案件)', start_date, end_date)
        df = df.reset_index(drop=True)

        df_ss = df[(df['サポート区分 (関連) (サポート案件)'] == 'SS')]
        df_tvs = df[(df['サポート区分 (関連) (サポート案件)'] == 'TVS')]
        df_kmn = df[(df['サポート区分 (関連) (サポート案件)'] == '顧問先')]
        df_hhd = df[(df['サポート区分 (関連) (サポート案件)'] == 'HHD')]

        df_wfc_over20_ss, df_wfc_over30_ss, df_wfc_over40_ss, df_wfc_over60_ss = self.convert_to_pending_num(df_ss)
        df_wfc_over20_tvs, df_wfc_over30_tvs, df_wfc_over40_tvs, df_wfc_over60_tvs = self.convert_to_pending_num(df_tvs)
        df_wfc_over20_kmn, df_wfc_over30_kmn, df_wfc_over40_kmn, df_wfc_over60_kmn = self.convert_to_pending_num(df_kmn)
        df_wfc_over20_hhd, df_wfc_over30_hhd, df_wfc_over40_hhd, df_wfc_over60_hhd = self.convert_to_pending_num(df_hhd)

        wfc_over20_ss = self.create_wfc_list(df_wfc_over20_ss)
        wfc_over30_ss = self.create_wfc_list(df_wfc_over30_ss)
        wfc_over40_ss = self.create_wfc_list(df_wfc_over40_ss)
        wfc_over60_ss = self.create_wfc_list(df_wfc_over60_ss)
        
        wfc_over20_tvs = self.create_wfc_list(df_wfc_over20_tvs)
        wfc_over30_tvs = self.create_wfc_list(df_wfc_over30_tvs)
        wfc_over40_tvs = self.create_wfc_list(df_wfc_over40_tvs)
        wfc_over60_tvs = self.create_wfc_list(df_wfc_over60_tvs)

        wfc_over20_kmn = self.create_wfc_list(df_wfc_over20_kmn)
        wfc_over30_kmn = self.create_wfc_list(df_wfc_over30_kmn)
        wfc_over40_kmn = self.create_wfc_list(df_wfc_over40_kmn)
        wfc_over60_kmn = self.create_wfc_list(df_wfc_over60_kmn)

        wfc_over20_hhd = self.create_wfc_list(df_wfc_over20_hhd)
        wfc_over30_hhd = self.create_wfc_list(df_wfc_over30_hhd)
        wfc_over40_hhd = self.create_wfc_list(df_wfc_over40_hhd)
        wfc_over60_hhd = self.create_wfc_list(df_wfc_over60_hhd)

        result = {
            'wfc_over20_ss': wfc_over20_ss,
            'wfc_over30_ss': wfc_over30_ss,
            'wfc_over40_ss': wfc_over40_ss,
            'wfc_over60_ss': wfc_over60_ss,
            'wfc_over20_tvs': wfc_over20_tvs,
            'wfc_over30_tvs': wfc_over30_tvs,
            'wfc_over40_tvs': wfc_over40_tvs,
            'wfc_over60_tvs': wfc_over60_tvs,
            'wfc_over20_kmn': wfc_over20_kmn,
            'wfc_over30_kmn': wfc_over30_kmn,
            'wfc_over40_kmn': wfc_over40_kmn,
            'wfc_over60_kmn': wfc_over60_kmn,
            'wfc_over20_hhd': wfc_over20_hhd,
            'wfc_over30_hhd': wfc_over30_hhd,
            'wfc_over40_hhd': wfc_over40_hhd,
            'wfc_over60_hhd': wfc_over60_hhd,
        }
        return result
    
    def group_activities_by_callback_duration(self, df):
        cb_0_20 = df[(df['時間差'] <= TOWENTY_MINUTES)].shape[0]
        cb_20_30 = df[(df['時間差'] > TOWENTY_MINUTES) & (df['時間差'] <= THIRTY_MINUTES)].shape[0]
        cb_30_40 = df[(df['時間差'] > THIRTY_MINUTES) & (df['時間差'] <= FORTY_MINUTES)].shape[0]
        cb_40_60 = df[(df['時間差'] > FORTY_MINUTES) & (df['時間差'] <= SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_60over = df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_not_include =df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')].shape[0]

        return cb_0_20, cb_20_30, cb_30_40, cb_40_60, cb_60over, cb_not_include
    
    
    
    def convert_to_pending_num(self, df):
        wfc_over_20 = df[df['お待たせ時間'] >= TOWENTY_MINUTES]
        wfc_over30 = df[df['お待たせ時間'] >= THIRTY_MINUTES]
        wfc_over40 = df[df['お待たせ時間'] >= FORTY_MINUTES]
        wfc_over60 = df[df['お待たせ時間'] >= SIXTY_MINUTES]

        return wfc_over_20, wfc_over30, wfc_over40, wfc_over60
    
    def callback_classification_by_group(self, df):
        df_cb_0_20 = df[(df['時間差'] <= TOWENTY_MINUTES)]
        df_cb_20_30 = df[(df['時間差'] > TOWENTY_MINUTES) & (df['時間差'] <= THIRTY_MINUTES)]
        df_cb_30_40 = df[(df['時間差'] > THIRTY_MINUTES) & (df['時間差'] <= FORTY_MINUTES)]
        df_cb_40_60 = df[(df['時間差'] > FORTY_MINUTES) & (df['時間差'] <= SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')]
        df_cb_60over = df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')]
        df_cb_not_include =df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')]

        return df_cb_0_20, df_cb_20_30, df_cb_30_40, df_cb_40_60, df_cb_60over, df_cb_not_include
    
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
