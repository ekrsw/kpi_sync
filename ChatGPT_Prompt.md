# プロジェクト構成
kpi_sync/
├── data/
│   ├── TS_todays_activity.xlsx
│   ├── TS_todays_close.xlsx
│   └── TS_todays_support.xlsx
├── src/
│   ├── calculator/
│   │   ├── kpi_calculator.py
│   │   └── operator_calculator.py
│   ├── excel_processors/
│   │   ├── activity_processor.py
│   │   ├── base.py
│   │   ├── close_processor.py
│   │   ├── excel_sync.py
│   │   └── support_processor.py
│   ├── controller.py
│   ├── scraper.py
│   └── views.py
├── .env
├── main.py
├── requirements.txt
└── settings.py

# コード詳細
## src/calculator/kpi_calculator.py
```
import logging

logger = logging.getLogger(__name__)

class KpiCalculator:
    TEMPLATE_MAP = {
        'SS': 'TEMPLATE_SS',
        'TVS': 'TEMPLATE_TVS',
        'KMN': 'TEMPLATE_KMN',
        'HHD': 'TEMPLATE_HHD'
    }

    # 留守電キー
    IVR_MAP = {
        'SS': 'ivr_ss',
        'TVS': 'ivr_tvs',
        'KMN': 'ivr_kmn',
        'HHD': 'ivr_hhd'
    }

    # 直受け対応件数キー
    DIRECT_MAP = {
        'SS': 'direct_ss',
        'TVS': 'direct_tvs',
        'KMN': 'direct_kmn',
        'HHD': 'direct_hhd'
    }

    # コールバック対応件数キー(待ち時間別)
    CB_MAP = {
        '0_20': {'SS': 'cb_0_20_ss', 'TVS': 'cb_0_20_tvs', 'KMN': 'cb_0_20_kmn', 'HHD': 'cb_0_20_hhd'},
        '20_30': {'SS': 'cb_20_30_ss', 'TVS': 'cb_20_30_tvs', 'KMN': 'cb_20_30_kmn', 'HHD': 'cb_20_30_hhd'},
        '30_40': {'SS': 'cb_30_40_ss', 'TVS': 'cb_30_40_tvs', 'KMN': 'cb_30_40_kmn', 'HHD': 'cb_30_40_hhd'},
        '40_60': {'SS': 'cb_40_60_ss', 'TVS': 'cb_40_60_tvs', 'KMN': 'cb_40_60_kmn', 'HHD': 'cb_40_60_hhd'},
        '60over': {'SS': 'cb_60over_ss', 'TVS': 'cb_60over_tvs', 'KMN': 'cb_60over_kmn', 'HHD': 'cb_60over_hhd'}
    }

    # お待たせ対応件数キー(待ち時間別)
    WFC_MAP = {
        '20over': {'SS': 'wfc_over20_ss', 'TVS': 'wfc_over20_tvs', 'KMN': 'wfc_over20_kmn', 'HHD': 'wfc_over20_hhd'},
        '30over': {'SS': 'wfc_over30_ss', 'TVS': 'wfc_over30_tvs', 'KMN': 'wfc_over30_kmn', 'HHD': 'wfc_over30_hhd'},
        '40over': {'SS': 'wfc_over40_ss', 'TVS': 'wfc_over40_tvs', 'KMN': 'wfc_over40_kmn', 'HHD': 'wfc_over40_hhd'},
        '60over': {'SS': 'wfc_over60_ss', 'TVS': 'wfc_over60_tvs', 'KMN': 'wfc_over60_kmn', 'HHD': 'wfc_over60_hhd'}
    }

    def __init__(self, data: dict):
        self.data = data

    def _select_template(self, group: str) -> str:
        if group not in self.TEMPLATE_MAP:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
        return self.TEMPLATE_MAP[group]

    def _get_ivr_key(self, group: str) -> str:
        if group not in self.IVR_MAP:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
        return self.IVR_MAP[group]

    def _get_direct_key(self, group: str) -> str:
        if group not in self.DIRECT_MAP:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
        return self.DIRECT_MAP[group]

    def _get_cb_key(self, group: str, time_range: str) -> str:
        if group not in self.CB_MAP[time_range]:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
        return self.CB_MAP[time_range][group]
    
    def _get_wfc_key(self, group: str, time_range: str) -> str:
        if group not in self.WFC_MAP[time_range]:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
        return self.WFC_MAP[time_range][group]

    @staticmethod
    def _calc_rate(a: int, b: int, wfc: int = 0) -> float:
        denominator = b + wfc
        return a / denominator if denominator != 0 else 0.0
    @staticmethod
    def _calc_count() -> int:
        pass

    def total_calls(self, group: str) -> int:
        """ 11_総着信数 (int): reporter_着信数 """
        template = self._select_template(group)
        return self.data[template]['total_calls']

    def ivr_interruptions(self, group: str) -> int:
        """ 12_自動音声ガイダンス途中切断数 (int): reporter_IVR応答前放棄呼数 + reporter_IVR切断数 """
        template = self._select_template(group)
        return (self.data[template]['IVR_interruptions_before_response']
                + self.data[template]['ivr_interruptions'])

    def abandoned_during_operator(self, group: str) -> int:
        """ 14_オペレーター呼出途中放棄数 (int): reporter_ACD放棄呼数 """
        template = self._select_template(group)
        return self.data[template]['abandoned_during_operator']

    def voicemails(self, group: str) -> int:
        """ 16_留守電数 (int): S_留守電 """
        ivr_key = self._get_ivr_key(group)
        return self.data[ivr_key]

    def abandoned_in_ivr(self, group: str) -> int:
        """ 15_留守電放棄件数 (int): reporter_タイムアウト数 - 16_留守電数 """
        template = self._select_template(group)
        return self.data[template]['time_out'] - self.voicemails(group)

    def abandoned_calls(self, group: str) -> int:
        """ 13_放棄呼数 (int): 14 + 15 """
        return self.abandoned_during_operator(group) + self.abandoned_in_ivr(group)

    def responses(self, group: str) -> int:
        """ 17_応答件数 (int): 11 - 12 - 13 """
        return self.total_calls(group) - self.ivr_interruptions(group) - self.abandoned_calls(group)

    def response_rate(self, group: str) -> float:
        """ 応答率 """
        return self._calc_rate(self.responses(group), self.total_calls(group))

    def phone_inquiries(self, group: str) -> int:
        """ 18_電話問い合わせ件数 (int): 16 + 17 """
        return self.voicemails(group) + self.responses(group)

    def direct_handling(self, group: str) -> int:
        """ 21_直受け対応件数 (int): support_case_直受け """
        direct_key = self._get_direct_key(group)
        return self.data[direct_key]

    def direct_handling_rate(self, group: str) -> float:
        """ 直受率: 21 / 18 """
        return self._calc_rate(self.direct_handling(group), self.phone_inquiries(group))

    def callback_count_0_to_20_min(self, group: str) -> int:
        """ 23_お待たせ0分～20分対応件数 (int) """
        return self.data[self._get_cb_key(group, '0_20')]

    def cumulative_callback_under_20_min(self, group: str) -> int:
        """ 24_お待たせ20分以内累計対応件数 (int): 21 + 23 """
        return self.direct_handling(group) + self.callback_count_0_to_20_min(group)

    def callback_count_20_to_30_min(self, group: str) -> int:
        """ 25_お待たせ20分～30分対応件数 (int) """
        return self.data[self._get_cb_key(group, '20_30')]

    def cumulative_callback_under_30_min(self, group: str) -> int:
        """ 26_お待たせ30分以内累計対応件数 (int): 24 + 25 """
        return self.cumulative_callback_under_20_min(group) + self.callback_count_20_to_30_min(group)

    def callback_count_30_to_40_min(self, group: str) -> int:
        """ 27_お待たせ30分～40分対応件数 (int) """
        return self.data[self._get_cb_key(group, '30_40')]

    def cumulative_callback_under_40_min(self, group: str) -> int:
        """ 28_お待たせ40分以内累計対応件数 (int): 26 + 27 """
        return self.cumulative_callback_under_30_min(group) + self.callback_count_30_to_40_min(group)

    def callback_count_40_to_60_min(self, group: str) -> int:
        """ 29_お待たせ40分～60分対応件数 (int) """
        return self.data[self._get_cb_key(group, '40_60')]

    def cumulative_callback_under_60_min(self, group: str) -> int:
        """ 30_お待たせ60分以内累計対応件数 (int): 28 + 29 """
        return self.cumulative_callback_under_40_min(group) + self.callback_count_40_to_60_min(group)

    def callback_count_over_60_min(self, group: str) -> int:
        """ 31_お待たせ60分以上対応件数 (int) """
        return self.data[self._get_cb_key(group, '60over')]

    def waiting_for_callback_count_over_20min(self, group: str) -> int:
        """ お待たせ20分以上対応件数 (int) """
        return len(self.data[self._get_wfc_key(group, '20over')])
    
    def waiting_for_callback_count_over_30min(self, group: str) -> int:
        """ お待たせ30分以上対応件数 (int) """
        return len(self.data[self._get_wfc_key(group, '30over')])
    
    def waiting_for_callback_count_over_40min(self, group: str) -> int:
        """ お待たせ40分以上対応件数 (int) """
        return len(self.data[self._get_wfc_key(group, '40over')])
    
    def waiting_for_callback_count_over_60min(self, group: str) -> int:
        """ お待たせ60分以上対応件数 (int) """
        return len(self.data[self._get_wfc_key(group, '60over')])
    
    def waiting_for_callback_list_over_20min(self, group: str) -> list:
        """ お待たせ20分以上対応リスト (list) """
        return self.data[self._get_wfc_key(group, '20over')]
    
    def waiting_for_callback_list_over_30min(self, group: str) -> list:
        """ お待たせ30分以上対応リスト (list) """
        return self.data[self._get_wfc_key(group, '30over')]
    
    def waiting_for_callback_list_over_40min(self, group: str) -> list:
        """ お待たせ40分以上対応リスト (list) """
        return self.data[self._get_wfc_key(group, '40over')]
    
    def waiting_for_callback_list_over_60min(self, group: str) -> list:
        """ お待たせ60分以上対応リスト (list) """
        return self.data[self._get_wfc_key(group, '60over')]
    
    def cumulative_callback_rate_under_20_min(self, group: str) -> float:
        """ 20分以内折返し率 (float) """
        den = self.cumulative_callback_under_60_min(group) + self.callback_count_over_60_min(group)
        return self._calc_rate(self.cumulative_callback_under_20_min(group), den + self.waiting_for_callback_count_over_20min(group))
    
    def cumulative_callback_rate_under_30_min(self, group: str) -> float:
        """ 30分以内折返し率 (float) """
        den = self.cumulative_callback_under_60_min(group) + self.callback_count_over_60_min(group)
        return self._calc_rate(self.cumulative_callback_under_30_min(group), den + self.waiting_for_callback_count_over_30min(group))
    
    def cumulative_callback_rate_under_40_min(self, group: str) -> float:
        """ 40分以内折返し率 (float) """
        den = self.cumulative_callback_under_60_min(group) + self.callback_count_over_60_min(group)
        return self._calc_rate(self.cumulative_callback_under_40_min(group), den + self.waiting_for_callback_count_over_40min(group))
    
    def cumulative_callback_rate_under_60_min(self, group: str) -> float:
        """ 60分以内折返し率 (float) """
        den = self.cumulative_callback_under_60_min(group) + self.callback_count_over_60_min(group)
        return self._calc_rate(self.cumulative_callback_under_60_min(group), den + self.waiting_for_callback_count_over_60min(group))
    
    def get_all_metrics(self, group: str) -> dict:
        return {
            "総着信数": self.total_calls(group),
            "自動音声ガイダンス途中切断数": self.ivr_interruptions(group),
            "放棄呼数": self.abandoned_calls(group),
            "オペレーター呼出途中放棄数": self.abandoned_during_operator(group),
            "留守電放棄件数": self.abandoned_in_ivr(group),
            "留守電数": self.voicemails(group),
            "応答件数": self.responses(group),
            "応答率": self.response_rate(group),
            "電話問い合わせ件数": self.phone_inquiries(group),
            "直受け対応件数": self.direct_handling(group),
            "直受け率": self.direct_handling_rate(group),
            "お待たせ0分～20分対応件数": self.callback_count_0_to_20_min(group),
            "お待たせ20分以内累計対応件数": self.cumulative_callback_under_20_min(group),
            "お待たせ20分～30分対応件数": self.callback_count_20_to_30_min(group),
            "お待たせ30分以内累計対応件数": self.cumulative_callback_under_30_min(group),
            "お待たせ30分～40分対応件数": self.callback_count_30_to_40_min(group),
            "お待たせ40分以内累計対応件数": self.cumulative_callback_under_40_min(group),
            "お待たせ40分～60分対応件数": self.callback_count_40_to_60_min(group),
            "お待たせ60分以内累計対応件数": self.cumulative_callback_under_60_min(group),
            "お待たせ60分以上対応件数": self.callback_count_over_60_min(group),
            "お待たせ20分以上対応件数": self.waiting_for_callback_count_over_20min(group),
            "お待たせ30分以上対応件数": self.waiting_for_callback_count_over_30min(group),
            "お待たせ40分以上対応件数": self.waiting_for_callback_count_over_40min(group),
            "お待たせ60分以上対応件数": self.waiting_for_callback_count_over_60min(group),
            "お待たせ20分以上対応リスト": self.waiting_for_callback_list_over_20min(group),
            "お待たせ30分以上対応リスト": self.waiting_for_callback_list_over_30min(group),
            "お待たせ40分以上対応リスト": self.waiting_for_callback_list_over_40min(group),
            "お待たせ60分以上対応リスト": self.waiting_for_callback_list_over_60min(group),
            "20分以内折返し率": self.cumulative_callback_rate_under_20_min(group),
            "30分以内折返し率": self.cumulative_callback_rate_under_30_min(group),
            "40分以内折返し率": self.cumulative_callback_rate_under_40_min(group),
            "60分以内折返し率": self.cumulative_callback_rate_under_60_min(group)
        }
```
## src/excel_processors/activity_processor.py
```
from .base import BaseProcessor
import pandas as pd
import datetime
import settings
import logging

logger = logging.getLogger(__name__)


# シリアル値の定義
TOWENTY_MINUTES = settings.SERIAL_20_MINUTES
THIRTY_MINUTES = settings.SERIAL_30_MINUTES
FORTY_MINUTES = settings.SERIAL_40_MINUTES
SIXTY_MINUTES = settings.SERIAL_60_MINUTES

class ActivityProcessor(BaseProcessor):
    def process(self) -> dict:
        """
        Activityファイルのデータを指定された条件でフィルタリングおよび整形します。
        Parameters
        ----------
        None

        Returns
        -------
        dict
            活動データのフィルタリングおよび整形結果
        """
        df = self.df.copy()
        result = {}
        try:
            # '件名'列を文字列型に変換
            df['件名'] = df['件名'].astype(str)

            # 件名に「【受付】」が含まれていないもののみ残す
            df = df[~df['件名'].str.contains('【受付】', na=False)]
            logger.info(f"'【受付】'が含まれるカラムを削除しています。 Remaining rows: {df.shape[0]}")

            # 日付範囲でフィルタリング
            start_date = datetime.date.today()
            end_date = datetime.date.today()
            df = self.filtered_by_date_range(df, '登録日時 (関連) (サポート案件)', start_date, end_date)
            logger.info(f"指定された日付範囲でフィルタリングしています。 Remaining rows: {df.shape[0]}")

            # 案件番号でソートし、最も早い日時を残して重複を削除
            df = df.sort_values(by=['案件番号 (関連) (サポート案件)', '登録日時'])
            df = df.drop_duplicates(subset='案件番号 (関連) (サポート案件)', keep='first')
            logger.info(f"案件番号でソートし、最も早い日時のデータを残して重複を削除しています。 Remaining rows: {df.shape[0]}")

            # '登録日時'と'登録日時 (関連) (サポート案件)'をdatetime型に変換
            # self.df['登録日時'] = pd.to_datetime(self.df['登録日時'], unit='d', origin='1899-12-30')
            # self.df['登録日時 (関連) (サポート案件)'] = pd.to_datetime(self.df['登録日時 (関連) (サポート案件)'], unit='d', origin='1899-12-30')

            # 時間差を計算
            df['時間差'] = df['登録日時'] - df['登録日時 (関連) (サポート案件)']
            df['時間差'] = df['時間差'].fillna(0.0)
            logger.info("'時間差'を計算しています。")

            # 受付タイプが折返しor留守電のものを抽出
            _df_ss_tvs_kmn = df[(df['受付タイプ (関連) (サポート案件)'] == '折返し') | (df['受付タイプ (関連) (サポート案件)'] == '留守電')]
            _df_hhd = df[(df['受付タイプ (関連) (サポート案件)'] == 'HHD入電（折返し）') | (df['受付タイプ (関連) (サポート案件)'] == '留守電')]
            logger.info(f"受付タイプが'折返し'または'留守電'のデータを抽出しています。")

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
    
    def waiting_for_callback(self,
                             start_date: datetime.date,
                             end_date: datetime.date) -> dict:
        """
        滞留案件をクループ別、滞留時間別に集計し、案件リストを文字列として作成。
        辞書に格納して返却する。
        
        Parameters
        ----------
        start_date : datetime.date  
            集計開始日
        end_date : datetime.date
            集計終了日
        
        Returns
        -------
        dict
        """
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
    
    def group_activities_by_callback_duration(self, df: pd.DataFrame) -> tuple:
        cb_0_20 = df[(df['時間差'] <= TOWENTY_MINUTES)].shape[0]
        cb_20_30 = df[(df['時間差'] > TOWENTY_MINUTES) & (df['時間差'] <= THIRTY_MINUTES)].shape[0]
        cb_30_40 = df[(df['時間差'] > THIRTY_MINUTES) & (df['時間差'] <= FORTY_MINUTES)].shape[0]
        cb_40_60 = df[(df['時間差'] > FORTY_MINUTES) & (df['時間差'] <= SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_60over = df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_not_include =df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')].shape[0]

        return cb_0_20, cb_20_30, cb_30_40, cb_40_60, cb_60over, cb_not_include
    
    
    
    def convert_to_pending_num(self, df: pd.DataFrame) -> tuple:
        wfc_over_20 = df[df['お待たせ時間'] >= TOWENTY_MINUTES]
        wfc_over30 = df[df['お待たせ時間'] >= THIRTY_MINUTES]
        wfc_over40 = df[df['お待たせ時間'] >= FORTY_MINUTES]
        wfc_over60 = df[df['お待たせ時間'] >= SIXTY_MINUTES]

        return wfc_over_20, wfc_over30, wfc_over40, wfc_over60
    
    def callback_classification_by_group(self, df: pd.DataFrame) -> tuple:
        df_cb_0_20 = df[(df['時間差'] <= TOWENTY_MINUTES)]
        df_cb_20_30 = df[(df['時間差'] > TOWENTY_MINUTES) & (df['時間差'] <= THIRTY_MINUTES)]
        df_cb_30_40 = df[(df['時間差'] > THIRTY_MINUTES) & (df['時間差'] <= FORTY_MINUTES)]
        df_cb_40_60 = df[(df['時間差'] > FORTY_MINUTES) & (df['時間差'] <= SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')]
        df_cb_60over = df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')]
        df_cb_not_include =df[(df['時間差'] > SIXTY_MINUTES) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')]

        return df_cb_0_20, df_cb_20_30, df_cb_30_40, df_cb_40_60, df_cb_60over, df_cb_not_include
    
    def create_wfc_list(self, df: pd.DataFrame) -> list:
        return list(df.loc[:, '案件番号 (関連) (サポート案件)'])

    @staticmethod
    def datetime_to_serial(dt: datetime.datetime, base_date=datetime.datetime(1899, 12, 30)) -> float:
        """
        datetimeオブジェクトをシリアル値に変換する。

        :param dt: 変換するdatetimeオブジェクト
        :param base_date: シリアル値の基準日（デフォルトは1899年12月30日）
        :return: シリアル値
        """
        return (dt - base_date).total_seconds() / (24 * 60 * 60)
```
## src/excel_processors/base.py
```
import pandas as pd
import datetime
import logging
import settings

logger = logging.getLogger(__name__)

class BaseProcessor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df = pd.DataFrame()

    def load_data(self) -> None:
        """
        Excelファイルを読み込み、DataFrameに格納します。
        """
        try:
            self.df = pd.read_excel(self.file_path)
            logger.info(f"Loaded {self.file_path} with {self.df.shape[0]} rows.")
        except Exception as e:
            logger.error(f"Failed to load {self.file_path}: {e}")
            raise

    def save_data(self, output_file: str) -> None:
        """
        DataFrameをExcelファイルとして保存します。
        """
        try:
            self.df.to_excel(output_file, index=False)
            logger.info(f"Data saved to {output_file}.")
        except Exception as e:
            logger.error(f"Failed to save data to {output_file}: {e}")
            raise

    def process(self) -> None:
        """
        データ処理のメインメソッド。各サブクラスでオーバーライドします。
        """
        raise NotImplementedError("Subclasses should implement this method.")

    def filtered_by_date_range(self, df: pd.DataFrame,
                               date_column: str,
                               start_date: datetime.date,
                               end_date: datetime.date) -> pd.DataFrame:
        """
        start_dateからend_dateの範囲のデータを抽出

        Parameters
        ----------
        df : pd.DataFrame
            フィルタリング対象のDataFrame
        date_column : str
            フィルタリング対象の日付カラム名
        start_date : datetime.date
            抽出開始日
        end_date : datetime.date
            抽出終了日
        
        Returns
        -------
        pd.DataFrame
        """
        start_date_serial = self.datetime_to_serial(datetime.datetime.combine(start_date, datetime.time.min))
        end_date_serial = self.datetime_to_serial(datetime.datetime.combine(end_date + datetime.timedelta(days=1), datetime.time.min))
        
        filtered_df = df[
            (df[date_column] >= start_date_serial) &
            (df[date_column] < end_date_serial)
        ].reset_index(drop=True)
        
        logger.debug(f"Filtered DataFrame from {start_date} to {end_date}: {filtered_df.shape[0]} rows")
        return filtered_df
    
    def current_time_to_serial(self, base_date=datetime.datetime(1899, 12, 30)) -> float:
        """
        現在日時をシリアル値に変換する。

        Parameters
        ----------
        base_date : datetime.datetime
            シリアル値の基準日（デフォルトは1899年12月30日）
        
        Returns
        -------
        float
        """
        current_time = datetime.datetime.now()
        serial_value = (current_time - base_date).total_seconds() / (24 * 60 * 60)
        return serial_value
    
    @staticmethod
    def datetime_to_serial(dt: datetime.datetime, base_date=datetime.datetime(1899, 12, 30)) -> float:
        """
        datetimeオブジェクトをシリアル値に変換する。

        Parameters
        ----------
        dt : datetime.datetime
            変換するdatetimeオブジェクト
        base_date : datetime.datetime
            シリアル値の基準日（デフォルトは1899年12月30日）
        
        Returns
        -------
        float
        """
        dt = datetime.datetime(dt.year, dt.month, dt.day)
        return (dt - base_date).total_seconds() / (24 * 60 * 60)
    
    @staticmethod
    def serial_to_datetime(serial, base_date=datetime.datetime(1899, 12, 30)) -> datetime.datetime:
        """
        シリアル値をdatetimeオブジェクトに変換する。

        Parameters
        ----------
        serial : float
            変換するシリアル値
        base_date : datetime.datetime
            シリアル値の基準日（デフォルトは1899年12月30日）
        
        Returns
        -------
        datetime.datetime
        """
        return base_date + datetime.timedelta(days=serial)
```
## src/excel_processors/close_processor.py
```
```
## src/excel_processors/excel_sync.py
```
import logging
import openpyxl
import os
import pythoncom  # COM初期化に必要
import time
from typing import List
import win32com.client

import settings

# ロガーの設定
logger = logging.getLogger(__name__)

class SynchronizedExcelProcessor:
    def __init__(self, file_paths: List[str],
                 max_retries: int = settings.SYNC_MAX_RETRIES,
                 retry_delay: int = settings.SYNC_RETRY_DELAY,
                 refresh_interval: int = settings.REFRESH_INTERVAL
                 ) -> None:
        """
        Excelファイルの同期処理を管理するクラス。

        Parameters
        ----------
        file_paths : List[str]
            同期するExcelファイルのパスのリスト。
        max_retries : int, optional
            同期失敗時の最大リトライ回数（デフォルトは設定ファイルから）。
        retry_delay : int, optional
            リトライ間の待機時間（秒、デフォルトは設定ファイルから）。
        refresh_interval : int, optional
            CalculationState を確認する際の待機時間（秒、デフォルトは設定ファイルから）。。
        """
        self.file_paths = file_paths
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.refresh_interval = refresh_interval

    def process_file(self, file_path, stop_event) -> dict:
        """
        このクラスのメインのメソッド。
        これを実行するだけ。

        Parameters
        ----------
        file_path : str
            処理するExcelファイルのパス。
        stop_event : threading.Event
            処理を停止するためのイベント。
        
        Returns
        -------
        dict
            ファイルの処理結果。
        """
        self._sync_file(file_path, stop_event)

        if settings.ACTIVITY_FILE in file_path:
            from src.excel_processors.activity_processor import ActivityProcessor
            activity = ActivityProcessor(file_path)
            activity.load_data()
            result = activity.process()
            return result
        elif settings.CLOSE_FILE in file_path:
            result = {}
            return result
        elif settings.SUPPORT_FILE in file_path:
            from src.excel_processors.support_processor import SupportProcessor
            support = SupportProcessor(file_path)
            support.load_data()
            result = support.process()
            return result
        else:
            logger.error(f"ファイル名がPathに含まれていません。{file_path}")

    
    def _sync_file(self, file_path, stop_event) -> None:
        """
        個別のExcelファイルを処理します。

        Parameters
        ----------
        file_path : str
            処理するExcelファイルのパス。
        """
        logger.debug(f"{file_path}の処理を開始します。")
        if stop_event.is_set():
            logger.info(f"{file_path}の処理が停止されました。")
            return
        if not os.path.exists(file_path):
            logger.warning(f"ファイルが存在しません。: {file_path}")
            return
        
        try:
            # COMライブラリを初期化
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            logger.debug("COMライブラリを初期化しました。")

            # Excelアプリケーション作成
            excel = self._create_excel_app()
            logger.debug('Excelアプリケーションを作成しました。')

            logger.info(f"{file_path}の同期を開始します。")
            retries = 0

            while retries < self.max_retries and not stop_event.is_set():
                try:
                    workbook = excel.Workbooks.Open(file_path)
                    logger.debug("ワークブックを開きました。")
                    workbook.RefreshAll()
                    logger.debug("ワークブックを更新しています。")
                    time.sleep(self.refresh_interval)
                    workbook.Save()
                    logger.debug("ワークブックを保存しました。")
                    workbook.Close()
                    logger.info(f"{file_path}の同期が完了しました。")
                    break
                except Exception as e:
                    retries += 1
                    logger.info(f"{file_path}の同期中にエラーが発生しました。（{retries}回目）: {e}")
                    if retries >= self.max_retries:
                        logger.error(f"{file_path}の同期に失敗しました。最大リトライ回数に達しました。: {e}")
                        # リトライ回数を超えたのでwhileループを抜けます。
                        break

                    else:
                        logger.info(f"{file_path}の同期を再試行します。")
                        time.sleep(self.retry_delay)
                    
                    # Excelのクローズ処理を行います。
                    self._close_app(excel)

                    # 次の処理に備えてExcelアプリケーションを作成します。
                    excel = self._create_excel_app()
                    logger.info('Excelアプリケーションを作成しました。')

        except Exception as e:
            logger.error(f"{file_path}の同期処理中に予期しないエラーが発生しました。{e}")
        finally:
            # COMライブラリを終了
            pythoncom.CoUninitialize()
            logger.debug("COMライブラリを終了しました。")
        
        
    
    def _create_excel_app(self) -> win32com.client.CDispatch:
        """
        Excelアプリケーションを起動し、設定を行うヘルパーメソッド。

        Returns
        -------
        excel_app : COMObject
            起動したExcelアプリケーションオブジェクト。
        """
        logger.info("Excelアプリケーションを起動します。")
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        return excel_app
    
    @staticmethod
    def check_and_close(file_names: List[str]) -> None:
        """
        指定されたExcelファイルが開かれているかを確認し、開かれている場合はExcelアプリケーションを強制終了します。
        
        Parameters
        ----------
        file_names : List[str]
            確認するExcelファイルのパスのリスト。
        """
        for file_name in file_names:
            try:
                with open(file_name, "r+b"):
                    logger.debug(f"{file_name}は閉じられています。")
            except PermissionError:
                logger.info(f"{file_name}は開かれています。")
                try:
                    os.system("taskkill /F /IM excel.exe")  # Excelプロセスを強制終了
                    logger.info("Excelアプリケーションを強制終了しました。")
                except Exception as e:
                    logger.error(f"Excelを強制終了する際にエラーが発生しました: {e}")
            except Exception as e:
                logger.error(f"{file_name}の確認中にエラーが発生しました。Excelアプリケーションを強制終了します。: {e}")
                try:
                    os.system("taskkill /F /IM excel.exe")  # Excelプロセスを強制終了
                    logger.info("Excelアプリケーションを強制終了しました。")
                except Exception as e:
                    logger.error(f"Excelを強制終了する際にエラーが発生しました: {e}")

    def _close_app(self, excel) -> None:
        """
        Excelアプリケーションを終了するヘルパー関数。
        """
        try:
            excel.Quit()
            logger.info("Excelアプリケーションを終了しました。")
        except Exception as quit_e:
            logger.warning(f"Excelの終了中にエラーが発生しました。: {quit_e}")
    
```
## src/excel_processors/support_processor.py
```
from .base import BaseProcessor
import pandas as pd
import datetime
import settings
import logging

logger = logging.getLogger(__name__)

class SupportProcessor(BaseProcessor):
    def process(self) -> dict:
        """
        Supportファイルのデータを指定された条件でフィルタリングおよび整形します。

        Returns
        -------
        dict
            Supportファイルのデータを指定された条件でフィルタリングおよび整形した結果。
        """
        result = {}

        try:
            base_df = self.df.fillna('') # これを入れないとdf['かんたん！保守区分'] == ''が上手く判定されない
            
            # 日付範囲でフィルタリング
            start_date = datetime.date.today()
            end_date = datetime.date.today()
            
            # start_dateからend_dateの範囲のデータを抽出
            base_df = self.filtered_by_date_range(base_df, '登録日時', start_date, end_date)

            # 直受けの検索条件
            df = base_df.copy()
            df = df[
                ((df['受付タイプ'] == '直受け') | (df['受付タイプ'] == 'HHD入電（直受け）')) & 
                (~df['顛末コード'].isin(['折返し不要・ｷｬﾝｾﾙ', 'ﾒｰﾙ・FAX回答（送信）', 'SRB投稿（要望）', 'ﾒｰﾙ・FAX文書（受信）'])) &
                ((df['かんたん！保守区分'] == '会員') | (df['かんたん！保守区分'] == '')) &
                (df['回答タイプ'] != '2次T転送')
                ]
            
            # グループ別に分割
            result['direct_ss'] = df[df['サポート区分'] == 'SS'].shape[0]
            result['direct_tvs'] = df[df['サポート区分'] == 'TVS'].shape[0]
            result['direct_kmn'] = df[df['サポート区分'] == '顧問先'].shape[0]
            result['direct_hhd'] = df[df['サポート区分'] == 'HHD'].shape[0]

            # 留守電数の検索条件
            df = base_df.copy()
            df = df[
                (df['受付タイプ'] == '留守電') & 
                (~df['顛末コード'].isin(['折返し不要・ｷｬﾝｾﾙ', 'ﾒｰﾙ・FAX回答（送信）', 'SRB投稿（要望）', 'ﾒｰﾙ・FAX文書（受信）']))
                ]
            
            # グループ別に分割
            result['ivr_ss'] = df[df['サポート区分'] == 'SS'].shape[0]
            result['ivr_tvs'] = df[df['サポート区分'] == 'TVS'].shape[0]
            result['ivr_kmn'] = df[df['サポート区分'] == '顧問先'].shape[0]
            result['ivr_hhd'] = df[df['サポート区分'] == 'HHD'].shape[0]

            logger.debug(f"サポート案件処理結果: {result}")
            return result

        except Exception as e:
            logger.error(f"直受け件数、留守電数の算定中にエラーが発生しました。: {e}")
            raise
```
## src/controller.py
```
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import threading

from src.excel_processors.excel_sync import SynchronizedExcelProcessor
from src.calculator.kpi_calculator import KpiCalculator
from src.scraper import Scraper
import settings


logger = logging.getLogger(__name__)

def collect_data() -> dict:
    """
    Excelファイルの処理とスクレイピング処理を同期的に実行する。
    
    Returns
    -------
    dict
        処理結果を格納した辞書型オブジェクト
    """
    stop_event = threading.Event()

    excel_processor = SynchronizedExcelProcessor(
        file_paths=settings.EXCEL_FILES,
        max_retries=settings.SYNC_MAX_RETRIES,
        retry_delay=settings.SYNC_RETRY_DELAY,
        refresh_interval=settings.REFRESH_INTERVAL
    )

    # scraper処理をここに書く
    scraper = Scraper()

    # Excelが開いているかを確認して開いている場合はExcelを強制終了する。
    SynchronizedExcelProcessor.check_and_close(settings.EXCEL_FILES)

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        results = {}

        # Excelファイルの処理をタスクとして追加
        logger.debug("Excelファイルの処理をタスクとして追加しています。")
        for file_path in excel_processor.file_paths:
            futures.append(
                executor.submit(excel_processor.process_file, file_path, stop_event)
            )
        
        # scraping処理をタスクとして追加
        logger.debug("スクレイピングの処理をタスクとして追加しています。")
        futures.append(
            executor.submit(scraper.scrape_ctstage_report, settings.TEMPLATES, stop_event)
        )

        try:
            for future in as_completed(futures):
                result = future.result()
                if isinstance(result, dict):
                    results.update(result)
                else:
                    logger.error(f"処理結果が辞書型ではありません。: {result}")
            return results
        except KeyboardInterrupt:
            logger.info("停止信号を受け取りました。全てのタスクを停止します。")
            stop_event.set()
        
        except Exception as e:
            logger.error(f"エラーが発生しました。: {e}")
            stop_event.set()

def calculate_group_kpis_for_all_groups(data: dict) -> dict:
    """
    KPIを計算する。
    """
    kpi_calculator = KpiCalculator(data)
    results = {}
    results['SS'] = kpi_calculator.get_all_metrics('SS')
    results['TVS'] = kpi_calculator.get_all_metrics('TVS')
    results['KMN'] = kpi_calculator.get_all_metrics('KMN')
    results['HHD'] = kpi_calculator.get_all_metrics('HHD')

    return results

def calculate_operator_kpis(results: dict) -> dict:
    """
    オペレーター別のKPIを計算する。
    """
    return results

def orchestrate_workflow():
    results = collect_data()
    kpi_results = calculate_group_kpis_for_all_groups(results)
    for k, v in kpi_results.items():
        for k2, v2 in v.items():
            print(f"{k} {k2}: {v2}")
```
## src/scraper.py
```
from bs4 import BeautifulSoup
import datetime
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
from typing import List
import pandas as pd

import settings


logger = logging.getLogger(__name__)

class Base:
    def __init__(self,
                 url: str = settings.REPORTER_URL,
                 id: str = settings.REPORTER_ID) -> None:
        self.url = url
        self.id = id
        self.df = pd.DataFrame()
        self.driver = None

    def create_driver(self) -> None:
        try:
            if self.driver:
                self.close_driver()
            
            logger.info("driverを作成しています。")
            options = Options()

            # ブラウザを表示させるか
            if settings.HEADLESS_MODE:
                options.add_argument('--headless')
            
            # コマンドプロンプトのログを表示させない。
            options.add_argument('--disable-logging')
            options.add_argument('--disable-extensions')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--log-level=3')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])

            # driverの作成
            self.driver = webdriver.Chrome(options=options)
            self.driver.implicitly_wait(5)
        except Exception as e:
            logger.error(f"driverの作成に失敗しました。: {e}")
    
    def login(self) -> None:
        """レポーターにログイン"""
        try:
            logger.info("ログインを試みています。")
            self.driver.get(self.url)
            logger.debug(f"URL {self.url}にアクセスしました。")

            # ID入力
            self.driver.find_element(By.ID, 'logon-operator-id').send_keys(self.id)
            logger.debug("IDを入力しました。")

            # ログインボタンをクリック
            self.driver.find_element(By.ID, 'logon-btn').click()
            logger.debug("ログインボタンをクリックしました。")

        except Exception as e:
            logger.error(f"ログインに失敗しました。: {e}")
            raise

    def call_template(self, template: str) -> None:
        """
        テンプレート呼び出し

        Parameters
        ----------
        template : List[str]
            表示したいテンプレート。
        """
        try:
            # テンプレート呼び出し
            logger.info(f"テンプレートを呼び出しています。: [{template}]")
            self.driver.find_element(By.ID, 'template-title-span').click()
            el2 = self.driver.find_element(By.ID, 'template-download-select')
            s2 = Select(el2)
            s2.select_by_value(template)
            self.driver.find_element(By.ID, 'template-creation-btn').click()
        except Exception as e:
            logger.error(f"テンプレートの呼び出しに失敗しました。: {e}")

    def create_report(self, element_id: str = "0"):
        """
        レポート作成
        
        Parameters
        ----------
        element_id : str
            選択するタブによって、"0" or "1"
        """
        try:
            self.driver.find_element(By.ID, f'panel-td-create-report-{element_id}').click()
        except Exception as e:
            logger.error(f"レポートの作成に失敗しました。: {e}")

    def select_tabs(self, tab_element_id: str = "1"):
        """
        レポートのタブ切り替え

        Parameters
        ----------
        tab_element_id : str
            選択するタブによって、"1" or "2"
        """
        self.driver.find_element(By.ID, f'normal-title{tab_element_id}').click()
    
    def create_dateframe(self, list_name: str) -> pd.DataFrame:
        """
        
        'normal-list1-dummy-0'
        'normal-list2-dummy-1'
        """
        try:
            logger.debug("ページソースをエンコードしています。")
            html = self.driver.page_source.encode('utf-8')
        except Exception as e:
            logger.error(f"ページソースをUTF-8でエンコード中にエラーが発生しました。: {e}")
        try:
            logger.debug("HTMLをパースしています。")
            soup = BeautifulSoup(html, 'lxml')
        except Exception as e:
            logger.error(f"HTMLのパース中にエラーが発生しました。: {e}")
        
        data_table = []

        # headerのリストを作成
        header_table = soup.find(id=f'{list_name}-table-head-table')
        xmp = header_table.thead.tr.find_all('xmp')
        header_list = [i.string for i in xmp]
        data_table.append(header_list)

        # bodyのリストを作成
        body_table = soup.find(id=f'{list_name}-table-body-table')
        tr = body_table.tbody.find_all('tr')
        for td in tr:
            xmp = td.find_all('xmp')
            row = [i.string for i in xmp]
            data_table.append(row)
        
        # テーブルをDataFrameに変換
        df = pd.DataFrame(data_table[1:], columns=data_table[0])
        df.set_index(df.columns[0], inplace=True)
        return df

    def close_driver(self):
        """driverを閉じる"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("driverを正常に閉じました。")
            except Exception as e:
                logger.error("driverの閉鎖に失敗しました。: {e}")
            finally:
                self.driver = None


class Scraper(Base):
    def scrape_ctstage_report(self, templates: List[str], stop_event):
        results = {}
        if stop_event.is_set():
            logger.info(f"スクレイピング処理が停止されました。")
        try:
            self.create_driver()
            self.login()
            for template in templates:
                retries = 0
                while retries < settings.REPORTER_MAX_RETRIES and not stop_event.is_set():
                    try:
                        self.call_template(template)
                        self.create_report(element_id="0")
                        time.sleep(1)

                        df1 = self.create_dateframe('normal-list1-dummy-0')

                        self.select_tabs(tab_element_id="2")
                        self.create_report(element_id="1")
                        time.sleep(1)

                        df2 = self.create_dateframe('normal-list2-dummy-1')

                        # 必要なデータを辞書に保存
                        template_result = {
                            "total_calls": int(df1.iloc[0, 0]), # 総着信数
                            "IVR_interruptions_before_response": int(df1.iloc[0, 1]), # IVR応答前放棄呼数
                            "ivr_interruptions": int(df1.iloc[0, 2]), # IVR切断数
                            "time_out": int(df2.iloc[0, 0]), # タイムアウト数
                            "abandoned_during_operator": int(df2.iloc[0, 1]) # ACD放棄呼数
                        }
                        results[template] = template_result
                        break

                    except Exception as e:
                        retries += 1
                        logger.error(f"{template}のスクレイピング中にエラーが発生しました({retries}回目)。: {e}")
                        self.close_driver()
                        self.create_driver()
                        self.login()
            return results
                        
        except Exception as e:
            logger.error(f"スクレイピング中に予期しないエラーが発生しました。")
            return results
        finally:
            self.close_driver()
            logger.debug("Web Driverを終了しました。")
```
## src/views.py
```
```
