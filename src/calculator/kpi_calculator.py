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