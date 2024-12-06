import logging

logger = logging.getLogger(__name__)

class KpiCalculator:
    def __init__(self, data: dict):
        self.data = data

    def total_calls(self, group: str) -> int:
        """ 11_総着信数 (int): reporter_着信数 """
        template = self._select_template(group)
        return self.data[template]['total_calls']
    
    def ivr_interruptions(self, group: str) -> int:
        """ 12_自動音声ガイダンス途中切断数 (int): reporter_IVR応答前放棄呼数 + reporter_IVR切断数 """
        template = self._select_template(group)
        return self.data[template]['IVR_interruptions_before_response'] + self.data[template]['ivr_interruptions']
    
    def abandoned_during_operator(self, group: str) -> int:
        """ 14_オペレーター呼出途中放棄数 (int): reporter_ACD放棄呼数 """
        template = self._select_template(group)
        return self.data[template]['abandoned_during_operator']
    
    def voicemails(self, group: str) -> int:
        """ 16_留守電数 (int): S_留守電 """
        if group == 'SS':
            return self.data['ivr_ss']
        elif group == 'TVS':
            return self.data['ivr_tvs']
        elif group == 'KMN':
            return self.data['ivr_kmn']
        elif group == 'HHD':
            return self.data['ivr_hhd']
        else:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
    
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
        if group == 'SS':
            return self.data['direct_ss']
        elif group == 'TVS':
            return self.data['direct_tvs']
        elif group == 'KMN':
            return self.data['direct_kmn']
        elif group == 'HHD':
            return self.data['direct_hhd']
        else:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
    
    def direct_handling_rate(self, group: str) -> float:
        """ 直受率: 21 /18 """
        return self._calc_rate(self.direct_handling(group), self.phone_inquiries(group))
    
    def callback_count_0_to_20_min(self, group: str) -> int:
        """ 23_お待たせ0分～20分対応件数 (int) """
        if group == 'SS':
            return self.data['cb_0_20_ss']
        elif group == 'TVS':
            return self.data['cb_0_20_tvs']
        elif group == 'KMN':
            return self.data['cb_0_20_kmn']
        elif group == 'HHD':
            return self.data['cb_0_20_hhd']
        else:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")
    
    def cumulative_callback_under_20_min(self, group: str) -> int:
        """ 24_お待たせ20分以内累計対応件数 (int): 21 + 23 """
        return self.direct_handling(group) + self.callback_count_0_to_20_min(group)
    
    def _select_template(self, group: str) -> dict:
        if group == 'SS':
            return 'TEMPLATE_SS'
        elif group == 'TVS':
            return 'TEMPLATE_TVS'
        elif group == 'KMN':
            return 'TEMPLATE_KMN'
        elif group == 'HHD':
            return 'TEMPLATE_HHD'
        else:
            logger.error(f"グループが存在しません。: {group}")
            raise ValueError(f"グループが存在しません。: {group}")

    def _calc_rate(self,a: int, b: int, wfc: int = 0):
        if (b + wfc) != 0:
            return a / (b + wfc)
        else:
            return 0.0