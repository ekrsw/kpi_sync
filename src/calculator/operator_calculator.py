import pandas as pd

import logging

logger = logging.getLogger(__name__)


class OperatorCalculator:
    def __init__(self, df_operators, df_ctstage, df_close, df_shift) -> None:
        self.df_operators = df_operators

        self.sweet_to_name = df_operators.dropna(subset=['Sweet']).set_index('Sweet')['氏名'].to_dict()
        self.ctstage_to_name = df_operators.set_index('CTStage')['氏名'].to_dict()

        df_ctstage.index = df_ctstage.index.map(self._replace_ctstage_to_name)
        df_shift.index = df_shift.index.map(self._replace_sweetname_to_name)
        self.df_ctstage = df_ctstage.map(self._time_to_days)
        self.df_shift = df_shift
        self.df_close = df_close

    def calculate(self) -> pd.DataFrame:
        """
        オペレーター別のACW, ATT, CPHを計算する。
        """
        active_operators = self.df_operators[self.df_operators['active'] == 1]['氏名']
        self.df_ctstage = self.df_ctstage[self.df_ctstage.index.isin(active_operators)]
        self.df_close = self.df_close[self.df_close.index.isin(active_operators)]
        self.df_shift = self.df_shift[self.df_shift.index.isin(active_operators)]
       
        print(self.sweet_to_name)
    
    def _float_to_hms(value: float) -> str:
        '''1日を1としたfloat型を'hh:mm:ss'形式の文字列に変換'''

        # 1日が1なので、24を掛けて時間単位に変換
        hours = value * 24

        # 時間の整数部分
        h = int(hours)

        # 残りの部分を分単位に変換
        minutes = (hours - h) * 60

        # 分の整数部分
        m = int(minutes)

        # 残りの部分を秒単位に変換
        seconds = (minutes - m) * 60

        # 秒の整数部分
        s = int(seconds)

        return f"{h:02}:{m:02}:{s:02}"
    
    def _time_to_days(self, time_str: str) -> float:
        """hh:mm:ss 形式の時間を1日を1としたときの時間に変換する関数
            Args:
                time_str(str): hh:mm:ss 形式の時間
            
            return:
                float: 1日を1としたときの時間"""
        # t = dt.datetime.strptime(time_str, "%H:%M:%S")
        split_time = time_str.split(':')
        t = (float(split_time[0]) + float(split_time[1]) / 60 + float(split_time[2]) / 3600) / 24

        return t
    
    def _replace_sweetname_to_name(self, name):
        if name in self.sweet_to_name:
            return self.sweet_to_name[name]
        else:
            logger.warning(f"'{name}'がOperaotrsファイルのSweetカラムに存在しません。")
            return name

    def _replace_ctstage_to_name(self, name):
        if name in self.ctstage_to_name:
            return self.ctstage_to_name[name]
        else:
            logger.warning(f"'{name}'がOperatorsファイルのCTStageカラムに存在しません。")
            return name