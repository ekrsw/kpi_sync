import pandas as pd
import settings
import datetime
from src.processors.base import BaseProcessor

import logging

logger = logging.getLogger(__name__)


class ShiftProcessor:
    def __init__(self,
                 df_operators: pd.DataFrame,
                 file_path: str = settings.SHIFT_SCHEDULE):
        self.sweet_to_name = df_operators.set_index('Sweet')['氏名'].to_dict()
        self.ctstage_to_name = df_operators.set_index('CTStage')['氏名'].to_dict()
        self.df = pd.read_csv(file_path, skiprows=2, header=1, index_col=1, quotechar='"', encoding='shift_jis')

    def process(self) -> dict:
        date_str = datetime.date.today().strftime("%d")
        df = self.df.iloc[:, :-1]
        # "組織名"、"従業員ID"、"種別" の列を削除
        df = df.drop(columns=["組織名", "従業員ID", "種別"])
        df_shift = df[[date_str]]
        df_shift.columns = ["シフト"]
        
        return df_shift
