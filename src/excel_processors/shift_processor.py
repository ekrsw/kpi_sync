import pandas as pd
import settings
import datetime
from src.excel_processors.base import BaseProcessor

import logging

logger = logging.getLogger(__name__)


class ShiftProcessor:
    def __init__(self, file_path: str, df_operators: pd.DataFrame):
        self.sweet_to_name = df_operators.set_index('Sweet')['氏名'].to_dict()
        self.ctstage_to_name = df_operators.set_index('CTStage')['氏名'].to_dict()
        date_str = datetime.date.today().strftime("%d")
        df = pd.read_csv(file_path, skiprows=2, header=1, index_col=1, quotechar='"', encoding='shift_jis')
        df = df.iloc[:, :-1]
        # "組織名"、"従業員ID"、"種別" の列を削除
        df = df.drop(columns=["組織名", "従業員ID", "種別"])
        df_shift = df[[date_str]]
        df_shift.columns = ["シフト"]
        df_shift.index = df_shift.index.map(self._replace_sweetname_to_name)

    def process(self):
        pass
    
    def _replace_sweetname_to_name(self, name):
        if name in self.sweet_to_name:
            return self.sweet_to_name[name]
        else:
            logger.warning(f"'{name}'がOperaotrsファイルのSweetカラムに存在しません。")
            return name

