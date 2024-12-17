import pandas as pd

import logging

logger = logging.getLogger(__name__)


class OperatorCalculator:
    def __init__(self, df_ctstage, df_close, df_shift) -> None:
        self.df_ctstage = df_ctstage
        self.df_close = df_close
        self.df_shift = df_shift

    def calculate(self) -> pd.DataFrame:
        pass

    