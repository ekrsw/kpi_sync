import pandas as pd

import logging

logger = logging.getLogger(__name__)


class OperatorCalculator:
    def __init__(self, data: pd.DataFrame) -> None:
        self.data = data

    def calculate(self) -> pd.DataFrame:
        pass

    def _str_to_serial(self, str_date: str) -> pd.Timestamp:
        pass