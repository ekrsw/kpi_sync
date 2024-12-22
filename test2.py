from src.controller import collect_data, collect_and_calculate_operator_kpis
import logging
import time

logging.basicConfig(level=logging.INFO, format='%(asctime)s:%(levelname)s:%(name)s:%(message)s',)
logger = logging.getLogger(__name__)

start = time.time()
results = collect_data()
df = collect_and_calculate_operator_kpis(results['TEMPLATE_OP'])
print(df)

end = time.time()
time_diff = end - start
print(f"処理が正常に終了しました。（処理時間: {time_diff} 秒）")