import pandas as pd
from Accessdb import AccessHelper
import os
from datetime import datetime

# 1. 連接 Access 資料庫
db = AccessHelper()

# 2. 讀取 STscoreAnalyze 資料表，只取 total 的資料
sql = "SELECT 學年度, 學期, 課號, 課程名稱, 平均分數, 學生總數 FROM STscoreAnalyze WHERE 分數區間='total'"
df = pd.read_sql(sql, db.conn)

db.close()

# 3. 依學年度分組，每個學年度一個分頁
output_dir = 'output_files'
os.makedirs(output_dir, exist_ok=True)
today_str = datetime.today().strftime('%Y%m%d')
output_path = os.path.join(output_dir, f'STscoreAVG_{today_str}.xlsx')

with pd.ExcelWriter(output_path) as writer:
    for year in sorted(df['學年度'].unique()):
        df_year = df[df['學年度'] == year]
        # 只保留指定欄位
        df_year = df_year[['學年度', '學期', '課號', '課程名稱', '平均分數', '學生總數']]
        # 分頁名稱：學年度
        sheet_name = f"{year}學年度"
        df_year.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"匯出完成！已產生 {output_path}")