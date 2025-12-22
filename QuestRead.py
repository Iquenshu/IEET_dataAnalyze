import pandas as pd
from Accessdb import AccessHelper  # 匯入自訂的 Access 資料庫工具類

# 1. 設定 Excel 檔案路徑與資料表、欄位
excel_path = r'input_files\問券\離校問券資料\1141畢業生離校問卷題目(電機).xlsx'
table_name = 'Questionnaire'
columns = ['學年', '學期', '對象', '題型', '中文指標', '欄位序號']

# 2. 讀取 Excel 並只保留需要的欄位
df = pd.read_excel(excel_path)
df = df[columns]

# 3. 連接 Access（自動使用全專案共用的資料庫路徑）
db = AccessHelper()
repeat_count = 0
import_count = 0

# 4. 寫入 Access（避免重複，包含題型）
for _, row in df.iterrows():
    where = "學年=? AND 學期=? AND 對象=? AND 欄位序號=? AND 題型=?"
    params = (row['學年'], row['學期'], row['對象'], row['欄位序號'], row['題型'])
    if db.is_duplicate(table_name, where, params):
        repeat_count += 1
        continue
    db.insert_row(table_name, columns, tuple(row))
    import_count += 1

db.close()
print(f"匯入完成！重複資料：{repeat_count} 筆，匯入新資料：{import_count} 筆")