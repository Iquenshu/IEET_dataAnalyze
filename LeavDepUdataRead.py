import pandas as pd
import os
from Accessdb import AccessHelper

# 離系問卷資料讀取程式（大學部）

# 1. 設定檔案與資料表
data_path = r'input_files\問券\離系問券資料\data_大學部問券0805.xlsx' # 你也可以改成 .xlsx
table_name = 'LeavDepUdata'

# 2. 欄位清單（抓取與資料表一致的欄位，排除自動編號ids）
columns = [
    'sqnum', 'sem', 'stname', 'uid', 'stemail', 'stphone', 'career', 'advisor',
    'A11','A12','A13','A14','A15','A21','A22','A23','A24','A25','A26','A27','A28','A29','A210','A211',
    'A31','A32','A33','A34','A35','A36','update_time'
]

# 3. 依副檔名自動選擇讀取方式
ext = os.path.splitext(data_path)[1].lower()
if ext == '.csv':
    df = pd.read_csv(data_path, encoding='utf-8', dtype=str)
elif ext in ['.xls', '.xlsx']:
    df = pd.read_excel(data_path, dtype=str)
else:
    raise ValueError('不支援的檔案格式')

df = df[columns]

# 4. 數字欄位轉型
for col in ['A11','A12','A13','A14','A15','A21','A22','A23','A24','A25','A26','A27','A28','A29','A210','A211']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 5. 寫入Access，比對uid、sem和sqnum避免資料重複
db = AccessHelper()
repeat_count = 0
import_count = 0

for _, row in df.iterrows():
    where = "uid=? AND sem=? AND sqnum=?"
    params = (row['uid'], row['sem'], row['sqnum'])
    if db.is_duplicate(table_name, where, params):
        repeat_count += 1
        continue
    db.insert_row(table_name, columns, tuple(row[col] for col in columns))
    import_count += 1

db.close()
print(f"匯入完成！重複資料：{repeat_count} 筆，匯入新資料：{import_count} 筆")