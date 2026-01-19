import pandas as pd
import os
from Accessdb import AccessHelper
import time

# 學生成績讀取程式

# 1. 設定檔案與資料表
data_path = r'input_files\學生成績\電機系100-108學年度大學部碩士班開課學生成績(所有修課學生)1140829test.xlsx'
table_name = 'STscore'

# 2. 欄位清單
columns = [
    '學年度', '學期', '開課系所代碼', '開課系所', '課號', '課程名稱', '必選修',
    '學號', '姓名', '學分數', '成績', '等第成績'
]

print(f"正在讀取 Excel 檔案... (請稍候)")
start_time = time.time()

# 讀取 Excel
ext = os.path.splitext(data_path)[1].lower()
if ext == '.csv':
    df = pd.read_csv(data_path, encoding='utf-8', dtype=str)
else:
    df = pd.read_excel(data_path, dtype=str)

df = df[columns]
# 數字轉型與過濾
df['學分數'] = pd.to_numeric(df['學分數'], errors='coerce')
df['成績'] = pd.to_numeric(df['成績'], errors='coerce')
df = df[df['學號'].fillna('').str.strip() != '']

print(f"Excel 讀取完成，共 {len(df)} 筆。正在載入資料庫比對索引...")

# --- 極速比對與穩定寫入 ---
db = AccessHelper()

# 1. 抓取現有資料指紋
existing_keys = set()
try:
    cursor = db.conn.cursor()
    cursor.execute(f"SELECT 學號, 課號, 學年度, 學期 FROM {table_name}")
    rows = cursor.fetchall()
    for r in rows:
        # 建立指紋: (學號, 課號, 學年度, 學期)
        key = (str(r[0]), str(r[1]), str(r[2]), str(r[3]))
        existing_keys.add(key)
    print(f"資料庫現有 {len(existing_keys)} 筆資料，開始進行記憶體比對...")
except Exception as e:
    print("讀取資料庫索引失敗，程式停止。", e)
    db.close()
    exit()

# 2. 準備要寫入的資料
rows_to_insert = []
duplicate_count = 0

for _, row in df.iterrows():
    current_key = (str(row['學號']), str(row['課號']), str(row['學年度']), str(row['學期']))
    
    if current_key in existing_keys:
        duplicate_count += 1
    else:
        # 準備插入的資料 Tuple
        rows_to_insert.append(tuple(row[col] for col in columns))
        # 加入指紋避免 Excel 內重複
        existing_keys.add(current_key)

# 3. 使用「交易模式」批次寫入 (穩定且快速)
import_count = len(rows_to_insert)
if import_count > 0:
    print(f"正在寫入 {import_count} 筆新資料 (使用交易模式)...")
    
    try:
        # 手動開啟 Cursor 進行交易控制
        cursor = db.conn.cursor()
        
        # 關閉自動提交，開啟交易模式
        db.conn.autocommit = False 
        
        insert_sql = f"INSERT INTO {table_name} ({','.join(columns)}) VALUES ({','.join(['?']*len(columns))})"
        
        # 執行多筆寫入 (不使用 fast_executemany，避免 Access 崩潰)
        cursor.executemany(insert_sql, rows_to_insert)
        
        # 一次提交所有變更
        db.conn.commit()
        
        # 恢復自動提交
        db.conn.autocommit = True
        
        print(f"寫入成功！")
        
    except Exception as e:
        db.conn.rollback() # 發生錯誤則回滾
        print("寫入失敗，已還原變更。錯誤訊息：", e)
else:
    print("沒有需要寫入的新資料。")

db.close()
end_time = time.time()
print(f"處理完成！耗時 {end_time - start_time:.2f} 秒")
print(f"重複資料(略過)：{duplicate_count} 筆，成功匯入：{import_count} 筆")