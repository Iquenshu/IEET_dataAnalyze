import pandas as pd
import pyodbc
import os
import time

# ==========================================
# 1. 設定檔案與資料表
# ==========================================
db_path = 'IEETdatabase.accdb'
data_path = r'input_files\學生成績\電機系109-113學年度大學部及碩士班博士班學生所有成績.xlsx'
table_name = 'STscore'
BATCH_SIZE = 1000  # 設定每 1000 筆寫入一次並顯示進度

# ==========================================
# 2. 資料庫連線工具
# ==========================================
def get_db_connection():
    full_db_path = os.path.abspath(db_path)
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        rf'DBQ={full_db_path};'
    )
    return pyodbc.connect(conn_str)

def clean_key_str(val):
    """將資料庫或 Excel 的值統一轉為乾淨的字串 (去除 .0 與 空白)"""
    if pd.isna(val) or val is None:
        return ""
    # 先轉字串 -> 去除前後空白 -> 如果有 .0 結尾則去除
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# ==========================================
# 3. 主程式邏輯
# ==========================================
def import_scores():
    print(f"正在讀取 Excel 檔案: {os.path.basename(data_path)} ...")
    
    if not os.path.exists(data_path):
        print(f"錯誤：找不到檔案 {data_path}")
        return

    # 1. 讀取 Excel
    try:
        df = pd.read_excel(data_path)
    except Exception as e:
        print(f"Excel 讀取失敗: {e}")
        return

    # 2. 欄位處理
    df.columns = [c.strip() for c in df.columns]
    
    rename_map = {'系所代碼': '開課系所代碼', '系所': '開課系所'}
    df.rename(columns=rename_map, inplace=True)

    # 處理重複成績
    if '成績.1' in df.columns:
        print("-> 偵測到雙重成績欄位，使用第二欄作為最終成績。")
        df['成績'] = df['成績.1']

    # 補齊必選修
    if '必選修' not in df.columns:
        df['必選修'] = '' 

    # 3. 資料清洗
    required_cols = [
        '學年度', '學期', '開課系所代碼', '開課系所', '課號', '課程名稱', '必選修',
        '學號', '姓名', '學分數', '成績', '等第成績'
    ]
    
    # 確保只有需要的欄位
    df_import = df[required_cols].copy()

    # 型別轉換
    df_import['學分數'] = pd.to_numeric(df_import['學分數'], errors='coerce').fillna(0)
    df_import['成績'] = pd.to_numeric(df_import['成績'], errors='coerce').fillna(0)
    
    # 強制將關鍵欄位轉為乾淨字串 (比對用)
    for col in ['學年度', '學期', '學號', '課號', '開課系所代碼']:
        df_import[col] = df_import[col].apply(clean_key_str)

    df_import['等第成績'] = df_import['等第成績'].astype(str).replace('nan', '').str.strip()

    # 4. 退選處理 (成績999)
    withdraw_mask = (df_import['成績'] == 999) & (df_import['等第成績'] == '')
    if withdraw_mask.sum() > 0:
        print(f"-> 標記 {withdraw_mask.sum()} 筆退選資料 (成績999)。")
        df_import.loc[withdraw_mask, '等第成績'] = '退選'

    # ==========================================
    # 5. 資料庫比對 (嚴格模式)
    # ==========================================
    conn = get_db_connection()
    cursor = conn.cursor()
    
    print("正在讀取資料庫現有資料 (建立比對指紋)...")
    
    # 抓取現有的 Key: 學年, 學期, 課號, 學號
    cursor.execute("SELECT [學年度], [學期], [課號], [學號] FROM STscore")
    existing_keys = set()
    
    rows = cursor.fetchall()
    for row in rows:
        # 使用相同的 clean_key_str 邏輯處理資料庫取出的資料
        key = (
            clean_key_str(row[0]), # 學年度
            clean_key_str(row[1]), # 學期
            clean_key_str(row[2]), # 課號
            clean_key_str(row[3])  # 學號
        )
        existing_keys.add(key)
    
    print(f"資料庫現有 {len(existing_keys)} 筆不重複成績紀錄。")

    # 準備要寫入的資料
    data_to_insert = []
    duplicate_count = 0
    excel_internal_dupes = 0
    
    # 用來檢查 Excel 內部是否有重複 (有些 Excel 本身就會重複列)
    current_batch_keys = set()

    for row in df_import.itertuples(index=False):
        # 建立這筆資料的 Key (順序需與 required_cols 對應)
        # 0:學年度, 1:學期, 4:課號, 7:學號
        current_key = (str(row[0]), str(row[1]), str(row[4]), str(row[7]))
        
        if current_key in existing_keys:
            duplicate_count += 1
        elif current_key in current_batch_keys:
            excel_internal_dupes += 1
        else:
            data_to_insert.append(tuple(row))
            current_batch_keys.add(current_key) # 加入暫存，避免本次匯入重複

    # ==========================================
    # 6. 分批寫入 (Batch Insert)
    # ==========================================
    total_insert = len(data_to_insert)
    
    if total_insert == 0:
        print("-" * 30)
        print("沒有需要寫入的新資料。")
        print(f"資料庫重複: {duplicate_count} 筆")
        print(f"Excel內部重複: {excel_internal_dupes} 筆")
    else:
        print(f"準備寫入 {total_insert} 筆新資料...")
        if duplicate_count > 0:
            print(f"(已過濾掉 {duplicate_count} 筆資料庫重複資料)")
        
        insert_sql = """
            INSERT INTO STscore (
                [學年度], [學期], [開課系所代碼], [開課系所], [課號], [課程名稱], [必選修],
                [學號], [姓名], [學分數], [成績], [等第成績]
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        
        try:
            conn.autocommit = False
            
            for i in range(0, total_insert, BATCH_SIZE):
                batch = data_to_insert[i : i + BATCH_SIZE]
                cursor.executemany(insert_sql, batch)
                conn.commit() # 每一批次提交一次
                
                # 顯示進度
                current_count = min(i + BATCH_SIZE, total_insert)
                print(f"進度: {current_count} / {total_insert} ... 完成")
                
            print("-" * 30)
            print(f"全數匯入完成！共新增 {total_insert} 筆資料。")
            
        except Exception as e:
            conn.rollback()
            print(f"寫入過程中發生錯誤: {e}")
        finally:
            conn.close()

if __name__ == "__main__":
    import_scores()