import os
import pyodbc
import pandas as pd

# ==============================================================================
# 【使用者設定區】 方便您日後修改路徑
# ==============================================================================
# 1. 目的地資料庫路徑 (IEET 總資料庫)
target_db = r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\IEETdatabase.accdb" 

# 2. 所有的來源歷年學位考試資料庫檔案列表
source_dbs = [
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_109學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_110學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_111學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_112學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_113學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_114學年度1150604.mdb",
]

# 3. 目的地資料表名稱 
target_table_name = "ThesisDefens_Data"

# Access 連線字串範本
conn_str_template = r"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={};"
# ==============================================================================


def main():
    if not os.path.exists(target_db):
        print(f"錯誤：找不到目的地資料庫 {target_db}")
        return

    # 連線至目的地資料庫
    target_conn = pyodbc.connect(conn_str_template.format(target_db))
    target_cursor = target_conn.cursor()

    # --- 步驟 A: 讀取目的地現有的資料，用來做重複檢查 ---
    print("正在讀取目的地現有資料以進行重複檢查...")
    
    existing_records = set()
    try:
        target_cursor.execute(f"SELECT 學號, 論文題目 FROM {target_table_name}")
        for row in target_cursor.fetchall():
            if row[0] and row[1]:
                existing_records.add(f"{str(row[0]).strip()}_{str(row[1]).strip()}")
        print(f"目前資料庫中已有 {len(existing_records)} 筆學位口試紀錄。")
    except pyodbc.Error as e:
        print(f"讀取目的地資料表時發生錯誤 (請確認 Access 內欄位已改為 '身份')：\n{e}")
        target_conn.close()
        return

    # --- 步驟 B: 逐一讀取來源檔案並處理 ---
    total_inserted = 0

    for src_db in source_dbs:
        if not os.path.exists(src_db):
            print(f"【警告】找不到檔案: {src_db}，跳過此檔案。")
            continue
            
        print(f"\n正在處理檔案: {os.path.basename(src_db)}")
        
        try:
            src_conn = pyodbc.connect(conn_str_template.format(src_db))
            
            # 讀取舊檔資料 (來源欄位為「身份」)
            query = """
                SELECT 
                    學號, 姓名, 共同指導, 論文題目, 口試日期, 指導教授, 口試委員, 身份
                FROM 
                    [口試]
            """
            df = pd.read_sql(query, src_conn)
            src_conn.close()
            
            if df.empty:
                print("此檔案的 [口試] 資料表內無任何資料。")
                continue

            # 清洗資料
            df = df.fillna("").astype(str).apply(lambda x: x.str.strip())
            df = df.replace('None', '')

            # 逐筆檢查並插入資料
            for _, row in df.iterrows():
                match_key = f"{row['學號']}_{row['論文題目']}"
                
                if match_key in existing_records:
                    continue
                
                # 【已修正】目的地欄位也同步改為「身份」
                insert_sql = f"""
                    INSERT INTO {target_table_name} 
                    (學號, 姓名, 身份, 共同指導, 論文題目, 口試日期, 指導教授, 口試委員) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """
                target_cursor.execute(insert_sql, (
                    row['學號'],
                    row['姓名'],
                    row['身份'],
                    row['共同指導'],
                    row['論文題目'],
                    row['口試日期'],
                    row['指導教授'],
                    row['口試委員']
                ))
                
                existing_records.add(match_key)
                total_inserted += 1
                
            target_conn.commit()
            print(f"檔案 {os.path.basename(src_db)} 匯入完成。")

        except Exception as e:
            print(f"處理檔案 {os.path.basename(src_db)} 時發生錯誤: {e}")
            target_conn.rollback()

    print(f"\n==========================================")
    print(f"全部處理完畢！本次共成功新匯入 {total_inserted} 筆歷年學位口試資料。")
    print(f"==========================================")

    target_cursor.close()
    target_conn.close()

if __name__ == "__main__":
    main()