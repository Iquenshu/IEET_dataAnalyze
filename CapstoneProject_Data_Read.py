import os
import re
import pyodbc
import pandas as pd

# ==================== 1. 設定路徑與檔案 ====================
# 目的地資料庫
target_db = r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\IEETdatabase.accdb" 

# 所有的來源歷年資料庫檔案列表 (以後有新檔案，直接加在陣列後面即可)
source_dbs = [
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\專題報名資料\RIMT2018_2020.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\專題報名資料\RIMT2018_2021.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\專題報名資料\RIMT2018_2022.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\專題報名資料\RIMT2018_2023to2025.mdb",
    # r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\專題報名資料\RIMT2024_2026.mdb", # 未來新加的檔放這裡
]

# 新建立的資料表名稱
target_table_name = "CapstoneProject_Data"

# Access 連線字串範本 (已修正為雙大括號，防止 .format() 發生 KeyError)
conn_str_template = r"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={};"

# ==================== 2. 學年度轉換邏輯 ====================
def convert_to_academic_year(id_no):
    """
    從專題組編號 (如 Project202000022) 提取年份並轉為學年度
    根據您的規則：2020年 -> 109學年度 (偏離值為 1911)
    """
    match = re.search(r'\d{4}', str(id_no))
    if match:
        year = int(match.group())
        academic_year = year - 1911  # 2020 - 1911 = 109
        return f"{academic_year}學年度"
    return "未知學年度"

# ==================== 3. 主程式開始 ====================
def main():
    if not os.path.exists(target_db):
        print(f"錯誤：找不到目的地資料庫 {target_db}")
        return

    # 連線至目的地資料庫
    target_conn = pyodbc.connect(conn_str_template.format(target_db))
    target_cursor = target_conn.cursor()

    # --- 步驟 A: 如果表不存在，則自動建立新表 (含複合主鍵防止重複) ---
    try:
        create_table_sql = f"""
        CREATE TABLE {target_table_name} (
            學年度 TEXT(50),
            專題組編號 TEXT(100),
            參賽分組 TEXT(50),
            專題名稱 TEXT(255),
            學生姓名 TEXT(50),
            學號 TEXT(50),
            指導教授 TEXT(50),
            PRIMARY KEY (專題組編號, 學號)
        );
        """
        target_cursor.execute(create_table_sql)
        target_conn.commit()
        print(f"成功建立新資料表：{target_table_name}")
    except pyodbc.Error as e:
        # 如果表已經存在，會跳到這裡，我們直接忽略它
        pass

    # --- 步驟 B: 讀取目的地現有的資料，用來做重複檢查 ---
    print("正在讀取目的地現有資料以進行重複檢查...")
    target_cursor.execute(f"SELECT 專題組編號, 學號 FROM {target_table_name}")
    # 將現有的 (專題組編號, 學號) 存在一個 set 裡面，比對速度極快
    existing_records = {f"{row[0]}_{row[1]}" for row in target_cursor.fetchall()}
    print(f"目前資料庫中已有 {len(existing_records)} 筆學生專題紀錄。")

    # --- 步驟 C: 逐一讀取來源檔案並處理 ---
    total_inserted = 0

    for src_db in source_dbs:
        if not os.path.exists(src_db):
            print(f"【警告】找不到檔案: {src_db}，跳過此檔案。")
            continue
            
        print(f"\n正在處理檔案: {os.path.basename(src_db)}")
        
        try:
            src_conn = pyodbc.connect(conn_str_template.format(src_db))
            
            # 使用 SQL JOIN 直接在來源資料庫把兩張表串起來
            # 使用 LEFT JOIN 確保就算某一組漏填指導教授，學生的資料依然能撈出來
            query = """
                SELECT 
                    R.IdNo AS 專題組編號,
                    R.GroupN AS 參賽分組,
                    R.ledname AS 專題名稱,
                    R.stname AS 學生姓名,
                    R.stdid AS 學號,
                    A.tchName AS 指導教授
                FROM 
                    RIMT2018 AS R
                LEFT JOIN 
                    Advisor AS A ON R.IdNo = A.IdNo
            """
            df = pd.read_sql(query, src_conn)
            src_conn.close()
            
            if df.empty:
                print("此檔案內無有效資料。")
                continue

            # 清洗資料：去除前後空格，避免因為空白導致重複檢查失效
            df = df.astype(str).apply(lambda x: x.str.strip())

            # 插入資料
            for _, row in df.iterrows():
                # 檢查這筆資料是否已經在目的地資料庫中
                match_key = f"{row['專題組編號']}_{row['學號']}"
                
                if match_key in existing_records:
                    # 資料已存在，跳過不處理
                    continue
                
                # 計算學年度
                academic_year = convert_to_academic_year(row['專題組編號'])
                
                # 寫入目的地
                insert_sql = f"""
                    INSERT INTO {target_table_name} 
                    (學年度, 專題組編號, 參賽分組, 專題名稱, 學生姓名, 學號, 指導教授) 
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """
                target_cursor.execute(insert_sql, (
                    academic_year,
                    row['專題組編號'],
                    row['參賽分組'],
                    row['專題名稱'],
                    row['學生姓名'],
                    row['學號'],
                    row['指導教授'] if row['指導教授'] != 'None' else ''
                ))
                
                # 將新加入的 key 放進集合，防止同一個檔案內有重複行
                existing_records.add(match_key)
                total_inserted += 1
                
            target_conn.commit()
            print(f"檔案 {os.path.basename(src_db)} 處理完成。")

        except Exception as e:
            print(f"處理檔案 {os.path.basename(src_db)} 時發生錯誤: {e}")
            target_conn.rollback()

    print(f"\n==========================================")
    print(f"全部處理完畢！本次共成功新匯入 {total_inserted} 筆資料。")
    print(f"==========================================")

    target_cursor.close()
    target_conn.close()

if __name__ == "__main__":
    main()