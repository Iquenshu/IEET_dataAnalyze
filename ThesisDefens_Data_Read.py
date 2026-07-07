import os
import pyodbc
import pandas as pd

# ==============================================================================
# 【使用者設定區】
# ==============================================================================
target_db = r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\IEETdatabase.accdb" 

source_dbs = [
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_108學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_109學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_110學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_111學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_112學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_113學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_114學年度1150604.mdb",
]

target_table_name = "ThesisDefens_Data"
conn_str_template = r"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={};"
# ==============================================================================

def main():
    if not os.path.exists(target_db):
        print(f"錯誤：找不到目的地資料庫 {target_db}")
        return

    target_conn = pyodbc.connect(conn_str_template.format(target_db))
    target_cursor = target_conn.cursor()

    print("正在讀取目的地現有資料以進行重複與覆蓋檢查...")
    existing_records = set()
    try:
        target_cursor.execute(f"SELECT 學號, 論文題目 FROM {target_table_name}")
        for row in target_cursor.fetchall():
            if row[0] and row[1]:
                # 將 學號 和 論文題目 去除空白後存入，作為唯一識別 key
                existing_records.add(f"{str(row[0]).strip()}_{str(row[1]).strip()}")
        print(f"目前資料庫中已有 {len(existing_records)} 筆學位口試紀錄。")
    except pyodbc.Error as e:
        print(f"讀取目的地資料表時發生錯誤：\n{e}")
        target_conn.close()
        return

    total_inserted = 0
    total_updated = 0

    for src_db in source_dbs:
        if not os.path.exists(src_db):
            print(f"【警告】找不到檔案: {src_db}，跳過此檔案。")
            continue
            
        print(f"\n正在處理檔案: {os.path.basename(src_db)}")
        
        try:
            src_conn = pyodbc.connect(conn_str_template.format(src_db))
            query = "SELECT * FROM [口試]"
            df = pd.read_sql(query, src_conn)
            src_conn.close()
            
            if df.empty:
                print("此檔案的 [口試] 資料表內無任何資料。")
                continue

            df.columns = [str(c).strip() for c in df.columns]

            if "身分" in df.columns and "身份" not in df.columns:
                df = df.rename(columns={"身分": "身份"})

            mandatory_fields = ["學號", "姓名", "身份", "共同指導", "論文題目", "口試日期", "指導教授", "口試委員"]
            for field in mandatory_fields:
                if field not in df.columns:
                    df[field] = ""

            df = df.fillna("").astype(str).apply(lambda x: x.str.strip())
            df = df.replace('None', '')

            file_inserted_count = 0
            file_updated_count = 0

            # 逐筆檢查
            for _, row in df.iterrows():
                if not row['學號'] or not row['論文題目']:
                    continue

                match_key = f"{row['學號']}_{row['論文題目']}"
                
                try:
                    # 判斷機制：如果「學號 + 論文題目」已經存在，就執行 UPDATE 更新其他欄位
                    if match_key in existing_records:
                        update_sql = f"""
                            UPDATE {target_table_name}
                            SET 姓名 = ?, 身份 = ?, 共同指導 = ?, 口試日期 = ?, 指導教授 = ?, 口試委員 = ?
                            WHERE 學號 = ? AND 論文題目 = ?
                        """
                        target_cursor.execute(update_sql, (
                            row['姓名'], row['身份'], row['共同指導'], row['口試日期'], 
                            row['指導教授'], row['口試委員'], row['學號'], row['論文題目']
                        ))
                        target_conn.commit()
                        total_updated += 1
                        file_updated_count += 1
                    
                    # 如果不存在，才執行 INSERT 新增
                    else:
                        insert_sql = f"""
                            INSERT INTO {target_table_name} 
                            (學號, 姓名, 身份, 共同指導, 論文題目, 口試日期, 指導教授, 口試委員) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """
                        target_cursor.execute(insert_sql, (
                            row['學號'], row['姓名'], row['身份'], row['共同指導'],
                            row['論文題目'], row['口試日期'], row['指導教授'], row['口試委員']
                        ))
                        target_conn.commit()
                        existing_records.add(match_key)
                        total_inserted += 1
                        file_inserted_count += 1

                except pyodbc.Error as row_err:
                    target_conn.rollback()
                    if "23000" in str(row_err):
                        # 如果連 update 都有主鍵限制衝突，就跳過
                        continue
                    else:
                        print(f"學號 {row['學號']} 處理時發生資料庫錯誤: {row_err}")

            print(f"檔案 {os.path.basename(src_db)} 處理完成：新加入 {file_inserted_count} 筆，更新 {file_updated_count} 筆。")

        except Exception as e:
            print(f"處理檔案 {os.path.basename(src_db)} 時發生嚴重錯誤: {e}")

    print(f"\n==========================================")
    print(f"全部處理完畢！")
    print(f"本次共成功【新匯入】 {total_inserted} 筆歷年學位口試資料。")
    print(f"本次共成功【覆蓋更新】 {total_updated} 筆手動修改的資料。")
    print(f"==========================================")

    target_cursor.close()
    target_conn.close()

if __name__ == "__main__":
    main()