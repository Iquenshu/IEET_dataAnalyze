import os
import pyodbc

# 所有的來源歷年檔案列表 (從您的錯誤日誌中抓取的實際檔名)
source_dbs = [
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_109學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_110學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_111學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_112學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_113學年度.mdb",
    r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\input_files\學位考試資料\paper_114學年度1150604.mdb",
]

conn_str_template = r"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={};"

print("==================================================")
print("  開始掃描所有歷年檔案 [口試] 表的真實欄位名稱")
print("==================================================\n")

for src_db in source_dbs:
    if not os.path.exists(src_db):
        print(f"❌ 找不到檔案: {os.path.basename(src_db)}")
        continue
        
    try:
        conn = pyodbc.connect(conn_str_template.format(src_db))
        cursor = conn.cursor()
        
        # 只讀取結構，不撈資料
        cursor.execute("SELECT * FROM [口試] WHERE 1=0")
        real_columns = [column[0] for column in cursor.description]
        
        print(f"📄 檔案: {os.path.basename(src_db)}")
        print(f"   實際欄位: {real_columns}\n")
        
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"💥 檔案 {os.path.basename(src_db)} 讀取失敗，錯誤訊息: {e}\n")

print("==================================================")
print("掃描完畢！請檢查哪一個檔案的欄位字體或名稱有落差。")
print("==================================================")