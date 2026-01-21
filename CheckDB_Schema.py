import pyodbc
import os

# 資料庫路徑
db_path = 'IEETdatabase.accdb'

def check_schema():
    if not os.path.exists(db_path):
        print(f"找不到資料庫檔案: {db_path}")
        return

    full_db_path = os.path.abspath(db_path)
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        rf'DBQ={full_db_path};'
    )

    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        # 取得所有資料表
        tables = [t.table_name for t in cursor.tables(tableType='TABLE')]
        
        print(f"=== 資料庫 {db_path} 結構檢測 ===")
        
        for table in tables:
            print(f"\n[資料表: {table}]")
            # 取得該表的所有欄位
            columns = [row.column_name for row in cursor.columns(table=table)]
            print("  欄位清單:", columns)
            
            # 特別檢查 Courses 表的關鍵欄位
            if table == 'Courses':
                required = ['id', 'academic_year', 'semester', 'dept_code', 'course_code']
                print("  -> 檢查關鍵欄位:")
                for r in required:
                    if r in columns:
                        print(f"     [OK] {r}")
                    else:
                        print(f"     [MISSING] 找不到欄位 '{r}' !!!")

        conn.close()
        
    except Exception as e:
        print(f"讀取資料庫失敗: {e}")

if __name__ == "__main__":
    check_schema()