import pandas as pd
import pyodbc
import os
import numpy as np

# 課程資料匯入程式（含課程分類、SDGs、核心能力）
# ==========================================
# 1. 檔案設定 (請再次確認您的路徑)
# ==========================================
db_path = 'IEETdatabase.accdb'

# 分類表 (Excel檔)
class_file = r'input_files\課程分類表\課程分類表1150119.xlsx'

# 原始課程資料 (Excel檔)
raw_file = r'input_files\開課課程資料\電機系109-113學年度開課課程資料(工程認證用)匯入.xlsx'

# ==========================================
# 2. 工具函式
# ==========================================
def get_db_connection():
    full_db_path = os.path.abspath(db_path)
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        rf'DBQ={full_db_path};'
    )
    return pyodbc.connect(conn_str)

def read_file_robust(filepath):
    """智慧讀取函式"""
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"找不到檔案: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()
    print(f"正在讀取: {os.path.basename(filepath)}...")
    
    if ext in ['.xlsx', '.xls']:
        return pd.read_excel(filepath)
    else:
        try:
            return pd.read_csv(filepath, encoding='utf-8')
        except:
            return pd.read_csv(filepath, encoding='big5')

def clean_boolean(val):
    """處理各種勾選標記轉為 Bit (True/False)"""
    if pd.isna(val): return False
    s = str(val).strip().upper()
    return s in ['1', 'V', 'TRUE', 'YES', 'Y', '1.0']

def clean_smc(val):
    """SMC 欄位轉布林"""
    if pd.isna(val): return False
    try:
        return True if int(float(val)) == 1 else False
    except:
        return False

# ==========================================
# 3. 主匯入邏輯 (完全針對您的資料庫結構修正)
# ==========================================
def import_data():
    conn = None
    try:
        # --- A. 讀取並整理分類表 ---
        df_class = read_file_robust(class_file)
        df_class.columns = [c.strip() for c in df_class.columns]
        
        # 自動對應欄位
        col_name = next((c for c in df_class.columns if '課程名稱' in c or 'course_name' in c), None)
        col_math = next((c for c in df_class.columns if '數學' in c or 'math' in c), None)
        col_eng = next((c for c in df_class.columns if '工程' in c or 'eng' in c), None)
        col_gen = next((c for c in df_class.columns if '通識' in c or 'general' in c), None)

        if not all([col_name, col_math, col_eng, col_gen]):
            print("錯誤：分類表欄位無法辨識。")
            return

        class_map = {}
        for _, row in df_class.iterrows():
            c_name = str(row[col_name]).strip()
            class_map[c_name] = {
                'math': clean_boolean(row[col_math]),
                'eng': clean_boolean(row[col_eng]),
                'gen': clean_boolean(row[col_gen])
            }
        print(f"分類表載入完成 ({len(class_map)} 筆)。")

        # --- B. 讀取原始課程資料 ---
        df_raw = read_file_robust(raw_file)
        df_raw.columns = [c.strip() for c in df_raw.columns]
        print(f"原始課程資料載入完成 ({len(df_raw)} 筆)。")

        # --- C. 寫入資料庫 ---
        conn = get_db_connection()
        cursor = conn.cursor()
        
        group_keys = ['學年度', '學期', '開課單位代碼', '課號']
        grouped = df_raw.groupby(group_keys)
        
        print("開始寫入 Access 資料庫 (Courses, Course_SDGs, Course_Competencies)...")
        count_new = 0
        count_update = 0
        
        for keys, group in grouped:
            # 確保型別正確，避免 param error
            year = int(keys[0])
            sem = int(keys[1])
            dept_code = str(keys[2])
            course_code = str(keys[3])
            
            first_row = group.iloc[0]
            course_name = str(first_row['課程名稱']).strip()
            credits_val = float(first_row['學分數']) if pd.notna(first_row['學分數']) else 0.0
            
            # 取得分類
            cls = class_map.get(course_name, {'math': False, 'eng': False, 'gen': False})
            
            # 1. 檢查課程是否存在 (使用 [] 避免關鍵字錯誤)
            cursor.execute("""
                SELECT [id] FROM [Courses] 
                WHERE [academic_year]=? AND [semester]=? AND [dept_code]=? AND [course_code]=?
            """, (year, sem, dept_code, course_code))
            
            row_exist = cursor.fetchone()
            
            if row_exist:
                # --- 更新模式 ---
                course_id = row_exist[0]
                cursor.execute("""
                    UPDATE [Courses] 
                    SET [is_math_science]=?, [is_eng_prof]=?, [is_general]=? 
                    WHERE [id]=?
                """, (cls['math'], cls['eng'], cls['gen'], course_id))
                
                # 刪除舊的子表資料 (以便重新插入)
                cursor.execute("DELETE FROM [Course_SDGs] WHERE [course_id]=?", (course_id,))
                cursor.execute("DELETE FROM [Course_Competencies] WHERE [course_id]=?", (course_id,))
                count_update += 1
            else:
                # --- 新增模式 ---
                cursor.execute("""
                    INSERT INTO [Courses] (
                        [academic_year], [semester], [dept_code], [course_code], 
                        [dept_name], [course_name], [is_required], [credits], [instructor],
                        [is_math_science], [is_eng_prof], [is_general]
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    year, sem, dept_code, course_code,
                    first_row['開課單位'], course_name, first_row['必選修'], credits_val, first_row['授課教師'],
                    cls['math'], cls['eng'], cls['gen']
                ))
                cursor.execute("SELECT @@IDENTITY")
                course_id = cursor.fetchone()[0]
                count_new += 1

            # 2. 處理 SDGs (水平寫入：一門課一筆資料，含17個欄位)
            # 收集 17 個 SDG 的布林值
            sdg_values = []
            has_any_sdg = False
            for i in range(1, 18):
                col_sdg = f'SDG{i}'
                val = clean_boolean(first_row.get(col_sdg, 0))
                if val: has_any_sdg = True
                sdg_values.append(val)
            
            # 只有當至少有一個 SDG 被勾選時才寫入 (或者是您希望全0也要寫入？這裡假設全0不寫省空間)
            # 如果希望就算全0也要寫，請移除 if has_any_sdg: 檢查
            if has_any_sdg:
                sql_sdg = """
                    INSERT INTO [Course_SDGs] (
                        [course_id], 
                        [sdg_1], [sdg_2], [sdg_3], [sdg_4], [sdg_5], 
                        [sdg_6], [sdg_7], [sdg_8], [sdg_9], [sdg_10], 
                        [sdg_11], [sdg_12], [sdg_13], [sdg_14], [sdg_15], [sdg_16], [sdg_17]
                    ) VALUES (?, ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """
                cursor.execute(sql_sdg, [course_id] + sdg_values)

            # 3. 處理 Core Competencies (垂直寫入：每行一筆能力，但 SMC 欄位是水平的)
            for _, row in group.iterrows():
                comp_desc = str(row.get('核心能力', '')).strip()
                if not comp_desc or comp_desc.lower() == 'nan':
                    continue
                
                # 判斷能力類型
                cap_type = 'General' if ('通識' in comp_desc or '全校' in comp_desc) else 'EE'
                
                # 讀取 SMC_0 ~ SMC_10
                smc_values = []
                for k in range(11):
                    val = clean_smc(row.get(f'SMC_{k}', 0))
                    smc_values.append(val)
                
                # 寫入子表 (使用正確的 smc_0, smc_1... 欄位名稱)
                sql_comp = """
                    INSERT INTO [Course_Competencies] (
                        [course_id], [capability_type], [competency_desc],
                        [smc_0], [smc_1], [smc_2], [smc_3], [smc_4], 
                        [smc_5], [smc_6], [smc_7], [smc_8], [smc_9], [smc_10]
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                cursor.execute(sql_comp, [course_id, cap_type, comp_desc] + smc_values)

        conn.commit()
        print("-" * 30)
        print(f"作業完成！")
        print(f"新增課程數: {count_new}")
        print(f"更新課程數: {count_update}")
        print("資料庫已同步至最新狀態。")

    except Exception as e:
        print(f"發生錯誤: {e}")
        # 印出錯誤詳情以便除錯
        import traceback
        traceback.print_exc()
        if conn: conn.rollback()
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    import_data()