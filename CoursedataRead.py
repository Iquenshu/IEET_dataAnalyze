import pandas as pd
import pyodbc
import os
import numpy as np

# 課程資料匯入程式（含課程分類、SDGs、核心能力）
# ==========================================
# 1. 檔案與資料庫路徑設定 (置於最前端方便隨時更動)
# ==========================================
db_path = 'IEETdatabase.accdb'

# 📌 [提示] 請在此處替換為您最新由資料庫生成、且經關鍵字優化後的分類表路徑
class_file = r'input_files\課程分類表\課程分類表_20260610.xlsx'

# 原始開課課程資料 (維持不變)
raw_file = r'input_files\開課課程資料\電機系109-113學年度開課課程資料(工程認證用)匯入.xlsx'


# ==========================================
# 2. 工具函式 (完全保留原有邏輯，並微調強化相容性)
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
    """處理各種勾選標記與全新1/0註記轉為 Bit (True/False)"""
    if pd.isna(val): return False
    # 💡 擴充支援：若分類表內直接為純數字 1 或 0，直接進行判定
    if isinstance(val, (int, float, np.number)):
        return int(val) == 1
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
# 3. 主匯入邏輯 (修正：改以【課號】作為核心 Mapping 依據)
# ==========================================
def import_data():
    conn = None
    try:
        # --- A. 讀取並整理分類表 ---
        df_class = read_file_robust(class_file)
        df_class.columns = [c.strip() for c in df_class.columns]
        
        # 💡 [關鍵修正] 自動對應欄位：新增課號 (course_code) 偵測，用以取代舊版的課名比對
        col_code = next((c for c in df_class.columns if '課號' in c or 'course_code' in c), None)
        col_math = next((c for c in df_class.columns if '數學' in c or 'is_math' in c), None)
        col_science = next((c for c in df_class.columns if '科學' in c or 'science' in c), None)
        col_eng = next((c for c in df_class.columns if '工程' in c or 'eng' in c), None)
        col_gen = next((c for c in df_class.columns if '通識' in c or 'general' in c), None)

        # 簡單檢查關鍵欄位
        if not col_code:
            print("❌ 錯誤：分類表中找不到 '課號' 欄位，無法進行精準資料對應。")
            return

        # 💡 [關鍵修正] 建立以【課號】為 Key 的對應字典，確保全半形與空格不會導致比對失效
        class_map = {}
        for _, row in df_class.iterrows():
            c_code = str(row[col_code]).strip()
            class_map[c_code] = {
                'math': clean_boolean(row.get(col_math, 0)),
                'science': clean_boolean(row.get(col_science, 0)),
                'eng': clean_boolean(row.get(col_eng, 0)),
                'gen': clean_boolean(row.get(col_gen, 0))
            }
        print(f"➔ 分類表載入完成，共取得 ({len(class_map)} 筆) 獨立課號分類對應。")

        # --- B. 讀取原始課程資料 ---
        df_raw = read_file_robust(raw_file)
        df_raw.columns = [c.strip() for c in df_raw.columns]
        print(f"原始課程資料載入完成 ({len(df_raw)} 筆)。")

        # --- C. 寫入與更新資料庫 ---
        conn = get_db_connection()
        cursor = conn.cursor()
        
        group_keys = ['學年度', '學期', '開課單位代碼', '課號']
        grouped = df_raw.groupby(group_keys)
        
        print("開始寫入 Access 資料庫 (Courses, Course_SDGs, Course_Competencies)...")
        count_new = 0
        count_update = 0
        
        for keys, group in grouped:
            # 確保型別正確
            year = int(keys[0])
            sem = int(keys[1])
            dept_code = str(keys[2])
            course_code = str(keys[3])
            
            first_row = group.iloc[0]
            course_name = str(first_row['課程名稱']).strip()
            credits_val = float(first_row['學分數']) if pd.notna(first_row['學分數']) else 0.0
            
            # 💡 [關鍵修正] 取得分類：改以 course_code (課號) 向 class_map 字典提領四大領域註記
            cls = class_map.get(course_code, {'math': False, 'science': False, 'eng': False, 'gen': False})
            
            # 1. 檢查課程是否存在
            cursor.execute("""
                SELECT [id] FROM [Courses] 
                WHERE [academic_year]=? AND [semester]=? AND [dept_code]=? AND [course_code]=?
            """, (year, sem, dept_code, course_code))
            
            row_exist = cursor.fetchone()
            
            if row_exist:
                # --- 更新模式：將從新分類表提領出的四大分類寫回資料庫 ---
                course_id = row_exist[0]
                cursor.execute("""
                    UPDATE [Courses] 
                    SET [is_math]=?, [is_science]=?, [is_eng_prof]=?, [is_general]=? 
                    WHERE [id]=?
                """, (cls['math'], cls['science'], cls['eng'], cls['gen'], course_id))
                
                # 刪除舊的子表資料 (以便重新插入，維持原始設計)
                cursor.execute("DELETE FROM [Course_SDGs] WHERE [course_id]=?", (course_id,))
                cursor.execute("DELETE FROM [Course_Competencies] WHERE [course_id]=?", (course_id,))
                count_update += 1
            else:
                # --- 新增模式：若為新開課程，直接連同分類一併寫入 ---
                cursor.execute("""
                    INSERT INTO [Courses] (
                        [academic_year], [semester], [dept_code], [course_code], 
                        [dept_name], [course_name], [is_required], [credits], [instructor],
                        [is_math], [is_science], [is_eng_prof], [is_general]
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    year, sem, dept_code, course_code,
                    first_row['開課單位'], course_name, first_row['必選修'], credits_val, first_row['授課教師'],
                    cls['math'], cls['science'], cls['eng'], cls['gen']
                ))
                cursor.execute("SELECT @@IDENTITY")
                course_id = cursor.fetchone()[0]
                count_new += 1

            # 2. 處理 SDGs (完整保留原有邏輯)
            sdg_values = []
            has_any_sdg = False
            for i in range(1, 18):
                col_sdg = f'SDG{i}'
                val = clean_boolean(first_row.get(col_sdg, 0))
                if val: has_any_sdg = True
                sdg_values.append(val)
            
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

            # 3. 處理 Core Competencies (完整保留原有邏輯)
            for _, row in group.iterrows():
                comp_desc = str(row.get('核心能力', '')).strip()
                if not comp_desc or comp_desc.lower() == 'nan':
                    continue
                
                cap_type = 'General' if ('通識' in comp_desc or '全校' in comp_desc) else 'EE'
                
                smc_values = []
                for k in range(11):
                    val = clean_smc(row.get(f'SMC_{k}', 0))
                    smc_values.append(val)
                
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
        print(f"🎉 作業完成！")
        print(f"新增課程數: {count_new}")
        print(f"更新課程數: {count_update}")
        print("Access 資料庫 Courses 表已依據【最新課號分類規範】全數同步完畢。")

    except Exception as e:
        print(f"发生错误: {e}")
        import traceback
        traceback.print_exc()
        if conn: conn.rollback()
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    import_data()