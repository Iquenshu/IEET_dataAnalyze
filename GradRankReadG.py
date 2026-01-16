import pandas as pd
import os
import numpy as np
from Accessdb import AccessHelper

# ==========================================
# [設定區] 研究所讀取程式
# ==========================================
TARGET_FOLDER = r'input_files\畢業總成績排名\碩士班' 
TABLE_NAME = 'GradRankG'  # 存入研究所資料表 (已移除 Class 欄位)
# ==========================================

def clean_int(val):
    if pd.isna(val) or str(val).strip() == '': return None
    try: return int(float(str(val).strip()))
    except: return None

def clean_float(val):
    if pd.isna(val) or str(val).strip() == '': return None
    try: return float(str(val).strip())
    except: return None

def import_grad_rank(file_path):
    file_name = os.path.basename(file_path)
    
    # [關鍵過濾] 只處理檔名包含 "碩士" 或 "電機碩" 的檔案
    if "碩士" not in file_name and "電機碩" not in file_name:
        return 

    print(f"\n📂 [研究所] 正在處理: {file_name} ...")

    if not os.path.exists(file_path):
        print(f"❌ 錯誤：找不到檔案 {file_path}")
        return

    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xls', '.xlsx']:
            df = pd.read_excel(file_path, dtype=str)
        elif ext == '.csv':
            try: df = pd.read_csv(file_path, encoding='utf-8', dtype=str)
            except: df = pd.read_csv(file_path, encoding='cp950', dtype=str)
        else: return
    except Exception as e:
        print(f"❌ 讀取失敗: {e}")
        return

    # 1. 欄位映射 (移除 '班別': 'Class')
    col_map_basic = {
        '學年': 'AcademicYear', '學期': 'Semester', '系所名稱': 'DeptName',
        '年級': 'Grade', '名次': 'Rank', # 注意：這裡已經沒有班別
        '學號': 'StudentID', '姓名': 'stName', '入學管道': 'EntryChannel',
        '總學分數': 'TotalCredits', '總平均分數': 'TotalAvg', 'GPA': 'GPA',
        '註記1': 'Note1', '註記2': 'Note2'
    }
    
    semester_cols_map = {}
    chinese_nums = ['一', '二', '三', '四', '五', '六', '七']
    for i, ch_num in enumerate(chinese_nums):
        y = i + 1
        semester_cols_map[f'第{ch_num}學年上學期學分數'] = f'Y{y}S1_Cred'
        semester_cols_map[f'第{ch_num}學年上學期學平均成績'] = f'Y{y}S1_Avg'
        semester_cols_map[f'第{ch_num}學年下學期學分數'] = f'Y{y}S2_Cred'
        semester_cols_map[f'第{ch_num}學年下學期學平均成績'] = f'Y{y}S2_Avg'

    full_map = {**col_map_basic, **semester_cols_map}
    
    int_db_cols = ['AcademicYear', 'Semester', 'Rank']
    float_db_cols = ['TotalCredits', 'TotalAvg', 'GPA'] + list(semester_cols_map.values())

    # 2. 定義寫入順序 (移除 Class)
    db_columns_ordered = [
        'AcademicYear', 'Semester', 'DeptName', 'Grade', 'Rank', 
        'StudentID', 'stName', 'EntryChannel', 
        'TotalCredits', 'TotalAvg', 'GPA', 'Note1', 'Note2'
    ]
    for y in range(1, 8):
        db_columns_ordered.extend([f'Y{y}S1_Cred', f'Y{y}S1_Avg', f'Y{y}S2_Cred', f'Y{y}S2_Avg'])

    db = AccessHelper()
    success_count = 0
    duplicate_count = 0
    error_count = 0

    print("開始寫入資料庫...")

    for idx, row in df.iterrows():
        sid = row.get('學號')
        if pd.isna(sid) or str(sid).strip() == '': continue
        
        insert_values = []
        for db_col in db_columns_ordered:
            target_csv_col = None
            for k, v in full_map.items():
                if v == db_col:
                    target_csv_col = k
                    break
            
            val = None
            if target_csv_col and target_csv_col in df.columns:
                raw_val = row[target_csv_col]
                
                if db_col in int_db_cols: val = clean_int(raw_val)
                elif db_col in float_db_cols: val = clean_float(raw_val)
                else: 
                    if pd.isna(raw_val) or str(raw_val).strip() == '': val = None
                    else: val = str(raw_val).strip()
            insert_values.append(val)

        # 3. 防重複邏輯 (學號+學年+學期)
        # 這裡不需要再處理 Class，邏輯變得很乾淨
        idx_sid = db_columns_ordered.index('StudentID')
        idx_ay = db_columns_ordered.index('AcademicYear')
        idx_sem = db_columns_ordered.index('Semester')
        
        params = (insert_values[idx_sid], insert_values[idx_ay], insert_values[idx_sem])
        
        if db.is_duplicate(TABLE_NAME, "StudentID=? AND AcademicYear=? AND Semester=?", params):
            duplicate_count += 1
            continue

        try:
            db.insert_row(TABLE_NAME, db_columns_ordered, tuple(insert_values))
            success_count += 1
        except Exception as e:
            err_msg = str(e)
            if '3022' in err_msg or '重複' in err_msg or '23000' in err_msg:
                duplicate_count += 1
            else:
                short_err = err_msg.split(']')[0] if ']' in err_msg else err_msg
                print(f"⚠️ 寫入錯誤 (學號: {sid}): {short_err}...")
                error_count += 1

    db.close()
    print(f"✅ 完成 {file_name}。新增: {success_count}，重複略過: {duplicate_count}，失敗: {error_count}")

if __name__ == "__main__":
    if os.path.exists(TARGET_FOLDER):
        print(f"--- [研究所] 開始掃描資料夾: {TARGET_FOLDER} ---")
        for file in os.listdir(TARGET_FOLDER):
            full_path = os.path.join(TARGET_FOLDER, file)
            if os.path.isfile(full_path) and file.lower().endswith(('.xlsx', '.xls', '.csv')):
                import_grad_rank(full_path)
    else:
        print(f"提示：資料夾不存在 ({TARGET_FOLDER})")