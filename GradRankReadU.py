import pandas as pd
import os
import numpy as np
from Accessdb import AccessHelper

#  大學部畢業總成績排名讀取程式

# ==========================================
# [設定區] 大學部讀取程式
# ==========================================
TARGET_FOLDER = r'input_files\畢業總成績排名\大學部' 
TABLE_NAME = 'GradRankU'  # 存入大學部資料表
# ==========================================

def clean_int(val):
    if pd.isna(val) or str(val).strip() == '': return None
    try: return int(float(str(val).strip()))
    except: return None

def clean_float(val):
    if pd.isna(val) or str(val).strip() == '': return None
    try: return float(str(val).strip())
    except: return None

def import_undergrad_rank(file_path):
    file_name = os.path.basename(file_path)
    
    # [關鍵過濾] 只處理檔名包含 "大學部" 的檔案
    if "大學部" not in file_name:
        return 

    print(f"\n📂 [大學部] 正在處理: {file_name} ...")

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

    # 基本欄位映射
    col_map_basic = {
        '學年': 'AcademicYear', '學期': 'Semester', '系所名稱': 'DeptName',
        '年級': 'Grade', '班別': 'Class', '名次': 'Rank',
        '學號': 'StudentID', '姓名': 'stName', '入學管道': 'EntryChannel',
        '總學分數': 'TotalCredits', '總平均分數': 'TotalAvg', 'GPA': 'GPA',
        '註記1': 'Note1', '註記2': 'Note2'
    }
    
    # 產生學期成績欄位 (Y1S1...Y7S2)
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

    # 定義寫入順序
    db_columns_ordered = [
        'AcademicYear', 'Semester', 'DeptName', 'Grade', 'Class', 'Rank', 
        'StudentID', 'stName', 'EntryChannel', 
        'TotalCredits', 'TotalAvg', 'GPA', 'Note1', 'Note2'
    ]
    # 加入所有學期欄位
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

        # 防重複邏輯 (大學部專用：學號+學年+學期+班別)
        try:
            # 直接 Insert，讓資料庫的主鍵(PK)去擋重複
            # Access 若遇到主鍵衝突會拋出錯誤，我們只要捕捉它即可
            db.insert_row(TABLE_NAME, db_columns_ordered, tuple(insert_values))
            success_count += 1
        
        except Exception as e:
            err_msg = str(e)
            # 捕捉主鍵重複錯誤 (Access 錯誤代碼通常包含 '3022' 或文字敘述)
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
        print(f"--- [大學部] 開始掃描資料夾: {TARGET_FOLDER} ---")
        for file in os.listdir(TARGET_FOLDER):
            full_path = os.path.join(TARGET_FOLDER, file)
            if os.path.isfile(full_path) and file.lower().endswith(('.xlsx', '.xls', '.csv')):
                import_undergrad_rank(full_path)
    else:
        print(f"提示：資料夾不存在 ({TARGET_FOLDER})")