import pandas as pd
import os
from Accessdb import AccessHelper
from openpyxl.styles import PatternFill

# ==========================================
# 設定
# ==========================================
db = AccessHelper()
OUTPUT_DIR = r'output_files\課程分類'
# [修正] 移除 v3，檔名改為：課程分類表_由資料庫生成_日期.xlsx
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f'課程分類表_由資料庫生成_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx')

# [設定] 參考用的手動校正表路徑
reference_class_file = r'input_files\課程分類表\課程分類表_20260211.xlsx'

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
    print(f"建立目錄: {OUTPUT_DIR}")

# ==========================================
# 分類邏輯 (完整保留原始結構、關鍵字與 0/1 數值)
# ==========================================
def classify_course_strict(row):
    name = str(row['課程名稱']).strip()
    name_upper = name.upper()
    
    # 預設值 (所有分類為 0，若都沒中，最後預設 general=1)
    is_math = 0
    is_science = 0
    is_eng = 0
    is_general = 0
    
    # ---------------------------------------------------------
    # 1. 數學 (Mathematics) - 最優先
    # ---------------------------------------------------------
    math_keywords = [
        '數學', '微積分', '線性代數', '機率', '統計', '會計', '計算', 
        '微分方程', '複變', '離散', '數值分析', '幾何', '代數',
        'CALCULUS', 'ALGEBRA', 'PROBABILITY', 'STATISTICS', 'MATH'
    ]
    
    # ---------------------------------------------------------
    # 2. 工程專業 (Engineering) - 次優先
    # ---------------------------------------------------------
    eng_keywords = [
        '電機', '資訊', '程式', '通訊', '晶片', '電力', '計算機', '工程', 
        '多媒體', '書報討論', '電子', '電路', '系統', '控制', '半導體', 
        '設計', '實習', '專題', '實驗', '邏輯', '微處理', 'VLSI', 'FPGA', 
        'JAVA', 'PYTHON', 'C++', 'AI', '機器學習', '演算法', '資料結構', 
        '網路', '訊號', '電波', '光電', '類比', '數位',
        'ELECTRIC', 'ELECTRONIC', 'SYSTEM', 'SIGNAL', 'CONTROL', 'COMMUNICATION',
        'NETWORK', 'SEMICONDUCTOR', 'CHIP', 'DESIGN', 'PROJECT', 'LAB'
    ]

    # ---------------------------------------------------------
    # 3. 基礎科學 (Science)
    # ---------------------------------------------------------
    sci_keywords = [
        '物理', '化學', '生物', '力學', '熱力學', '電磁', '量子', '光學', 
        'PHYSICS', 'CHEMISTRY', 'BIOLOGY', 'MECHANICS'
    ]
    
    # ---------------------------------------------------------
    # 4. 通識 (General) - 最後判斷
    # ---------------------------------------------------------
    gen_keywords = [
        '文', '語', '史', '生活', '寫作', '經濟', '政治', '管理', '人文', 
        '服務學習', '國防', '公共', '倫理', '運動', '食品', '身心', '文化', 
        '音樂', '藝術', '體育', '軍訓', '博雅', '社會', '哲學', '心理', '法學', 
        '通識', '跨域', '講座', '溝通', '思考', '概論', '導論'
    ]

    # --- 判斷流程 (互斥) ---
    if any(k in name_upper for k in math_keywords):
        is_math = 1
    elif any(k in name_upper for k in eng_keywords):
        is_eng = 1
    elif any(k in name_upper for k in sci_keywords):
        is_science = 1
    elif '學' in name and not any(k in name for k in ['文學', '史學', '哲學', '法學', '美學', '語言學', '心理學', '社會學', '政治學', '經濟學', '管理學']): 
        is_science = 1
    elif any(k in name_upper for k in gen_keywords):
        is_general = 1
    else:
        is_general = 1

    return is_math, is_science, is_eng, is_general

# ==========================================
# 主程式
# ==========================================
def generate_classification_from_db():
    # 讀取參考用的舊分類檔
    ref_data = {}
    if os.path.exists(reference_class_file):
        print(f"正在讀取參考分類表: {reference_class_file}")
        df_ref = pd.read_excel(reference_class_file)
        for _, row in df_ref.iterrows():
            # 建立以「課號_課程名稱」為 key 的字典，保留原始 0/1 分類
            key = f"{str(row['課號']).strip()}_{str(row['課程名稱']).strip()}"
            ref_data[key] = {
                'is_math': row.get('is_math', 0),
                'is_science': row.get('is_science', 0),
                'is_eng_prof': row.get('is_eng_prof', 0),
                'is_general': row.get('is_general', 0)
            }

    print("從資料庫 STscore 讀取課程清單...")
    sql = "SELECT DISTINCT 課號, 課程名稱 FROM STscore"
    try:
        df_courses = pd.read_sql(sql, db.conn)
    except Exception as e:
        print(f"資料庫讀取失敗: {e}")
        return

    if df_courses.empty:
        print("警告：STscore 資料表是空的！")
        return

    print(f"共取得 {len(df_courses)} 筆不重複課程。")
    print("正在進行繼承與自動分類...")

    results = []
    for _, row in df_courses.iterrows():
        c_code = str(row['課號']).strip()
        c_name = str(row['課程名稱']).strip()
        key = f"{c_code}_{c_name}"
        
        if key in ref_data:
            # 1. 繼承原有 0/1 資料
            results.append({
                '課號': c_code,
                '課程名稱': c_name,
                'is_math': ref_data[key]['is_math'],
                'is_science': ref_data[key]['is_science'],
                'is_eng_prof': ref_data[key]['is_eng_prof'],
                'is_general': ref_data[key]['is_general'],
                'is_new': False
            })
        else:
            # 2. 執行原始自動分類邏輯 (返回 0/1)
            m, s, e, g = classify_course_strict(row)
            results.append({
                '課號': c_code,
                '課程名稱': c_name,
                'is_math': m,
                'is_science': s,
                'is_eng_prof': e,
                'is_general': g,
                'is_new': True
            })
        
    df_out = pd.DataFrame(results)
    
    # 排序優化：維持原始排序邏輯
    df_out = df_out.sort_values(
        by=['is_general', 'is_math', 'is_science', 'is_eng_prof', '課程名稱'], 
        ascending=[False, False, False, False, True]
    )
    
    # 寫入 Excel 並處理格式
    print(f"正在寫入檔案: {OUTPUT_FILE}")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        # 排除輔助欄位後寫入
        df_export = df_out.drop(columns=['is_new'])
        df_export.to_excel(writer, index=False, sheet_name='課程分類')
        
        ws = writer.book['課程分類']
        red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        
        # 針對新課程標註紅色
        for i, is_new in enumerate(df_out['is_new'], start=2):
            if is_new:
                for col in range(1, 7): # A-F 欄
                    ws.cell(row=i, column=col).fill = red_fill

    print("完成！紅色標註為新增課程，分類數據維持 0 與 1。")

if __name__ == "__main__":
    generate_classification_from_db()
    db.close()