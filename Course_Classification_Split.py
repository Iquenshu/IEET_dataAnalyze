import pandas as pd
import os
from Accessdb import AccessHelper

# ==========================================
# 設定
# ==========================================
db = AccessHelper()
OUTPUT_DIR = r'output_files\課程分類'
# 檔名加上 v3 以示區別
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f'課程分類表_由資料庫生成_v3_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx')

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
    print(f"建立目錄: {OUTPUT_DIR}")

# ==========================================
# 分類邏輯 (依據您的新需求)
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
    # 包含: XX數學, XX會計, XX統計, XX計算, 微積分, 線性代數...
    math_keywords = [
        '數學', '微積分', '線性代數', '機率', '統計', '會計', '計算', 
        '微分方程', '複變', '離散', '數值分析', '幾何', '代數',
        'CALCULUS', 'ALGEBRA', 'PROBABILITY', 'STATISTICS', 'MATH'
    ]
    
    # ---------------------------------------------------------
    # 2. 工程專業 (Engineering) - 次優先
    # ---------------------------------------------------------
    # 包含: 電機, 資訊, 程式, 通訊, 晶片, 電力, 計算機, 工程, 多媒體, 書報討論
    # 額外加入: 電子, 電路, 系統, 控制, 半導體, 專題, 實習, 實驗 (常見工程課)
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
    # 包含: 物理, 化學, 生物, 力學, 熱力學, 電磁學, 量子, 光學
    # 以及名稱包含 "學" 的 (但要排除掉已經被歸類為工程的電子學/電路學等)
    sci_keywords = [
        '物理', '化學', '生物', '力學', '熱力學', '電磁', '量子', '光學', 
        'PHYSICS', 'CHEMISTRY', 'BIOLOGY', 'MECHANICS'
    ]
    
    # ---------------------------------------------------------
    # 4. 通識 (General) - 最後判斷
    # ---------------------------------------------------------
    # 包含: 文, 語, 史, 生活, 寫作, 經濟, 政治, 管理, 人文, 服務學習, 
    # 國防, 公共, 倫理, 運動, 食品, 身心, 文化, 音樂, 藝術...
    gen_keywords = [
        '文', '語', '史', '生活', '寫作', '經濟', '政治', '管理', '人文', 
        '服務學習', '國防', '公共', '倫理', '運動', '食品', '身心', '文化', 
        '音樂', '藝術', '體育', '軍訓', '博雅', '社會', '哲學', '心理', '法學', 
        '通識', '跨域', '講座', '溝通', '思考', '概論', '導論'
    ]

    # --- 判斷流程 (互斥) ---
    
    # 1. 數學
    if any(k in name_upper for k in math_keywords):
        is_math = 1
        
    # 2. 工程 (優先於 "XX學" 的科學判斷，以免 "電子學" 變科學)
    elif any(k in name_upper for k in eng_keywords):
        is_eng = 1
        
    # 3. 科學
    elif any(k in name_upper for k in sci_keywords):
        is_science = 1
    elif '學' in name and not any(k in name for k in ['文學', '史學', '哲學', '法學', '美學', '語言學', '心理學', '社會學', '政治學', '經濟學', '管理學']): 
        # 如果包含 "學" 且不是人文社科類的學，歸入科學 (例如: 材料科學, 宇宙學)
        # 這是一個廣泛的規則，可能會抓到一些漏網之魚
        is_science = 1
        
    # 4. 通識 (關鍵字判斷)
    elif any(k in name_upper for k in gen_keywords):
        is_general = 1
        
    # 5. 預設歸入通識
    else:
        is_general = 1

    return is_math, is_science, is_eng, is_general

# ==========================================
# 主程式
# ==========================================
def generate_classification_from_db():
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
    print("正在進行自動分類 (v3)...")

    results = []
    for _, row in df_courses.iterrows():
        m, s, e, g = classify_course_strict(row)
        results.append({
            '課號': row['課號'],
            '課程名稱': row['課程名稱'],
            'is_math': m,
            'is_science': s,
            'is_eng_prof': e,
            'is_general': g
        })
        
    df_out = pd.DataFrame(results)
    
    # 排序優化：先排通識，再排專業，方便檢查
    df_out = df_out.sort_values(
        by=['is_general', 'is_math', 'is_science', 'is_eng_prof', '課程名稱'], 
        ascending=[False, False, False, False, True]
    )
    
    print(f"正在寫入檔案: {OUTPUT_FILE}")
    df_out.to_excel(OUTPUT_FILE, index=False)
    print("完成！請打開 Excel 檢查並手動微調分類結果，然後覆蓋回 input_files 資料夾。")

if __name__ == "__main__":
    generate_classification_from_db()
    db.close()