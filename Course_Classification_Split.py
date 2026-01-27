import pandas as pd
import os

# ==========================================
# 檔案路徑設定
# ==========================================
# 輸入檔案 (學生成績檔，用來抓取所有開過的課程名稱)
input_file = r'input_files\學生成績\電機系109-113學年度大學部及碩士班博士班學生所有成績.xlsx'
# 輸出檔案 (新的分類表)
output_file = r'output_files\課程分類\課程分類表.xlsx'

def classify_course_split(row):
    """
    更新後的分類邏輯：將數學與科學分開
    """
    name = str(row['課程名稱']).strip()
    code = str(row['課號']).strip().upper()
    
    is_math = 0
    is_science = 0
    is_eng = 0
    is_general = 0
    
    # ---------------------------------------------------------
    # 1. 數學 (Mathematics)
    # ---------------------------------------------------------
    math_keywords = [
        '微積分', '工程數學', '線性代數', '機率', '統計', '微分方程', '複變', 
        '離散數學', '數值分析', '幾何', '代數',
        'Calculus', 'Engineering Math', 'Linear Algebra', 'Probability', 
        'Statistics', 'Differential Equations', 'Discrete Math', 'Numerical Analysis'
    ]
    if any(k in name for k in math_keywords):
        is_math = 1

    # ---------------------------------------------------------
    # 2. 基礎科學 (Basic Science)
    # ---------------------------------------------------------
    science_keywords = [
        '普通物理', '普通化學', '普通生物', '物理實驗', '化學實驗', '生物實驗', 
        '物理', '化學', '生物', '力學', '電磁學', # 電磁學有時算物理，但在電機系通常算工程，這邊需小心
        'Physics', 'Chemistry', 'Biology'
    ]
    # 特別排除：電磁學在電機系通常歸類為工程專業，而非普通物理
    # 若您希望 "電磁學" 算基礎科學，請移除下方的排除
    exclude_science = ['電磁學', 'Electromagnetics']

    if any(k in name for k in science_keywords):
        if not any(ex in name for ex in exclude_science):
            is_science = 1

    # ---------------------------------------------------------
    # 3. 工程專業 (Engineering Professional)
    # ---------------------------------------------------------
    eng_keywords = [
        '電路', '電子', '電磁', '訊號', '系統', '控制', '通訊', '電力', '電機', '程式', '邏輯', '半導體', 
        '晶片', '微處理', '網路', '資訊', '演算', '結構', '專題', '實驗', '實習', '工程', 'AI', '智慧', 
        '光電', '積體電路', '天線', '微波', '嵌入式', '物聯網', '機器學習', '類神經', 'FPGA', 'VLSI',
        'Circuit', 'Electronics', 'Signal', 'System', 'Control', 'Communication', 'Power', 'Program',
        'Semiconductor', 'Chip', 'Microprocessor', 'Network', 'Algorithm'
    ]
    
    exclude_eng = [
        '經濟', '管理', '社會', '心理', '法律', '會計', '文學', '藝術', '歷史', '文化', '哲學', '宗教', 
        '政治', '行銷', '財務', '金融', '商學', '法規', '憲法'
    ]
    eng_exceptions = ['工程經濟', '工程管理', '電信法規', '工程倫理']

    is_potential_eng = any(k in name for k in eng_keywords)
    has_exclude_word = any(k in name for k in exclude_eng)
    is_exception = any(k in name for k in eng_exceptions)

    if is_potential_eng:
        if has_exclude_word and not is_exception:
            is_eng = 0
        else:
            is_eng = 1

    # ---------------------------------------------------------
    # 4. 通識課程 (General Education)
    # ---------------------------------------------------------
    gen_keywords = [
        '國文', '英文', '日文', '外文', '語言', '歷史', '地理', '體育', '通識', '服務學習', '軍訓', '全民國防',
        '哲學', '藝術', '心理', '社會', '經濟', '法學', '憲法', '公民', '生涯', '導師', 'Chinese', 'English', 
        'History', 'Physical', 'General', 'Art', 'Music', 'Management', 'Economics', 'Law'
    ]
    
    if any(k in name for k in gen_keywords) or code.startswith('G') or code.startswith('C'):
        is_general = 1
    
    # [預設規則] 如果既不是數學、也不是科學、也不是工程，就歸類為通識
    if is_math == 0 and is_science == 0 and is_eng == 0:
        is_general = 1

    # [優先權規則] 若被歸類為工程，通常不算通識
    if is_eng == 1:
        is_general = 0

    return is_math, is_science, is_eng, is_general

def generate_split_classification_list():
    if not os.path.exists(input_file):
        print(f"錯誤：找不到輸入檔案: {input_file}")
        return

    print("讀取成績資料檔中...")
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"讀取失敗：{e}")
        return

    df.columns = [c.strip() for c in df.columns]
    
    if '課程名稱' not in df.columns:
        print("錯誤：找不到 '課程名稱' 欄位。")
        return

    print("正在分類課程 (數學/科學/工程/通識)...")
    unique_courses = df.drop_duplicates(subset=['課程名稱'])[['課號', '課程名稱']].copy()
    unique_courses = unique_courses.sort_values('課程名稱')
    
    results = []
    for _, row in unique_courses.iterrows():
        m, s, e, g = classify_course_split(row)
        results.append({
            'course_name': row['課程名稱'],
            'is_math': m,      # 新欄位
            'is_science': s,   # 新欄位
            'is_eng_prof': e,
            'is_general': g
        })
        
    df_out = pd.DataFrame(results)
    
    # 建立輸出資料夾
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        df_out.to_excel(output_file, index=False)
        print(f"\n新版分類表已產生！")
        print(f"包含欄位: is_math, is_science, is_eng_prof, is_general")
        print(f"檔案位於：{output_file}")
    except PermissionError:
        print(f"\n存檔失敗：請確認檔案 {output_file} 是否已被開啟。")

if __name__ == "__main__":
    generate_split_classification_list()