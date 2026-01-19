import pandas as pd
import os

# ==========================================
# 檔案路徑設定
# ==========================================
# 輸入檔案
input_file = r'input_files\學生成績\電機系109-113學年度大學部及碩士班博士班學生所有成績.xlsx'
# 輸出檔案
output_file = r'output_files\課程分類\課程分類表.xlsx'

def classify_course_strict(row):
    """
    根據電機系/工學院觀點進行嚴格分類
    原則：非電機/工程相關專業，一律歸類為通識
    """
    name = str(row['課程名稱']).strip()
    code = str(row['課號']).strip().upper()
    
    is_math = 0
    is_eng = 0
    is_general = 0
    
    # ---------------------------------------------------------
    # 1. 數學及基礎科學 (Math & Basic Science)
    # ---------------------------------------------------------
    # 僅包含工學院認可的基礎學科
    math_keywords = [
        '微積分', '普通物理', '普通化學', '普通生物', '工程數學', '線性代數', '機率', '統計', 
        '微分方程', '複變', '離散數學', '數值分析', '物理實驗', '化學實驗', '生物實驗',
        'Calculus', 'Physics', 'Chemistry', 'Biology', 'Engineering Math', 'Linear Algebra', 
        'Probability', 'Statistics', 'Differential Equations', 'Discrete Math'
    ]
    # 排除像是 "數學欣賞", "物理與生活" 這類通識課 (如果有需要可加負面表列)
    if any(k in name for k in math_keywords):
        is_math = 1

    # ---------------------------------------------------------
    # 2. 工程專業 (Engineering Professional)
    # ---------------------------------------------------------
    # 正面表列：電機與工程常用詞彙
    eng_keywords = [
        '電路', '電子', '電磁', '訊號', '系統', '控制', '通訊', '電力', '電機', '程式', '邏輯', '半導體', 
        '晶片', '微處理', '網路', '資訊', '演算', '結構', '專題', '實驗', '實習', '工程', 'AI', '智慧', 
        '光電', '積體電路', '天線', '微波', '嵌入式', '物聯網', '機器學習', '類神經', 'FPGA', 'VLSI',
        'Circuit', 'Electronics', 'Signal', 'System', 'Control', 'Communication', 'Power', 'Program',
        'Semiconductor', 'Chip', 'Microprocessor', 'Network', 'Algorithm'
    ]
    
    # 負面表列：看起來像專業但非工學院專業 (若出現這些詞，取消工程分類)
    # 例如：管理系統(Management)、經濟分析(Economics)、法律系統(Law)
    exclude_keywords = [
        '經濟', '管理', '社會', '心理', '法律', '會計', '文學', '藝術', '歷史', '文化', '哲學', '宗教', 
        '政治', '行銷', '財務', '金融', '商學', '法規', '憲法'
    ]
    
    # 例外豁免：雖然有負面詞，但確實是工程課
    eng_exceptions = ['工程經濟', '工程管理', '電信法規', '工程倫理']

    # 判斷邏輯
    is_potential_eng = any(k in name for k in eng_keywords)
    has_exclude_word = any(k in name for k in exclude_keywords)
    is_exception = any(k in name for k in eng_exceptions)

    if is_potential_eng:
        if has_exclude_word and not is_exception:
            is_eng = 0 # 雖有工程關鍵字，但包含排除詞且非例外 -> 不是工程
        else:
            is_eng = 1

    # ---------------------------------------------------------
    # 3. 通識課程 (General Education)
    # ---------------------------------------------------------
    # 明確通識關鍵字
    gen_keywords = [
        '國文', '英文', '日文', '外文', '語言', '歷史', '地理', '體育', '通識', '服務學習', '軍訓', '全民國防',
        '哲學', '藝術', '心理', '社會', '經濟', '法學', '憲法', '公民', '生涯', '導師', 'Chinese', 'English', 
        'History', 'Physical', 'General', 'Art', 'Music', 'Management', 'Economics', 'Law'
    ]
    
    # 條件 A: 包含通識關鍵字
    # 條件 B: 課號以 G, C, A 開頭 (常見通識代碼，視學校而定)
    # 條件 C: **最重要** 既不是數理，也不是工程，就預設歸類為通識
    
    if any(k in name for k in gen_keywords) or code.startswith('G') or code.startswith('C'):
        is_general = 1
    
    # [關鍵修正] 預設歸類為通識 (Default to General)
    # 如果它完全沒被分類到數理或工程，它就是通識 (包含外系的專業課)
    if is_math == 0 and is_eng == 0:
        is_general = 1

    # [防呆] 如果被歸類為工程，通常就不算通識 (除非是跨領域通識，這裡先從嚴認定)
    if is_eng == 1:
        is_general = 0

    return is_math, is_eng, is_general

def generate_strict_classification_list():
    if not os.path.exists(input_file):
        print(f"錯誤：找不到輸入檔案: {input_file}")
        return

    print("讀取成績資料檔中 (Excel模式)...")
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"讀取失敗：{e}")
        return

    df.columns = [c.strip() for c in df.columns]
    
    # 確保有'課程名稱'欄位
    if '課程名稱' not in df.columns:
        print("錯誤：找不到 '課程名稱' 欄位。")
        return

    # 取出不重複課程
    print("正在擷取並分類課程...")
    unique_courses = df.drop_duplicates(subset=['課程名稱'])[['課號', '課程名稱']].copy()
    unique_courses = unique_courses.sort_values('課程名稱')
    
    results = []
    for _, row in unique_courses.iterrows():
        m, e, g = classify_course_strict(row)
        results.append({
            'course_name': row['課程名稱'],
            'is_math_science': m,
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
        print(f"\n分類表 (嚴格版) 已成功產生！")
        print(f"檔案位於：{output_file}")
        print("-" * 30)

    except PermissionError:
        print(f"\n存檔失敗：請確認檔案 {output_file} 是否已被開啟。")

if __name__ == "__main__":
    generate_strict_classification_list()