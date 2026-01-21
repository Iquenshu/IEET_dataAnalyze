import pyodbc
import pandas as pd
import os
import warnings

# 忽略 SQLAlchemy 的警告
warnings.filterwarnings("ignore", category=UserWarning)

# ==========================================
# 1. 設定與規則定義
# ==========================================
db_path = 'IEETdatabase.accdb'

# [重要] 資料庫欄位對照表 (完全依照您提供的 Schema 設定)
col_map = {
    # 基本資料
    'course_id':     'course_id',
    'academic_year': 'academic_year',
    'semester':      'semester',
    'course_code':   'course_code',
    'course_name':   'course_name',
    
    # 平均分數 (依照您提供的 Schema)
    'course_score_AVG': 'course_score_AVG', 

    # 核心能力 (K1~K5)
    'K1': 'has_SO_K1',
    'K2': 'has_SO_K2',
    'K3': 'has_SO_K3',
    'K4': 'has_SO_K4',
    'K5': 'has_SO_K5',
    
    # 教育目標 (PEO)
    'PEO_Theory': 'is_PEO_Theory',      # 學識理論
    'PEO_Tech':   'is_PEO_Skill',       # 專業技術
    'PEO_Team':   'is_PEO_Ethics',      # 團隊合作與工程倫理
    'PEO_Innov':  'is_PEO_innovation',  # 獨立思考與創新
    'PEO_Global': 'is_PEO_Global'       # 國際視野
}

# 教育目標判定規則 (任一符合即算符合)
peo_rules = {
    'PEO_Theory': [1, 2, 4, 5],
    'PEO_Tech':   [1, 2, 3],
    'PEO_Team':   [3, 4],
    'PEO_Innov':  [1, 2, 3, 4, 5],
    'PEO_Global': [4, 5]
}

# ==========================================
# 2. 資料庫連線
# ==========================================
def get_db_connection():
    full_db_path = os.path.abspath(db_path)
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        rf'DBQ={full_db_path};'
    )
    return pyodbc.connect(conn_str)

# ==========================================
# 3. 功能模組
# ==========================================

def calculate_course_averages(conn):
    """
    從 STscore 計算每門課的平均分
    規則：排除 '退選' 與 分數 999
    """
    print("正在計算各課程平均分數 (排除退選)...")
    
    # 讀取成績資料
    sql = "SELECT [學年度], [學期], [課號], [成績], [等第成績] FROM STscore"
    try:
        df_score = pd.read_sql(sql, conn)
    except Exception as e:
        print(f"讀取 STscore 失敗: {e}")
        return pd.DataFrame() 

    # 資料清洗
    df_score['成績'] = pd.to_numeric(df_score['成績'], errors='coerce')
    
    # 排除無效成績 (999 或 退選)
    # 注意：這裡將所有相關欄位轉為字串再去除空白，確保比對準確
    df_valid = df_score[
        (df_score['成績'] != 999) & 
        (df_score['等第成績'].astype(str).str.strip() != '退選') &
        (df_score['成績'].notna())
    ].copy()

    if len(df_valid) == 0:
        print("警告：沒有有效的成績資料可供計算。")
        return pd.DataFrame()

    # 型別標準化 (確保能跟 Courses 表對上)
    try:
        # 將學年度/學期轉為整數，去除 .0
        df_valid['學年度'] = pd.to_numeric(df_valid['學年度'], errors='coerce').fillna(0).astype(int)
        df_valid['學期'] = pd.to_numeric(df_valid['學期'], errors='coerce').fillna(0).astype(int)
        df_valid['課號'] = df_valid['課號'].astype(str).str.strip()
    except Exception as e:
        print(f"型別轉換錯誤: {e}")

    # 分組計算平均
    avg_df = df_valid.groupby(['學年度', '學期', '課號'])['成績'].mean().reset_index()
    avg_df.rename(columns={'成績': 'avg_score'}, inplace=True)
    avg_df['avg_score'] = avg_df['avg_score'].round(2)
    
    print(f"已計算 {len(avg_df)} 門課程的平均成績。")
    return avg_df

def build_matrix():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. 檢查資料表結構是否存在 (防呆)
    try:
        # 測試讀取關鍵欄位
        check_col = col_map['K1'] 
        cursor.execute(f"SELECT TOP 1 [{check_col}] FROM Course_Matrix")
    except Exception as e:
        print("錯誤：無法存取 Course_Matrix 資料表或欄位名稱不符。")
        print(f"請確認 Access 資料表 [Course_Matrix] 中是否有 [{check_col}] 這個欄位？")
        print(f"系統錯誤訊息: {e}")
        return

    # 2. 清空舊數據 (保留結構，不刪除表)
    print("正在清空 Course_Matrix 舊有數據...")
    try:
        cursor.execute("DELETE FROM Course_Matrix")
        cursor.commit()
    except Exception as e:
        print(f"清除資料失敗: {e}")
        return

    # 3. 取得平均分數表
    df_avgs = calculate_course_averages(conn)

    # 4. 分析課程核心能力 (K1-K5)
    print("正在整合課程、能力指標與平均分數...")
    sql_courses = """
        SELECT 
            C.id AS course_id, C.academic_year, C.semester, C.course_code, C.course_name,
            CC.competency_desc,
            CC.smc_0, CC.smc_1, CC.smc_2, CC.smc_3, CC.smc_4, 
            CC.smc_5, CC.smc_6, CC.smc_7, CC.smc_8, CC.smc_9, CC.smc_10
        FROM Courses AS C
        LEFT JOIN Course_Competencies AS CC ON C.id = CC.course_id
    """
    try:
        df_courses = pd.read_sql(sql_courses, conn)
    except Exception as e:
        print(f"讀取 Courses 失敗: {e}")
        return
    
    # 5. 聚合運算 (Aggregation)
    matrix_data = {} 
    
    print(f"正在處理 {len(df_courses)} 筆課程能力資料...")
    
    for _, row in df_courses.iterrows():
        c_id = row['course_id']
        
        # 初始化
        if c_id not in matrix_data:
            matrix_data[c_id] = {
                'course_id': c_id,
                'academic_year': int(row['academic_year']),
                'semester': int(row['semester']),
                'course_code': str(row['course_code']).strip(),
                'course_name': row['course_name'],
                # K1~K5
                1: False, 2: False, 3: False, 4: False, 5: False,
                # 平均分 (初始化為 None)
                'course_score_AVG': None
            }
        
        # 處理 K 能力
        comp_desc = str(row['competency_desc']).strip()
        if comp_desc and comp_desc != 'None':
            k_num = 0
            try:
                # 嘗試解析開頭數字
                first_char = comp_desc[0]
                if first_char in '１２３４５': # 全形
                    k_num = {'１':1, '２':2, '３':3, '４':4, '５':5}[first_char]
                elif first_char.isdigit(): 
                    k_num = int(first_char)
            except: 
                pass
            
            # 檢查 SMC 是否有勾選 (Access CheckBox: True/-1/1)
            has_assessment = any(row[f'smc_{i}'] for i in range(11))
            
            if has_assessment and 1 <= k_num <= 5:
                matrix_data[c_id][k_num] = True

    # 6. 合併平均分數
    # 建立快速查詢表: (year, sem, code) -> score
    avg_lookup = {}
    if not df_avgs.empty:
        for _, row in df_avgs.iterrows():
            key = (row['學年度'], row['學期'], row['課號'])
            avg_lookup[key] = row['avg_score']

    # 填入平均分
    for c_id, data in matrix_data.items():
        key = (data['academic_year'], data['semester'], data['course_code'])
        if key in avg_lookup:
            data['course_score_AVG'] = avg_lookup[key]

    # 7. 寫入資料庫
    print("正在寫入資料庫 Course_Matrix ...")
    
    # 準備欄位順序 (排除 matrix_id，因為它是自動編號)
    columns = [
        col_map['course_id'], col_map['academic_year'], col_map['semester'], 
        col_map['course_code'], col_map['course_name'], col_map['course_score_AVG'],
        col_map['K1'], col_map['K2'], col_map['K3'], col_map['K4'], col_map['K5'],
        col_map['PEO_Theory'], col_map['PEO_Tech'], col_map['PEO_Team'], 
        col_map['PEO_Innov'], col_map['PEO_Global']
    ]
    
    # 加上 [] 保護欄位名稱
    safe_columns = [f"[{c}]" for c in columns]
    placeholders = ",".join(["?"] * len(columns))
    insert_sql = f"INSERT INTO Course_Matrix ({','.join(safe_columns)}) VALUES ({placeholders})"
    
    conn.autocommit = False
    count = 0
    
    try:
        for c_id, data in matrix_data.items():
            # K 列表
            my_k_list = [k for k in range(1, 6) if data[k]]
            
            # PEO 判定 (交集)
            peo_results = {}
            for peo_key, req_k in peo_rules.items():
                peo_results[peo_key] = bool(set(my_k_list) & set(req_k))
            
            # 參數準備 (確保順序與 columns 一致)
            params = (
                data['course_id'], data['academic_year'], data['semester'], 
                data['course_code'], data['course_name'], data['course_score_AVG'],
                data[1], data[2], data[3], data[4], data[5],
                peo_results['PEO_Theory'], peo_results['PEO_Tech'], 
                peo_results['PEO_Team'], peo_results['PEO_Innov'], 
                peo_results['PEO_Global']
            )
            
            cursor.execute(insert_sql, params)
            count += 1
            
        conn.commit()
        print("-" * 30)
        print(f"成功更新 {count} 筆資料 (含平均分數、能力指標、教育目標)。")
        
    except Exception as e:
        conn.rollback()
        print(f"寫入過程發生錯誤: {e}")
        # 進階除錯資訊
        import traceback
        traceback.print_exc()
    
    conn.close()

if __name__ == "__main__":
    if os.path.exists(db_path):
        build_matrix()
    else:
        print(f"找不到資料庫: {db_path}")