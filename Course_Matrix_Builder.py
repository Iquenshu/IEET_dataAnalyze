import pyodbc
import pandas as pd
import os

# ==========================================
# 1. 設定與規則定義
# ==========================================
db_path = 'IEETdatabase.accdb'

# 教育目標 (PEO) 判定規則
# 根據您的描述：只要具備列表中的 "任一項" 核心能力 (Intersection)，即符合該目標
peo_rules = {
    # 欄位名稱 (英) : [需要的核心能力 K1~K5]
    'PEO_Theory':   [1, 2, 4, 5],     # 學識理論
    'PEO_Tech':     [1, 2, 3],        # 專業技術
    'PEO_Team':     [3, 4],           # 團隊合作與工程倫理
    'PEO_Innov':    [1, 2, 3, 4, 5],  # 獨立思考與創新 (全部都算)
    'PEO_Global':   [4, 5]            # 國際視野
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
# 3. 建立資料表結構
# ==========================================
def init_matrix_table(cursor):
    # 為了確保欄位正確，這裡採用「重建」策略
    # 如果您想保留手動修改的表，請註解掉 DROP TABLE，並確保下方 INSERT 欄位名稱與您的一致
    try:
        cursor.execute("DROP TABLE Course_Matrix")
        cursor.commit()
        print("舊有的 Course_Matrix 已清除。")
    except:
        pass # 表不存在，忽略錯誤

    print("正在建立新資料表 Course_Matrix ...")
    
    # 建立符合 5 項教育目標的結構
    # K1~K5 代表核心能力
    sql = """
    CREATE TABLE Course_Matrix (
        id COUNTER CONSTRAINT PrimaryKey PRIMARY KEY,
        course_id LONG,
        academic_year LONG,
        semester LONG,
        course_code TEXT(50),
        course_name TEXT(255),
        
        has_K1 BIT, has_K2 BIT, has_K3 BIT, has_K4 BIT, has_K5 BIT,
        
        PEO_Theory BIT,   -- 學識理論
        PEO_Tech BIT,     -- 專業技術
        PEO_Team BIT,     -- 團隊合作與工程倫理
        PEO_Innov BIT,    -- 獨立思考與創新
        PEO_Global BIT    -- 國際視野
    )
    """
    cursor.execute(sql)
    cursor.commit()

# ==========================================
# 4. 主分析邏輯
# ==========================================
def build_matrix():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. 初始化資料表
    init_matrix_table(cursor)
    
    print("正在分析課程核心能力 (K1-K5)...")
    
    # 2. 抓取資料：Courses + Course_Competencies
    # 我們需要知道每一門課，到底勾選了哪些 SMC (評量方式)
    sql = """
        SELECT 
            C.id AS course_id, C.academic_year, C.semester, C.course_code, C.course_name,
            CC.competency_desc,
            CC.smc_0, CC.smc_1, CC.smc_2, CC.smc_3, CC.smc_4, 
            CC.smc_5, CC.smc_6, CC.smc_7, CC.smc_8, CC.smc_9, CC.smc_10
        FROM Courses AS C
        LEFT JOIN Course_Competencies AS CC ON C.id = CC.course_id
    """
    
    df = pd.read_sql(sql, conn)
    
    print(f"原始資料共 {len(df)} 筆能力細項，開始聚合運算...")
    
    # 用 Dictionary 來整合每一門課的數據 (Key = course_id)
    matrix_data = {} 
    
    for _, row in df.iterrows():
        c_id = row['course_id']
        
        # 初始化該課程 (如果尚未存在於 dict)
        if c_id not in matrix_data:
            matrix_data[c_id] = {
                'course_id': c_id,
                'academic_year': row['academic_year'],
                'semester': row['semester'],
                'course_code': row['course_code'],
                'course_name': row['course_name'],
                # 預設所有 K 能力為 False
                'has_K1': False, 'has_K2': False, 'has_K3': False, 
                'has_K4': False, 'has_K5': False
            }
        
        # 解析這一行代表哪一個 K (1~5)
        comp_desc = str(row['competency_desc']).strip()
        if not comp_desc or comp_desc == 'None':
            continue
            
        # 假設描述是以數字開頭 (例如 "1.整合...", "3.具備...")
        try:
            # 取第一個字元轉數字，並確保它是半形數字
            first_char = comp_desc[0]
            # 針對全形數字簡單轉換 (防止資料有 １. xxx)
            if first_char in '１２３４５':
                mapping = {'１':1, '２':2, '３':3, '４':4, '５':5}
                k_num = mapping[first_char]
            else:
                k_num = int(first_char)
        except:
            continue # 無法辨識編號，跳過
            
        # 檢查 SMC 是否有勾選 (Access True可能是 -1 或 1)
        has_assessment = False
        for i in range(11):
            val = row[f'smc_{i}']
            if val: # Truthy check
                has_assessment = True
                break
        
        # 如果有評量方式，且編號在 1~5 之間，標記為具備該能力
        if has_assessment and 1 <= k_num <= 5:
            matrix_data[c_id][f'has_K{k_num}'] = True

    # 3. 計算 PEO 符合度並寫入
    print("正在計算 5 大教育目標符合度...")
    
    insert_sql = """
        INSERT INTO Course_Matrix (
            course_id, academic_year, semester, course_code, course_name,
            has_K1, has_K2, has_K3, has_K4, has_K5,
            PEO_Theory, PEO_Tech, PEO_Team, PEO_Innov, PEO_Global
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    
    conn.autocommit = False # 開啟交易模式加速寫入
    count = 0
    
    for c_id, data in matrix_data.items():
        # A. 收集該課程具備的 K 列表 (例如 [1, 3])
        my_k_list = []
        for k in range(1, 6):
            if data[f'has_K{k}']:
                my_k_list.append(k)
        
        # B. 判定 PEO (交集邏輯：只要有任一項符合)
        # set(my_k_list) & set(rule_list) 會取出交集，如果有值轉 bool 就是 True
        is_theory = bool(set(my_k_list) & set(peo_rules['PEO_Theory']))
        is_tech   = bool(set(my_k_list) & set(peo_rules['PEO_Tech']))
        is_team   = bool(set(my_k_list) & set(peo_rules['PEO_Team']))
        is_innov  = bool(set(my_k_list) & set(peo_rules['PEO_Innov']))
        is_global = bool(set(my_k_list) & set(peo_rules['PEO_Global']))
        
        # C. 執行寫入
        cursor.execute(insert_sql, (
            data['course_id'], data['academic_year'], data['semester'], 
            data['course_code'], data['course_name'],
            data['has_K1'], data['has_K2'], data['has_K3'], 
            data['has_K4'], data['has_K5'],
            is_theory, is_tech, is_team, is_innov, is_global
        ))
        count += 1
        
    conn.commit()
    conn.close()
    print("-" * 30)
    print(f"矩陣表 (Course_Matrix) 建置完成！共處理 {count} 門課程。")
    print("欄位說明：")
    print("PEO_Theory : 學識理論")
    print("PEO_Tech   : 專業技術")
    print("PEO_Team   : 團隊合作與工程倫理")
    print("PEO_Innov  : 獨立思考與創新")
    print("PEO_Global : 國際視野")

if __name__ == "__main__":
    build_matrix()