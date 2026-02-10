import pandas as pd
from Accessdb import AccessHelper
import os
import random
from datetime import datetime
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
import warnings

# ==========================================
# 1. 環境設定
# ==========================================
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

db = AccessHelper()
today_str = datetime.today().strftime('%Y%m%d')

BASE_DIR = 'output_files'
SUB_DIR = '畢業生成績分析'
OUTPUT_DIR_PATH = os.path.join(BASE_DIR, SUB_DIR)

if not os.path.exists(OUTPUT_DIR_PATH):
    os.makedirs(OUTPUT_DIR_PATH)
    print(f"建立目錄: {OUTPUT_DIR_PATH}")

# ==========================================
# 2. 輔助函數
# ==========================================

def get_id_column(df):
    candidates = ['student_id', '學號', 'id', 'std_id']
    for col in df.columns:
        if col.lower() in candidates:
            return col
    # Fallback to column index 0 if not found, but it's risky.
    # Better to return None and let caller handle.
    return None

def find_col(df, keywords):
    for col in df.columns:
        if any(k in col.lower() for k in keywords):
            return col
    return None

def check_bool(val):
    return str(val).strip().lower() in ['1', 'true', 'yes']

# ==========================================
# 3. 資料載入與學生篩選
# ==========================================
print("讀取排名資料表...")

df_rank_u_all = pd.read_sql("SELECT * FROM GradRankU", db.conn)
df_rank_g_all = pd.read_sql("SELECT * FROM GradRankG", db.conn)

# 強制識別關鍵欄位
# U
id_col_u = 'StudentID' # 明確指定
ay_col_u = 'AcademicYear'
class_col = 'Class'
rank_col_u = 'Rank'
grade_col_u = 'Grade'
# GradRankU 通常有 TotalStudents 欄位嗎? 假設沒有則自動計算
total_col_u = find_col(df_rank_u_all, ['TotalStudents', '總人數'])

# G
id_col_g = 'StudentID'
ay_col_g = 'AcademicYear'
rank_col_g = 'Rank'
grade_col_g = 'Grade'
total_col_g = find_col(df_rank_g_all, ['TotalStudents', '總人數'])

# 篩選函式
def filter_students(df, ay_col, class_col, rank_col, total_col, is_undergrad=True):
    # 1. 學年 109-114
    target_years = [str(y) for y in range(109, 115)]
    df_filtered = df[df[ay_col].astype(str).isin(target_years)].copy()
    
    # 2. 班級 = '全' (僅大學部)
    if is_undergrad and class_col in df.columns:
        df_filtered = df_filtered[df_filtered[class_col] == '全']
        
    # 3. 計算總人數 (若無欄位)
    # 假設學號是唯一 key
    # 這裡的邏輯：如果沒有 total_col，我們假設 rank 的最大值就是總人數 (雖然不完全精確但可用)
    # 或者用 count
    if not total_col:
        # 計算該學年該班級的總人數
        # 簡單起見，我們用 groupby 學年來算 count
        counts = df_filtered.groupby(ay_col)[id_col_u if is_undergrad else id_col_g].count().reset_index()
        counts.rename(columns={ (id_col_u if is_undergrad else id_col_g) : 'calc_total'}, inplace=True)
        df_filtered = df_filtered.merge(counts, on=ay_col, how='left')
        total_col = 'calc_total'
        
    # 4. 篩選前 33%
    df_filtered[rank_col] = pd.to_numeric(df_filtered[rank_col], errors='coerce')
    df_filtered[total_col] = pd.to_numeric(df_filtered[total_col], errors='coerce')
    
    # 避免 total 為 0
    df_filtered = df_filtered[df_filtered[total_col] > 0]
    
    # Rank <= Total * 0.33
    df_top = df_filtered[df_filtered[rank_col] <= (df_filtered[total_col] * 0.33)].copy()
    
    return df_top, total_col

print("正在篩選大學部學生...")
df_u_pool, total_col_u = filter_students(df_rank_u_all, ay_col_u, class_col, rank_col_u, total_col_u, is_undergrad=True)

print("正在篩選碩士班學生...")
df_g_pool, total_col_g = filter_students(df_rank_g_all, ay_col_g, None, rank_col_g, total_col_g, is_undergrad=False)

# 隨機抽選
targets_u = {}
targets_g = {}

def sample_students(df, ay_col, id_col, rank_col, total_col, grade_col):
    result = {}
    sampled_years = []
    
    years = sorted(df[ay_col].unique())
    for year in years:
        pool = df[df[ay_col] == year]
        if not pool.empty:
            sample = pool.sample(n=min(2, len(pool)))
            
            records = []
            for _, row in sample.iterrows():
                rec = {
                    'sid': row[id_col],
                    'rank': row[rank_col],
                    'total': row[total_col],
                    'grad_year': row[ay_col],
                    'grad_grade': row[grade_col] if grade_col in row else None
                }
                records.append(rec)
            
            result[year] = records
            sampled_years.append(str(year))
            
    return result, sampled_years

u_data, u_years = sample_students(df_u_pool, ay_col_u, id_col_u, rank_col_u, total_col_u, grade_col_u)
g_data, g_years = sample_students(df_g_pool, ay_col_g, id_col_g, rank_col_g, total_col_g, grade_col_g)

print(f"大學部抽選學年: {u_years}, 共 {sum(len(v) for v in u_data.values())} 人")
print(f"碩士班抽選學年: {g_years}, 共 {sum(len(v) for v in g_data.values())} 人")

# 收集所有 ID
all_u_ids = [x['sid'] for yr in u_data for x in u_data[yr]]
all_g_ids = [x['sid'] for yr in g_data for x in g_data[yr]]
all_target_ids = all_u_ids + all_g_ids

if not all_target_ids:
    print("無符合條件學生，程式結束。")
    exit()

# ==========================================
# 4. 載入課程地圖與成績
# ==========================================
print("讀取外部課程分類表...")

# 讀取分類表
map_file = r'input_files\課程分類表\課程分類表1150127.xlsx'
df_ext_map = pd.DataFrame()

if os.path.exists(map_file):
    try:
        # 雖然檔名是 xlsx 但內容可能是 csv? 
        # 根據之前經驗，如果是 csv 內容存成 xlsx，pd.read_excel 可能會失敗
        # 先嘗試 read_excel
        try:
            df_ext_map = pd.read_excel(map_file)
        except:
            df_ext_map = pd.read_csv(map_file)
            
        # 標準化欄位名稱
        # 預期: course_name, is_math, is_science, is_eng_prof, is_general
        # 如果是中文: 課程名稱, 數學, 基礎科學, 工程專業, 通識
        # 統一轉小寫比對
        df_ext_map.columns = [c.strip().lower() for c in df_ext_map.columns]
        
        # 找出對應欄位
        name_col = next((c for c in df_ext_map.columns if 'name' in c or '名稱' in c), None)
        
        if name_col:
            df_ext_map.set_index(name_col, inplace=True)
        else:
            print("分類表找不到課程名稱欄位！")
            df_ext_map = pd.DataFrame() # Reset
            
    except Exception as e:
        print(f"讀取分類表失敗: {e}")
else:
    print(f"找不到分類表檔案: {map_file}")

# 讀取成績
print("讀取 STscore...")
df_scores_raw = pd.read_sql("SELECT * FROM STscore", db.conn)

# 找出成績表欄位
sc_id_col = find_col(df_scores_raw, ['學號', 'StudentID'])
sc_ay_col = find_col(df_scores_raw, ['學年度', 'Year'])
sc_sem_col = find_col(df_scores_raw, ['學期', 'Semester'])
sc_code_col = find_col(df_scores_raw, ['課號', 'Code'])
sc_name_col = find_col(df_scores_raw, ['名稱', 'Name', 'Title'])
sc_credit_col = find_col(df_scores_raw, ['學分', 'Credit'])

# 篩選
df_scores = df_scores_raw[df_scores_raw[sc_id_col].isin(all_target_ids)].copy()

# ==========================================
# 5. 分析與匯出
# ==========================================

def calculate_relative_grade(course_year, grad_year, grad_grade):
    try:
        diff = int(grad_year) - int(course_year)
        grade = int(grad_grade) - diff
        if grade < 1: return "Pre"
        return str(grade)
    except:
        return "?"

def get_course_category(c_name, df_map):
    # Default
    res = {'math': False, 'sci': False, 'eng': False, 'gen': False}
    
    if df_map.empty or pd.isna(c_name):
        return res
        
    c_name = str(c_name).strip()
    
    if c_name in df_map.index:
        info = df_map.loc[c_name]
        # Handle duplicates in map? loc might return DataFrame
        if isinstance(info, pd.DataFrame):
            info = info.iloc[0]
            
        # Check columns
        # Assuming map has columns like 'is_math', 'is_science'...
        # Or check content
        
        def is_true(k):
            # Try to find column containing k
            col = next((c for c in info.index if k in c), None)
            if col:
                return check_bool(info[col])
            return False

        res['math'] = is_true('math') or is_true('數學')
        res['sci'] = is_true('science') or is_true('基礎科學')
        res['eng'] = is_true('eng') or is_true('工程')
        res['gen'] = is_true('general') or is_true('通識')
        
    return res

missing_courses = set()

def process_student_data(s_rec, default_grade):
    sid = s_rec['sid']
    rank_str = f"{s_rec['rank']}/{s_rec['total']}"
    
    grad_year = s_rec['grad_year']
    grad_grade = s_rec['grad_grade'] if s_rec['grad_grade'] else default_grade
    
    # Filter student scores
    my_scores = df_scores[df_scores[sc_id_col] == sid].copy()
    my_scores = my_scores.sort_values(by=[sc_ay_col, sc_sem_col])
    
    sem_data = {} # Key: "1(1)" -> list of course strings
    stats = {'數學': 0, '基礎科學': 0, '工程專業': 0, '通識': 0, '總計': 0}
    
    for _, row in my_scores.iterrows():
        ay = row[sc_ay_col]
        sem = row[sc_sem_col]
        code = row[sc_code_col]
        name = row[sc_name_col]
        credit = row[sc_credit_col]
        
        try:
            credit_val = float(credit)
        except:
            credit_val = 0
            
        # Rel Grade
        rel_grade = calculate_relative_grade(ay, grad_year, grad_grade)
        sem_key = f"{rel_grade}({sem})"
        
        # Category
        cats = get_course_category(name, df_ext_map)
        
        if cats['math']: stats['數學'] += credit_val
        elif cats['sci']: stats['基礎科學'] += credit_val
        elif cats['eng']: stats['工程專業'] += credit_val
        elif cats['gen']: stats['通識'] += credit_val
        
        # Check missing
        if not any(cats.values()) and not df_ext_map.empty:
             # Only flag if map exists but course not found/categorized
             if str(name).strip() not in df_ext_map.index:
                 missing_courses.add(f"{code} {name}")
        
        stats['總計'] += credit_val
        
        # Format
        fmt = f"{name}({code}, {int(credit_val) if credit_val.is_integer() else credit_val})"
        
        if sem_key not in sem_data:
            sem_data[sem_key] = []
        sem_data[sem_key].append(fmt)
        
    return sem_data, stats, rank_str, sid

def write_report(data, filename, default_grade):
    full_path = os.path.join(OUTPUT_DIR_PATH, filename)
    print(f"寫入 Excel: {full_path}")
    
    with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
        for year in sorted(data.keys()):
            sheet_name = f"{year}學年"
            ws = writer.book.create_sheet(sheet_name)
            
            cursor = 1
            
            for s_rec in data[year]:
                sem_data, stats, rank_str, sid = process_student_data(s_rec, default_grade)
                
                # Header
                header = f"學號: {sid}  |  排名: {rank_str}"
                ws.cell(row=cursor, column=1, value=header).font = Font(bold=True, size=12)
                cursor += 1
                
                # Table Headers
                headers = ["年級(學期)", "課程(課號、學分)", "學分統計"]
                for i, h in enumerate(headers, 1):
                    cell = ws.cell(row=cursor, column=i, value=h)
                    cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                    cell.alignment = Alignment(horizontal='center')
                    cell.border = Border(bottom=Side(style='thin'))
                cursor += 1
                
                # Rows
                def sort_key(k):
                    try:
                        g, s = k.replace(')', '').split('(')
                        return (int(g), int(s))
                    except:
                        return (99, 99)
                
                sorted_sems = sorted(sem_data.keys(), key=sort_key)
                
                stat_text = (
                    f"總學分: {stats['總計']}\n"
                    f"數學: {stats['數學']}\n"
                    f"基礎科學: {stats['基礎科學']}\n"
                    f"工程專業: {stats['工程專業']}\n"
                    f"通識: {stats['通識']}"
                )
                
                first_row = True
                
                if not sorted_sems:
                    ws.cell(row=cursor, column=1, value="無修課資料")
                    cursor += 1
                
                for sem in sorted_sems:
                    courses_str = "、".join(sem_data[sem])
                    ws.cell(row=cursor, column=1, value=sem).alignment = Alignment(horizontal='center', vertical='center')
                    ws.cell(row=cursor, column=2, value=courses_str).alignment = Alignment(wrap_text=True, vertical='center')
                    
                    if first_row:
                        ws.cell(row=cursor, column=3, value=stat_text).alignment = Alignment(wrap_text=True, vertical='top')
                        first_row = False
                    
                    cursor += 1
                
                cursor += 2
            
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 80
            ws.column_dimensions['C'].width = 25
            
        if 'Sheet' in writer.book.sheetnames:
            writer.book.remove(writer.book['Sheet'])

# 匯出
write_report(u_data, f"大學部_畢業生成績分析_{today_str}.xlsx", 4)
write_report(g_data, f"碩士班_畢業生成績分析_{today_str}.xlsx", 2)

# 輸出缺失課程
if missing_courses:
    print("-" * 30)
    print(f"注意：有 {len(missing_courses)} 門課程在分類表中找不到。")
    with open(os.path.join(OUTPUT_DIR_PATH, "missing_courses.txt"), "w", encoding='utf-8') as f:
        f.write("\n".join(sorted(list(missing_courses))))
    print("清單已存至 missing_courses.txt")

db.close()
print("完成！")