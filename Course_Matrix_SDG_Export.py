import pandas as pd
from Accessdb import AccessHelper
import os
from datetime import datetime
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
import warnings

# [修正] 忽略 Pandas 與 SQLAlchemy 的警告訊息，讓輸出畫面乾淨
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# ==========================================
# 1. 設定與準備
# ==========================================
db = AccessHelper()
today_str = datetime.today().strftime('%Y%m%d')

# [設定] 檔案輸出路徑
BASE_DIR = 'output_files'
SUB_DIR = '課程核心能力及SDG關聯表'
OUTPUT_DIR_PATH = os.path.join(BASE_DIR, SUB_DIR)

if not os.path.exists(OUTPUT_DIR_PATH):
    os.makedirs(OUTPUT_DIR_PATH)
    print(f"已建立資料夾: {OUTPUT_DIR_PATH}")

# 定義完整名稱字典
k_mapping = {
    'has_SO_K1': 'K1 能夠整合、組織電機專業理論來分析、表達問題之能力',
    'has_SO_K2': 'K2 能夠運用電機專業知識解決及實作電機工程問題之能力',
    'has_SO_K3': 'K3 具備分工、協調、重視團隊合作精神、遵守工程倫理以達成工作目標之能力',
    'has_SO_K4': 'K4 能夠激發自己潛能、融合他人智慧，具備獨立思考以及研究創新之能力',
    'has_SO_K5': 'K5 具備吸收電機新知、掌握國際發展趨勢，隨時接受競爭挑戰之能力'
}

sdg_mapping = {
    'sdg_1': 'SDG1 消除貧窮',
    'sdg_2': 'SDG2 消除飢餓',
    'sdg_3': 'SDG3 良好健康與福祉',
    'sdg_4': 'SDG4 教育品質',
    'sdg_5': 'SDG5 性別平等',
    'sdg_6': 'SDG6 乾淨水源與公共衛生',
    'sdg_7': 'SDG7 可負擔乾淨能源',
    'sdg_8': 'SDG8 優質工作與經濟成長',
    'sdg_9': 'SDG9 工業、創新和基礎建設',
    'sdg_10': 'SDG10 減少不平等',
    'sdg_11': 'SDG11 永續城市',
    'sdg_12': 'SDG12 責任消費與生產',
    'sdg_13': 'SDG13 氣候行動',
    'sdg_14': 'SDG14 海洋生態',
    'sdg_15': 'SDG15 陸域生態',
    'sdg_16': 'SDG16 和平、正義和穩健的制度',
    'sdg_17': 'SDG17 促進目標實現的全球夥伴關係'
}

k_cols = list(k_mapping.keys())
sdg_cols = list(sdg_mapping.keys())

# ==========================================
# 2. 資料讀取
# ==========================================
print("正在讀取並整合資料庫 (Courses + Matrix + SDGs)...")

sql = '''
SELECT 
    C.academic_year, C.semester, C.dept_code,
    C.course_code, C.course_name, C.is_required, C.credits, C.instructor,
    M.has_SO_K1, M.has_SO_K2, M.has_SO_K3, M.has_SO_K4, M.has_SO_K5,
    S.sdg_1, S.sdg_2, S.sdg_3, S.sdg_4, S.sdg_5, S.sdg_6, 
    S.sdg_7, S.sdg_8, S.sdg_9, S.sdg_10, S.sdg_11, S.sdg_12, 
    S.sdg_13, S.sdg_14, S.sdg_15, S.sdg_16, S.sdg_17
FROM (Courses AS C
LEFT JOIN Course_Matrix AS M ON C.id = M.course_id)
LEFT JOIN Course_SDGs AS S ON C.id = S.course_id
ORDER BY C.academic_year DESC, C.semester, C.course_code
'''

df_raw = pd.read_sql(sql, db.conn)

# 排除第3學期
df_raw = df_raw[df_raw['semester'] != 3]

# ==========================================
# 3. 資料處理函式
# ==========================================

def format_check_numeric(val):
    """
    格式化檢查：轉換為數字 1 或 0
    """
    if pd.isna(val):
        return 0
    try:
        if isinstance(val, (bool, int, float)):
            return 1 if int(val) == 1 else 0
        
        s_val = str(val).lower().strip()
        if s_val in ['1', 'true', '1.0', 'yes']:
            return 1
        return 0
    except:
        return 0

def get_required_str(val):
    """必選修轉換"""
    try:
        if isinstance(val, bool):
            return "必修" if val else "選修"
        s_val = str(val).lower().strip()
        if s_val in ['1', 'true', '1.0', 'yes']:
            return "必修"
        return "選修"
    except:
        return "選修"

def process_data_for_table(df_sem, target_cols, mapping_dict):
    """
    處理主表資料：
    1. 轉換 1/0
    2. [修正] 排除全 0 的列 (row_sum == 0 則不加入 output)
    """
    output_rows = []
    
    # 目標顯示欄位
    display_labels = [mapping_dict[c] for c in target_cols]
    
    # 統計用計數器
    totals = {label: 0 for label in display_labels}
    
    for _, row in df_sem.iterrows():
        # 先計算該列在目標欄位的總分
        row_sum = 0
        current_vals = []
        for col_db in target_cols:
            val = format_check_numeric(row[col_db])
            row_sum += val
            current_vals.append(val)
            
        # [關鍵修正] 如果這門課在這個表中完全沒有勾選 (總分為0)，則跳過，不顯示在主表
        if row_sum == 0:
            continue

        # 建立 Excel 資料列
        excel_row = {
            '學年': row['academic_year'],
            '學期': row['semester'],
            '課程名稱及課號': f"{row['course_name']} ({row['course_code']})",
            '必選修': get_required_str(row['is_required']),
            '學分': row['credits']
        }
        
        # 填入 1/0 並累加總計
        for val, col_label in zip(current_vals, display_labels):
            excel_row[col_label] = val
            totals[col_label] += val
                
        output_rows.append(excel_row)
        
    # 轉為 DataFrame
    df_out = pd.DataFrame(output_rows)
    
    # 確保欄位順序 (防呆：若全空則建立空表)
    base_cols = ['學年', '學期', '課程名稱及課號', '必選修', '學分']
    if not df_out.empty:
        df_out = df_out[base_cols + display_labels]
    else:
        df_out = pd.DataFrame(columns=base_cols + display_labels)
    
    return df_out, totals, display_labels

def get_missing_string(df_sem, check_cols):
    """
    找出全為 0 的課程，並回傳逗號分隔的字串
    格式: 課名 (課號), 課名 (課號)...
    """
    missing_items = []
    
    for _, row in df_sem.iterrows():
        row_sum = 0
        for col in check_cols:
            row_sum += format_check_numeric(row[col])
        
        if row_sum == 0:
            # [修正] 只取課名與課號
            item = f"{row['course_name']} ({row['course_code']})"
            missing_items.append(item)
    
    if not missing_items:
        return "(無)"
        
    # 用逗號連接
    return ", ".join(missing_items)

# ==========================================
# 4. Excel 寫入函式 (含美化)
# ==========================================
def write_table_block(ws, df_data, totals, labels, start_row, header_color):
    """
    寫入主表 (包含標題、資料、總計列)
    """
    # 1. 寫入資料
    rows = list(dataframe_to_rows(df_data, index=False, header=True))
    
    for r_idx, row in enumerate(rows):
        current_excel_row = start_row + r_idx
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=current_excel_row, column=c_idx, value=value)
            
            # 樣式
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            # 標題列樣式
            if r_idx == 0:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
                ws.row_dimensions[current_excel_row].height = 60 # 標題高一點方便換行
    
    # 2. 寫入「總計」列
    last_row = start_row + len(df_data) + 1 
    ws.cell(row=last_row, column=1, value="總計").font = Font(bold=True)
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
    
    # 總計標題格式
    cell_total_title = ws.cell(row=last_row, column=1)
    cell_total_title.alignment = Alignment(horizontal='center')
    cell_total_title.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    cell_total_title.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # 填入統計值
    start_col_idx = 6 # 從第6欄開始是動態欄位
    for i, label in enumerate(labels):
        col_idx = start_col_idx + i
        val = totals[label]
        cell = ws.cell(row=last_row, column=col_idx, value=val)
        
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 回傳下一個表格的起始位置 (空3行)
    return last_row + 4

def write_missing_list_section(ws, missing_k_str, missing_sdg_str, start_row):
    """
    寫入未關聯清單 (兩列：K 與 SDG)
    """
    # --- 列 1: 核心能力 ---
    ws.cell(row=start_row, column=1, value="未關聯任何核心能力之課程").font = Font(bold=True)
    cell_title_k = ws.cell(row=start_row, column=1)
    cell_title_k.alignment = Alignment(horizontal='left', vertical='center') # 改垂直置中
    cell_title_k.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    # 內容 (合併欄位顯示長字串)
    cell_content_k = ws.cell(row=start_row, column=2, value=missing_k_str)
    cell_content_k.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=15) # 合併多一點欄位
    
    # 自動調整列高 (簡單估算：字串長度 / 概略寬度)
    # 這裡給一個基本高度，如果字串太長 Excel 打開時 wrap_text 會自動處理顯示
    ws.row_dimensions[start_row].height = 30
    
    # --- 列 2: SDGs ---
    next_row = start_row + 1
    ws.cell(row=next_row, column=1, value="未關聯任何 SDGs 之課程").font = Font(bold=True)
    cell_title_sdg = ws.cell(row=next_row, column=1)
    cell_title_sdg.alignment = Alignment(horizontal='left', vertical='center')
    cell_title_sdg.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    # 內容
    cell_content_sdg = ws.cell(row=next_row, column=2, value=missing_sdg_str)
    cell_content_sdg.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells(start_row=next_row, start_column=2, end_row=next_row, end_column=15)
    
    ws.row_dimensions[next_row].height = 30
    
    return next_row + 2

def set_column_widths(ws, label_count):
    ws.column_dimensions['A'].width = 8  # 學年
    ws.column_dimensions['B'].width = 6  # 學期
    ws.column_dimensions['C'].width = 40 # 課程名稱
    ws.column_dimensions['D'].width = 10 # 必選修
    ws.column_dimensions['E'].width = 6  # 學分
    
    # 動態欄位寬度
    for i in range(label_count):
        col_letter = chr(64 + 6 + i)
        # 超過 Z 的欄位處理
        if 6 + i > 26:
            first = chr(64 + (6 + i - 1) // 26)
            second = chr(64 + (6 + i - 1) % 26 + 1)
            col_letter = first + second
            
        ws.column_dimensions[col_letter].width = 15

# ==========================================
# 5. 主流程
# ==========================================
def process_dept_export(df_dept, dept_name):
    if df_dept.empty:
        print(f"{dept_name} 無資料。")
        return

    output_filename = f'{dept_name}_課程核心能力及SDG關聯表_{today_str}.xlsx'
    full_path = os.path.join(OUTPUT_DIR_PATH, output_filename)
    
    print(f"正在建立檔案: {full_path}")
    
    with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
        
        # 排序
        df_dept['sort_key'] = df_dept['academic_year'] * 10 + df_dept['semester']
        unique_sems = df_dept[['academic_year', 'semester', 'sort_key']].drop_duplicates().sort_values('sort_key', ascending=False)
        
        for _, row in unique_sems.iterrows():
            year = int(row['academic_year'])
            sem = int(row['semester'])
            
            sheet_name = f"{year}-{sem}"
            
            # 篩選學期資料
            df_sem = df_dept[
                (df_dept['academic_year'] == year) & 
                (df_dept['semester'] == sem)
            ].copy()
            
            # 排序：必修在前
            df_sem = df_sem.sort_values(by=['is_required', 'course_code'], ascending=[False, True])
            
            ws = writer.book.create_sheet(sheet_name)
            
            current_row = 1
            
            # --- 表格 1: 核心能力 (橘色) ---
            df_k, totals_k, labels_k = process_data_for_table(df_sem, k_cols, k_mapping)
            current_row = write_table_block(ws, df_k, totals_k, labels_k, current_row, "ED7D31")
            
            # --- 表格 2: SDGs (綠色) ---
            df_sdg, totals_sdg, labels_sdg = process_data_for_table(df_sem, sdg_cols, sdg_mapping)
            current_row = write_table_block(ws, df_sdg, totals_sdg, labels_sdg, current_row, "70AD47")
            
            # --- 下方清單: 未關聯課程 ---
            missing_k_str = get_missing_string(df_sem, k_cols)
            missing_sdg_str = get_missing_string(df_sem, sdg_cols)
            
            current_row += 1 # 增加間隔
            write_missing_list_section(ws, missing_k_str, missing_sdg_str, current_row)
            
            # 設定欄寬
            set_column_widths(ws, 17)
            
        if 'Sheet' in writer.book.sheetnames:
            writer.book.remove(writer.book['Sheet'])
            
    print(f"匯出成功！")

# 執行
if not df_raw.empty:
    # 資料清理
    df_raw['dept_code'] = df_raw['dept_code'].astype(str).str.strip()
    
    # 大學部
    df_undergrad = df_raw[df_raw['dept_code'] == 'B301'].copy()
    process_dept_export(df_undergrad, "大學部")
    
    # 碩士班
    df_grad = df_raw[df_raw['dept_code'] == 'M301'].copy()
    process_dept_export(df_grad, "碩士班")
else:
    print("資料庫無任何課程資料。")

db.close()
print("-" * 30)
print(f"全部完成！請檢查: {OUTPUT_DIR_PATH}")