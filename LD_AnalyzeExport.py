import pandas as pd
# import matplotlib.pyplot as plt  # (已註解) 暫不使用圖表
from Accessdb import AccessHelper
# from openpyxl.drawing.image import Image as XLImage # (已註解)
import os
import uuid
from datetime import datetime
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================
# 1. 設定與準備
# ==========================================
db = AccessHelper()
today_str = datetime.today().strftime('%Y%m%d')

# ---------------------------------------------------------
# [設定] 檔案輸出路徑
# ---------------------------------------------------------
# 您可以在這裡修改輸出的主資料夾與子資料夾名稱
BASE_DIR = 'output_files'          # 主資料夾
SUB_DIR = '離系問券統計'           # 子資料夾 (將檔案放在這裡)

# 組合完整路徑: output_files\離系問券統計
OUTPUT_DIR_PATH = os.path.join(BASE_DIR, SUB_DIR)

# 確保資料夾存在 (如果沒有會自動建立)
if not os.path.exists(OUTPUT_DIR_PATH):
    os.makedirs(OUTPUT_DIR_PATH)
    print(f"已建立資料夾: {OUTPUT_DIR_PATH}")
# ---------------------------------------------------------

# 讀取資料庫
print("正在讀取資料庫...")
df_u = pd.read_sql("SELECT * FROM LDUdataAnalyze", db.conn)
df_g = pd.read_sql("SELECT * FROM LDGdataAnalyze", db.conn)
df_q = pd.read_sql("SELECT * FROM LeavDepQuest", db.conn)

def get_quest_map(row):
    """解析問卷題目對照表"""
    big_map = {}
    small_map = {}
    # 假設大題是 A10, A20, A30
    for i in range(1, 4):
        big_key = f"A{i}0"
        if big_key in row:
            big_map[big_key] = row[big_key]
        # 假設小題是 A11~A19...
        for j in range(1, 12):
            small_key = f"A{i}{j}"
            if small_key in row:
                small_map[small_key] = row[small_key]
    return big_map, small_map

# 取得題目對照
quest_row_u = df_q[df_q['QuestType'] == 'UdataZH'].iloc[0]
quest_row_g = df_q[df_q['QuestType'] == 'GdataZH'].iloc[0]
big_map_u, small_map_u = get_quest_map(quest_row_u)
big_map_g, small_map_g = get_quest_map(quest_row_g)

# ==========================================
# 2. Excel 寫入核心邏輯
# ==========================================
def write_sheet_data(ws, df_source, big_map, small_map, startrow=0, label=""):
    """
    將 DataFrame 資料寫入指定的 Worksheet
    """
    # 寫入標題 (例如: 109學年總和)
    ws.cell(row=startrow+1, column=1, value=label)
    startrow += 1
    
    # [修改] 新的表頭名稱
    header = ['題目', '完全未達成', '小部分達成', '部分達成', '大部分達成', '完全達成', '總數']

    # 遍歷大題 (A10, A20...)
    for big_key, big_title in big_map.items():
        # 跳過 A30 (通常是簡答題)
        if big_key == "A30":
            continue
            
        # 寫入大題標題
        ws.cell(row=startrow+1, column=1, value=f"【{big_title}】")
        startrow += 1
        
        # 找出該大題下的所有小題
        qids = [k for k in small_map if k.startswith(big_key[:2]) and not k.startswith("A30")]
        
        table_data = []
        for qid in qids:
            # 篩選該題目的數據
            row = df_source[df_source['qid'] == qid]
            if row.empty:
                continue
            
            # 取第一筆 (因為經過篩選或加總後應該只有一筆)
            row_data = row.iloc[0]
            
            table_data.append([
                small_map.get(qid, qid),
                row_data['count_1'],
                row_data['count_2'],
                row_data['count_3'],
                row_data['count_4'],
                row_data['count_5'],
                row_data['total']
            ])
            
        if table_data:
            # 轉為 DataFrame 以便寫入
            table_df = pd.DataFrame(table_data, columns=header)
            
            # 使用 openpyxl 逐行寫入 (包含表頭)
            for r_idx, row in enumerate(dataframe_to_rows(table_df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=startrow + r_idx, column=c_idx, value=value)
            
            # 更新下一個表格的起始列 (資料列數 + 標題列 + 間隔)
            startrow += len(table_data) + 3
            
    return startrow

def export_excel(df, big_map, small_map, filename, sheet_suffix):
    # 組合完整檔案路徑
    full_output_path = os.path.join(OUTPUT_DIR_PATH, filename)
    print(f"準備寫入檔案: {full_output_path}")

    with pd.ExcelWriter(full_output_path) as writer:
        
        # --- [分頁 1] 全部學年統計 (109-113總計) ---
        # 邏輯：排除 'T' 結尾的列 (避免重複計算)，將所有原始學期資料加總
        df_raw = df[~df['sem'].str.endswith('T')].copy()
        
        if not df_raw.empty:
            # 依據題目 (qid) 加總所有 count 欄位
            df_all_years_sum = df_raw.groupby('qid')[
                ['count_1', 'count_2', 'count_3', 'count_4', 'count_5', 'total']
            ].sum().reset_index()
            
            ws_all = writer.book.create_sheet("全部學年統計")
            write_sheet_data(
                ws_all, 
                df_all_years_sum, 
                big_map, 
                small_map, 
                startrow=0, 
                label=f"歷年總計 (所有學年)"
            )
        else:
            print("警告：沒有原始學期資料可供加總，跳過「全部學年統計」分頁。")

        # --- [分頁 2...] 各學年分頁 ---
        # 取得所有學年 (前3碼)
        years = sorted(list(set(df['sem'].str[:3])))
        
        for year in years:
            sheet_name = f"{year}_{sheet_suffix}"
            ws = writer.book.create_sheet(sheet_name)
            current_row = 0
            
            # 1. 嘗試抓取該學年的總計列 ('XXXT')
            df_year_total = df[df['sem'] == f"{year}T"]
            
            # 如果資料庫裡沒有 'T' 列，我們就自己算！
            if df_year_total.empty:
                # 抓取該學年所有學期 (e.g., 1091, 1092)
                df_year_raw = df[(df['sem'].str.startswith(year)) & (~df['sem'].str.endswith('T'))]
                if not df_year_raw.empty:
                    df_year_total = df_year_raw.groupby('qid')[
                        ['count_1', 'count_2', 'count_3', 'count_4', 'count_5', 'total']
                    ].sum().reset_index()
            
            # 寫入學年總表
            if not df_year_total.empty:
                current_row = write_sheet_data(ws, df_year_total, big_map, small_map, startrow=current_row, label=f"{year}學年總和")
                current_row += 1 # 空一行
            
            # 2. 寫入各學期明細 (01, 02)
            for sem_suffix in ["01", "02"]:
                sem = f"{year}{sem_suffix}"
                df_sem = df[df['sem'] == sem]
                if not df_sem.empty:
                    current_row = write_sheet_data(ws, df_sem, big_map, small_map, startrow=current_row, label=f"{sem}學期")
                    current_row += 1

# ==========================================
# 3. 執行匯出
# ==========================================

print(f"輸出目標資料夾: {OUTPUT_DIR_PATH}")

# 輸出大學部
export_excel(df_u, big_map_u, small_map_u, f'大學部離系問券統計_{today_str}.xlsx', '大學部')

# 輸出研究所
export_excel(df_g, big_map_g, small_map_g, f'研究所離系問券統計_{today_str}.xlsx', '研究所')

db.close()
print("-" * 30)
print(f"執行完成！請檢查資料夾: {OUTPUT_DIR_PATH}")