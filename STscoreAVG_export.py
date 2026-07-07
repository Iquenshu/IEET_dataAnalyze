import pandas as pd
from Accessdb import AccessHelper
import os
from datetime import datetime

# ==============================================================================
# 【使用者設定區】 可自由調整您想統計的學年起迄區間
# ==============================================================================
START_YEAR = 109  # 起始學年
END_YEAR = 114    # 結束學年
# ==============================================================================

# 1. 連接 Access 資料庫
db = AccessHelper()

# 2. 從原始成績表 STscore 讀取資料並過濾掉 999 退選成績
sql = "SELECT 學年度, 學期, 課號, 課程名稱, 成績 FROM STscore WHERE 成績 <> 999"
print("正在從 STscore 讀取原始成績資料...")
df_score = pd.read_sql(sql, db.conn)

db.close()

if df_score.empty:
    print("錯誤：沒有讀取到任何有效成績資料。")
    exit()

# 將學年度轉為數字型態以方便進行區間比對
df_score['學年度_num'] = df_score['學年度'].astype(int)

# ==============================================================================
# 3. 處理「各學年獨立」的統計資料
# ==============================================================================
print("正在計算各學年課程平均分數與及格率...")
records_yearly = []

for (year, sem, course_id, course_name), group in df_score.groupby(['學年度', '學期', '課號', '課程名稱']):
    total_students = group['成績'].count()
    avg_score = group['成績'].mean()
    pass_students = group[group['成績'] >= 60]['成績'].count()
    pass_rate = (pass_students / total_students) * 100 if total_students > 0 else 0.0
    
    records_yearly.append({
        '學年度': year, '學期': sem, '課號': course_id, '課程名稱': course_name,
        '平均分數': round(avg_score, 2), '學生總數': total_students,
        '及格率': f"{round(pass_rate, 1)}%"
    })

df_yearly_result = pd.DataFrame(records_yearly)

# ==============================================================================
# 4. 處理「指定區間年（跨學年合併）」的統計資料
# ==============================================================================
print(f"正在計算指定區間 ({START_YEAR} ~ {END_YEAR}學年) 的合併平均分數與及格率...")
records_range = []

# 篩選出符合使用者設定區間的原始成績
df_range = df_score[(df_score['學年度_num'] >= START_YEAR) & (df_score['學年度_num'] <= END_YEAR)].copy()

if not df_range.empty:
    # 跨學年合併時，依「課號、課程名稱」進行分組
    for (course_id, course_name), group in df_range.groupby(['課號', '課程名稱']):
        total_students = group['成績'].count()
        avg_score = group['成績'].mean()
        pass_students = group[group['成績'] >= 60]['成績'].count()
        pass_rate = (pass_students / total_students) * 100 if total_students > 0 else 0.0
        
        records_range.append({
            '課號': course_id,
            '課程名稱': course_name,
            '區間總平均分數': round(avg_score, 2),
            '區間學生總人數': total_students,
            '區間總及格率': f"{round(pass_rate, 1)}%"
        })

df_range_result = pd.DataFrame(records_range)

# ==============================================================================
# 5. 寫入 Excel 檔案 (包含年度分頁與專屬區間分頁)
# ==============================================================================
output_dir = 'output_files'
os.makedirs(output_dir, exist_ok=True)
today_str = datetime.today().strftime('%Y%m%d')
output_path = os.path.join(output_dir, f'STscoreAVG_{today_str}.xlsx')

print(f"正在寫入 Excel 檔案: {output_path}")
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    
    # A. 優先寫入「指定區間年總計」分頁
    if not df_range_result.empty:
        df_range_result = df_range_result.sort_values(by=['課號'])
        range_sheet_name = f"{START_YEAR}-{END_YEAR}年總計"
        df_range_result.to_excel(writer, sheet_name=range_sheet_name, index=False)
        print(f"  -> 已成功建立區間分頁: {range_sheet_name}")
    else:
        print(f"  【警告】在 {START_YEAR} 到 {END_YEAR} 學年間找不到任何成績資料，未建立區間分頁。")

    # B. 依序寫入各學年度的獨立分頁
    for year in sorted(df_yearly_result['學年度'].unique(), reverse=True):
        df_year = df_yearly_result[df_yearly_result['學年度'] == year].copy()
        df_year = df_year.sort_values(by=['學期', '課號'])
        
        df_year = df_year[['學年度', '學期', '課號', '課程名稱', '平均分數', '學生總數', '及格率']]
        sheet_name = f"{year}學年度"
        df_year.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"匯出完成！已產生 {output_path}")