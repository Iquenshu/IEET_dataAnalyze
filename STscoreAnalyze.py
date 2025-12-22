import pandas as pd
from Accessdb import AccessHelper

# 1. 設定來源資料表與分析結果資料表名稱
source_table = 'STscore'         # 原始成績資料表
analyze_table = 'STscoreAnalyze' # 分析結果資料表

# 2. 連接 Access 資料庫
db = AccessHelper()

# 3. 從 STscore 讀取所有成績資料
sql = f"SELECT * FROM {source_table}"
try:
    df = pd.read_sql(sql, db.conn)  # 讀取成績資料為 pandas DataFrame
except Exception as e:
    print("讀取成績資料失敗，請檢查資料庫連線或 SQL 語法。錯誤訊息：", e)
    db.close()
    exit()

# 檢查原始成績資料
print("原始成績資料筆數：", df.shape[0])
if df.shape[0] == 0:
    print("警告：原始成績資料為空，請確認 STscore 資料表有資料。")
    db.close()
    exit()
print("原始成績資料前5筆：")
print(df.head())

# 4. 先過濾掉分數為999（退選）的學生
df_valid = df[df['成績'] != 999]
print("未退選有效成績資料筆數：", df_valid.shape[0])

# 5. 設定分數區間與標籤（分布統計用）
bins = [0,10,20,30,40,50,60,70,80,90,100]
labels = ["0-9","10-19","20-29","30-39","40-49","50-59","60-69","70-79","80-89","90-100"]

records = []  # 用來存放所有分析結果

# 6. 依「學年度、學期、課號、課程名稱」分組，統計各分數區間人數、平均分數、學生總數
for (year, sem, course_id, course_name), group in df_valid.groupby(['學年度', '學期', '課號', '課程名稱']):
    # 計算 total 平均分數（該課該學期所有未退選學生的平均）
    total_avg_score = group['成績'].mean()
    total_students = group['成績'].count()
    # 加入 total 統計
    records.append({
        '學年度': str(year),
        '學期': str(sem),
        '課號': str(course_id),
        '課程名稱': str(course_name),
        '分數區間': 'total',
        '人數': int(total_students),
        '平均分數': float(total_avg_score) if pd.notnull(total_avg_score) else 0.0,
        '學生總數': int(total_students)
    })
    # 各分數區間統計
    dist = pd.cut(group['成績'], bins=bins, labels=labels, right=False)
    for label in labels:
        group_label = group[dist == label]
        label_count = group_label.shape[0]
        label_avg = group_label['成績'].mean() if label_count > 0 else 0.0
        records.append({
            '學年度': str(year),
            '學期': str(sem),
            '課號': str(course_id),
            '課程名稱': str(course_name),
            '分數區間': str(label),
            '人數': int(label_count),
            '平均分數': float(label_avg) if pd.notnull(label_avg) else 0.0,
            '學生總數': int(total_students)
        })

# 檢查分析結果
print("分析結果筆數：", len(records))
if len(records) == 0:
    print("警告：分析結果為空，請檢查分組欄位或原始資料內容。")
    db.close()
    exit()
print("分析結果前5筆：")
print(records[:5])

# 7. 逐筆寫入 Access 資料表，避免重複
columns = ['學年度', '學期', '課號', '課程名稱', '分數區間', '人數', '平均分數', '學生總數']
success_count = 0
fail_count = 0

for row in records:
    values = (
        row['學年度'],
        row['學期'],
        row['課號'],
        row['課程名稱'],
        row['分數區間'],
        row['人數'],
        row['平均分數'],
        row['學生總數']
    )
    where = "學年度=? AND 學期=? AND 課號=? AND 分數區間=?"
    params = (row['學年度'], row['學期'], row['課號'], row['分數區間'])
    if db.is_duplicate(analyze_table, where, params):
        # 主鍵重複時用 UPDATE 更新資料
        try:
            update_sql = f"""
            UPDATE {analyze_table}
            SET 課程名稱=?, 人數=?, 平均分數=?, 學生總數=?
            WHERE 學年度=? AND 學期=? AND 課號=? AND 分數區間=?
            """
            update_params = (
                row['課程名稱'], row['人數'], row['平均分數'], row['學生總數'],
                row['學年度'], row['學期'], row['課號'], row['分數區間']
            )
            db.cursor.execute(update_sql, update_params)
            db.conn.commit()
            print("已更新重複資料：", values)
            success_count += 1
        except Exception as e:
            print("更新失敗：", values)
            print("錯誤訊息：", e)
            fail_count += 1
        continue
    try:
        db.insert_row(analyze_table, columns, values)
        success_count += 1
    except Exception as e:
        print("寫入失敗：", values)
        print("錯誤訊息：", e)
        fail_count += 1

db.close()

print(f"成績分布分析完成，成功寫入 {success_count} 筆，失敗 {fail_count} 筆。")