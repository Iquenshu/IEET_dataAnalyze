import pandas as pd
from Accessdb import AccessHelper

# 1. 讀取 LeavDepUdata 資料表
db = AccessHelper()
table_name = 'LeavDepUdata'
analyze_table = 'LDUdataAnalyze'

# 2. 需要統計的題目欄位
qid_list = [
    'A11','A12','A13','A14','A15','A21','A22','A23','A24','A25','A26','A27','A28','A29','A210','A211'
]

# 3. 讀取所有資料
sql = f"SELECT sem, {','.join(qid_list)} FROM {table_name}"
df = pd.read_sql(sql, db.conn)

# 4. 統計各學年學期各題目各答案數量（橫式結構，每列一個題目所有答案次數）
records = []
df['year'] = df['sem'].str[:3]  # 取學年
df['semester'] = df['sem'].str[3:]  # 取學期

def get_counts(group, qid):
    # 回傳1~5各答案的次數（即使為0也記錄）
    counts = [int((group[qid] == ans).sum()) for ans in range(1, 6)]
    total = sum(counts)
    return counts, total

for qid in qid_list:
    # 各學年學期
    for sem, group in df.groupby('sem'):
        counts, total = get_counts(group, qid)
        records.append({'sem': sem, 'qid': qid,
                        'count_1': counts[0], 'count_2': counts[1], 'count_3': counts[2], 'count_4': counts[3], 'count_5': counts[4],
                        'total': total})
    # 學年總和
    for year, group in df.groupby('year'):
        sem_year = f"{year}T"
        counts, total = get_counts(group, qid)
        records.append({'sem': sem_year, 'qid': qid,
                        'count_1': counts[0], 'count_2': counts[1], 'count_3': counts[2], 'count_4': counts[3], 'count_5': counts[4],
                        'total': total})

# 5. 寫入 LDUdataAnalyze 資料表（橫式結構）
columns = ['sem', 'qid', 'count_1', 'count_2', 'count_3', 'count_4', 'count_5', 'total']
repeat_count = 0
import_count = 0

for row in records:
    where = "sem=? AND qid=?"
    params = (row['sem'], row['qid'])
    if db.is_duplicate(analyze_table, where, params):
        repeat_count += 1
        continue
    db.insert_row(analyze_table, columns, tuple(row[col] for col in columns))
    import_count += 1

db.close()
print(f"統計完成！重複資料：{repeat_count} 筆，匯入新資料：{import_count} 筆")