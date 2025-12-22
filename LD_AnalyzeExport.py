import pandas as pd
import matplotlib.pyplot as plt
from Accessdb import AccessHelper
from openpyxl.drawing.image import Image as XLImage
import os
import uuid
from datetime import datetime

db = AccessHelper()

# 取得今天日期字串，格式：YYYYMMDD
today_str = datetime.today().strftime('%Y%m%d')

# 1. 讀取統計資料與題目對照表
df_u = pd.read_sql("SELECT * FROM LDUdataAnalyze", db.conn)
df_g = pd.read_sql("SELECT * FROM LDGdataAnalyze", db.conn)
df_q = pd.read_sql("SELECT * FROM LeavDepQuest", db.conn)

def get_quest_map(row):
    big_map = {}
    small_map = {}
    for i in range(1, 4):
        big_key = f"A{i}0"
        if big_key in row:
            big_map[big_key] = row[big_key]
        for j in range(1, 12):
            small_key = f"A{i}{j}"
            if small_key in row:
                small_map[small_key] = row[small_key]
    return big_map, small_map

quest_row_u = df_q[df_q['QuestType'] == 'UdataZH'].iloc[0]
quest_row_g = df_q[df_q['QuestType'] == 'GdataZH'].iloc[0]
big_map_u, small_map_u = get_quest_map(quest_row_u)
big_map_g, small_map_g = get_quest_map(quest_row_g)

def export_excel(df, big_map, small_map, filename, label):
    output_path = os.path.join('output_files', filename)
    temp_img_files = []
    temp_dir = 'temp'
    os.makedirs(temp_dir, exist_ok=True)

    with pd.ExcelWriter(output_path) as writer:
        years = sorted(set(df['sem'].str[:3]))
        for year in years:
            sheet_name = f"{year}_{label}"
            ws = writer.book.create_sheet(sheet_name)
            startrow = 0
            ws.cell(row=startrow+1, column=1, value=f"{year}學年總和")
            startrow += 1
            # 只抓 A10, A20, A21... 不抓 A30（避免簡答題進入統計）
            for big_key, big_title in big_map.items():
                # 跳過 A30（問券簡答題）
                if big_key == "A30":
                    continue
                ws.cell(row=startrow+1, column=1, value=f"【{big_title}】")
                startrow += 1
                # 整合同大題目所有小題目成一個表格
                qids = [k for k in small_map if k.startswith(big_key[:2]) and not k.startswith("A30")]
                table_data = []
                for qid in qids:
                    row = df[(df['sem'] == f"{year}T") & (df['qid'] == qid)]
                    if row.empty:
                        continue
                    row = row.iloc[0]
                    table_data.append([
                        small_map[qid],
                        row['count_1'],
                        row['count_2'],
                        row['count_3'],
                        row['count_4'],
                        row['count_5'],
                        row['total']
                    ])
                table_df = pd.DataFrame(table_data, columns=['題目', '答案1', '答案2', '答案3', '答案4', '答案5', '總數'])
                table_df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)
                # --- 圖表產生區 ---
                # 每個圖表最多顯示5個題目
                for i in range(0, len(table_data), 5):
                    group = table_data[i:i+5]
                    if not group:
                        continue
                    img_path = os.path.join(temp_dir, f'temp_{big_key}_{year}_{i}_{uuid.uuid4().hex}.png')
                    plt.figure(figsize=(7,3))
                    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
                    # 橫向多題目，每題一組長條
                    x = range(1, 6)
                    for idx, item in enumerate(group):
                        plt.bar(
                            [xi + idx*0.15 for xi in x],  # 每題往右偏移
                            item[1:6],
                            width=0.15,
                            label=f"A{big_key[1]}{i+idx+1}",
                            color=colors[idx % len(colors)],
                            alpha=0.8
                        )
                        # 在每個長條頂端加上數字標籤（縮小字體）
                        for xi, count in zip([xi + idx*0.15 for xi in x], item[1:6]):
                            plt.text(xi, count, str(count), ha='center', va='bottom', fontsize=7)
                    plt.xlabel('答案', fontproperties="Microsoft JhengHei")
                    plt.ylabel('次數', fontproperties="Microsoft JhengHei")
                    # 圖表標題顯示「學年總和」+ 題目編號
                    qid_labels = [f"A{big_key[1]}{i+j+1}" for j in range(len(group))]
                    plt.title(f"{year}學年總和 {','.join(qid_labels)}", fontproperties="Microsoft JhengHei", fontsize=10)
                    plt.legend(loc='best', fontsize=8)
                    plt.subplots_adjust(left=0.15, right=0.95, top=0.85, bottom=0.2)
                    plt.tight_layout()
                    plt.savefig(img_path, dpi=120)
                    plt.close()
                    img = XLImage(img_path)
                    img_cell = f"H{startrow+2+i*6}"
                    ws.add_image(img, img_cell)
                    temp_img_files.append(img_path)
                startrow += len(table_data) + 2
            # --- 上下學期 ---
            for sem in [f"{year}01", f"{year}02"]:
                df_sem = df[df['sem'] == sem]
                if df_sem.empty:
                    continue
                ws.cell(row=startrow+1, column=1, value=f"{sem}學期")
                startrow += 1
                for big_key, big_title in big_map.items():
                    if big_key == "A30":
                        continue
                    ws.cell(row=startrow+1, column=1, value=f"【{big_title}】")
                    startrow += 1
                    qids = [k for k in small_map if k.startswith(big_key[:2]) and not k.startswith("A30")]
                    table_data = []
                    for qid in qids:
                        row = df_sem[df_sem['qid'] == qid]
                        if row.empty:
                            continue
                        row = row.iloc[0]
                        table_data.append([
                            small_map[qid],
                            row['count_1'],
                            row['count_2'],
                            row['count_3'],
                            row['count_4'],
                            row['count_5'],
                            row['total']
                        ])
                    table_df = pd.DataFrame(table_data, columns=['題目', '答案1', '答案2', '答案3', '答案4', '答案5', '總數'])
                    table_df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)
                    for i in range(0, len(table_data), 5):
                        group = table_data[i:i+5]
                        if not group:
                            continue
                        img_path = os.path.join(temp_dir, f'temp_{big_key}_{sem}_{i}_{uuid.uuid4().hex}.png')
                        plt.figure(figsize=(7,3))
                        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
                        x = range(1, 6)
                        for idx, item in enumerate(group):
                            plt.bar(
                                [xi + idx*0.15 for xi in x],
                                item[1:6],
                                width=0.15,
                                label=f"A{big_key[1]}{i+idx+1}",
                                color=colors[idx % len(colors)],
                                alpha=0.8
                            )
                            for xi, count in zip([xi + idx*0.15 for xi in x], item[1:6]):
                                plt.text(xi, count, str(count), ha='center', va='bottom', fontsize=7)
                        plt.xlabel('答案', fontproperties="Microsoft JhengHei")
                        plt.ylabel('次數', fontproperties="Microsoft JhengHei")
                        qid_labels = [f"A{big_key[1]}{i+j+1}" for j in range(len(group))]
                        plt.title(f"{sem} {','.join(qid_labels)}", fontproperties="Microsoft JhengHei", fontsize=10)
                        plt.legend(loc='best', fontsize=8)
                        plt.subplots_adjust(left=0.15, right=0.95, top=0.85, bottom=0.2)
                        plt.tight_layout()
                        plt.savefig(img_path, dpi=120)
                        plt.close()
                        img = XLImage(img_path)
                        img_cell = f"H{startrow+2+i*6}"
                        ws.add_image(img, img_cell)
                        temp_img_files.append(img_path)
                    startrow += len(table_data) + 2

    # --- 清理暫存圖片檔案 ---
    for img_path in temp_img_files:
        if os.path.exists(img_path):
            os.remove(img_path)

# 3. 分別輸出大學部與研究所
export_excel(df_u, big_map_u, small_map_u, f'大學部離系問券統計_{today_str}.xlsx', '大學部')
export_excel(df_g, big_map_g, small_map_g, f'研究所離系問券統計_{today_str}.xlsx', '研究所')

db.close()
print("匯出完成！已產生大學部與研究所統計Excel檔案。")