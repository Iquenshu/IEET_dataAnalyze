import os
import re
import pyodbc
import pandas as pd
from datetime import datetime
from openpyxl.styles import Alignment, Border, Side, Font

# ==============================================================================
# 【使用者設定區】 方便您日後修改路徑
# ==============================================================================
# 1. 目的地 Excel 輸出資料夾路徑
OUTPUT_DIR = r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\output_files\畢業生論文清單"

# 2. 您的 IEET 來源資料庫路徑
TARGET_DB = r"D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\IEETdatabase.accdb"

# 3. 資料來源資料表名稱
SRC_TABLE_NAME = "ThesisDefens_Data"
# ==============================================================================

# Access 連線字串
CONN_STR = r"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={};".format(TARGET_DB)


def calculate_academic_year(date_str):
    """
    從口試日期自動推算學年度 (純數字字串，如 '109')
    """
    date_str = str(date_str).strip().replace('.', '/').replace('-', '/')
    parts = date_str.split('/')
    if len(parts) >= 2:
        try:
            year = int(parts[0])
            month = int(parts[1])
            if year > 1900:
                year -= 1911
            
            # 8月(含)以後屬於當學年度；7月(含)以前屬於前一學年度
            return str(year) if month >= 8 else str(year - 1)
        except ValueError:
            pass

    if len(date_str) == 7 and date_str.isdigit():
        try:
            year = int(date_str[:3])
            month = int(date_str[3:5])
            return str(year) if month >= 8 else str(year - 1)
        except ValueError:
            pass

    return "未知"


def clean_for_excel(text):
    """
    濾除引發 openpyxl 崩潰的 XML 非法控制字元
    """
    if not isinstance(text, str):
        return text
    # 移除非法 ASCII 控制字元 (0-31 區間)
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    return cleaned.strip()


def excel_formatting(ws, num_rows):
    """
    針對 Excel 進行美化排版
    """
    header_font = Font(name='微軟正黑體', size=12, bold=True)
    data_font = Font(name='微軟正黑體', size=11)
    
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # 1. 標題列美化
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    # 2. 資料列美化 (修改為 1 到 8，涵蓋全新第 7 欄)
    for row in range(2, num_rows + 2):
        for col in range(1, 8):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border
            
            # 欄位對齊：學年度(1)、指導教授(3)、身分(5)、學號(6)、口試日期(7) 置中；其餘靠左
            if col in [1, 3, 5, 6, 7]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # 3. 設定最適欄位寬度 (加入 G 欄口試日期寬度)
    ws.column_dimensions['A'].width = 10  # 學年度
    ws.column_dimensions['B'].width = 15  # 研究生姓名
    ws.column_dimensions['C'].width = 15  # 指導教授
    ws.column_dimensions['D'].width = 50  # 論文題目
    ws.column_dimensions['E'].width = 12  # 身分
    ws.column_dimensions['F'].width = 16  # 學號
    ws.column_dimensions['G'].width = 15  # 口試日期 (便於對照檢查)


def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(TARGET_DB):
        print(f"錯誤：找不到資料庫檔案 {TARGET_DB}")
        return

    print("正在讀取資料庫...")
    conn = pyodbc.connect(CONN_STR)
    query = f"SELECT * FROM {SRC_TABLE_NAME}"
    df = pd.read_sql(query, conn)
    conn.close()

    if df.empty:
        print("資料庫內沒有任何資料，取消匯出。")
        return

    # 清除欄位前後可能隱藏的空白
    df.columns = df.columns.str.strip()

    # 彈性對應更名
    if '身份' in df.columns and '身分' not in df.columns:
        df = df.rename(columns={'身份': '身分'})
    if '姓名' in df.columns:
        df = df.rename(columns={'姓名': '研究生姓名'})

    # 清洗資料並濾除非法 XML 字元
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).apply(clean_for_excel)
        df[col] = df[col].replace('None', '')

    print("正在計算學年度與整理欄位...")
    # 1. 計算純數字學年度
    df['學年度'] = df['口試日期'].apply(calculate_academic_year)

    # 2. 篩選出您指定的 7 個欄位 (最後加上口試日期)
    try:
        df_final = df[['學年度', '研究生姓名', '指導教授', '論文題目', '身分', '學號', '口試日期']].copy()
    except KeyError as e:
        print(f"欄位篩選失敗，資料表內似乎缺少了特定的欄位: {e}")
        return

    # 3. 排序 (依學年度、學號排序)
    df_final = df_final.sort_values(by=['學年度', '學號'], ascending=[True, True])

    # 4. 產出檔名
    today_str = datetime.now().strftime("%Y%m%d")
    file_name = f"畢業生論文清單_{today_str}.xlsx"
    excel_path = os.path.join(OUTPUT_DIR, file_name)

    print("正在寫入 Excel 檔案 (單一工作表)...")
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name="畢業生論文清單", index=False)
        ws = writer.sheets["畢業生論文清單"]
        excel_formatting(ws, len(df_final))

    print(f"\n==============================================")
    print(f"大功告成！含【口試日期】的對照清單已成功產出。")
    print(f"檔案路徑: {excel_path}")
    print(f"==============================================")


if __name__ == "__main__":
    main()