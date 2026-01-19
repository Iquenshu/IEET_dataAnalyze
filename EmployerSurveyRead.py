import pandas as pd
import os
from Accessdb import AccessHelper

# 畢業生流向雇主問卷匯入程式

def import_employer_survey():
    # 1. 設定檔案路徑 
    # (請確認此路徑指向您電腦上的實際檔案)
    data_path = r'input_files\畢業生流向雇主問券\雇主問卷匯入用1140724.xlsx'
    table_name = 'EmployerSurvey'

    print(f"正在讀取檔案: {data_path} ...")
    
    if not os.path.exists(data_path):
        print(f"錯誤：找不到檔案 {data_path}")
        return

    # 2. 依副檔名自動選擇讀取方式 (修正錯誤的關鍵)
    ext = os.path.splitext(data_path)[1].lower()
    
    try:
        if ext in ['.xls', '.xlsx']:
            print("偵測到 Excel 格式，使用 read_excel 讀取...")
            df = pd.read_excel(data_path)
        elif ext == '.csv':
            print("偵測到 CSV 格式，使用 read_csv 讀取...")
            try:
                df = pd.read_csv(data_path, encoding='utf-8')
            except UnicodeDecodeError:
                print("UTF-8 讀取失敗，嘗試 CP950...")
                df = pd.read_csv(data_path, encoding='cp950')
        else:
            print(f"不支援的檔案格式: {ext}")
            return
    except Exception as e:
        print(f"讀取檔案發生錯誤: {e}")
        return

    # 3. 定義欄位對照表 (Key: Excel原標題關鍵字, Value: Access欄位名)
    col_mapping = {
        '填寫順序': 'id',
        '1.您認為本系教育目標［學識理論］': 'Q1_Theory_Imp',
        '2.您認為本系教育目標［專業技術］': 'Q2_Tech_Imp',
        '3.您認為本系教育目標［團隊精神與工程倫理］': 'Q3_Team_Imp',
        '4.您認為本系教育目標［獨立思考與創新］': 'Q4_Innov_Imp',
        '5.您認為本系教育目標［國際視野］': 'Q5_Global_Imp',
        '6.您目前是否有帶領過': 'Has_Hired',
        '7.您所帶領員工在「學識理論」': 'Q7_Theory_Perf',
        '8.您所帶領員工在「專業技術」': 'Q8_Tech_Perf',
        '9.您所帶領員工在「團隊精神與工程倫理」': 'Q9_Team_Perf',
        '10.您所帶領員工在「獨立思考與創新」': 'Q10_Innov_Perf',
        '11.您所帶領員工在「國際視野」': 'Q11_Global_Perf',
        '12.如有任何建議': 'Suggestions',
        '資料建立日期': 'Fill_Date'
    }

    # 4. 重新命名欄位並篩選
    db_data = pd.DataFrame()
    
    print("正在處理欄位對應...")
    for map_key, db_col in col_mapping.items():
        found = False
        for csv_col in df.columns:
            if map_key in csv_col: # 使用模糊比對
                db_data[db_col] = df[csv_col]
                found = True
                break
        if not found:
            # 針對非必要欄位可顯示警告，但不中斷
            print(f"提醒：找不到包含「{map_key}」的欄位，將填入空白。")
            db_data[db_col] = None

    # 處理空值 (NaN 轉 None)
    db_data = db_data.where(pd.notnull(db_data), None)

    # 5. 寫入 Access 資料庫
    db = AccessHelper()
    columns = list(db_data.columns)
    
    repeat_count = 0
    import_count = 0

    print(f"開始寫入資料庫 [{table_name}] ...")
    for idx, row in db_data.iterrows():
        try:
            # 防重複檢查：使用 'id' (填寫順序)
            if row['id'] is None:
                continue
                
            where = "id=?"
            params = (row['id'],)
            
            if db.is_duplicate(table_name, where, params):
                repeat_count += 1
                continue
            
            # 準備插入資料
            values = tuple(row[col] for col in columns)
            db.insert_row(table_name, columns, values)
            import_count += 1
            
        except Exception as e:
            print(f"第 {idx+1} 筆資料寫入失敗: {e}")

    db.close()
    print("="*30)
    print(f"匯入完成！")
    print(f"成功新增：{import_count} 筆")
    print(f"重複略過：{repeat_count} 筆")
    print("="*30)

if __name__ == "__main__":
    import_employer_survey()