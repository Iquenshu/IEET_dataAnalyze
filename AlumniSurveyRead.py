import pandas as pd
import os
from Accessdb import AccessHelper

def import_alumni_survey():
    # 1. 設定檔案路徑
    # 這裡假設您將檔案放在 "input_files\畢業系友流向問券" 資料夾下
    # 請修改檔名以符合您實際放入的檔案
    folder_path = r'input_files\畢業系友流向問券'
    file_name = '電機系畢業系友流向問卷匯入1140728.xlsx' 
    data_path = os.path.join(folder_path, file_name)
    
    table_name = 'AlumniSurvey'

    print(f"正在準備讀取檔案: {data_path} ...")
    
    if not os.path.exists(data_path):
        print(f"錯誤：找不到檔案，請確認路徑與檔名是否正確。")
        print(f"預期路徑: {data_path}")
        return

    # 2. 自動判斷副檔名並讀取
    ext = os.path.splitext(data_path)[1].lower()
    df = pd.DataFrame()
    
    try:
        if ext in ['.xls', '.xlsx']:
            print("偵測到 Excel 格式，使用 read_excel 讀取...")
            # 有時候 Excel 會有隱藏的工作表，預設讀取第一張 (sheet_name=0)
            df = pd.read_excel(data_path, sheet_name=0)
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
        print(f"讀取檔案發生嚴重錯誤: {e}")
        return

    # 3. 定義欄位對照表 (Key: 題目關鍵字, Value: 資料庫欄位)
    # 只要標題包含 Key 中的字串就會被抓取
    col_mapping = {
        '填寫順序': 'id',
        '會員帳號': 'MemberID',
        '1.畢業學制': 'Degree',
        '2.畢業學年度': 'GradYear',
        '3.目前任職公司的人數規模': 'CompanySize',
        '4.目前任職公司之行業別': 'IndustryType',
        '5.目前任職公司之產業屬性': 'IndustryAttr',
        '6.目前任職職務或研究之屬性': 'JobType',
        '7.目前任職職務的業務範圍': 'JobScope',
        '8.目前職務是否擔任主管': 'IsManager',
        '9.承8題': 'ManageCount',
        '10.工作團隊人數': 'TeamSize',
        '11.所屬工作團隊': 'PatentCount',
        # 教育目標重要性 (12-16)
        '12.您認為中山電機系教育目標［學識理論］': 'Q1_Theory_Imp',
        '13.您認為中山電機系教育目標［專業技術］': 'Q2_Tech_Imp',
        '14.您認為中山電機系教育目標［團隊精神與工程倫理］': 'Q3_Team_Imp',
        '15.您認為中山電機系教育目標［獨立思考與創新］': 'Q4_Innov_Imp',
        '16.您認為中山電機系教育目標［國際視野］': 'Q5_Global_Imp',
        # 自我工作態度評價 (17-21)
        '17.您對畢業迄今的自我工作態度評價［學識理論］': 'Q1_Theory_Sat',
        '18.您對畢業迄今的自我工作態度評價［專業技術］': 'Q2_Tech_Sat',
        '19.您對畢業迄今的自我工作態度評價［團隊精神與工程倫理］': 'Q3_Team_Sat',
        '20.您對畢業迄今的自我工作態度評價［獨立思考與創新］': 'Q4_Innov_Sat',
        '21.您對畢業迄今的自我工作態度評價［國際視野］': 'Q5_Global_Sat',
        # 其他
        '22.如有任何建議': 'Suggestions',
        '資料建立日期': 'Fill_Date'
    }

    # 4. 建立要寫入資料庫的 DataFrame
    db_data = pd.DataFrame()
    
    print("正在對應欄位...")
    for map_key, db_col in col_mapping.items():
        found = False
        for csv_col in df.columns:
            if map_key in str(csv_col): # 模糊比對
                db_data[db_col] = df[csv_col]
                found = True
                break
        if not found:
            # 針對非必要欄位不報錯，僅提醒
            # print(f"  提醒：找不到包含「{map_key}」的欄位 (Access欄位: {db_col})，將留空。")
            db_data[db_col] = None

    # 處理空值
    db_data = db_data.where(pd.notnull(db_data), None)

    # 5. 寫入 Access
    db = AccessHelper()
    columns = list(db_data.columns)
    
    repeat_count = 0
    import_count = 0
    error_count = 0

    print(f"開始寫入資料庫 [{table_name}] ...")
    
    for idx, row in db_data.iterrows():
        try:
            # 防重複機制：使用 'id' (填寫順序) 
            if row['id'] is None:
                continue
                
            where = "id=?"
            params = (row['id'],)
            
            if db.is_duplicate(table_name, where, params):
                repeat_count += 1
                continue
            
            # 插入資料
            values = tuple(row[col] for col in columns)
            db.insert_row(table_name, columns, values)
            import_count += 1
            
        except Exception as e:
            error_count += 1
            print(f"第 {idx+1} 筆寫入錯誤: {e}")

    db.close()
    
    print("="*40)
    print(f"匯入總結 - {table_name}")
    print(f"成功新增: {import_count}")
    print(f"重複略過: {repeat_count}")
    if error_count > 0:
        print(f"寫入失敗: {error_count}")
    print("="*40)

if __name__ == "__main__":
    # 確保資料夾存在，若不存在自動建立 (方便使用者)
    if not os.path.exists(r'input_files\畢業系友流向問券'):
        os.makedirs(r'input_files\畢業系友流向問券')
        print(r"已建立資料夾: input_files\畢業系友流向問券，請將檔案放入後再執行。")
    else:
        import_alumni_survey()