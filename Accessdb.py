import pyodbc

# 預設 Access 資料庫路徑（全專案共用）
DEFAULT_ACCESS_PATH = r'D:\113年後資料\系辦辦公相關\IEET認證\python程式\PythonIEET\PythonIEET\IEETdatabase.accdb'

class AccessHelper:
    """
    AccessHelper 類別：提供 Access 資料庫的連線、查詢重複、插入資料等常用功能。
    """

    def __init__(self, db_path=DEFAULT_ACCESS_PATH):
        """
        初始化 AccessHelper，建立資料庫連線。
        db_path: Access 資料庫檔案路徑，預設為全域 DEFAULT_ACCESS_PATH。
        """
        self.conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            rf'DBQ={db_path};'
        )
        self.conn = pyodbc.connect(self.conn_str)
        self.cursor = self.conn.cursor()

    def is_duplicate(self, table, where_clause, params):
        """
        檢查資料是否重複。
        table: 資料表名稱
        where_clause: SQL WHERE 條件（不含 WHERE）
        params: 條件對應的參數 tuple
        回傳 True 表示有重複，False 表示無重複
        """
        sql = f"SELECT COUNT(*) FROM {table} WHERE {where_clause}"
        self.cursor.execute(sql, params)
        return self.cursor.fetchone()[0] > 0

    def insert_row(self, table, columns, values):
        """
        插入一筆資料到指定資料表。
        table: 資料表名稱
        columns: 欄位名稱 list
        values: 欄位值 tuple
        """
        placeholders = ','.join(['?'] * len(columns))
        sql = f"INSERT INTO {table} ({','.join(columns)}) VALUES ({placeholders})"
        self.cursor.execute(sql, values)
        self.conn.commit()

    def bulk_insert(self, table, columns, rows):
        """
        批次插入多筆資料到指定資料表。
        table: 資料表名稱
        columns: 欄位名稱 list
        rows: 欄位值的 list of tuple
        使用 fast_executemany 提升效率。
        """
        placeholders = ','.join(['?'] * len(columns))
        sql = f"INSERT INTO {table} ({','.join(columns)}) VALUES ({placeholders})"
        self.cursor.fast_executemany = True  # 提升批次插入效率
        self.cursor.executemany(sql, rows)
        self.conn.commit()

    def close(self):
        """
        關閉資料庫連線
        """
        self.conn.close()