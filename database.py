import sqlite3
import os
import sys
from datetime import datetime

# --- 1. 路徑處理邏輯 (統一管理) ---
def get_db_path():
    """獲取資料庫絕對路徑，確保不論從哪啟動都能找到檔案"""
    # 如果是打包後的環境 (.app / .exe)
    if getattr(sys, 'frozen', False):
        # 這裡根據你的需求，通常打包後資料庫會放在程式旁邊
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        # 開發環境：使用 database.py 所在的資料夾
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(BASE_DIR, "patients.db")

# 設定全域變數 DB_NAME 供所有函數使用
DB_NAME = get_db_path()
print(f"📍 目前使用的資料庫路徑: {DB_NAME}")

def connect_db():
    """建立資料庫連線"""
    return sqlite3.connect(DB_NAME)

# --- 2. 資料庫初始化 ---
def initialize_db():
    """初始化資料庫：建立資料表並檢查欄位"""
    conn = connect_db()
    cursor = conn.cursor()
    # 注意：這裡欄位名稱維持你的 "condiction" (拼錯但需維持一致以免報錯)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS patients (
        id TEXT PRIMARY KEY,
        name TEXT,
        drug TEXT,
        condiction TEXT,
        last_date TEXT,
        doctor TEXT,
        return_date TEXT,
        note TEXT 
    )
    """)
    # 檢查是否需要補上 note 欄位
    cursor.execute("PRAGMA table_info(patients)")
    columns = [col[1] for col in cursor.fetchall()]
    if "note" not in columns:
        cursor.execute("ALTER TABLE patients ADD COLUMN note TEXT")
        conn.commit()
    conn.close()

# --- 3. 資料操作函數 ---

def get_patient_by_id(p_id):
    """根據 ID 取得單一個案 (包含自動偵測欄位與型態轉換)"""
    try:
        conn = connect_db()
        cursor = conn.cursor()
        
        # 1. 偵測欄位名稱 (防止第一欄不叫 id)
        cursor.execute("PRAGMA table_info(patients)")
        columns = [info[1] for info in cursor.fetchall()]
        id_col = columns[0] if columns else "id"
        
        target_id = str(p_id).strip()
        
        # 2. 嘗試字串查詢
        cursor.execute(f"SELECT * FROM patients WHERE {id_col}=?", (target_id,))
        data = cursor.fetchone()
        
        # 3. 失敗則嘗試整數查詢
        if not data:
            try:
                cursor.execute(f"SELECT * FROM patients WHERE {id_col}=?", (int(target_id),))
                data = cursor.fetchone()
            except: pass
                
        conn.close()
        return data
    except Exception as e:
        print(f"❌ 查詢失敗: {e}")
        return None

def get_all_patients():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM patients")
    data = cursor.fetchall()
    conn.close()
    return data

def add_new_case(data_tuple):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO patients (id, name, drug, condiction, last_date, doctor, return_date, note) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, data_tuple)
    conn.commit()
    conn.close()

def update_patient_data(data):
    """
    data 順序: (姓名, 藥物, 條件, 結束日, 醫師, 回診日, 備註, 病歷號)
    """
    try:
        conn = connect_db()
        cursor = conn.cursor()
        sql = """
        UPDATE patients 
        SET name=?, drug=?, condiction=?, last_date=?, doctor=?, return_date=?, note=?
        WHERE id=?
        """
        cursor.execute(sql, data)
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        return success
    except Exception as e:
        print(f"❌ 更新失敗: {e}")
        return False

def delete_patient(pid):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM patients WHERE id = ?", (pid,))
    conn.commit()
    conn.close()

def search_patients(keyword):
    conn = connect_db()
    cursor = conn.cursor()
    query = "SELECT * FROM patients WHERE name LIKE ? OR id LIKE ?"
    cursor.execute(query, (f'%{keyword}%', f'%{keyword}%'))
    results = cursor.fetchall()
    conn.close()
    return results