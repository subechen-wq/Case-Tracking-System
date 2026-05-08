# ===== 自動安裝套件 =====
import os
import sys

# 取得程式執行時的真實路徑 (相容 PyInstaller 打包後與開發環境)
if getattr(sys, 'frozen', False):
    # 如果是打包後的環境
    bundle_dir = os.path.dirname(sys.executable)
else:
    # 如果是開發環境 (VS Code)
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

# 強制指向執行檔同目錄下的資料庫
db_path = os.path.join(bundle_dir, "patients.db")

import subprocess
import sys

def install(package):
    subprocess.check_call([
        sys.executable,
        "-m",
        "pip",
        "install",
        package
    ])

try:
 from openpyxl import Workbook
except ImportError:
 install("openpyxl")
 from openpyxl import Workbook
 
try:
    from tkcalendar import Calendar
except ImportError:
    install("tkcalendar")
    from tkcalendar import Calendar


import sqlite3
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox, ttk
from openpyxl import Workbook
from tkinter import filedialog

# ===== DB =====
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

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

cursor.execute("PRAGMA table_info(patients)")
columns = [col[1] for col in cursor.fetchall()]

if "note" not in columns:
 cursor.execute("ALTER TABLE patients ADD COLUMN note TEXT")
 conn.commit()


# ===== 功能函數 =====
def show_login():
 login_window = tk.Toplevel()
 login_window.title("登入系統")
 login_window.protocol("WM_DELETE_WINDOW", exit_program)
 login_window.geometry("420x360")
 login_window.configure(bg="#e9eef5")
 login_window.grab_set()

 # ===== 陰影 =====
 shadow = tk.Frame(login_window, bg="#cfd6e0")
 shadow.place(relx=0.5, rely=0.5, anchor="center", width=320, height=260)

 # ===== 卡片 =====
 card = tk.Frame(login_window, bg="white")
 card.place(relx=0.5, rely=0.5, anchor="center", width=300, height=240)

 card.pack_propagate(False)

 # ===== 標題 =====
 tk.Label(
 card,
 text="🔐 系統登入",
 font=("Microsoft JhengHei", 18, "bold"),
 bg="white",
 fg="#2c3e50"
 ).pack(pady=(15, 10))

 # ===== 帳號 =====
 frame_user = tk.Frame(card, bg="white")
 frame_user.pack(fill="x", padx=25, pady=5)

 tk.Label(frame_user, text="👤帳號", bg="white", font=("Arial", 12)).pack(side="left")

 entry_user = tk.Entry(frame_user, bd=0, font=("Microsoft JhengHei", 11))
 entry_user.pack(side="left", fill="x", expand=True)

 tk.Frame(card, height=1, bg="#dcdde1").pack(fill="x", padx=25)

 # ===== 密碼 =====
 frame_pass = tk.Frame(card, bg="white")
 frame_pass.pack(fill="x", padx=25, pady=10)

 tk.Label(frame_pass, text="🔑密碼", bg="white", font=("Arial", 12)).pack(side="left")

 entry_pass = tk.Entry(frame_pass, show="*", bd=0, font=("Microsoft JhengHei", 11))
 entry_pass.pack(side="left", fill="x", expand=True)

 tk.Frame(card, height=1, bg="#dcdde1").pack(fill="x", padx=25)

 # ===== 錯誤提示 =====
 error_label = tk.Label(card, text="", fg="#e74c3c", bg="white", font=("Microsoft JhengHei", 9))
 error_label.pack(pady=5)

 # ===== 登入邏輯 =====
 def login():
    global current_user_role, current_username, user_label

    username = entry_user.get()
    password = entry_pass.get()

    if username in USERS and USERS[username]["password"] == password:
        current_user_role = USERS[username]["role"]
        current_username = username

        user_label.config(text=f"👤 {current_username}")

        login_window.destroy()
        root.deiconify()
        root.lift()

        root.after(1000, apply_permissions) 
        root.after(150, refresh_table)
        root.after(300, show_due_alert)

    else:
        error_label.config(text="❌ 帳號或密碼錯誤")

 # ===== 登入按鈕 =====
 btn_login = tk.Label(
 card,
 text="登入",
 bg="#3498db",
 fg="white",
 font=("Microsoft JhengHei", 12, "bold"),
 cursor="hand2",
 width=20,
 pady=6
 )
 btn_login.pack(pady=15)

 btn_login.bind("<Button-1>", lambda e: login())

 # hover 動畫
 btn_login.bind("<Enter>", lambda e: btn_login.config(bg="#2980b9"))
 btn_login.bind("<Leave>", lambda e: btn_login.config(bg="#3498db"))

 # Enter 快速登入
 login_window.bind("<Return>", lambda e: login())

 entry_user.focus()


def exit_program():
    """優雅地關閉程式，避免 Mac 彈出錯誤報告"""
    try:
        # 1. 先關閉資料庫連線
        if 'conn' in globals():
            conn.close()
    except:
        pass
    
    # 2. 銷毀所有視窗
    root.destroy()
    
    # 3. 使用 os._exit(0) 或是直接讓 mainloop 結束
    # 在 Mac 上，直接 destroy 掉 root 通常就能讓程式安靜地結束
    # 如果還是跳警告，就用下面這行：
    import os
    os._exit(0)


def apply_permissions():
    global btn_delete
    try:
        # 增加判斷，如果按鈕還沒準備好，就跳到 except 
        if 'btn_delete' in globals() and btn_delete.winfo_exists():
            if current_user_role == "nurse":
                btn_delete.config(state="normal")
            else:
                btn_delete.config(state="disabled")
    except (NameError, AttributeError, tk.TclError):
        # 找不到按鈕就不報錯，半秒後再試
        root.after(500, apply_permissions)

def format_id(pid):
 return str(pid).strip().zfill(3)

#登出按鈕
def logout():
    global current_user_role

    current_user_role = None
    user_label.config(text="未登入")

    for row in tree.get_children():
        tree.delete(row)

    root.withdraw()  # 隱藏主畫面
    show_login()     # 回登入畫面
    on_closing()

#重新登入按鈕
def relogin():
 logout()

# ===== CRUD =====

import pandas as pd

def import_excel_general():
    path = filedialog.askopenfilename(
        title="選取肝炎試辦 Excel 檔案",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not path:
        return

    try:
        df = pd.read_excel(path)
        
        # 欄位對照
        col_map = {
            '病歷號碼': 'id',
            '病患姓名': 'name',
            '試辦種類': 'drug',
            '主治醫師': 'doctor',
            '用藥結束日': 'last_date'
        }

        # 檢查必要欄位
        if not all(c in df.columns for c in col_map.keys()):
            messagebox.showerror("格式錯誤", "檔案缺少必要欄位，請檢查 Excel 表頭。")
            return

        import_count = 0
        skip_count = 0
        
        for _, row in df.iterrows():
            # 1. 處理病歷號
            raw_id = str(row['病歷號碼']).split('.')[0] if pd.notna(row['病歷號碼']) else ""
            if not raw_id: continue
            pid = format_id(raw_id)
            
            # 2. 基本資料
            name = str(row['病患姓名']) if pd.notna(row['病患姓名']) else ""
            doctor = str(row['主治醫師']) if pd.notna(row['主治醫師']) else ""
            
            # 3. 藥物名稱轉換邏輯
            raw_drug = str(row['試辦種類']) if pd.notna(row['試辦種類']) else ""
            matched_drug = "其他"
            
            # 定義規則：當 Excel 的字串中「包含」Key 時，轉換為系統標準的 Value
            drug_rules = {
                "Tenofovir alafenamide 25mg": "Tenofovir alafenamide 25mg",
                "Tenofovir 300mg": "Tenofovir 300mg",
                "Entecavir Tab. 0.5mg": "Entecavir 0.5",
                "Baraclude 0.5mg": "Entecavir 0.5",
                "Entecavir Tab. 1mg": "Entecavir 1.0",
                "Baraclude 1mg": "Entecavir 1.0",
                "Sebivo": "Telbivudine",
                "Maviret": "Maviret",
                "Epclusa": "Epclusa"
            }
            
            for key, standard_name in drug_rules.items():
                if key in raw_drug: # 使用包含判定 (Substring matching)
                    matched_drug = standard_name
                    break
            
            # 4. 處理日期 (如果為空則存入空字串)
            if pd.notna(row['用藥結束日']):
                if hasattr(row['用藥結束日'], 'strftime'):
                    last_date = row['用藥結束日'].strftime('%Y-%m-%d')
                else:
                    last_date = str(row['用藥結束日'])[:10]
            else:
                last_date = ""  # 👈 修改點：無日期時顯示空白

            # 5. 寫入資料庫
            try:
                cursor.execute("""
                    INSERT INTO patients (id, name, drug, condiction, last_date, doctor, return_date, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (pid, name, matched_drug, "計畫匯入", last_date, doctor, "", "Excel 批次匯入"))
                import_count += 1
            except sqlite3.IntegrityError:
                skip_count += 1
                continue

        conn.commit()
        refresh_table()
        messagebox.showinfo("匯入完成", f"成功匯入：{import_count} 筆\n重複跳過：{skip_count} 筆")

    except Exception as e:
        messagebox.showerror("系統錯誤", f"無法讀取檔案：{e}")

def add_patient():
    if not validate_inputs():
        return

    if entry_id.get().strip() == "":
        messagebox.showerror("錯誤", "請輸入病歷號")
        return

    pid = format_id(entry_id.get())
    name = entry_name.get()
    drug = combo_drug.get()
    condiction = combo_condiction.get()
    last_date = entry_date.get()
    doctor = combo_doctor.get()
    return_date = entry_return_date.get()
    note = entry_note.get()

    try:
        cursor.execute("""
        INSERT INTO patients (id, name, drug, condiction, last_date, doctor, return_date, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (pid, name, drug, condiction, last_date, doctor, return_date, note))

        conn.commit()
        refresh_table()
        messagebox.showinfo("成功", "資料已新增")
        clear_entries()

    except Exception as e:
        messagebox.showerror("錯誤", f"資料錯誤或重複\n{e}")

def open_calendar(target_entry):
    top = tk.Toplevel(root)
    top.title("選擇日期")
    top.geometry("300x300")

    # ⭐ Mac穩定關鍵（不要grab_set）
    top.transient(root)
    top.lift()
    top.focus_force()

    cal = Calendar(top, date_pattern="yyyy-mm-dd")
    cal.pack(pady=10)

    def select_date():
        selected = cal.get_date()

        target_entry.delete(0, tk.END)
        target_entry.insert(0, selected)

        top.destroy()

    tk.Button(top, text="確定", command=select_date).pack(pady=10)

# ===== 去識別化 =====
def mask_name(name):
    if len(name) <= 1:
        return name
    elif len(name) == 2:
        return name[0] + "O"
    else:
        return name[0] + "O" + name[-1]

def update_patient():
 pid = format_id(entry_id.get())

 cursor.execute("""
 UPDATE patients
 SET name=?, drug=?, condiction=?, last_date=?, doctor=?, return_date=?, note=?
 WHERE id=?
 """, (
 entry_name.get(),
 combo_drug.get(),
 combo_condiction.get(),
 entry_date.get(),
 combo_doctor.get(),
 entry_return_date.get(),
 entry_note.get(),
 pid
 ))

 conn.commit()
 refresh_table()


def delete_patient():
 selected = tree.selection()

 if not selected:
    messagebox.showwarning("提示", "請先選擇資料")
    return

 pid = format_id(tree.item(selected[0])["values"][0])

 print("👉 DELETE PID:", pid)

 cursor.execute("DELETE FROM patients WHERE id=?", (pid,))

 print("👉 影響筆數:", cursor.rowcount)

 conn.commit()

 refresh_table()

def search_patient():
    keyword = entry_search.get()

    for row in tree.get_children():
        tree.delete(row)

    cursor.execute("""
    SELECT id, name, drug, condiction, last_date, doctor, return_date, note FROM patients 
    WHERE id LIKE ? OR name LIKE ?
    """, (f"%{keyword}%", f"%{keyword}%"))

    for row in cursor.fetchall():
        pid, name, drug, condiction, last_date, doctor, return_date, note = row
        next_date, status = calculate_status(last_date, drug)

        if "🔴" in status:
            tag = "red"
        elif "🟡" in status:
            tag = "yellow"
        else:
            tag = "green"

        display_name = safe_name(name)

        tree.insert("", "end",
            values=(pid, name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )

def filter_by_status():
    selected_status = combo_status.get()
    keyword = entry_search.get()

    for row in tree.get_children():
        tree.delete(row)

    cursor.execute("""
    SELECT id, name, drug, condiction, last_date, doctor, return_date, note 
    FROM patients 
    WHERE id LIKE ? OR name LIKE ?
    """, (f"%{keyword}%", f"%{keyword}%"))

    for row in cursor.fetchall():
        pid, name, drug, condiction, last_date, doctor, return_date, note = row
        next_date, status = calculate_status(last_date, drug)

        # 狀態過濾
        if selected_status != "全部":
            if selected_status not in status:
                continue

        # 顏色
        if "🔴" in status:
            tag = "red"
        elif "🟡" in status:
            tag = "yellow"
        else:
            tag = "green"

        display_name = safe_name(name)

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, status, note),
            tags=(tag,)
        )

def apply_filter():
    global count_label
    keyword = entry_search.get().strip()
    status_filter = combo_search_status.get()
    drug_filter = combo_drug_filter.get()

    print("DEBUG status:", status_filter)

    for row in tree.get_children():
        tree.delete(row)

    cursor.execute("""
    SELECT id, name, drug, condiction, last_date, doctor, return_date, note 
    FROM patients
    WHERE id LIKE ? OR name LIKE ?
    """, (f"%{keyword}%", f"%{keyword}%"))

    for pid, name, drug, condiction, last_date, doctor, return_date, note in cursor.fetchall():

        display_name = safe_name(name)

        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        # 狀態篩選
        if status_filter != "全部" and status_filter != display_status:
            continue

        # 藥物篩選
        if drug_filter != "全部" and drug != drug_filter:
            continue

        # 顏色
        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )
        
        count = len(tree.get_children())
        count_label.config(text=f"目前總筆數：{count} 筆 (篩選後)")

    print("status changed")


# ===== 藥物追蹤天數設定 =====
DRUG_FOLLOWUP_DAYS = {
 "Entecavir 1.0": 30,
 "Entecavir 0.5": 30,
 "Tenofovir 300mg": 30,
 "Tenofovir alafenamide 25mg": 30,
 "Telbivudine": 30,
 "Maviret": 7,
 "Epclusa": 7
}

def calculate_status(end_date_str, drug):
    # 新增：如果日期是空的，直接回傳正常狀態
    if not end_date_str or end_date_str == "":
        return "尚未設定", "NORMAL"
        
    today = datetime.today()
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

    days = DRUG_FOLLOWUP_DAYS.get(drug, 30)
    overdue_date = end_date + timedelta(days=days)

    days_passed = (today - end_date).days

    if days_passed > days:
        return overdue_date.date(), "OVERDUE"
    elif days_passed >= days - 10:
        return overdue_date.date(), "ALERT"
    else:
        return overdue_date.date(), "NORMAL"


def check_due_patients():
    alert_list = []

    cursor.execute("""
    SELECT id, name, drug, last_date 
    FROM patients
    """)

    for pid, name, drug, last_date in cursor.fetchall():
        try:
            next_date, status = calculate_status(last_date, drug)
            display_name = safe_name(name)

            if status == "ALERT":
                days_left = (next_date - datetime.today().date()).days
                alert_list.append(f"{pid} {display_name}（{days_left}天內逾期）")

            elif status == "OVERDUE":
                alert_list.append(f"{pid} {display_name}（已逾期）")

        except:
            continue

    return alert_list

def show_due_alert():
 alert_list = check_due_patients()

 if not alert_list:
    return # 沒有就不跳

 message = "⚠️ 即將逾期提醒\n\n"
 message += "\n".join(alert_list[:10]) # 最多顯示10筆

 if len(alert_list) > 10:
    message += f"\n\n...還有 {len(alert_list)-10} 筆"

 messagebox.showwarning("追蹤提醒", message)

def status_display(status):
 if status == "OVERDUE":
        return "🔴已逾期"
 elif status == "ALERT":
        return "🟡提醒"
 else:
        return "🟢正常" 

# ===== 顯示 =====
FULL_ACCESS_USERS = ["admin"]
USERS = {
 "admin": {"password": "admin123", "role": "admin"},
 "A71002": {"password": "1234", "role": "nurse"}
}

current_user_role = None
current_username = None 

def get_display_name(name):
    try:
        if current_username in FULL_ACCESS_USERS:
            return name
        else:
            return mask_name(name)
    except:
        return mask_name(name)

def safe_name(name):
 return get_display_name(name)

def clear_entries():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_note.delete(0, tk.END)
    
    entry_date.delete(0, tk.END)
    today = datetime.today()

    # 清除回診日期
    entry_return_date.delete(0, tk.END)
    entry_return_date.delete(0, tk.END)

    combo_drug.set("請選擇藥物")
    combo_condiction.set("請選擇用藥條件")
    combo_doctor.set("請選擇醫師")

def clear_placeholder(event):
    if entry_return_date.get() == "請選擇日期":
        entry_return_date.delete(0, tk.END)

def refresh_table():
    global count_label # 確保能存取標籤
    tree.delete(*tree.get_children())

    keyword = entry_search.get().strip()

    if keyword:
        cursor.execute("""
        SELECT id, name, drug, condiction, last_date, doctor, return_date, note 
        FROM patients
        WHERE name LIKE ?
        """, ('%' + keyword + '%',))
    else:
        cursor.execute("""
        SELECT id, name, drug, condiction, last_date, doctor, return_date, note 
        FROM patients
        """)

    for pid, name, drug, condiction, last_date, doctor, return_date, note in cursor.fetchall():

        display_name = safe_name(name)

        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )
        count = len(tree.get_children())
        count_label.config(text=f"目前總筆數：{count} 筆")

def refresh_all():
    clear_entries()
    entry_search.delete(0, tk.END)

    for row in tree.get_children():
        tree.delete(row)

    combo_search_status.current(0)
    combo_drug_filter.current(0)
    combo_export_status.current(0)

    refresh_table()
    show_due_alert()
 
def on_select(event):
    selected = tree.selection()
    if not selected:
        return

    item = selected[0]
    values = tree.item(item, "values")

    print("DEBUG:", values)

    entry_id.delete(0, tk.END)
    entry_id.insert(0, values[0])

    # ⭐ 這裡不要用 values[1]（因為是遮罩）
    cursor.execute("SELECT name FROM patients WHERE id=?", (values[0],))
    real_name = cursor.fetchone()[0]

    entry_name.delete(0, tk.END)
    entry_name.insert(0, real_name)

    combo_drug.set(values[2])
    combo_condiction.set(values[3])

    entry_date.delete(0, tk.END)
    if values[4]:
        entry_date.insert(0, values[4])

    combo_doctor.set(values[5])

    entry_return_date.delete(0, tk.END)
    if values[6]:
        entry_return_date.insert(0, values[6])

    entry_note.delete(0, tk.END)
    entry_note.insert(0, values[9])

def validate_inputs():
    drug = combo_drug.get()
    condiction = combo_condiction.get()
    last_date = entry_date.get()

    if drug == "" or drug == "請選擇藥物":
        messagebox.showerror("錯誤", "請選擇藥物")
        return False

    if condiction == "" or condiction == "請選擇用藥條件":
        messagebox.showerror("錯誤", "請選擇用藥條件")
        return False

    try:
        datetime.strptime(last_date, "%Y-%m-%d")
    except:
        messagebox.showerror("錯誤", "日期格式需為 YYYY-MM-DD")
        return False

    return True

def search_by_drug():
    selected_drug = combo_drug_filter.get()

    for row in tree.get_children():
        tree.delete(row)

    cursor.execute("SELECT id, name, drug, condiction, last_date, doctor, return_date, note FROM patients")

    for pid, name, drug, condiction, last_date, doctor, return_date, note in cursor.fetchall():

        next_date, status = calculate_status(last_date, drug)

        # 只篩選藥物
        if selected_drug != "全部":
            if drug != selected_drug:
                continue

        # 顏色
        tag = "green"
        if "🔴" in status:
            tag = "red"
        elif "🟡" in status:
            tag = "yellow"

        display_name = safe_name(name)

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, status, note),
            tags=(tag,)
        )
        
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "肝炎用藥追蹤名單"

    ws.append(["病歷號", "姓名", "藥物", "用藥條件", "用藥結束日", "醫師", "回診日期", "到期日", "狀態", "備註"])

    selected_status = combo_export_status.get()

    cursor.execute("SELECT id, name, drug, condiction, last_date, doctor, return_date, note FROM patients")

    for pid, name, drug, condiction, last_date, doctor, return_date, note in cursor.fetchall():
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        # 修正篩選
        if selected_status != "全部" and selected_status != display_status:
            continue

        ws.append([
            pid,
            safe_name(name),
            drug,
            condiction,
            last_date,
            doctor,
            return_date,
            str(next_date),
            display_status,
            note
        ])

    default_name = f"肝炎用藥追蹤名單_{datetime.today().strftime('%Y%m%d')}.xlsx"

    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=default_name
    )

    if path:
        wb.save(path)
        messagebox.showinfo("完成", "匯出成功")

# ===== GUI =====
root = tk.Tk()
root.withdraw() # 👈 先隱藏主畫面
root.title("個案追蹤管理系統")
root.protocol("WM_DELETE_WINDOW", exit_program)
root.geometry("1200x650")
root.configure(bg="#e9eef5")

import sys

def on_closing():
    """當使用者關閉視窗或點擊登出時，徹底結束程式"""
    try:
        # 1. 確保資料庫連線有被關閉
        if 'conn' in globals():
            conn.close()
            print("資料庫連線已關閉")
    except Exception as e:
        print(f"關閉連線時發生錯誤: {e}")
    finally:
        # 2. 銷毀所有視窗並徹底退出進程
        root.destroy()
        sys.exit(0)

# 關鍵：告訴視窗，當點擊左上角的「紅色叉叉」時，執行上面的 on_closing 函數
root.protocol("WM_DELETE_WINDOW", on_closing)

# ===== 標題 =====
header = tk.Frame(root, bg="#1f2d3d", height=60)
header.pack(fill="x")

tk.Label(
 header,
 text="🧾 B/C肝用藥追蹤管理系統",
 font=("Microsoft JhengHei", 18, "bold"),
 fg="white",
 bg="#1f2d3d"
).pack(side="left", padx=20, pady=10)

# ===== 主容器 =====
main_frame = tk.Frame(root, bg="#eef2f7")
main_frame.pack(fill="both", expand=True, padx=15, pady=10)

# ===== 卡片：輸入區 =====
input_card = tk.Frame(main_frame, bg="white", bd=1, relief="solid")
input_card.pack(fill="x", pady=5)

for i in range(10):
 input_card.grid_columnconfigure(i, weight=1)

# ===== 第一排 =====
tk.Label(input_card, text="病歷號", bg="white").grid(row=0, column=0, sticky="e", padx=5, pady=5)
entry_id = tk.Entry(input_card)
entry_id.grid(row=0, column=1, sticky="ew", padx=5)

tk.Label(input_card, text="姓名", bg="white").grid(row=0, column=2, sticky="e")
entry_name = tk.Entry(input_card)
entry_name.grid(row=0, column=3, sticky="ew", padx=5)

tk.Label(input_card, text="藥物", bg="white").grid(row=0, column=4, sticky="e")
combo_drug = ttk.Combobox(input_card, values=[
    "請選擇藥物",
    "Entecavir 1.0",
    "Entecavir 0.5",
    "Tenofovir 300mg",
    "Tenofovir alafenamide 25mg",
    "Telbivudine", # 👈 新增
    "Maviret",
    "Epclusa"
], state="readonly")
combo_drug.grid(row=0, column=5, sticky="ew", padx=5)
combo_drug.current(0)

tk.Label(input_card, text="用藥條件", bg="white").grid(row=0, column=6, sticky="e")
combo_condiction = ttk.Combobox(input_card, values=[
 "請選擇用藥條件",
 "e+，肝代償不全",
 "e-，肝代償不全",
 "器官移植長期使用",
 "癌症患者化療中發作長期使用",
 "化療預防性治療",
 "肝硬化長期使用",
 "懷孕婦女(27週)預防垂直感染",
 "器官移植預防性治療",
 "器官移植長期使用",
 "e+，ALT>=5X",
 "e+，2X<=ALT<5",
 "e-，ALT>=2X",
 "肝癌並接受根除性治療",
 "免疫抑制劑治療",
 "e+，ALT",
 "e-，ALT",
 "DAA"
], state="readonly")
combo_condiction.grid(row=0, column=7, sticky="ew", padx=5)
combo_condiction.current(0)

# ===== 第二排 =====
tk.Label(input_card, text="用藥結束日", bg="white").grid(row=1, column=0, sticky="e")

entry_date = tk.Entry(input_card)
entry_date.grid(row=1, column=1, sticky="ew", padx=5)

btn_date = tk.Button(
    input_card,
    text="📅",
    command=lambda: open_calendar(entry_date)
)
btn_date.grid(row=1, column=2, sticky="w")

tk.Label(input_card, text="醫師", bg="white").grid(row=1, column=2, sticky="e")
combo_doctor = ttk.Combobox(input_card, values=[
 "請選擇醫師",
 "Alice",
 "Bob",
 "Cathy",
 "Denil"
], state="readonly")
combo_doctor.grid(row=1, column=3, sticky="ew", padx=5)
combo_doctor.current(0)

tk.Label(input_card, text="回診日期", bg="white").grid(row=1, column=4, sticky="e")

entry_return_date = tk.Entry(input_card)
entry_return_date.grid(row=1, column=5, sticky="ew", padx=5)

btn_return = tk.Button(
    input_card,
    text="📅",
    command=lambda: open_calendar(entry_return_date)
)
btn_return.grid(row=1, column=6, sticky="w")

# 清空顯示
entry_return_date.delete(0, tk.END)

tk.Label(input_card, text="備註", bg="white").grid(row=1, column=6, sticky="e")
entry_note = tk.Entry(input_card)
entry_note.grid(row=1, column=7, columnspan=3, sticky="ew", padx=5)

# ===== 查詢區 =====
search_card = tk.Frame(main_frame, bg="white", bd=1, relief="solid")
search_card.pack(fill="x", pady=5)

for i in range(10):
 search_card.grid_columnconfigure(i, weight=1)

tk.Label(search_card, text="關鍵字", bg="white").grid(row=0, column=0, sticky="e")
entry_search = tk.Entry(search_card)
entry_search.grid(row=0, column=1, sticky="ew", padx=5)

tk.Label(search_card, text="狀態查詢", bg="white").grid(row=0, column=2, sticky="e")
combo_search_status = ttk.Combobox(search_card, values=[
 "全部", "🟢正常", "🟡提醒", "🔴已逾期"
], state="readonly")
combo_search_status.grid(row=0, column=3, sticky="ew")
combo_search_status.current(0)
entry_search.bind("<KeyRelease>", lambda e: apply_filter())

tk.Label(search_card, text="藥物篩選", bg="white").grid(row=0, column=4, sticky="e")
combo_drug_filter = ttk.Combobox(search_card, values=[
 "全部",
 "Entecavir 1.0",
 "Entecavir 0.5",
 "Tenofovir 300mg",
 "Tenofovir alafenamide 25mg",
 "Maviret",
 "Epclusa"
], state="readonly")
combo_drug_filter.grid(row=0, column=5, sticky="ew")
combo_drug_filter.current(0)

# 匯出
tk.Label(search_card, text="匯出狀態").grid(row=0, column=8)

combo_export_status = ttk.Combobox(search_card, values=[
 "全部",
 "🟢正常",
 "🟡提醒",
 "🔴已逾期"
], state="readonly")

combo_export_status.grid(row=0, column=9)
combo_export_status.current(0)

# ===== 按鈕區 =====
btn_frame = tk.Frame(main_frame, bg="#eef2f7")
btn_frame.pack(fill="x", pady=5)

style = ttk.Style()
style.theme_use("clam") # 讓顏色可控

style.map("Treeview",
    background=[("selected", "#3498db")],
    foreground=[("selected", "white")]
)

def create_btn(text, cmd, color):
 style_name = f"{text}.TButton"

 style.configure(style_name, background=color, foreground="white")

 btn = ttk.Button(btn_frame, text=text, command=cmd, style=style_name)

 return btn

create_btn("➕ 新增", add_patient, "#27ae60").pack(side="left", padx=6)
create_btn("✏️ 修改", update_patient, "#2980b9").pack(side="left", padx=6)
create_btn("🗑️ 刪除", delete_patient, "#c0392b").pack(side="left", padx=6)
create_btn("🔍 查詢", apply_filter, "#d68910").pack(side="left", padx=6)
create_btn("🔄更新", refresh_all, "#7f8c8d").pack(side="left", padx=6)
create_btn("📥 匯入試辦名單", import_excel_general, "#16a085").pack(side="left", padx=6)
create_btn("📤 匯出", export_excel, "#8e44ad").pack(side="left", padx=6)

# 👉 右側帳號操作
create_btn("🚪 登出", logout, "#2c3e50").pack(side="right", padx=6)
create_btn("🔐 重新登入", relogin, "#34495e").pack(side="right", padx=6)

# ===== 表格區 =====
user_label = tk.Label(
 header,
 text="未登入", 
 font=("Microsoft JhengHei", 10),
 fg="white",
 bg="#1f2d3d"
)
user_label.pack(side="right", padx=20)

table_frame = tk.Frame(main_frame)
table_frame.pack(fill="both", expand=True)

count_label = tk.Label(
    main_frame, 
    text="目前總筆數：0 筆", 
    font=("Microsoft JhengHei", 10, "bold"),
    bg="#eef2f7",
    fg="#2c3e50"
)
count_label.pack(side="bottom", anchor="w", pady=5)

# 設定 table_frame 的 grid 權重，確保 treeview 佔據剩餘空間
table_frame.grid_rowconfigure(0, weight=1)
table_frame.grid_columnconfigure(0, weight=1)

columns = ("病歷號", "姓名", "藥物", "用藥條件", "結束日期", "醫師", "回診日期", "到期日", "狀態", "備註")

tree = ttk.Treeview(table_frame, columns=columns, show="headings")
tree.grid(row=0, column=0, sticky="nsew") # 使用 sticky="nsew" 填滿

# 設定欄位
for col in columns:
    tree.heading(col, text=col, anchor="center") 
    tree.column(col, anchor="w", width=120)

tree.column("備註", width=400)

style.configure("Treeview",
 font=("Microsoft JhengHei", 10),
 rowheight=28
)

style.configure("Treeview.Heading",
 font=("Microsoft JhengHei", 10, "bold")
)

tree.tag_configure("red", background="#ffdddd")
tree.tag_configure("yellow", background="#fff3cd")
tree.tag_configure("green", background="#e8f5e9")

# ===== 滾動條 (正確綁定) =====

# 垂直捲軸
scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
scroll_y.grid(row=0, column=1, sticky="ns") # ns 表示由北向南延伸
tree.configure(yscrollcommand=scroll_y.set)

# 水平捲軸
scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
scroll_x.grid(row=1, column=0, sticky="ew") # ew 表示由東向西延伸
tree.configure(xscrollcommand=scroll_x.set)

# ===== 綁定 =====
tree.bind("<<TreeviewSelect>>", on_select)

combo_search_status.bind("<<ComboboxSelected>>", lambda e: apply_filter())
combo_drug_filter.bind("<<ComboboxSelected>>", lambda e: apply_filter())

# 啟動
show_login()
root.mainloop()

conn.close()