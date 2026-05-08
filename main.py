import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime, timedelta

# --- 1. 第三方套件 (Standard Libraries 之後) ---
import pandas as pd              # 如果你有用到 Excel 匯入
from openpyxl import Workbook
from tkcalendar import Calendar

# --- 2. 自定義模組 (最後引入) ---
import database  
import ui_pages

# --- 3. 程式初始化動作 ---
# 確保資料庫在介面出來前就先準備好
database.initialize_db()


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
        
        # (這裡的欄位對照和檢查邏輯保持不變...)
        col_map = {'病歷號碼': 'id', '病患姓名': 'name', '試辦種類': 'drug', '主治醫師': 'doctor', '用藥結束日': 'last_date'}
        if not all(c in df.columns for c in col_map.keys()):
            messagebox.showerror("格式錯誤", "檔案缺少必要欄位")
            return

        import_count = 0
        skip_count = 0
        
        for _, row in df.iterrows():
            # (這裡的資料處理邏輯保持不變：pid, name, matched_drug, last_date 的產生方式...)
            # ... [省略中間處理代碼] ...
            
            # 準備打包資料
            new_data = (pid, name, matched_drug, "計畫匯入", last_date, doctor, "", "Excel 批次匯入")

            # --- 💡 這裡改為呼叫 database.py 的函數 ---
            if database.try_add_patient(new_data):
                import_count += 1
            else:
                skip_count += 1

        # 結束後刷新畫面
        refresh_table()
        messagebox.showinfo("匯入完成", f"成功匯入：{import_count} 筆\n重複跳過：{skip_count} 筆")

    except Exception as e:
        messagebox.showerror("系統錯誤", f"無法讀取檔案：{e}")


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


def delete_patient():
    selected = tree.selection()

    if not selected:
        messagebox.showwarning("提示", "請先選擇資料")
        return

    # 取得 ID 並格式化
    pid = format_id(tree.item(selected[0])["values"][0])

    print("👉 準備刪除 PID:", pid)

    # --- 這裡改為呼叫 database.py 的函數 ---
    database.delete_patient(pid) 

    # 既然 database.py 已經 commit 並 close 了，主程式只要重新整理畫面就好
    refresh_table()
    messagebox.showinfo("成功", f"個案 {pid} 已刪除")

def search_patient():
    keyword = entry_search.get()

    # 清空現有表格
    for row in tree.get_children():
        tree.delete(row)

    # --- 💡 這裡改為呼叫 database.py 的函數 ---
    data = database.search_patients(keyword)

    # 處理並填入資料
    for row in data:
        pid, name, drug, condiction, last_date, doctor, return_date, note = row
        
        # 這裡的邏輯建議直接呼叫你原本寫好的計算函數
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)
        
        # 決定顏色標籤 (tag)
        if "OVERDUE" in raw_status:
            tag = "red"
        elif "ALERT" in raw_status:
            tag = "yellow"
        else:
            tag = "green"

        display_name = safe_name(name)

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )

def filter_by_status():
    selected_status = combo_status.get()
    keyword = entry_search.get()

    # 1. 清空表格
    for row in tree.get_children():
        tree.delete(row)

    # 2. --- 💡 呼叫 database.py 取得資料 ---
    data = database.search_patients(keyword)

    # 3. 跑一個迴圈處理資料就好
    for pid, name, drug, condiction, last_date, doctor, return_date, note in data:
        
        # 這裡處理計算邏輯
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status) # 建議轉換成文字顯示

        # 狀態過濾
        if selected_status != "全部":
            # 這裡建議比對 display_status (例如 "逾期")
            if selected_status != display_status:
                continue

        # 顏色判定
        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        display_name = safe_name(name)

        # 4. 插入表格
        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )

def apply_filter():
    global count_label
    keyword = entry_search.get().strip()
    status_filter = combo_search_status.get()
    drug_filter = combo_drug_filter.get()

    # 1. 清空表格
    for row in tree.get_children():
        tree.delete(row)

    # 2. --- 💡 這裡改為呼叫 database.py ---
    # 拿到初步符合關鍵字的資料
    data = database.search_patients(keyword)

    # 3. 在主程式進行「狀態」與「藥物」的細部篩選 (這屬於 UI 顯示邏輯)
    for pid, name, drug, condiction, last_date, doctor, return_date, note in data:
        display_name = safe_name(name)
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        # 狀態篩選邏輯保持不變
        if status_filter != "全部" and status_filter != display_status:
            continue

        # 藥物篩選邏輯保持不變
        if drug_filter != "全部" and drug != drug_filter:
            continue

        # 顏色決定
        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )
        
    # 4. 更新筆數 (移到迴圈外，效能更好)
    count = len(tree.get_children())
    count_label.config(text=f"目前總筆數：{count} 筆 (篩選後)")


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

    # --- 💡 1. 呼叫 database.py 取得所有資料 ---
    data = database.get_all_patients()

    # --- 💡 2. 進行邏輯判斷 ---
    for row in data:
        # 注意：根據我們 database.py 的設計，row 的順序是：
        # (id, name, drug, condiction, last_date, doctor, return_date, note)
        pid = row[0]
        name = row[1]
        drug = row[2]
        last_date = row[4]

        try:
            # 這裡的計算邏輯保持不變
            next_date, raw_status = calculate_status(last_date, drug)
            display_name = safe_name(name)

            if raw_status == "ALERT":
                days_left = (next_date - datetime.today().date()).days
                alert_list.append(f"{pid} {display_name}（{days_left} 天內逾期）")

            elif raw_status == "OVERDUE":
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

def clear_placeholder(event):
    if entry_return_date.get() == "請選擇日期":
        entry_return_date.delete(0, tk.END)

def refresh_table():
    global count_label
    tree.delete(*tree.get_children())

    keyword = entry_search.get().strip()

    # 1. 取得原始資料
    if keyword:
        data = database.search_patients(keyword)
    else:
        data = database.get_all_patients()

    # 2. 處理資料並填入表格
    for pid, name, drug, condiction, last_date, doctor, return_date, note in data:
        
        # --- 💡 這些計算邏輯必須留在迴圈內，因為每個人都不一樣 ---
        display_name = safe_name(name) 
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        # 根據狀態決定顏色
        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )
    
    # 3. 💡 全部處理完後，最後才更新總筆數 (在迴圈外)
    count = len(tree.get_children())
    count_label.config(text=f"目前總筆數：{count} 筆")

def refresh_all():
    
    # 清空搜尋關鍵字
    entry_search.delete(0, tk.END)

    #  重設所有篩選下拉選單回到「全部」
    combo_search_status.current(0)
    combo_drug_filter.current(0)
    combo_export_status.current(0)

    #  重新整理表格資料
    # 不需要手動執行 tree.delete，因為 refresh_table() 內部通常已經寫了清空邏輯
    refresh_table()
    
    #  再次檢查是否有逾期提醒
    show_due_alert()
 
def on_select(event):
    selected_items = tree.selection()
    if not selected_items:
        return

    # 取得選中行的資料
    item_attr = tree.item(selected_items[0])
    values = item_attr['values']
    
    # 你可以保留這一行，在下方的狀態列或 Console 顯示目前選中誰
    print(f"目前選中：{values[1]} (病歷號: {values[0]})")

def on_item_double_click(event):
    selected = tree.selection()
    if not selected: return
    
    # 確保抓到的 ID 是乾淨的
    raw_id = tree.item(selected[0])['values'][0]
    patient_id = str(raw_id).strip() # 轉成字串並去空格
    
    print(f"目前選中：{tree.item(selected[0])['values'][1]} (病歷號: {patient_id})")
    
    ui_pages.show_edit_window(root, patient_id, refresh_table)

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

    # 1. 清空表格
    for row in tree.get_children():
        tree.delete(row)

    # 2. --- 💡 呼叫 database.py 取得所有資料 ---
    data = database.get_all_patients()

    # 3. 處理與過濾
    for pid, name, drug, condiction, last_date, doctor, return_date, note in data:

        # 這裡的計算邏輯
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        # 藥物篩選邏輯
        if selected_drug != "全部":
            if drug != selected_drug:
                continue

        # 顏色決定 (建議統一使用 raw_status 判斷，比判斷符號更精準)
        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        display_name = safe_name(name)

        tree.insert("", "end",
            values=(pid, display_name, drug, condiction, last_date, doctor, return_date, next_date, display_status, note),
            tags=(tag,)
        )
        
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "肝炎用藥追蹤名單"

    ws.append(["病歷號", "姓名", "藥物", "用藥條件", "用藥結束日", "醫師", "回診日期", "到期日", "狀態", "備註"])

    selected_status = combo_export_status.get()

    # --- 💡 這裡改為呼叫 database.py ---
    data = database.get_all_patients()

    for row in data:
        # row 順序: (id, name, drug, condiction, last_date, doctor, return_date, note)
        pid, name, drug, condiction, last_date, doctor, return_date, note = row
        
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        # 篩選邏輯
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
root.withdraw() # 👈 先隱藏，避免畫面閃爍
root.title("個案追蹤管理系統")

# 1. 確保資料庫已就緒
try:
    database.initialize_db() # 呼叫我們搬到 database.py 的初始化函數
except Exception as e:
    messagebox.showerror("啟動失敗", f"無法連結資料庫：{e}")
    sys.exit()

# 2. 設定關閉程式的協議
root.protocol("WM_DELETE_WINDOW", exit_program)

# 3. 視窗尺寸與背景
root.geometry("1200x650")
root.configure(bg="#e9eef5")

# 4. 在所有組件(Widget)都建立好後，記得在 main.py 最後面補上：
# root.deiconify()

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

create_btn("➕ 新增", lambda: ui_pages.show_add_patient_window(root, refresh_table), "#27ae60").pack(side="left", padx=6)
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

tree.column("病歷號", width=100, anchor="center")
tree.column("姓名", width=100, anchor="center")
tree.column("藥物", width=200)
tree.column("用藥條件", width=200)
tree.column("結束日期", width=120, anchor="center")
tree.column("醫師", width=100, anchor="center")
tree.column("回診日期", width=120, anchor="center")
tree.column("到期日", width=120, anchor="center")
tree.column("狀態", width=100, anchor="center")
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
tree.bind("<Double-1>", on_item_double_click)

# 綁定雙擊事件
tree.bind("<Double-1>", on_item_double_click)

combo_search_status.bind("<<ComboboxSelected>>", lambda e: apply_filter())
combo_drug_filter.bind("<<ComboboxSelected>>", lambda e: apply_filter())

# --- 視窗關閉的提醒邏輯 ---
def on_closing():
    # 彈出確認視窗
    if messagebox.askyesno("結束程式", "要關閉程式了嗎？\n\n別忘了備份 patients.db 到雲端喔！"):
        root.destroy()

# --- 啟動程序 ---

# 1. 綁定關閉視窗的協定 (要在 mainloop 之前)
root.protocol("WM_DELETE_WINDOW", on_closing)

# 2. 啟動登入畫面 (或是直接顯示主畫面)
show_login()

# 3. 進入主迴圈
root.mainloop()

# 4. 這裡不需要 conn.close()，因為資料庫連線在 database.py 裡已經處理好了