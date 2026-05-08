# ===== 自動安裝套件 =====
import subprocess
import sys

def install(package):
    subprocess.check_call([
        sys.executable,
        "-m",
        "pip",
        "install",
        package,
        "--break-system-packages"
    ])

try:
    from openpyxl import Workbook
except ImportError:
    install("openpyxl")
    from openpyxl import Workbook


import sqlite3
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox, ttk
from openpyxl import Workbook
from tkinter import filedialog

# ===== DB =====
conn = sqlite3.connect("patients.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS patients (
    id TEXT PRIMARY KEY,
    name TEXT,
    drug TEXT,
    last_date TEXT,
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
    login_window.geometry("350x280")
    login_window.configure(bg="#ecf0f1")
    login_window.grab_set()

    # ===== 卡片 =====
    card = tk.Frame(login_window, bg="white", bd=0)
    card.place(relx=0.5, rely=0.5, anchor="center", width=300, height=220)

    # ===== 標題 =====
    tk.Label(
        card,
        text="🔐 系統登入",
        font=("Microsoft JhengHei", 16, "bold"),
        bg="white",
        fg="#2c3e50"
    ).pack(pady=10)

    # ===== 帳號 =====
    tk.Label(card, text="帳號", bg="white", anchor="w").pack(fill="x", padx=20)
    entry_user = tk.Entry(card, bd=1, relief="solid", font=("Microsoft JhengHei", 10))
    entry_user.pack(padx=20, pady=5, fill="x")

    # ===== 密碼 =====
    tk.Label(card, text="密碼", bg="white", anchor="w").pack(fill="x", padx=20)
    entry_pass = tk.Entry(card, show="*", bd=1, relief="solid", font=("Microsoft JhengHei", 10))
    entry_pass.pack(padx=20, pady=5, fill="x")

    # ===== 登入功能 =====
    def login():
        global current_user_role, current_username

        username = entry_user.get()
        password = entry_pass.get()

        if username in USERS and USERS[username]["password"] == password:
            current_user_role = USERS[username]["role"]
            current_username = username   

            login_window.destroy()
            root.deiconify()

            entry_search.delete(0, tk.END)

            root.after(100, refresh_table)
            root.after(500, show_due_alert)

            apply_permissions()

            refresh_table()
            root.update_idletasks()

            root.after(500, show_due_alert)

        else:
            messagebox.showerror("錯誤", "帳號或密碼錯誤")

    # ===== 按鈕 hover 效果 =====
    def on_enter(e):
        btn_login.config(bg="#2980b9")

    def on_leave(e):
        btn_login.config(bg="#3498db")

    # ===== 登入按鈕 =====
    btn_login = tk.Button(
        card,
        text="登入",
        command=login,
        bg="#3498db",
        fg="white",
        font=("Microsoft JhengHei", 10, "bold"),
        relief="flat",
        cursor="hand2"
    )
    btn_login.pack(pady=15, ipadx=10, ipady=5)

    btn_login.bind("<Enter>", on_enter)
    btn_login.bind("<Leave>", on_leave)

    # ===== Enter鍵登入 =====
    login_window.bind("<Return>", lambda e: login())

    # 預設 focus
    entry_user.focus()

    error_label = tk.Label(card, text="", fg="red", bg="white")
    error_label.pack()

    # 登入失敗時：
    error_label.config(text="帳號或密碼錯誤")

def apply_permissions():
    if current_user_role == "nurse":
        btn_delete.config(state="normal")   # 可刪
    else:
        btn_delete.config(state="disabled") # admin不可刪

    refresh_table()

def format_id(pid):
    return str(pid).strip().zfill(3)

#登出按鈕
def logout():
    global current_user_role

    current_user_role = None

    for row in tree.get_children():
        tree.delete(row)

    root.withdraw()   #  隱藏主畫面
    show_login()      #  回登入畫面

#重新登入按鈕
def relogin():
    logout()

# ===== CRUD =====
def add_patient():
    if not validate_inputs():
        return

    pid = format_id(entry_id.get())
    name = entry_name.get()
    drug = combo_drug.get()
    last_date = entry_date.get()

    try:
        return_date = entry_return_date.get()
        note = entry_note.get()

        cursor.execute("""
        INSERT INTO patients (id, name, drug, last_date, return_date, note)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (pid, name, drug, last_date, return_date, note))
        conn.commit()
        refresh_table()
    except:
        messagebox.showerror("錯誤", "資料錯誤或病歷號重複")

    messagebox.showinfo("成功", "資料已新增")
    clear_entries()    

    
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
        SET name=?, drug=?, last_date=?, return_date=?, note=?
        WHERE id=?
    """, (
        entry_name.get(),
        combo_drug.get(),
        entry_date.get(),
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
        SELECT id, name, drug, last_date, return_date, note FROM patients 
        WHERE id LIKE ? OR name LIKE ?
    """, (f"%{keyword}%", f"%{keyword}%"))

    for row in cursor.fetchall():
        pid, name, drug, last_date, return_date, note = row
        next_date, status = calculate_status(last_date, drug)

        if "🔴" in status:
            tag = "red"
        elif "🟡" in status:
            tag = "yellow"
        else:
            tag = "green"

        display_name = safe_name(name)    

        tree.insert("", "end",
            values=(pid, display_name, drug, last_date, return_date, next_date, status, note),
            tags=(tag,)
        )

def filter_by_status():
    selected_status = combo_status.get()
    keyword = entry_search.get()

    for row in tree.get_children():
        tree.delete(row)

    cursor.execute("""
        SELECT id, name, drug, last_date, return_date, note 
        FROM patients 
        WHERE id LIKE ? OR name LIKE ?
    """, (f"%{keyword}%", f"%{keyword}%"))

    for row in cursor.fetchall():
        pid, name, drug, last_date, return_date, note = row
        next_date, status = calculate_status(last_date, drug)

        # 👉 狀態過濾
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
            values=(pid, display_name, drug, last_date, return_date, next_date, status, note),
            tags=(tag,)
        )

def apply_filter():
    global entry_search
    keyword = entry_search.get().strip()
    status_filter = combo_search_status.get()
    drug_filter = combo_drug_filter.get()

    print("DEBUG status:", status_filter)  # ⭐ 先看有沒有抓到

    for row in tree.get_children():
        tree.delete(row)

    cursor.execute("""
        SELECT id, name, drug, last_date, return_date, note 
        FROM patients
        WHERE id LIKE ? OR name LIKE ?
    """, (f"%{keyword}%", f"%{keyword}%"))

    for pid, name, drug, last_date, return_date, note in cursor.fetchall():

        display_name = safe_name(name)

        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        #  狀態篩選
        if status_filter != "全部" and status_filter != display_status:
            continue

        #  藥物篩選
        if drug_filter != "全部" and drug != drug_filter:
            continue

        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        tree.insert("", "end",
            values=(pid, display_name, drug, last_date, return_date, next_date, display_status, note),
            tags=(tag,)
        )

        print("status changed")


# ===== 藥物追蹤天數設定 =====
DRUG_FOLLOWUP_DAYS = {
    "Entecavir 1.0": 30,
    "Entecavir 0.5": 30,
    "Tenofovir 300mg": 30,
    "Tenofovir alafenamide 25mg": 30,
    "Maviret": 7,
    "Epclusa": 7
}

def calculate_status(end_date_str, drug):
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
        return  # 沒有就不跳

    message = "⚠️ 即將逾期提醒\n\n"
    message += "\n".join(alert_list[:10])  # 最多顯示10筆

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
FULL_ACCESS_USERS = ["A71002"]
USERS = {
    "A71002": {"password": "1234", "role": "nurse"},
    "admin": {"password": "admin123", "role": "admin"}
}

current_user_role = None
current_username = None   

def get_display_name(name):
    try:
        if current_username in FULL_ACCESS_USERS:
            return name
        return mask_name(name)
    except:
        return mask_name(name)

def safe_name(name):
    return get_display_name(name)

def clear_entries():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_date.delete(0, tk.END)
    entry_return_date.delete(0, tk.END)
    entry_note.delete(0, tk.END)
   
    combo_drug.set("請選擇藥物")

def refresh_table():
    tree.delete(*tree.get_children())
    entry_search.delete(0, tk.END)

    keyword = entry_search.get().strip()

    if keyword:
        cursor.execute("""
            SELECT id, name, drug, last_date, return_date, note 
            FROM patients
            WHERE name LIKE ?
        """, ('%' + keyword + '%',))
    else:
        cursor.execute("""
            SELECT id, name, drug, last_date, return_date, note 
            FROM patients
        """)

    for pid, name, drug, last_date, return_date, note in cursor.fetchall():

        display_name = safe_name(name)

        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        tag = "green"
        if raw_status == "OVERDUE":
            tag = "red"
        elif raw_status == "ALERT":
            tag = "yellow"

        tree.insert("", "end",
            values=(pid, display_name, drug, last_date, return_date, next_date, display_status, note),
            tags=(tag,)
        )

def refresh_all():
    clear_entries()

    # 清空 tree
    for row in tree.get_children():
        tree.delete(row)

    # 清空搜尋欄
    entry_search.delete(0, tk.END)

    # 重置 combobox
    combo_status.current(0)
    combo_search_status.current(0)
    combo_drug_filter.current(0)
    combo_export_status.current(0)

    # 清空選取狀態（避免殘留選取）
    tree.selection_remove(tree.selection())

    # 重新載入資料
    refresh_table()
    show_due_alert()
    
def on_select(event):
    selected = tree.selection()
    if not selected:
        return

    values = tree.item(selected[0])["values"]

    entry_id.delete(0, tk.END)
    entry_id.insert(0, values[0])

    entry_name.delete(0, tk.END)
    entry_name.insert(0, values[1])

    combo_drug.set(values[2])

    entry_date.delete(0, tk.END)
    entry_date.insert(0, values[3])

    entry_return_date.delete(0, tk.END)
    entry_return_date.insert(0, values[4])

    entry_note.delete(0, tk.END)
    entry_note.insert(0, values[7])

def validate_inputs():
    drug = combo_drug.get()
    last_date = entry_date.get()

    if drug == "" or drug == "請選擇藥物":
        messagebox.showerror("錯誤", "請選擇藥物")
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

    cursor.execute("SELECT id, name, drug, last_date, return_date, note FROM patients")

    for pid, name, drug, last_date, return_date, note in cursor.fetchall():
        next_date, status = calculate_status(last_date, drug)

        # 👉 只篩選藥物
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
            values=(pid, display_name, drug, last_date, return_date, next_date, status, note),
            tags=(tag,)
        )

def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "個案資料"

    ws.append(["病歷號", "姓名", "藥物", "結束日期", "回診日期", "追蹤日期", "狀態"])

    selected_status = combo_export_status.get()

    cursor.execute("SELECT id, name, drug, last_date, return_date, note FROM patients")

    for pid, name, drug, last_date, return_date, note in cursor.fetchall():
        next_date, raw_status = calculate_status(last_date, drug)
        display_status = status_display(raw_status)

        #  修正篩選
        if selected_status != "全部" and selected_status != display_status:
            continue

        display_name = safe_name(name)

        ws.append([
            pid,
            safe_name(name),
            drug,
            last_date,
            return_date,
            str(next_date),
            display_status
        ])

    path = filedialog.asksaveasfilename(defaultextension=".xlsx")

    if path:
        wb.save(path)
        messagebox.showinfo("完成", "匯出成功")


# ===== GUI =====
root = tk.Tk()
root.withdraw()   # 👈 先隱藏主畫面
root.title("個案追蹤管理系統")
root.geometry("1140x600")
root.configure(bg="#f4f6f9")

title = tk.Label(
    root,
    text="🧾 B/C肝用藥追蹤管理系統",
    font=("Microsoft JhengHei", 20, "bold"),
    bg="#1f2d3d",
    fg="white",
    pady=15
)
title.pack(fill="x")

frame = tk.Frame(root, bg="white", bd=2, relief="ridge")
frame.pack(padx=15, pady=10, fill="x")
frame.configure(highlightbackground="#dcdde1", highlightthickness=1)

# ===== 輸入欄位 =====
tk.Label(frame, text="病歷號").grid(row=0, column=0)
entry_id = tk.Entry(frame, font=("Microsoft JhengHei", 10), bd=1, relief="solid")
entry_id.grid(row=0, column=1, padx=5, pady=5)
entry_name = tk.Entry(frame, font=("Microsoft JhengHei", 10))
entry_name.grid(row=0, column=3)
entry_date = tk.Entry(frame, font=("Microsoft JhengHei", 10))
entry_date.grid(row=0, column=7)

tk.Label(frame, text="姓名").grid(row=0, column=2)

tk.Label(frame, text="備註").grid(row=1, column=4)

entry_note = tk.Entry(frame, width=30)
entry_note.grid(row=1, column=5, padx=5)

def bind_enter(event=None):
    apply_filter()
    return "break"

tk.Label(frame, text="藥物").grid(row=0, column=4)
combo_drug = ttk.Combobox(frame, values=[
    "請選擇藥物",
    "Entecavir 1.0",
    "Entecavir 0.5",
    "Tenofovir 300mg",
    "Tenofovir alafenamide 25mg",
    "Maviret",
    "Epclusa"
], state="readonly")
combo_drug.grid(row=0, column=5)
combo_drug.current(0)

tk.Label(frame, text="結束日期").grid(row=0, column=6)
entry_date.grid(row=0, column=7)

tk.Label(frame, text="關鍵字").grid(row=0, column=8)

entry_search = tk.Entry(frame)
entry_search.grid(row=0, column=9)
tk.Button(frame, text="查詢", command=apply_filter).grid(row=0, column=10, padx=5)

for widget in [entry_id, entry_name]:
    widget.bind("<KeyRelease>", lambda e: apply_filter())

# ===== 第二排 =====
tk.Label(frame, text="回診日期").grid(row=1, column=0)
entry_return_date = tk.Entry(frame)
entry_return_date.grid(row=1, column=1)

for widget in [entry_search, entry_id, entry_name, entry_date]:
    widget.bind("<Return>", bind_enter)

entry_return_date.bind("<Return>", bind_enter)

tk.Label(frame, text="狀態查詢").grid(row=1, column=2)

combo_search_status = ttk.Combobox(frame, values=[
    "全部",
    "🟢正常",
    "🟡提醒",
    "🔴已逾期"
], state="readonly")

combo_search_status.grid(row=1, column=3)
combo_search_status.current(0)

tk.Label(frame, text="藥物篩選").grid(row=1, column=6)

combo_drug_filter = ttk.Combobox(frame, values=[
    "全部",
    "Entecavir 1.0",
    "Entecavir 0.5",
    "Tenofovir 300mg",
    "Tenofovir alafenamide 25mg",
    "Maviret",
    "Epclusa"
], state="readonly")

combo_drug_filter.current(0)
combo_drug_filter.grid(row=1, column=7)

# 查詢欄位
combo_search_status.bind("<<ComboboxSelected>>", lambda e: apply_filter())
combo_drug_filter.bind("<<ComboboxSelected>>", lambda e: apply_filter())

def create_btn(parent, text, command, style):
    btn = tk.Label(
        parent,
        text=text,
        bg=style["bg"],
        fg="white",
        font=("Microsoft JhengHei", 11, "bold"),
        padx=14,
        pady=7,
        cursor="hand2",
        bd=0
    )

    def on_enter(e):
        btn.config(bg=style["hover"])

    def on_leave(e):
        btn.config(bg=style["bg"])

    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    btn.bind("<Button-1>", lambda e: command())

    return btn

BTN_STYLE = {
    "add": {"bg": "#2ecc71", "hover": "#27ae60"},
    "update": {"bg": "#3498db", "hover": "#2980b9"},
    "delete": {"bg": "#e74c3c", "hover": "#c0392b"},
    "search": {"bg": "#f39c12", "hover": "#d68910"},
    "refresh": {"bg": "#95a5a6", "hover": "#7f8c8d"},
    "export": {"bg": "#9b59b6", "hover": "#8e44ad"},
    "logout": {"bg": "#34495e", "hover": "#2c3e50"},
    "relogin": {"bg": "#16a085", "hover": "#138d75"}
}

btn_frame = tk.Frame(frame, bg="white")
btn_frame.grid(row=2, column=0, columnspan=10, pady=10)

create_btn(btn_frame, "➕ 新增", add_patient, BTN_STYLE["add"]).pack(side="left", padx=5)
create_btn(btn_frame, "✏️ 修改", update_patient, BTN_STYLE["update"]).pack(side="left", padx=5)
create_btn(btn_frame, "🗑️ 刪除", delete_patient, BTN_STYLE["delete"]).pack(side="left", padx=5)
create_btn(btn_frame, "🔍 查詢", apply_filter, BTN_STYLE["search"]).pack(side="left", padx=5)
create_btn(btn_frame, "🔄 刷新", refresh_all, BTN_STYLE["refresh"]).pack(side="left", padx=5)
create_btn(btn_frame, "📤 匯出", export_excel, BTN_STYLE["export"]).pack(side="left", padx=5)

top_right = tk.Frame(root, bg="#1f2d3d")
top_right.place(relx=1.0, x=-10, y=8, anchor="ne")

btn_logout = create_btn(top_right, "登出", logout, BTN_STYLE["logout"])
btn_logout.pack(side="left", padx=5)

btn_relogin = create_btn(top_right, "重新登入", relogin, BTN_STYLE["relogin"])
btn_relogin.pack(side="left", padx=5)

style = ttk.Style()
style.theme_use("clam")  

style.configure("Treeview",
    font=("Microsoft JhengHei", 10),
    rowheight=30,
    background="white",
    fieldbackground="white"
)

style.configure("Treeview.Heading",
    font=("Microsoft JhengHei", 11, "bold"),
    background="#34495e",
    foreground="white"
)

last_hover = None

def on_motion(event):
    global last_hover

    row = tree.identify_row(event.y)

    if row == last_hover:
        return

    if last_hover:
        tags = list(tree.item(last_hover, "tags"))
        if "hover" in tags:
            tags.remove("hover")
        tree.item(last_hover, tags=tags)

    if row:
        tags = list(tree.item(row, "tags"))
        if "hover" not in tags:
            tags.append("hover")
        tree.item(row, tags=tags)
        last_hover = row

    if row in tree.selection():
        return

def on_leave(event):
    tree.tag_remove("hover", *tree.get_children())

style.map("Treeview",
    background=[("selected", "#74b9ff")],
    foreground=[("selected", "black")]
)

# ===== 表格區（升級版）=====
table_frame = tk.Frame(root)
table_frame.pack(fill="both", expand=True)

columns = ("病歷號", "姓名", "藥物", "結束日期", "回診日期", "追蹤日期", "狀態", "備註")

table_frame.grid_rowconfigure(0, weight=1)
table_frame.grid_columnconfigure(0, weight=1)

# ===== Treeview =====
tree = ttk.Treeview(
    table_frame,
    columns=columns,
    show="headings",
    selectmode="browse"
)

# ===== Scrollbar =====
y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)

tree.configure(
    yscrollcommand=y_scroll.set,
    xscrollcommand=x_scroll.set
)

# ===== 排版 =====
tree.grid(row=0, column=0, sticky="nsew")
y_scroll.grid(row=0, column=1, sticky="ns")
x_scroll.grid(row=1, column=0, sticky="ew")

# ===== 欄位設定（升級：自動縮放感）=====
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120, minwidth=80, anchor="center")

tree.column("姓名", width=100)
tree.column("備註", width=200)

# ===== tag 顏色優化 =====
tree.tag_configure("hover", background="#eef6ff")
tree.tag_configure("red", background="#ffe5e5")
tree.tag_configure("yellow", background="#fff7d6")
tree.tag_configure("green", background="#e8f7ea")

# ===== hover 效果優化 =====
last_hover = None

def on_motion(event):
    global last_hover

    row = tree.identify_row(event.y)
    if row == last_hover:
        return

    if last_hover:
        tree.item(last_hover, tags=tuple(
            t for t in tree.item(last_hover, "tags") if t != "hover"
        ))

    if row:
        tags = list(tree.item(row, "tags"))
        if "hover" not in tags:
            tags.append("hover")
        tree.item(row, tags=tags)
        last_hover = row

def on_leave(event):
    global last_hover
    last_hover = None

tree.bind("<Motion>", on_motion)
tree.bind("<Leave>", on_leave)

# ===== ⭐ 滾輪支援（Windows + Mac 通用）=====
def on_mousewheel(event):
    if event.delta:
        tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
    else:
        tree.yview_scroll(int(event.num == 4 and -1 or 1), "units")

tree.bind_all("<MouseWheel>", on_mousewheel)      # Windows / Mac
tree.bind_all("<Button-4>", on_mousewheel)        # Linux
tree.bind_all("<Button-5>", on_mousewheel)


# 匯出
tk.Label(frame, text="匯出狀態").grid(row=1, column=8)

combo_export_status = ttk.Combobox(frame, values=[
    "全部",
    "🟢正常",
    "🟡提醒",
    "🔴已逾期"
], state="readonly")

combo_export_status.grid(row=1, column=9)
combo_export_status.current(0)

# 綁定點選事件
tree.bind("<<TreeviewSelect>>", on_select)


#程式啟動時先登入
show_login()
root.mainloop()

conn.close()