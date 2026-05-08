import tkinter as tk
from tkinter import ttk, messagebox
import database
from datetime import datetime

# --- 1. 新增個案的視窗 ---
def show_add_patient_window(parent, refresh_callback):
    add_win = tk.Toplevel(parent)
    add_win.title("🏥 新增收案登記")
    add_win.geometry("500x700")
    add_win.configure(bg="#f8f9fa")
    add_win.grab_set()

    tk.Label(add_win, text="🏥 新增個案資料", font=("Microsoft JhengHei", 16, "bold"), bg="#f8f9fa").pack(pady=15)

    form_frame = tk.Frame(add_win, bg="#f8f9fa")
    form_frame.pack(padx=30, fill="both", expand=True)

    # 【重要修改】病歷號：在新增模式下，必須是 normal 才能輸入！
    tk.Label(form_frame, text="病歷號碼:", bg="#f8f9fa").grid(row=0, column=0, sticky="w", pady=5)
    ent_id = tk.Entry(form_frame, width=30)
    ent_id.config(state="normal") # 👈 確保這裡是白色可輸入狀態
    ent_id.grid(row=0, column=1, pady=5)

    # 姓名
    tk.Label(form_frame, text="姓名:", bg="#f8f9fa").grid(row=1, column=0, sticky="w", pady=5)
    ent_name = tk.Entry(form_frame, width=30)
    ent_name.grid(row=1, column=1, pady=5)

    # 藥物 (Combobox)
    tk.Label(form_frame, text="藥物:", bg="#f8f9fa").grid(row=2, column=0, sticky="w", pady=5)
    combo_drug = ttk.Combobox(form_frame, width=28, state="readonly", values=[
        "Entecavir 1.0", "Entecavir 0.5", "Tenofovir 300mg", 
        "Tenofovir alafenamide 25mg", "Telbivudine", "Maviret", "Epclusa"
    ])
    combo_drug.grid(row=2, column=1, pady=5)

    # 用藥條件 (Combobox)
    tk.Label(form_frame, text="用藥條件:", bg="#f8f9fa").grid(row=3, column=0, sticky="w", pady=5)
    combo_condition = ttk.Combobox(form_frame, width=28, state="readonly", values=[
        "e+，肝代償不全", "e-，肝代償不全", "癌症患者化療中發作長期使用", "化療預防性治療", "肝硬化長期使用", "DAA"
    ])
    combo_condition.grid(row=3, column=1, pady=5)

    # 用藥結束日 (預設今天)
    tk.Label(form_frame, text="用藥結束日:", bg="#f8f9fa").grid(row=4, column=0, sticky="w", pady=5)
    ent_date = tk.Entry(form_frame, width=30)
    ent_date.insert(0, datetime.today().strftime('%Y-%m-%d'))
    ent_date.grid(row=4, column=1, pady=5)

    # 主治醫師
    tk.Label(form_frame, text="主治醫師:", bg="#f8f9fa").grid(row=5, column=0, sticky="w", pady=5)
    combo_doctor = ttk.Combobox(form_frame, width=28, state="readonly", values=["Alice", "Bob", "Cathy", "Denil"])
    combo_doctor.grid(row=5, column=1, pady=5)

    # 回診日期
    tk.Label(form_frame, text="回診日期:", bg="#f8f9fa").grid(row=6, column=0, sticky="w", pady=5)
    ent_return_date = tk.Entry(form_frame, width=30)
    ent_return_date.grid(row=6, column=1, pady=5)

    # 備註
    tk.Label(form_frame, text="備註:", bg="#f8f9fa").grid(row=7, column=0, sticky="nw", pady=5)
    ent_note = tk.Entry(form_frame, width=30)
    ent_note.grid(row=7, column=1, pady=5)

    def handle_save():
        data = (
            ent_id.get().strip(),
            ent_name.get().strip(),
            combo_drug.get(),
            combo_condition.get(),
            ent_date.get().strip(),
            combo_doctor.get(),
            ent_return_date.get().strip(),
            ent_note.get().strip()
        )
        if not data[0] or not data[1]:
            messagebox.showerror("錯誤", "病歷號碼與姓名為必填！")
            return
        try:
            database.add_new_case(data)
            messagebox.showinfo("成功", "個案已成功收錄！")
            refresh_callback()
            add_win.destroy()
        except Exception as e:
            messagebox.showerror("失敗", f"存檔出錯：{e}")

    tk.Button(add_win, text="💾 確定新增", command=handle_save, bg="#27ae60", fg="white", 
              highlightbackground="#27ae60", font=("Microsoft JhengHei", 12, "bold"), width=15).pack(pady=30)


# --- 2. 修改個案的視窗 (雙擊表格彈出) ---
def show_edit_window(parent, patient_id, refresh_callback):
    data = database.get_patient_by_id(patient_id)
    if not data:
        messagebox.showerror("錯誤", f"找不到資料 (ID: {patient_id})")
        return

    # 🏥 順序檢查（請確保這與 database.py 中的 initialize_db 順序一致）
    # 0:ID, 1:名, 2:藥, 3:條件, 4:結束日, 5:醫, 6:回診, 7:備註
    pid, name, drug, condition, last_date, doctor, return_date, note = data

    edit_win = tk.Toplevel(parent)
    edit_win.title(f"修改資料 - {name}")
    edit_win.geometry("500x750")
    edit_win.configure(bg="#f8f9fa")
    edit_win.grab_set()

    form_frame = tk.Frame(edit_win, bg="#f8f9fa")
    form_frame.pack(padx=30, fill="both", expand=True)

    # 1. 病歷號 (唯讀)
    tk.Label(form_frame, text="病歷號碼:", bg="#f8f9fa").grid(row=0, column=0, sticky="w", pady=5)
    ent_id = tk.Entry(form_frame, width=30)
    ent_id.insert(0, pid)
    ent_id.config(state="readonly", readonlybackground="#e9ecef") 
    ent_id.grid(row=0, column=1, pady=5)

    # 2. 姓名
    tk.Label(form_frame, text="姓名:", bg="#f8f9fa").grid(row=1, column=0, sticky="w", pady=5)
    ent_name = tk.Entry(form_frame, width=30)
    ent_name.insert(0, name)
    ent_name.grid(row=1, column=1, pady=5)

    # 3. 藥物
    tk.Label(form_frame, text="藥物:", bg="#f8f9fa").grid(row=2, column=0, sticky="w", pady=5)
    combo_drug = ttk.Combobox(form_frame, width=28, values=[
        "Entecavir 1.0", "Entecavir 0.5", "Tenofovir 300mg", 
        "Tenofovir alafenamide 25mg", "Maviret", "Epclusa"
    ])
    combo_drug.set(drug)
    combo_drug.grid(row=2, column=1, pady=5)

    # 4. 用藥條件 (使用長列表，只定義一次)
    tk.Label(form_frame, text="用藥條件:", bg="#f8f9fa").grid(row=3, column=0, sticky="w", pady=5)
    combo_condition = ttk.Combobox(form_frame, width=28, values=[
        "e+，肝代償不全", "e-，肝代償不全", "器官移植長期使用", "癌症患者化療中發作長期使用",
        "化療預防性治療", "肝硬化長期使用", "懷孕婦女(27週)預防垂直感染", "器官移植預防性治療",
        "e+，ALT>=5X", "e+，2X<=ALT<5", "e-，ALT>=2X", "肝癌並接受根除性治療",
        "免疫抑制劑治療", "e+，ALT", "e-，ALT", "DAA"
    ])
    combo_condition.set(condition)
    combo_condition.grid(row=3, column=1, pady=5)

    # 5. 用藥結束日
    tk.Label(form_frame, text="用藥結束日:", bg="#f8f9fa").grid(row=4, column=0, sticky="w", pady=5)
    ent_date = tk.Entry(form_frame, width=30)
    ent_date.insert(0, last_date)
    ent_date.grid(row=4, column=1, pady=5)

    # 6. 醫師
    tk.Label(form_frame, text="主治醫師:", bg="#f8f9fa").grid(row=5, column=0, sticky="w", pady=5)
    combo_doctor = ttk.Combobox(form_frame, width=28, values=["Alice", "Bob", "Cathy", "Denil"])
    combo_doctor.set(doctor)
    combo_doctor.grid(row=5, column=1, pady=5)

    # 7. 回診日
    tk.Label(form_frame, text="回診日期:", bg="#f8f9fa").grid(row=6, column=0, sticky="w", pady=5)
    ent_return_date = tk.Entry(form_frame, width=30)
    ent_return_date.insert(0, return_date)
    ent_return_date.grid(row=6, column=1, pady=5)

    # 8. 備註
    tk.Label(form_frame, text="備註:", bg="#f8f9fa").grid(row=7, column=0, sticky="nw", pady=5)
    ent_note = tk.Entry(form_frame, width=30)
    ent_note.insert(0, note)
    ent_note.grid(row=7, column=1, pady=5)


    def handle_update():
        try:
            updated_info = (ent_name.get().strip(), combo_drug.get(), combo_condition.get(), 
                            ent_date.get().strip(), combo_doctor.get(), ent_return_date.get().strip(), 
                            ent_note.get().strip(), pid)
            database.update_patient_data(updated_info)
            messagebox.showinfo("成功", "資料已更新")
            refresh_callback()
            edit_win.destroy()
        except Exception as e:
            messagebox.showerror("更新失敗", f"錯誤：{e}")

    def handle_delete():
        if messagebox.askyesno("確認刪除", f"確定要刪除【{name}】嗎？"):
            database.delete_patient(pid)
            refresh_callback()
            edit_win.destroy()

    btn_frame = tk.Frame(edit_win, bg="#f8f9fa")
    btn_frame.pack(side="bottom", pady=40)

    tk.Button(btn_frame, text="💾 儲存修改", command=handle_update, bg="#3498db", fg="white", 
              highlightbackground="#3498db", font=("Microsoft JhengHei", 12, "bold"), width=12).grid(row=0, column=0, padx=10)
    
    tk.Button(btn_frame, text="🗑️ 刪除個案", command=handle_delete, bg="#e74c3c", fg="white", 
              highlightbackground="#e74c3c", font=("Microsoft JhengHei", 12, "bold"), 
              width=12).grid(row=0, column=1, padx=10) # 👈 這裡只留下 column=1