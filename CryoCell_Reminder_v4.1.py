# CryoCell_Reminder_v4.1_!final!

import tkinter as tk
from tkinter import messagebox, ttk
import sqlite3
import datetime
import threading
import time
import schedule
import pandas as pd
import smtplib
import json
import base64
import os
from email.mime.text import MIMEText

CONFIG_FILE = "email_config.json"
running = True


# ------------------ 配置管理 ------------------
def load_email_config():
    if not os.path.exists(CONFIG_FILE):
        return None
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
        if "EMAIL_PASSWORD" in data:
            data["EMAIL_PASSWORD"] = base64.b64decode(data["EMAIL_PASSWORD"]).decode("utf-8")
        return data

def save_email_config(email_user, email_receiver, email_password):
    data = {
        "EMAIL_USER": email_user,
        "EMAIL_RECEIVER": email_receiver,
        "EMAIL_PASSWORD": base64.b64encode(email_password.encode("utf-8")).decode("utf-8")
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

# ------------------ 邮箱信息输入 ------------------
def ask_email_info():
    temp_root = tk.Tk()
    temp_root.withdraw()  # 🔥 把Tk主窗口隐藏（不显示）

    popup = tk.Toplevel(temp_root)
    popup.title("设置邮箱账号")
    popup.geometry("400x400")
    popup.configure(bg="white")
    popup.resizable(False, False)

    center_window(popup)

    font_label = ("Arial", 11)

    frame = tk.Frame(popup, bg="white", padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="发件人邮箱:", bg="white", font=font_label).grid(row=0, column=0, pady=10, sticky="e")
    entry_user = tk.Entry(frame, font=("Arial", 11), width=30)
    entry_user.grid(row=0, column=1, pady=10, sticky="w")

    tk.Label(frame, text="收件人邮箱:", bg="white", font=font_label).grid(row=1, column=0, pady=10, sticky="e")
    entry_receiver = tk.Entry(frame, font=("Arial", 11), width=30)
    entry_receiver.grid(row=1, column=1, pady=10, sticky="w")

    tk.Label(frame, text="发件人邮箱密码:", bg="white", font=font_label).grid(row=2, column=0, pady=10, sticky="e")
    entry_password = tk.Entry(frame, font=("Arial", 11), width=30, show="*")
    entry_password.grid(row=2, column=1, pady=10, sticky="w")

    def confirm():
        email_user = entry_user.get().strip()
        email_receiver = entry_receiver.get().strip()
        email_password = entry_password.get().strip()
        if not email_user or not email_receiver or not email_password:
            messagebox.showwarning("提示", "请填写完整信息")
            return
        save_email_config(email_user, email_receiver, email_password)
        messagebox.showinfo("成功", "邮箱配置已保存")
        try:
            update_email_label()
        except:
            pass
        popup.destroy()
        temp_root.destroy()  # 🔥 填完销毁临时Tk窗口！

    btn_frame = tk.Frame(frame, bg="white")
    btn_frame.grid(row=3, column=0, columnspan=2, pady=20)

    tk.Button(btn_frame, text="保存配置", command=confirm, bg="#4CAF50", fg="white", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="测试发送", command=test_email, bg="#2196F3", fg="white", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=10)
    popup.grab_set()
    popup.wait_window()
#-----------------邮箱配置测试------------------------
def test_email():
    email_user = entry_user.get().strip()
    email_receiver = entry_receiver.get().strip()
    email_password = entry_password.get().strip()
    if not email_user or not email_receiver or not email_password:
        messagebox.showwarning("提示", "请填写完整信息后再测试发送")
        return
    try:
        smtp_server = "smtp.zju.edu.cn"
        smtp_port = 994
        msg = MIMEText("CryoCell Reminder 邮件测试成功！", "plain", "utf-8")
        msg["Subject"] = "CryoCell Reminder 邮件测试"
        msg["From"] = email_user
        msg["To"] = email_receiver

        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(email_user, email_password)
        server.sendmail(email_user, [email_receiver], msg.as_string())
        server.quit()

        messagebox.showinfo("成功", "测试邮件发送成功！✅")
    except Exception as e:
        messagebox.showerror("错误", f"测试邮件发送失败！\\n错误信息：{e}")


# ------------------- 数据库初始化 -------------------
def init_db():
    conn = sqlite3.connect("cell_freeze.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cell_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cell_name TEXT NOT NULL,
            freeze_date TEXT NOT NULL,
            remind_date TEXT NOT NULL,
            notified INTEGER DEFAULT 0,
            box_id INTEGER,
            position TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS freeze_boxes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            box_name TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()

# ------------------- 添加细胞记录 -------------------
def add_record():
    selected_pos = [[]]  # 保存选择的位置

    def choose_position():
        selected_box_name = box_var.get()
        if selected_box_name == "不指定":
            messagebox.showinfo("提示", "未选择冻存盒")
            return
        box_id = box_dict[selected_box_name]
        pos = open_position_selector(box_id)
        if pos:
            selected_pos[0] = pos
            pos_label.config(text=f"已选位置: {len(selected_pos[0])}个")

    def confirm_add():
        cell_name = entry_name.get().strip()
        if not cell_name:
            messagebox.showwarning("提示", "请输入细胞名称")
            return

        selected_box_name = box_var.get()
        if selected_box_name == "不指定":
            box_id = None
            positions = []
        else:
            box_id = box_dict[selected_box_name]
            positions = selected_pos[0]  # 获取多选位置列表

        # 验证至少选择一个位置（如果指定了冻存盒）
        if selected_box_name != "不指定" and not positions:
            messagebox.showwarning("提示", "请至少选择一个位置")
            return

        freeze_date = datetime.date.today()
    
        try:
            days = int(entry_days.get())
            if days <= 0:
                raise ValueError
        except ValueError:
            days = 14

        remind_date = freeze_date + datetime.timedelta(days=days)

        conn = sqlite3.connect("cell_freeze.db")
        cursor = conn.cursor()
    
        # 为每个位置插入记录
        for pos in positions:
            cursor.execute("""
                INSERT INTO cell_records 
                (cell_name, freeze_date, remind_date, box_id, position) 
                VALUES (?, ?, ?, ?, ?)
            """, (cell_name, freeze_date.isoformat(), remind_date.isoformat(), box_id, pos))
    
        conn.commit()
        conn.close()

        log_message(f"添加记录：{cell_name} ×{len(positions)}，提醒日期：{remind_date}")
        refresh_table()
        popup.destroy()

    popup = tk.Toplevel()
    popup.title("添加细胞记录")
    popup.geometry("400x350")  # 更宽更高
    popup.configure(bg="white")
    popup.resizable(True, True)

    # 居中
    popup.update_idletasks()
    w = popup.winfo_width()
    h = popup.winfo_height()
    ws = popup.winfo_screenwidth()
    hs = popup.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    popup.geometry(f"{w}x{h}+{x}+{y}")

    font_label = ("Arial", 11)
    font_entry = ("Arial", 11)

    form_frame = tk.Frame(popup, padx=20, pady=20, bg="white")
    form_frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(form_frame, text="细胞名称:", bg="white", font=font_label).grid(row=0, column=0, sticky="e", pady=10)
    entry_name = tk.Entry(form_frame, font=font_entry, width=30)
    entry_name.grid(row=0, column=1, pady=10, sticky="w")

    tk.Label(form_frame, text="提醒天数 (默认14):", bg="white", font=font_label).grid(row=1, column=0, sticky="e", pady=10)
    entry_days = tk.Entry(form_frame, font=font_entry, width=30)
    entry_days.insert(0, "14")
    entry_days.grid(row=1, column=1, pady=10, sticky="w")

    tk.Label(form_frame, text="冻存盒:", bg="white", font=font_label).grid(row=2, column=0, sticky="e", pady=10)
    box_var = tk.StringVar()
    boxes = get_boxes()
    box_dict = {"不指定": None}
    for b in boxes:
        box_dict[f"{b[1]} (ID {b[0]})"] = b[0]

    box_menu = ttk.Combobox(form_frame, textvariable=box_var, values=list(box_dict.keys()), font=("Arial", 11), width=28, state="readonly")
    box_menu.grid(row=2, column=1, pady=10, sticky="w")
    box_menu.current(0)

    tk.Button(form_frame, text="选择位置", command=choose_position, bg="#3399FF", fg="white", font=("Arial", 11)).grid(row=3, column=0, columnspan=2, pady=10)
    pos_label = tk.Label(form_frame, text="未选择位置", bg="white", font=("Arial", 10, "italic"), fg="#666666")
    pos_label.grid(row=4, column=0, columnspan=2, pady=5)

    tk.Button(form_frame, text="确认添加", command=confirm_add, bg="#4CAF50", fg="white", font=("Arial", 11, "bold")).grid(row=5, column=0, columnspan=2, pady=15)

#--------------冻存盒位置选择----------------------
def open_position_selector(box_id):
    selected_pos = set()  # 使用集合存储多选位置
    confirmed = [False]   # 确认状态

    selector = tk.Toplevel()
    selector.title("选择位置")
    selector.geometry("760x700")  # ✅ 设置窗口尺寸
    selector.configure(bg="white")
    selector.resizable(True, True)

    center_window(selector)  # ✅ 窗口居中

    # 添加确认按钮框架
    btn_frame = tk.Frame(selector, bg="white", pady=10)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X)

    # 确认按钮
    def on_confirm():
        confirmed[0] = True
        selector.destroy()

    tk.Button(btn_frame, text="确认选择", command=on_confirm, 
             bg="#4CAF50", fg="white", font=("Arial", 12, "bold")).pack()

    # 网格框架
    grid_frame = tk.Frame(selector, bg="white")
    grid_frame.pack(expand=True)

    # 获取已占用位置
    conn = sqlite3.connect("cell_freeze.db")
    cursor = conn.cursor()
    cursor.execute("SELECT position FROM cell_records WHERE box_id = ? AND position IS NOT NULL", (box_id,))
    occupied = {row[0] for row in cursor.fetchall()}
    conn.close()

    # 按钮点击处理
    def toggle_pos(pos, btn):
        if pos in selected_pos:
            selected_pos.remove(pos)
            btn.config(bg="#f0f0f0", fg="black")
        else:
            selected_pos.add(pos)
            btn.config(bg="#3399ff", fg="white")

    # 生成网格按钮
    buttons = {}
    for i, row_label in enumerate("ABCDEFGHI"):
        for j in range(1, 10):
            pos = f"{row_label}{j}"
            btn = tk.Button(grid_frame, text=pos, width=6, height=3,
                            bg="#f0f0f0" if pos not in occupied else "#cccccc",
                            relief="groove", font=("Arial", 10, "bold"),
                            state=tk.NORMAL if pos not in occupied else tk.DISABLED)
            btn.grid(row=i, column=j-1, padx=5, pady=5)
            if pos not in occupied:
                btn.configure(command=lambda p=pos, b=btn: toggle_pos(p, b))
            buttons[pos] = btn

    selector.wait_window()
    return list(selected_pos) if confirmed[0] else None

# ------------------- 新建冻存盒 -------------------
def create_box():
    popup = tk.Toplevel()
    popup.title("新建冻存盒")
    popup.geometry("320x180")
    popup.configure(bg="white")
    popup.resizable(True, True)

    # 居中
    popup.update_idletasks()
    w = popup.winfo_width()
    h = popup.winfo_height()
    ws = popup.winfo_screenwidth()
    hs = popup.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    popup.geometry(f"{w}x{h}+{x}+{y}")

    frame = tk.Frame(popup, bg="white", padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="冻存盒名称:", bg="white", font=("Arial", 11)).pack(pady=10)
    entry = tk.Entry(frame, font=("Arial", 11), width=25)
    entry.pack()

    def confirm_create():
        box_name = entry.get().strip()
        if not box_name:
            messagebox.showwarning("警告", "请输入冻存盒名称")
            return
        conn = sqlite3.connect("cell_freeze.db")
        cursor = conn.cursor()
        cursor.execute("INSERT INTO freeze_boxes (box_name) VALUES (?)", (box_name,))
        conn.commit()
        conn.close()
        log_message(f"新建冻存盒：{box_name}")
        messagebox.showinfo("成功", f"冻存盒 '{box_name}' 创建成功！")
        popup.destroy()

    tk.Button(frame, text="确认创建", command=confirm_create, bg="#4CAF50", fg="white", font=("Arial", 11, "bold")).pack(pady=15)


# ------------------- 导出盒子布局到Excel -------------------
def export_box():
    boxes = get_boxes()
    if not boxes:
        messagebox.showwarning("提示", "当前没有冻存盒")
        return

    box_dict = {f"{b[1]} (ID {b[0]})": b[0] for b in boxes}
    popup = tk.Toplevel()
    popup.title("选择盒子导出")
    popup.geometry("360x220")
    popup.configure(bg="white")
    popup.resizable(True, True)

    # 居中
    popup.update_idletasks()
    w = popup.winfo_width()
    h = popup.winfo_height()
    ws = popup.winfo_screenwidth()
    hs = popup.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    popup.geometry(f"{w}x{h}+{x}+{y}")

    frame = tk.Frame(popup, bg="white", padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="选择冻存盒:", bg="white", font=("Arial", 11)).pack(pady=10)
    box_var = tk.StringVar()
    combo = ttk.Combobox(frame, values=list(box_dict.keys()), textvariable=box_var, font=("Arial", 11), state="readonly", width=28)
    combo.pack()

    def do_export():
        selected = box_var.get()
        if not selected:
            return
        box_id = box_dict[selected]

        conn = sqlite3.connect("cell_freeze.db")
        df = pd.read_sql_query("SELECT position, cell_name FROM cell_records WHERE box_id = ?", conn, params=(box_id,))
        conn.close()

        grid = pd.DataFrame("", index=[chr(i) for i in range(65, 74)], columns=[str(j) for j in range(1, 10)])
        for _, row in df.iterrows():
            if row['position']:
                row_char = row['position'][0].upper()
                col_num = row['position'][1]
                if row_char in grid.index and col_num in grid.columns:
                    grid.at[row_char, col_num] = row['cell_name']

        box_name_clean = selected.split(" (ID")[0].replace(" ", "_")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{box_name_clean}_{timestamp}.xlsx"
        grid.to_excel(filename)
        messagebox.showinfo("导出成功", f"布局已导出到 {filename}")
        popup.destroy()

    tk.Button(frame, text="导出", command=do_export, bg="#4CAF50", fg="white", font=("Arial", 11, "bold")).pack(pady=20)

# ------------------- 刷新主表格 -------------------
def refresh_table():
    for row in table.get_children():
        table.delete(row)

    conn = sqlite3.connect("cell_freeze.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT 
            c.id,
            c.cell_name, 
            c.freeze_date,
            c.remind_date,
            c.notified,
            b.box_name,
            c.position
        FROM cell_records c
        LEFT JOIN freeze_boxes b ON c.box_id = b.id
    """)
    records = cursor.fetchall()
    conn.close()

    columns = ("细胞名称", "添加日期", "提醒日期", "状态", "冻存盒", "位置", "hidden_id")
    table.configure(columns=columns)
    table["displaycolumns"] = columns[:-1]

    query = search_var.get().lower()
    count = 0
    for r in records:
        status = "已提醒" if r[4] else "待提醒"
        values = (
            r[1],
            r[2],
            r[3],
            status,
            r[5] or "未分配",
            r[6] or "-",
            r[0]
        )
        if not query or query in r[1].lower():
            table.insert("", tk.END, values=values)
            count += 1

    cell_count_label.config(text=f"当前细胞数量：{count}")

# ------------------- 导出全部细胞记录到Excel -------------------
def export_to_excel():
    conn = sqlite3.connect("cell_freeze.db")
    df = pd.read_sql_query("""
        SELECT 
            c.cell_name AS '细胞名称',
            c.freeze_date AS '添加日期',
            c.remind_date AS '提醒日期',
            b.box_name AS '冻存盒',
            c.position AS '位置',
            CASE c.notified WHEN 1 THEN '已提醒' ELSE '待提醒' END AS '状态'
        FROM cell_records c
        LEFT JOIN freeze_boxes b ON c.box_id = b.id
    """, conn)
    conn.close()

    df.to_excel("cell_freeze_records.xlsx", index=False)
    messagebox.showinfo("导出成功", "记录已成功导出为 'cell_freeze_records.xlsx'")
    
# ------------------ 日志记录 ------------------
def log_message(message):
    with open("reminder_log.txt", "a", encoding="utf-8") as f:
        timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        f.write(f"{timestamp} {message}\n")

# ------------------ 邮件发送 ------------------
def send_reminders():
    config = load_email_config()
    if not config:
        log_message("未配置邮箱，无法发送提醒")
        return

    smtp_server = "smtp.zju.edu.cn"
    smtp_port = 994
    EMAIL_USER = config["EMAIL_USER"]
    EMAIL_RECEIVER = config["EMAIL_RECEIVER"]
    EMAIL_PASSWORD = config["EMAIL_PASSWORD"]

    conn = sqlite3.connect("cell_freeze.db")
    cursor = conn.cursor()

    today = datetime.date.today().isoformat()
    cursor.execute("SELECT id, cell_name, freeze_date FROM cell_records WHERE remind_date <= ? AND notified = 0", (today,))
    to_notify = cursor.fetchall()

    if not to_notify:
        conn.close()
        return

    for rec in to_notify:
        id_, name, freeze_date = rec
        subject = f"请将“{name}”细胞放入液氮罐"
        content = f"您在 {freeze_date} 放入 -80°C 的“{name}”细胞，现在应转移至液氮罐中。"

        msg = MIMEText(content, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = EMAIL_USER
        msg["To"] = EMAIL_RECEIVER

        try:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, [EMAIL_RECEIVER], msg.as_string())
            server.quit()
            cursor.execute("UPDATE cell_records SET notified = 1 WHERE id = ?", (id_,))
            log_message(f"已发送提醒邮件：{name}")
        except Exception as e:
            log_message(f"发送邮件失败：{name}，错误：{e}")

    conn.commit()
    conn.close()

# ------------------- 定时器，每天检查提醒 -------------------
def run_scheduler():
    while running:
        try:
            schedule.run_pending()
        except Exception as e:
            log_message(f"后台定时任务异常：{e}")
        time.sleep(60)

#-------------------关闭应用----------------
def on_closing():
    global running
    if messagebox.askokcancel("退出", "确定要退出CryoCell Reminder吗？"):
        running = False
        root.destroy()



# ------------------- 新增冻存盒管理界面 -------------------
def manage_boxes():
    class BoxManagerWindow:
        def __init__(self, parent):
            self.window = tk.Toplevel(parent)
            self.window.title("冻存盒管理")
            self.window.geometry("760x650")
            self.window.configure(bg="white")
            self.window.resizable(True, True)

            self.selected_pos = set()  # 🔥 🔥 🔥 这一行非常重要，初始化！

            # 居中弹窗
            self.center_window(self.window)

            # 样式初始化
            style = ttk.Style()
            style.configure("Occupied.TButton", background="#e6f0ff", foreground="#003366", font=('Arial', 9, 'bold'))
            style.configure("Empty.TButton", background="#f9f9f9", foreground="#999999", font=('Arial', 9))
            style.configure("Selected.TButton", background="#3399ff", foreground="white", font=('Arial', 9, 'bold'))

            # 冻存盒选择
            self.box_var = tk.StringVar()
            self.box_frame = tk.Frame(self.window, bg="white")
            self.box_frame.pack(pady=10)

            ttk.Label(self.box_frame, text="选择冻存盒:", background="white", font=("Arial", 11)).pack(side=tk.LEFT)
            self.box_combobox = ttk.Combobox(self.box_frame, textvariable=self.box_var, font=("Arial", 11), width=30, state="readonly")
            self.box_combobox.pack(side=tk.LEFT, padx=5)
            ttk.Button(self.box_frame, text="刷新列表", command=self.load_boxes).pack(side=tk.LEFT, padx=5)
            self.box_combobox.bind("<<ComboboxSelected>>", lambda e: self.update_grid())

            # 网格区
            self.grid_frame = tk.Frame(self.window, bg="white")
            self.grid_frame.pack(pady=10)

            # 按钮区
            self.btn_frame = tk.Frame(self.window, bg="white")
            self.btn_frame.pack(pady=10)
            ttk.Button(self.btn_frame, text="取出细胞", command=self.release_position).pack()

            self.create_grid()
            self.load_boxes()

        def center_window(self, win):
            win.update_idletasks()
            w = win.winfo_width()
            h = win.winfo_height()
            ws = win.winfo_screenwidth()
            hs = win.winfo_screenheight()
            x = (ws // 2) - (w // 2)
            y = (hs // 2) - (h // 2)
            win.geometry(f"{w}x{h}+{x}+{y}")

        def load_boxes(self):
            boxes = get_boxes()
            self.box_dict = {f"{b[1]} (ID {b[0]})": b[0] for b in boxes}
            self.box_combobox['values'] = list(self.box_dict.keys())
            if self.box_dict:
                self.box_combobox.current(0)
                self.update_grid()
            else:
                messagebox.showinfo("提示", "当前没有可用冻存盒")

        def create_grid(self):
            for widget in self.grid_frame.winfo_children():
                widget.destroy()
            self.buttons = {}
            for i, row in enumerate("ABCDEFGHI"):
                for j in range(1, 10):
                    pos = f"{row}{j}"
                    btn = ttk.Button(self.grid_frame, text=pos, width=9)
                    btn.grid(row=i, column=j-1, padx=3, pady=3)
                    btn.configure(command=lambda p=pos: self.select_position(p))
                    self.buttons[pos] = btn

        def update_grid(self):
            if not self.box_var.get():
                return

            try:
                box_id = self.box_dict[self.box_var.get()]
            except KeyError:
                return

            conn = sqlite3.connect("cell_freeze.db")
            cursor = conn.cursor()
            cursor.execute("""
                SELECT position, cell_name, freeze_date 
                FROM cell_records 
                WHERE box_id = ? AND position IS NOT NULL
            """, (box_id,))
            cells = {pos: (name, date) for pos, name, date in cursor.fetchall()}
            conn.close()

            for pos, btn in self.buttons.items():
                if pos in cells:
                    btn.config(text=f"{pos}\n{cells[pos][0]}\n({cells[pos][1]})", style="Occupied.TButton")
                else:
                    btn.config(text=pos, style="Empty.TButton")

        def select_position(self, pos):
            if pos in self.selected_pos:
                self.selected_pos.remove(pos)
                self.buttons[pos].config(style="Occupied.TButton" if "\\n" in self.buttons[pos].cget("text") else "Empty.TButton")
            else:
                self.selected_pos.add(pos)
                self.buttons[pos].config(style="Selected.TButton")

        def release_position(self):
            if not self.selected_pos:
                messagebox.showwarning("警告", "请先选择要操作的位置")
                return

            pos_list = "\n".join(self.selected_pos)
            confirm = messagebox.askyesno("确认操作", f"确定要取出以下位置的细胞吗？\n\n{pos_list}")
            if not confirm:
                return

            box_id = self.box_dict[self.box_var.get()]
            conn = sqlite3.connect("cell_freeze.db")
            cursor = conn.cursor()
        
            # 批量更新选中位置
            for pos in self.selected_pos:
                cursor.execute("""
                    UPDATE cell_records 
                    SET position = NULL 
                    WHERE box_id = ? AND position = ?
                """, (box_id, pos))
        
            conn.commit()
            conn.close()

            log_message(f"取出细胞：{len(self.selected_pos)}个位置")
            self.selected_pos.clear()
            self.update_grid()

    BoxManagerWindow(root)


# ------------------- GUI界面构建 -------------------
def create_gui():
    global table, root, cell_count_label, search_var, email_label
    root = tk.Tk()
    root.title("CryoCell Reminder v4.1 - 细胞冻存管理中心")
    root.geometry("1000x700")
    root.configure(bg="white")
    root.resizable(True, True)
    root.protocol("WM_DELETE_WINDOW", on_closing)


    center_window(root)

    toolbar = tk.Frame(root, bg="white")
    toolbar.pack(side=tk.TOP, fill=tk.X)

    btn_style = {"relief": tk.FLAT, "bg": "white", "fg": "#0066cc", "font": ("Arial", 10, "bold")}

    tk.Button(toolbar, text="作者信息", command=lambda: messagebox.showinfo("作者", "CryoCell Reminder v4.1\n作者: 威震八方@ZJU"), **btn_style).pack(side=tk.LEFT, padx=8, pady=5)
    tk.Button(toolbar, text="导出所有记录", command=export_to_excel, **btn_style).pack(side=tk.LEFT, padx=8, pady=5)
    tk.Button(toolbar, text="添加细胞", command=add_record, **btn_style).pack(side=tk.LEFT, padx=8, pady=5)

    # 冻存盒管理按钮 + 弹出式菜单
    box_manage_btn = tk.Button(toolbar, text="冻存盒管理", **btn_style)
    box_manage_btn.pack(side=tk.LEFT, padx=8, pady=5)

    box_menu = tk.Menu(root, tearoff=0)
    box_menu.add_command(label="新建冻存盒", command=create_box)
    box_menu.add_command(label="删除冻存盒", command=delete_box)
    box_menu.add_command(label="管理冻存盒", command=manage_boxes)

    def show_box_menu(event):
        box_menu.post(event.x_root, event.y_root)

    box_manage_btn.bind("<Button-1>", show_box_menu)

    tk.Button(toolbar, text="修改邮箱设置", command=ask_email_info, **btn_style).pack(side=tk.LEFT, padx=8, pady=5)

    email_label = tk.Label(toolbar, text="", bg="white", fg="#333333", font=("Arial", 10))
    email_label.pack(side=tk.RIGHT, padx=10)
    update_email_label()

    # 搜索区域
    search_frame = tk.Frame(root, bg="white")
    search_frame.pack(pady=5)
    tk.Label(search_frame, text="搜索细胞:", bg="white", font=("Arial", 11)).pack(side=tk.LEFT)
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Arial", 11), width=30)
    search_entry.pack(side=tk.LEFT, padx=5)
    search_var.trace("w", lambda *args: refresh_table())

    # 细胞数量显示
    cell_count_label = tk.Label(root, text="当前细胞数量：0", bg="white", font=("Arial", 11), fg="#333333")
    cell_count_label.pack(pady=5)

    # 数据表格
    columns = ("细胞名称", "添加日期", "提醒日期", "状态", "冻存盒", "位置", "hidden_id")
    table = ttk.Treeview(root, columns=columns, show="headings", selectmode="extended")

    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
    style.configure("Treeview", font=("Arial", 11))

    for col in columns:
        table.heading(col, text=col, command=lambda c=col: sort_by(c, False))
        table.column(col, width=150 if col != "细胞名称" else 220)
    table.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 右键菜单
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="删除记录", command=delete_record)

    def on_right_click(event):
        try:
            table.focus_set()  # 把焦点放到表格，防止有时候失焦
            row_id = table.identify_row(event.y)

            if row_id:  # 如果点到了一行
                if row_id not in table.selection():
                    table.selection_set(row_id)
            else:
                # 点击空白区域不改变已有选中
                pass

            # 无论如何弹出右键菜单
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    table.bind("<Button-3>", on_right_click)


    refresh_table()

    return root

#-------------------窗口居中------------------
def center_window(win):
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

# ------------------- 新增删除记录函数 -------------------
def delete_record():
    selected_items = table.selection()
    if not selected_items:
        messagebox.showwarning("提示", "请先选择要删除的记录")
        return

    if not messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selected_items)} 条记录吗？"):
        return

    conn = sqlite3.connect("cell_freeze.db")
    cursor = conn.cursor()

    for item in selected_items:
        record = table.item(item)
        record_id = record['values'][6]  # hidden_id列
        cursor.execute("DELETE FROM cell_records WHERE id = ?", (record_id,))
        table.delete(item)

    conn.commit()
    conn.close()
    refresh_table()
    messagebox.showinfo("删除成功", "选中的记录已成功删除！✅")


# ------------------- 修改refresh_table函数 -------------------
def refresh_table():
    for row in table.get_children():
        table.delete(row)

    conn = sqlite3.connect("cell_freeze.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT 
            c.id,
            c.cell_name, 
            c.freeze_date,
            c.remind_date,
            c.notified,
            b.box_name,
            c.position
        FROM cell_records c
        LEFT JOIN freeze_boxes b ON c.box_id = b.id
    """)
    records = cursor.fetchall()
    conn.close()

    columns = ("细胞名称", "添加日期", "提醒日期", "状态", "冻存盒", "位置", "hidden_id")
    table.configure(columns=columns)
    table["displaycolumns"] = columns[:-1]

    for col in columns[:-1]:
        table.heading(col, text=col)
        table.column(col, width=120 if col != "细胞名称" else 200)

    query = search_var.get().lower()
    count = 0
    for r in records:
        status = "已提醒" if r[4] else "待提醒"
        values = (
            r[1],  # 细胞名称
            r[2],  # 添加日期
            r[3],  # 提醒日期
            status,
            r[5] or "未分配",  # 冻存盒
            r[6] or "-",  # 位置
            r[0]   # hidden_id
        )
        
        # 扩展搜索范围
        search_text = f"{r[1]}{r[2]}{r[5]}{r[6]}".lower()
        if not query or query in search_text:
            table.insert("", tk.END, values=values)
            count += 1

    cell_count_label.config(text=f"当前细胞数量：{count}")

# ------------------- 工具函数 -------------------
def get_boxes():
    conn = sqlite3.connect("cell_freeze.db")
    cursor = conn.cursor()
    cursor.execute("SELECT id, box_name FROM freeze_boxes")
    boxes = cursor.fetchall()
    conn.close()
    return boxes

def log_message(message):
    with open("reminder_log.txt", "a", encoding="utf-8") as f:
        timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        f.write(f"{timestamp} {message}\n")

# ------------------- 冻存盒管理菜单 -------------------
def box_management_menu(event=None):
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="新建冻存盒", command=create_box)
    menu.add_command(label="删除冻存盒", command=delete_box)
    menu.add_command(label="管理冻存盒", command=manage_boxes)
    menu.add_command(label="导出盒子布局", command=export_box)
    menu.post(event.x_root, event.y_root)

# ------------------- 删除冻存盒功能 -------------------
def delete_box():
    boxes = get_boxes()
    if not boxes:
        messagebox.showwarning("提示", "当前没有冻存盒可删除！")
        return

    box_dict = {f"{b[1]} (ID {b[0]})": b[0] for b in boxes}

    popup = tk.Toplevel()
    popup.title("删除冻存盒")
    popup.geometry("360x220")
    popup.configure(bg="white")
    popup.resizable(True, True)

    # 居中
    popup.update_idletasks()
    w = popup.winfo_width()
    h = popup.winfo_height()
    ws = popup.winfo_screenwidth()
    hs = popup.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    popup.geometry(f"{w}x{h}+{x}+{y}")

    frame = tk.Frame(popup, bg="white", padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="选择要删除的冻存盒:", bg="white", font=("Arial", 11)).pack(pady=10)
    box_var = tk.StringVar()
    combo = ttk.Combobox(frame, values=list(box_dict.keys()), textvariable=box_var, font=("Arial", 11), state="readonly", width=28)
    combo.pack()

    def confirm_delete():
        selected = box_var.get()
        if not selected:
            return
        box_id = box_dict[selected]

        conn = sqlite3.connect("cell_freeze.db")
        cursor = conn.cursor()
        cursor.execute("SELECT cell_name, position FROM cell_records WHERE box_id = ?", (box_id,))
        cells = cursor.fetchall()

        if cells:
            cell_info = "\n".join([f"{c[0]} ({c[1]})" for c in cells if c[1]])
            confirm = messagebox.askyesno("警告", f"该盒子中还有以下细胞:\n\n{cell_info}\n\n仍然要删除？")
            if not confirm:
                conn.close()
                return

        cursor.execute("DELETE FROM freeze_boxes WHERE id = ?", (box_id,))
        cursor.execute("UPDATE cell_records SET box_id = NULL, position = NULL WHERE box_id = ?", (box_id,))
        conn.commit()
        conn.close()

        log_message(f"删除冻存盒：{selected}")
        messagebox.showinfo("删除成功", f"已删除 {selected}")
        popup.destroy()

    tk.Button(frame, text="确认删除", command=confirm_delete, bg="#ff4d4d", fg="white", font=("Arial", 11, "bold")).pack(pady=20)

# ------------------- 列排序功能 -------------------
def sort_by(col, descending):
    data = [(table.set(child, col), child) for child in table.get_children('')]
    try:
        data.sort(reverse=descending, key=lambda t: (t[0] if t[0] else "9999-99-99"))
    except:
        data.sort(reverse=descending)

    for index, (val, child) in enumerate(data):
        table.move(child, '', index)

    table.heading(col, command=lambda: sort_by(col, not descending))

# ------------------- 更新邮箱信息 -------------------
def update_email_label():
    config = load_email_config()
    try:
        if config and email_label.winfo_exists():
            email_label.config(text=f"当前邮箱: {config['EMAIL_USER']}")
    except NameError:
        # 如果email_label不存在，忽略（说明主界面还没建立）
        pass



# ------------------- 主程序入口 -------------------
if __name__ == "__main__":
    init_db()
    if not load_email_config():
        ask_email_info()
    threading.Thread(target=run_scheduler, daemon=True).start()
    app = create_gui()
    app.mainloop()

