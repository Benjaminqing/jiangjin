
# Time: 2024/1/9 10:31
import tkinter as tk
from tkinter import filedialog, messagebox
import sqlite3
import pandas as pd
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd


# 数据库操作函数
def db_connection():
    conn = sqlite3.connect('company.db')
    return conn

def view_data():
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM employee")
    rows = cursor.fetchall()
    update_display(rows)
    conn.close()

def add_data(name, employee_number, email, position):
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO employee (name, employee_number, email, position) VALUES (?, ?, ?, ?)",
                   (name, employee_number, email, position))
    conn.commit()
    conn.close()
    view_data()

def update_data(id, name, employee_number, email, position):
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE employee SET name = ?, employee_number = ?, email = ?, position = ? WHERE id = ?",
                   (name, employee_number, email, position, id))
    conn.commit()
    conn.close()
    view_data()

def delete_data(id):
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM employee WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    view_data()



# Excel文件读取部分

def open_file():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_paths:
        for file_path in file_paths:
            # 这里可以添加读取Excel文件的代码
            print(f"文件已打开：{file_path}")

# GUI界面部分
def drop(event):
    file_paths = root.tk.splitlist(event.data)
    if file_paths:
        for file_path in file_paths:
            # 这里可以添加读取Excel文件的代码
            print(f"文件已拖入：{file_path}")

root = TkinterDnD.Tk()
root.title("企业员工信息管理系统")

# 创建一个带虚线边框的拖放区域
drop_frame = tk.Label(root, text="请将excel文件拖入", relief="groove", bd=2, fg="blue")
drop_frame.config(width=40, height=10, borderwidth=2, relief='dashed')
drop_frame.pack(padx=10, pady=10, fill='both', expand=True)
drop_frame.drop_target_register(DND_FILES)
drop_frame.dnd_bind('<<Drop>>', drop)

root = tk.Tk()
root.title("企业员工信息管理系统")

# ...（其余的GUI设置，包括文本框、按钮等）

# 新增一个用于打开Excel文件的按钮
open_file_button = tk.Button(root, text="打开Excel文件", command=open_file)
open_file_button.grid(row=7, column=0, columnspan=2)

# 启动GUI事件循环
root.mainloop()
