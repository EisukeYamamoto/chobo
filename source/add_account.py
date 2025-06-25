import tkinter as tk
from tkinter import messagebox
import openpyxl
import os

EXCEL_FILE = "database.xlsx"

def load_account_names():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["口座一覧"]
    names = [row[1] for row in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()
    return names

def get_next_account_id():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["口座一覧"]
    ids = [row[0] for row in ws.iter_rows(min_row=2, values_only=True) if row[0] is not None]
    wb.close()
    return max(ids, default=0) + 1

def add_account():
    name = name_entry.get().strip()
    balance = balance_entry.get().strip()

    if not name or not balance:
        messagebox.showwarning("入力エラー", "すべての項目を入力してください")
        return

    if name in load_account_names():
        messagebox.showerror("重複エラー", "この口座名はすでに存在します")
        return

    try:
        balance = float(balance)
    except ValueError:
        messagebox.showerror("金額エラー", "初期残高は数値で入力してください")
        return

    account_id = get_next_account_id()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["口座一覧"]
    ws.append([account_id, name, balance])
    wb.save(EXCEL_FILE)
    wb.close()

    messagebox.showinfo("成功", f"{name} を追加しました")
    name_entry.delete(0, tk.END)
    balance_entry.delete(0, tk.END)

# --- GUI構築 ---
root = tk.Tk()
root.title("新しい口座の追加")

tk.Label(root, text="口座名").grid(row=0, column=0)
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1)

tk.Label(root, text="初期残高").grid(row=1, column=0)
balance_entry = tk.Entry(root)
balance_entry.grid(row=1, column=1)

tk.Button(root, text="口座を追加", command=add_account).grid(row=2, column=0, columnspan=2)

root.mainloop()
