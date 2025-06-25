import tkinter as tk
from tkinter import messagebox
import openpyxl
import os
import sys

EXCEL_FILE = "database.xlsx"

if not os.path.exists(EXCEL_FILE):
    messagebox.showerror("ファイルエラー", f"{EXCEL_FILE} が見つかりません。")
    sys.exit()

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

    messagebox.showinfo("成功", f"口座 '{name}' を登録しました\n初期残高: {balance:,.0f} 円")
    name_entry.delete(0, tk.END)
    balance_entry.delete(0, tk.END)

def back_to_menu():
    root.destroy()

# --- GUI構築 ---
root = tk.Tk()
root.title("新しい口座の追加")
root.geometry("500x300")
root.minsize(400, 250)

for i in range(4):
    root.grid_rowconfigure(i, weight=0)
root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(0, weight=0)
root.grid_columnconfigure(1, weight=1)

font_label = ("Arial", 12)
font_entry = ("Arial", 12)
font_button = ("Arial", 12)

# 口座名
tk.Label(root, text="口座名", font=font_label).grid(row=0, column=0, padx=10, pady=10, sticky="e")
name_entry = tk.Entry(root, font=font_entry)
name_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

# 初期残高
tk.Label(root, text="初期残高", font=font_label).grid(row=1, column=0, padx=10, pady=10, sticky="e")
balance_entry = tk.Entry(root, font=font_entry)
balance_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

# ボタン
tk.Button(root, text="口座を追加", command=add_account, font=font_button).grid(row=2, column=0, columnspan=2, pady=10)
tk.Button(root, text="メニューに戻る", command=back_to_menu, font=font_button).grid(row=3, column=0, columnspan=2, pady=10)

root.mainloop()