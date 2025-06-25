import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk, messagebox
import openpyxl
from datetime import datetime
import os
import sys

EXCEL_FILE = "database.xlsx"

# --- ファイル存在チェック ---
if not os.path.exists(EXCEL_FILE):
    messagebox.showerror("ファイルエラー", f"{EXCEL_FILE} が見つかりません。")
    sys.exit()

def load_accounts():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["口座一覧"]
    accounts = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        accounts[row[1]] = row[0]  # 口座名:口座ID
    wb.close()
    return accounts

def save_transaction(date_str, account_name, summary, deposit, withdrawal):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    account_id = load_accounts()[account_name]
    ws = wb["取引履歴"]
    ws.append([date_str, account_id, summary, deposit, withdrawal])
    wb.save(EXCEL_FILE)
    wb.close()

def register_transaction():
    date_str = date_entry.get().strip()
    account_name = account_combo.get()
    summary = summary_entry.get().strip()
    amount = amount_entry.get().strip()
    mode = mode_var.get()

    if not (date_str and account_name and summary and amount):
        messagebox.showwarning("入力エラー", "すべての項目を入力してください")
        return

    try:
        float_amount = float(amount)
    except ValueError:
        messagebox.showwarning("金額エラー", "金額は数値で入力してください")
        return

    deposit = float_amount if mode == "預入" else ""
    withdrawal = float_amount if mode == "引出" else ""

    try:
        save_transaction(date_str, account_name, summary, deposit, withdrawal)
        messagebox.showinfo("登録完了", "取引が登録されました")
        summary_entry.delete(0, tk.END)
        amount_entry.delete(0, tk.END)
    except Exception as e:
        messagebox.showerror("登録エラー", f"保存中にエラーが発生しました：\n{e}")

# --- GUIセットアップ ---
root = tk.Tk()
root.title("入出金登録")
root.geometry("300x250")

tk.Label(root, text="日付").grid(row=0, column=0)
date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
date_entry.grid(row=0, column=1)

tk.Label(root, text="口座").grid(row=1, column=0)
account_combo = ttk.Combobox(root, state="readonly")
account_combo.grid(row=1, column=1)

try:
    accounts = load_accounts()
    account_combo["values"] = list(accounts.keys())
except Exception as e:
    messagebox.showerror("口座読み込みエラー", f"口座一覧の読み込みに失敗しました\n{e}")
    root.destroy()
    sys.exit()

tk.Label(root, text="摘要").grid(row=2, column=0)
summary_entry = tk.Entry(root)
summary_entry.grid(row=2, column=1)

tk.Label(root, text="金額").grid(row=3, column=0)
amount_entry = tk.Entry(root)
amount_entry.grid(row=3, column=1)

mode_var = tk.StringVar(value="預入")
tk.Radiobutton(root, text="預入", variable=mode_var, value="預入").grid(row=4, column=0)
tk.Radiobutton(root, text="引出", variable=mode_var, value="引出").grid(row=4, column=1)

tk.Button(root, text="登録", command=register_transaction).grid(row=5, column=0, columnspan=2, pady=10)

root.mainloop()
