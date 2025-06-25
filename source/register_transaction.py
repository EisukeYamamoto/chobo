import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from datetime import datetime

EXCEL_FILE = "database.xlsx"

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
    date_str = date_entry.get()
    account_name = account_combo.get()
    summary = summary_entry.get()
    amount = amount_entry.get()
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

    save_transaction(date_str, account_name, summary, deposit, withdrawal)
    messagebox.showinfo("登録完了", "取引が登録されました")

    summary_entry.delete(0, tk.END)
    amount_entry.delete(0, tk.END)

# --- GUIセットアップ ---
root = tk.Tk()
root.title("入出金登録")

tk.Label(root, text="日付 (YYYY-MM-DD)").grid(row=0, column=0)
date_entry = tk.Entry(root)
date_entry.insert(0, datetime.today().strftime("%Y-%m-%d"))
date_entry.grid(row=0, column=1)

tk.Label(root, text="口座").grid(row=1, column=0)
account_combo = ttk.Combobox(root, state="readonly", values=list(load_accounts().keys()))
account_combo.grid(row=1, column=1)

tk.Label(root, text="摘要").grid(row=2, column=0)
summary_entry = tk.Entry(root)
summary_entry.grid(row=2, column=1)

tk.Label(root, text="金額").grid(row=3, column=0)
amount_entry = tk.Entry(root)
amount_entry.grid(row=3, column=1)

mode_var = tk.StringVar(value="預入")
tk.Radiobutton(root, text="預入", variable=mode_var, value="預入").grid(row=4, column=0)
tk.Radiobutton(root, text="引出", variable=mode_var, value="引出").grid(row=4, column=1)

tk.Button(root, text="登録", command=register_transaction).grid(row=5, column=0, columnspan=2)

root.mainloop()