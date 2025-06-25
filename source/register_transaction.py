import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from datetime import datetime
import os
from tkcalendar import DateEntry

EXCEL_FILE = "database.xlsx"

def show_register_window():
    if not os.path.exists(EXCEL_FILE):
        messagebox.showerror("ファイルエラー", f"{EXCEL_FILE} が見つかりません。")
        return

    def load_accounts():
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["口座一覧"]
        accounts = {row[1]: row[0] for row in ws.iter_rows(min_row=2, values_only=True)}
        wb.close()
        return accounts

    def save_transaction(date_str, account_name, summary, deposit, withdrawal):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        account_id = accounts[account_name]
        ws = wb["取引履歴"]
        ws.append([date_str, account_id, summary, deposit, withdrawal])
        wb.save(EXCEL_FILE)
        wb.close()

    def get_current_balance(account_name):
        account_id = accounts[account_name]
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws1 = wb["口座一覧"]
        ws2 = wb["取引履歴"]
        initial = 0
        for row in ws1.iter_rows(min_row=2, values_only=True):
            if row[0] == account_id:
                initial = float(row[2])
                break
        total_deposit = 0
        total_withdrawal = 0
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if row[1] == account_id:
                total_deposit += float(row[3] or 0)
                total_withdrawal += float(row[4] or 0)
        wb.close()
        return initial + total_deposit - total_withdrawal

    def register_transaction():
        date_str = date_entry.get()
        account_name = account_combo.get()
        summary = summary_entry.get().strip()
        amount = amount_entry.get().strip()
        mode = mode_var.get()

        if not (date_str and account_name and summary and amount):
            messagebox.showwarning("入力エラー", "すべての項目を入力してください")
            return

        try:
            input_date = datetime.strptime(date_str, "%Y-%m-%d")
            if input_date > datetime.today():
                messagebox.showwarning("日付エラー", "未来の日付は入力できません")
                return
        except ValueError:
            messagebox.showwarning("日付エラー", "日付の形式が正しくありません")
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
            current_balance = get_current_balance(account_name)
            messagebox.showinfo("登録完了", f"取引が登録されました\n現在の残高: {current_balance:,.0f} 円")
            summary_entry.delete(0, tk.END)
            amount_entry.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("登録エラー", f"保存中にエラーが発生しました：\n{e}")

    accounts = load_accounts()

    root = tk.Toplevel()
    root.title("入出金登録")
    root.geometry("500x450")
    root.minsize(400, 300)

    for i in range(7):
        root.grid_rowconfigure(i, weight=0)
    root.grid_rowconfigure(6, weight=1)
    root.grid_columnconfigure(0, weight=0)
    root.grid_columnconfigure(1, weight=1)

    font_label = ("Arial", 12)
    font_entry = ("Arial", 12)
    font_button = ("Arial", 12)

    tk.Label(root, text="日付", font=font_label).grid(row=0, column=0, padx=5, pady=5, sticky="e")
    date_entry = DateEntry(root, date_pattern='yyyy-mm-dd', font=font_entry)
    date_entry.set_date(datetime.today())
    date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    tk.Label(root, text="口座", font=font_label).grid(row=1, column=0, padx=5, pady=5, sticky="e")
    account_combo = ttk.Combobox(root, state="readonly", font=font_entry)
    account_combo["values"] = list(accounts.keys())
    account_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

    tk.Label(root, text="摘要", font=font_label).grid(row=2, column=0, padx=5, pady=5, sticky="e")
    summary_entry = tk.Entry(root, font=font_entry)
    summary_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

    tk.Label(root, text="金額", font=font_label).grid(row=3, column=0, padx=5, pady=5, sticky="e")
    amount_entry = tk.Entry(root, font=font_entry)
    amount_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

    mode_var = tk.StringVar(value="預入")
    mode_frame = tk.Frame(root)
    mode_frame.grid(row=4, column=0, columnspan=2, pady=10)
    tk.Radiobutton(mode_frame, text="預入", variable=mode_var, value="預入", font=font_entry).pack(side="left", padx=10)
    tk.Radiobutton(mode_frame, text="引出", variable=mode_var, value="引出", font=font_entry).pack(side="left", padx=10)

    tk.Button(root, text="登録", command=register_transaction, font=font_button).grid(row=5, column=0, columnspan=2, pady=10)
    tk.Button(root, text="閉じる", command=root.destroy, font=font_button).grid(row=6, column=0, columnspan=2, pady=10)
