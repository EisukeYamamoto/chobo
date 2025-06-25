import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import openpyxl
import pandas as pd
from datetime import datetime

EXCEL_FILE = "database.xlsx"

def load_accounts():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["口座一覧"]
    accounts = {row[1]: row[0] for row in ws.iter_rows(min_row=2, values_only=True)}  # 口座名:ID
    wb.close()
    return accounts

def search_transactions():
    try:
        start = datetime.strptime(start_entry.get(), "%Y-%m-%d")
        end = datetime.strptime(end_entry.get(), "%Y-%m-%d")
    except ValueError:
        messagebox.showerror("日付エラー", "日付は YYYY-MM-DD 形式で入力してください")
        return

    account_name = account_combo.get()
    if not account_name:
        messagebox.showerror("選択エラー", "口座を選択してください")
        return

    account_id = accounts[account_name]
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["取引履歴"]
    records = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        date_str, acc_id, summary, deposit, withdrawal = row
        if acc_id != account_id:
            continue
        try:
            tx_date = datetime.strptime(str(date_str), "%Y-%m-%d")
        except Exception:
            continue
        if start <= tx_date <= end:
            records.append([date_str, summary, deposit or "", withdrawal or ""])
    wb.close()

    for row in tree.get_children():
        tree.delete(row)
    for r in records:
        tree.insert("", tk.END, values=r)

    global current_results
    current_results = records

def export_to_excel():
    if not current_results:
        messagebox.showinfo("出力なし", "出力するデータがありません")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    df = pd.DataFrame(current_results, columns=["日付", "摘要", "預入", "引出"])
    df.to_excel(file_path, index=False)
    messagebox.showinfo("出力完了", f"{file_path} に出力しました")

# --- GUI構築 ---
root = tk.Tk()
root.title("取引履歴の参照と出力")

accounts = load_accounts()
current_results = []

tk.Label(root, text="口座").grid(row=0, column=0)
account_combo = ttk.Combobox(root, values=list(accounts.keys()), state="readonly")
account_combo.grid(row=0, column=1)

tk.Label(root, text="開始日").grid(row=1, column=0)
start_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
start_entry.grid(row=1, column=1)

tk.Label(root, text="終了日").grid(row=2, column=0)
end_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
end_entry.grid(row=2, column=1)

tk.Button(root, text="検索", command=search_transactions).grid(row=3, column=0, columnspan=2, pady=5)

tree = ttk.Treeview(root, columns=("日付", "摘要", "預入", "引出"), show="headings")
for col in ("日付", "摘要", "預入", "引出"):
    tree.heading(col, text=col)
    tree.column(col, width=100)
tree.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

tk.Button(root, text="Excelに出力", command=export_to_excel).grid(row=5, column=0, columnspan=2)

root.mainloop()
