import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import pandas as pd
from datetime import datetime
from tkcalendar import DateEntry

EXCEL_FILE = "database.xlsx"

def load_accounts():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["口座一覧"]
    accounts = {row[1]: row[0] for row in ws.iter_rows(min_row=2, values_only=True)}
    wb.close()
    return accounts

def get_initial_balance(account_name):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["口座一覧"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] == account_name:
            wb.close()
            return float(row[2])
    wb.close()
    return 0.0

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
    initial_balance = get_initial_balance(account_name)
    current_balance = initial_balance

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["取引履歴"]
    records = []
    total_deposit = 0
    total_withdrawal = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        date_str, acc_id, summary, deposit, withdrawal = row
        if acc_id != account_id:
            continue
        try:
            tx_date = datetime.strptime(str(date_str), "%Y-%m-%d")
        except Exception:
            continue
        if start <= tx_date <= end:
            deposit_amt = float(deposit or 0)
            withdrawal_amt = float(withdrawal or 0)
            current_balance += deposit_amt - withdrawal_amt
            records.append([
                date_str,
                summary,
                deposit if deposit else "",
                withdrawal if withdrawal else "",
                current_balance
            ])
            total_deposit += deposit_amt
            total_withdrawal += withdrawal_amt
    wb.close()

    for row in tree.get_children():
        tree.delete(row)
    for r in records:
        tree.insert("", tk.END, values=r)

    global current_results, final_balance
    current_results = records
    final_balance = current_balance
    deposit_var.set(f"{total_deposit:,.0f} 円")
    withdrawal_var.set(f"{total_withdrawal:,.0f} 円")
    balance_var.set(f"{current_balance:,.0f} 円")

def export_to_excel():
    if not current_results:
        messagebox.showinfo("出力なし", "出力するデータがありません")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    df = pd.DataFrame(current_results, columns=["日付", "摘要", "預入", "引出", "残高"])
    deposit_total = sum(float(r[2]) for r in current_results if r[2] != "")
    withdrawal_total = sum(float(r[3]) for r in current_results if r[3] != "")

    df_summary = pd.DataFrame([
        {"日付": "", "摘要": "", "預入": "", "引出": "", "残高": ""},
        {"日付": "合計預入", "摘要": "", "預入": deposit_total, "引出": "", "残高": ""},
        {"日付": "合計引出", "摘要": "", "預入": "", "引出": withdrawal_total, "残高": ""},
        {"日付": "残高", "摘要": "", "預入": final_balance, "引出": "", "残高": ""}
    ])
    df_out = pd.concat([df, df_summary], ignore_index=True)
    df_out.to_excel(file_path, index=False)
    messagebox.showinfo("出力完了", f"{file_path} に出力しました")

# --- GUI構築 ---
root = tk.Tk()
root.title("取引履歴の参照と出力（残高付き）")

accounts = load_accounts()
current_results = []
final_balance = 0.0

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

tree = ttk.Treeview(root, columns=("日付", "摘要", "預入", "引出", "残高"), show="headings")
for col in ("日付", "摘要", "預入", "引出", "残高"):
    tree.heading(col, text=col)
    tree.column(col, width=100)
tree.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

tk.Label(root, text="合計預入").grid(row=5, column=0)
deposit_var = tk.StringVar(value="―")
tk.Label(root, textvariable=deposit_var).grid(row=5, column=1)

tk.Label(root, text="合計引出").grid(row=6, column=0)
withdrawal_var = tk.StringVar(value="―")
tk.Label(root, textvariable=withdrawal_var).grid(row=6, column=1)

tk.Label(root, text="残高").grid(row=7, column=0)
balance_var = tk.StringVar(value="―")
tk.Label(root, textvariable=balance_var, font=("Arial", 12, "bold")).grid(row=7, column=1)

tk.Button(root, text="Excelに出力", command=export_to_excel).grid(row=8, column=0, columnspan=2, pady=10)

root.mainloop()
