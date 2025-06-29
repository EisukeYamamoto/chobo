import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import pandas as pd

EXCEL_FILE = "database.xlsx"

def show_balance_window():
    def load_balances():
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws_accounts = wb["口座一覧"]
            ws_transactions = wb["取引履歴"]

            accounts = {}
            for row in ws_accounts.iter_rows(min_row=2, values_only=True):
                account_id, name, initial, account_type = row[:4]
                display_name = f"{name}（{account_type}）"
                accounts[account_id] = {
                    "name": display_name,
                    "initial": float(initial),
                    "deposit": 0.0,
                    "withdrawal": 0.0
                }

            for row in ws_transactions.iter_rows(min_row=2, values_only=True):
                _, acc_id, _, deposit, withdrawal, *_ = row
                if acc_id in accounts:
                    accounts[acc_id]["deposit"] += float(deposit or 0)
                    accounts[acc_id]["withdrawal"] += float(withdrawal or 0)

            results = []
            total_balance = 0
            for acc in accounts.values():
                current = acc["initial"] + acc["deposit"] - acc["withdrawal"]
                results.append([
                    acc["name"],
                    acc["deposit"],
                    acc["withdrawal"],
                    current
                ])
                total_balance += current
            wb.close()
            return results, total_balance
        except Exception as e:
            messagebox.showerror("エラー", f"残高データの読み込みに失敗しました\n{e}")
            return [], 0

    def show_balances():
        nonlocal current_data, current_total
        records, total = load_balances()
        current_data = records
        current_total = total
        for row in tree.get_children():
            tree.delete(row)
        for r in records:
            tree.insert("", tk.END, values=[
                r[0],
                f"{r[1]:,.0f}",
                f"{r[2]:,.0f}",
                f"{r[3]:,.0f}"
            ])
        total_balance_text.set(f"全口座の残高合計：{total:,.0f} 円")

    def export_to_excel():
        if not current_data:
            messagebox.showinfo("出力なし", "出力するデータがありません")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excelファイル", "*.xlsx")])
        if not file_path:
            return
        df = pd.DataFrame(
            [row[:4] for row in current_data],
            columns=["口座", "合計預入", "合計引出", "現在残高"]
        )
        df_summary = pd.DataFrame([{
            "口座": "全口座の残高合計",
            "合計預入": "",
            "合計引出": "",
            "現在残高": current_total
        }])
        df_out = pd.concat([df, df_summary], ignore_index=True)
        df_out.to_excel(file_path, index=False)
        messagebox.showinfo("出力完了", f"{file_path} に出力しました")

    balance_win = tk.Toplevel()
    balance_win.title("残高一覧")
    balance_win.geometry("700x500")
    balance_win.minsize(500, 400)

    font_label = ("Arial", 12)
    font_tree = ("Arial", 11)
    font_button = ("Arial", 12)

    current_data = []
    current_total = 0
    total_balance_text = tk.StringVar(value="全口座の残高合計：― 円")

    balance_win.grid_rowconfigure(0, weight=1)
    balance_win.grid_rowconfigure(1, weight=0)
    balance_win.grid_rowconfigure(2, weight=0)
    balance_win.grid_columnconfigure(0, weight=1)

    tree = ttk.Treeview(balance_win, columns=("口座", "合計預入", "合計引出", "現在残高"), show="headings", height=20)
    style = ttk.Style()
    style.configure("Treeview.Heading", font=font_label)
    style.configure("Treeview", font=font_tree, rowheight=28)

    for col in ("口座", "合計預入", "合計引出", "現在残高"):
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center", stretch=True)

    tree.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

    tk.Label(balance_win, textvariable=total_balance_text, font=("Arial", 13, "bold")).grid(row=1, column=0, pady=5)

    btn_frame = tk.Frame(balance_win)
    btn_frame.grid(row=2, column=0, pady=10)

    tk.Button(btn_frame, text="更新", command=show_balances, font=font_button).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Excelに出力", command=export_to_excel, font=font_button).pack(side="left", padx=10)
    tk.Button(btn_frame, text="閉じる", command=balance_win.destroy, font=font_button).pack(side="left", padx=10)

    show_balances()
