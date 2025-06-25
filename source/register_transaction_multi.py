import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import openpyxl
import os

EXCEL_FILE = "database.xlsx"

def show_register_multi_window():
    if not os.path.exists(EXCEL_FILE):
        messagebox.showerror("ファイルエラー", f"{EXCEL_FILE} が見つかりません。")
        return

    def load_accounts():
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["口座一覧"]
        accounts = {row[1]: row[0] for row in ws.iter_rows(min_row=2, values_only=True)}
        wb.close()
        return accounts

    def save_transactions(account_id, entries):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["取引履歴"]
        for entry in entries:
            ws.append([entry['date'], account_id, entry['summary'], entry['deposit'], entry['withdrawal']])
        wb.save(EXCEL_FILE)
        wb.close()

    def get_current_balance(account_id):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws1 = wb["口座一覧"]
        ws2 = wb["取引履歴"]
        initial = 0
        for row in ws1.iter_rows(min_row=2, values_only=True):
            if row[0] == account_id:
                initial = float(row[2])
                break
        deposit_total = 0
        withdrawal_total = 0
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if row[1] == account_id:
                deposit_total += float(row[3] or 0)
                withdrawal_total += float(row[4] or 0)
        wb.close()
        return initial + deposit_total - withdrawal_total

    def show_confirmation(entries):
        popup = tk.Toplevel()
        popup.title("登録内容の確認")
        popup.geometry("600x300")

        font_header = ("Arial", 11, "bold")
        font_row = ("Arial", 11)

        header_frame = tk.Frame(popup)
        header_frame.pack(fill="x", padx=10, pady=5)

        for text, width in zip(["日付", "摘要", "預入", "引出"], [15, 25, 10, 10]):
            tk.Label(header_frame, text=text, width=width, font=font_header, anchor="w").pack(side="left")

        for entry in entries:
            row_frame = tk.Frame(popup)
            row_frame.pack(fill="x", padx=10)
            tk.Label(row_frame, text=entry["date"], width=15, font=font_row, anchor="w").pack(side="left")
            tk.Label(row_frame, text=entry["summary"], width=25, font=font_row, anchor="w").pack(side="left")
            tk.Label(row_frame, text=entry["deposit"] if entry["deposit"] != "" else "", width=10, font=font_row, anchor="e").pack(side="left")
            tk.Label(row_frame, text=entry["withdrawal"] if entry["withdrawal"] != "" else "", width=10, font=font_row, anchor="e").pack(side="left")

        tk.Button(popup, text="閉じる", font=("Arial", 11), command=popup.destroy).pack(pady=10)

    def register_all():
        selected_account = account_combo.get()
        if not selected_account:
            messagebox.showwarning("口座未選択", "口座を選択してください")
            return

        account_id = accounts[selected_account]
        all_entries = []

        for widgets in rows:
            date_str = widgets[0].get().strip()
            summary = widgets[1].get().strip()
            deposit = widgets[2].get().strip()
            withdrawal = widgets[3].get().strip()

            if not (date_str and summary and (deposit or withdrawal)):
                continue

            try:
                input_date = datetime.strptime(date_str, "%Y-%m-%d")
                if input_date > datetime.today():
                    messagebox.showwarning("日付エラー", f"{date_str} は未来日付です")
                    return
            except ValueError:
                messagebox.showwarning("日付エラー", f"{date_str} の形式が正しくありません")
                return

            try:
                deposit_val = float(deposit) if deposit else ""
                withdrawal_val = float(withdrawal) if withdrawal else ""
            except ValueError:
                messagebox.showwarning("金額エラー", f"{summary} の金額が不正です")
                return

            all_entries.append({
                'date': date_str,
                'summary': summary,
                'deposit': deposit_val,
                'withdrawal': withdrawal_val
            })

        if not all_entries:
            messagebox.showinfo("確認", "登録する有効な取引がありません")
            return

        try:
            save_transactions(account_id, all_entries)
            show_confirmation(all_entries)
            balance = get_current_balance(account_id)
            messagebox.showinfo("登録完了", f"{len(all_entries)} 件登録しました。\n現在の残高: {balance:,.0f} 円")
            for widgets in rows:
                for w in widgets:
                    w.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("保存エラー", f"保存中にエラーが発生しました：\n{e}")

    def add_row():
        row_frame = tk.Frame(entries_frame)
        row_frame.pack(fill="x", pady=2)

        date_entry = ttk.Entry(row_frame, font=font_entry, width=12)
        date_entry.insert(0, datetime.today().strftime("%Y-%m-%d"))
        date_entry.pack(side="left", padx=5)

        summary_entry = ttk.Entry(row_frame, font=font_entry, width=20)
        summary_entry.pack(side="left", padx=5)

        deposit_entry = ttk.Entry(row_frame, font=font_entry, width=10)
        deposit_entry.pack(side="left", padx=5)

        withdrawal_entry = ttk.Entry(row_frame, font=font_entry, width=10)
        withdrawal_entry.pack(side="left", padx=5)

        rows.append((date_entry, summary_entry, deposit_entry, withdrawal_entry))

    root = tk.Toplevel()
    root.title("複数行 入出金登録")
    root.geometry("850x600")

    font_label = ("Arial", 12)
    font_entry = ("Arial", 12)
    font_button = ("Arial", 12)

    accounts = load_accounts()
    rows = []

    tk.Label(root, text="口座名", font=font_label).pack(pady=(10, 0))
    account_combo = ttk.Combobox(root, font=font_entry, state="readonly")
    account_combo["values"] = list(accounts.keys())
    account_combo.pack(pady=5)

    header_frame = tk.Frame(root)
    header_frame.pack(fill="x", padx=20)
    for text, width in zip(["日付", "摘要", "預入", "引出"], [12, 20, 10, 10]):
        tk.Label(header_frame, text=text, width=width, anchor="w", font=font_label).pack(side="left", padx=5)

    entries_frame = tk.Frame(root)
    entries_frame.pack(fill="both", expand=True, padx=20, pady=10)

    add_row()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="＋ 行追加", command=add_row, font=font_button).pack(side="left", padx=10)
    tk.Button(btn_frame, text="登録", command=register_all, font=font_button).pack(side="left", padx=10)
    tk.Button(btn_frame, text="メニューに戻る", command=root.destroy, font=font_button).pack(side="left", padx=10)

