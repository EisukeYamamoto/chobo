import tkinter as tk
from tkinter import messagebox
import subprocess
import os
import sys

def open_script(script_name):
    try:
        # Python実行パスで指定
        python_exe = sys.executable
        subprocess.Popen([python_exe, script_name])
    except Exception as e:
        messagebox.showerror("エラー", f"起動に失敗しました: {e}")

root = tk.Tk()
root.title("帳簿管理アプリ メニュー")
root.geometry("300x200")

tk.Label(root, text="機能を選んでください", font=("Arial", 14)).pack(pady=20)

tk.Button(root, text="① 入出金を登録", command=lambda: open_script("register_transaction.py"), width=25).pack(pady=5)
tk.Button(root, text="② 新しい口座を追加", command=lambda: open_script("add_account.py"), width=25).pack(pady=5)
tk.Button(root, text="③ 取引履歴を参照・出力", command=lambda: open_script("view_transactions.py"), width=25).pack(pady=5)

root.mainloop()
