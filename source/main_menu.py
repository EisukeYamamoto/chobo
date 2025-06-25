import tkinter as tk
from view_transactions import show_transaction_window
from register_transaction import show_register_window
from add_account import show_add_account_window
from view_balances import show_balance_window

root = tk.Tk()
root.title("帳簿アプリ メニュー")
root.geometry("400x400")

font_title = ("Arial", 16, "bold")
font_button = ("Arial", 14)

tk.Label(root, text="帳簿メニュー", font=font_title).pack(pady=20)

tk.Button(root, text="入出金登録", command=show_register_window, font=font_button, width=20).pack(pady=10)
tk.Button(root, text="取引履歴参照", command=show_transaction_window, font=font_button, width=20).pack(pady=10)
tk.Button(root, text="口座追加", command=show_add_account_window, font=font_button, width=20).pack(pady=10)
tk.Button(root, text="残高一覧", command=show_balance_window, font=font_button, width=20).pack(pady=10)

tk.Button(root, text="終了", command=root.destroy, font=font_button, width=20).pack(pady=20)

root.mainloop()
