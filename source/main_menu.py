import tkinter as tk
from view_transactions import show_transaction_window
from edit_transaction import show_edit_transaction_window
from register_transaction import show_register_window
from register_transaction_multi import show_register_multi_window
from add_account import show_add_account_window
from view_balances import show_balance_window

version = "v1.0.0"

root = tk.Tk()
root.title("帳簿アプリ メニュー")
root.geometry("400x600")
root.configure(bg="#f0f0f5")  # 背景をやや明るめのグレーに

font_title = ("Arial", 16, "bold")
font_button = ("Arial", 14)

tk.Label(root, text="帳簿メニュー", font=font_title, bg="#f0f0f5").pack(pady=20)
tk.Label(root, text=version, font=("Arial", 10), fg="gray").pack(pady=(0, 10))

def create_colored_button(text, command, bg_color, fg_color="white"):
    return tk.Button(
        root,
        text=text,
        command=command,
        font=font_button,
        width=20,
        bg=bg_color,
        fg=fg_color,
        activebackground="#cccccc"
    )

create_colored_button("入出金登録", show_register_window, "#4caf50").pack(pady=10)       # 緑
create_colored_button("入出金一括登録", show_register_multi_window, "#2196f3").pack(pady=10)  # 青
create_colored_button("取引履歴参照", show_transaction_window, "#9c27b0").pack(pady=10)     # 紫
create_colored_button("取引履歴修正", show_edit_transaction_window, "#ff9800").pack(pady=10)  # オレンジ
create_colored_button("口座追加", show_add_account_window, "#3f51b5").pack(pady=10)         # インディゴ
create_colored_button("残高一覧", show_balance_window, "#009688").pack(pady=10)             # ティール
create_colored_button("終了", root.destroy, "#f44336").pack(pady=20)                        # 赤

root.mainloop()