"""
AIST 0.0.0.3 — Ассистент для просмотра Excel-таблиц.
Добавлено: кнопки Инструкция и Закрыть, аккуратная раскладка.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


# ============================================================
# НАСТРОЙКИ
# ============================================================

CONTACT_PHONE = "+7 (000) 000-00-00"
CONTACT_EMAIL = "aist.help@example.com"
CONTACT_TELEGRAM = "@aist_support"


# ============================================================
# ФУНКЦИИ
# ============================================================

def open_file():
    """Открывает Excel-файл и показывает его размеры."""
    path = filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
    )
    if not path:
        return

    try:
        df = pd.read_excel(path)
    except Exception as e:
        messagebox.showerror("Ошибка чтения",
                             f"Не удалось прочитать файл:\n{e}")
        return

    info = (
        f"Файл: {path}\n\n"
        f"Строк:   {len(df)}\n"
        f"Колонок: {len(df.columns)}\n\n"
        f"Названия колонок:\n" + "\n".join(f"  • {c}" for c in df.columns)
    )
    label_info.config(text=info)


def show_help():
    """Окно Помощь — контакты и горячие клавиши."""
    if hasattr(show_help, "win") and show_help.win.winfo_exists():
        show_help.win.lift()
        show_help.win.focus_force()
        return

    win = tk.Toplevel(root)
    win.title("Помощь — AIST")
    win.geometry("480x420")
    win.resizable(False, False)

    tk.Label(win, text="🆘  Помощь", font=("Arial", 16, "bold")) \
        .pack(pady=(15, 5))

    tk.Label(win, text="Если что-то не работает — обращайтесь:",
             font=("Arial", 10)) \
        .pack(pady=(0, 15))

    frame_contacts = tk.Frame(win, bg="white", relief="solid", bd=1)
    frame_contacts.pack(fill="x", padx=20, pady=5)

    contacts_text = (
        f"📞  Телефон:   {CONTACT_PHONE}\n\n"
        f"📧  Почта:     {CONTACT_EMAIL}\n\n"
        f"💬  Telegram:  {CONTACT_TELEGRAM}"
    )
    tk.Label(frame_contacts, text=contacts_text,
             font=("Consolas", 11), bg="white",
             justify="left", anchor="w") \
        .pack(padx=15, pady=15, fill="x")

    tk.Label(win, text="Горячие клавиши:",
             font=("Arial", 11, "bold")) \
        .pack(pady=(15, 5), anchor="w", padx=20)

    hotkeys_text = (
        "Alt + V    — открыть Excel-файл\n"
        "F1         — эта справка\n"
        "F2         — инструкция\n"
        "Ctrl + Q   — закрыть программу\n"
        "Esc        — закрыть это окно"
    )
    tk.Label(win, text=hotkeys_text,
             font=("Consolas", 10), justify="left", anchor="w") \
        .pack(padx=40, fill="x")

    tk.Button(win, text="Закрыть", width=15,
              command=win.destroy).pack(pady=20)

    win.bind("<Escape>", lambda e: win.destroy())
    show_help.win = win


def show_instructions():
    """Окно Инструкция — подробное руководство."""
    if hasattr(show_instructions, "win") and show_instructions.win.winfo_exists():
        show_instructions.win.lift()
        show_instructions.win.focus_force()
        return

    win = tk.Toplevel(root)
    win.title("Инструкция — AIST")
    win.geometry("600x560")
    win.resizable(False, False)

    tk.Label(win, text="📖  Инструкция", font=("Arial", 16, "bold")) \
        .pack(pady=(15, 5))

    # Текст инструкции — многострочный, с прокруткой
    frame_text = tk.Frame(win)
    frame_text.pack(fill="both", expand=True, padx=20, pady=10)

    text = tk.Text(frame_text, wrap="word", font=("Arial", 11),
                   relief="solid", bd=1, padx=10, pady=10)
    scroll = tk.Scrollbar(frame_text, command=text.yview)
    text.configure(yscrollcommand=scroll.set)

    text.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")

    instructions = """Добро пожаловать в AIST!
    AIST — это ассистент для работы с Excel-таблицами.
    Он помогает бухгалтеру открывать и смотреть таблицы
    без запуска самого Excel.


    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    📂  КАК ОТКРЫТЬ ФАЙЛ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    1. Нажмите кнопку «📂 Открыть Excel».
    2. В открывшемся окне выберите нужный файл.
    3. Нажмите «Открыть».

    AIST прочитает файл и покажет:
       • сколько в нём строк,
       • сколько колонок,
       • названия колонок.

    Быстрый способ — клавиши Alt + V.


    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    🆘  ЕСЛИ ЧТО-ТО НЕ РАБОТАЕТ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Нажмите «❓ Помощь» (или клавишу F1) —
    там будут телефон и почта для связи.


    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    ⌨️  ГОРЯЧИЕ КЛАВИШИ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

       Alt + V    открыть Excel-файл
       F1         помощь
       F2         эта инструкция
       Ctrl + Q   закрыть программу
       Esc        закрыть текущее окно


    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    🐦  О ПРОГРАММЕ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Версия: 0.0.0.3
    Проект развивается постепенно.
    Если есть пожелания — напишите нам
    (контакты в окне «Помощь»).
    """

    text.insert("1.0", instructions)
    text.config(state="disabled")  # запрещаем редактирование, только чтение

    tk.Button(win, text="Понятно", width=15,
              command=win.destroy).pack(pady=15)

    win.bind("<Escape>", lambda e: win.destroy())
    show_instructions.win = win


def close_app():
    """Закрывает программу. Спрашивает подтверждение."""
    if messagebox.askyesno("Выход", "Закрыть AIST?"):
        root.destroy()


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================

root = tk.Tk()
root.title("🐦 AIST — 0.0.0.3")
root.geometry("550x440")

title = tk.Label(root, text="AIST", font=("Arial", 20, "bold"))
title.pack(pady=(15, 0))

subtitle = tk.Label(root, text="Ассистент для работы с Excel",
                    font=("Arial", 10), fg="gray")
subtitle.pack(pady=(0, 15))

# --- Две строки кнопок ---
frame_top = tk.Frame(root)
frame_top.pack(pady=(5, 2))

btn_open = tk.Button(frame_top, text="📂 Открыть Excel",
                     font=("Arial", 11), width=18, pady=4,
                     command=open_file)
btn_open.pack(side="left", padx=3)

btn_help = tk.Button(frame_top, text="❓ Помощь",
                     font=("Arial", 11), width=18, pady=4,
                     command=show_help)
btn_help.pack(side="left", padx=3)

frame_bottom = tk.Frame(root)
frame_bottom.pack(pady=(2, 10))

btn_instr = tk.Button(frame_bottom, text="📖 Инструкция",
                      font=("Arial", 11), width=18, pady=4,
                      command=show_instructions)
btn_instr.pack(side="left", padx=3)

btn_close = tk.Button(frame_bottom, text="🚪 Закрыть",
                      font=("Arial", 11), width=18, pady=4,
                      command=close_app)
btn_close.pack(side="left", padx=3)

# --- Информационное поле ---
label_info = tk.Label(root, text="Файл не выбран",
                      font=("Consolas", 10),
                      justify="left", anchor="nw",
                      bg="white", relief="sunken", bd=1)
label_info.pack(fill="both", expand=True, padx=15, pady=15)

# --- Горячие клавиши ---
root.bind("<Alt-v>", lambda e: open_file())
root.bind("<F1>", lambda e: show_help())
root.bind("<F2>", lambda e: show_instructions())
root.bind("<Control-q>", lambda e: close_app())

root.mainloop()