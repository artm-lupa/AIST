"""
AIST 0.1.0.4 — Ассистент для просмотра Excel-таблиц.
Тема: "Тёплый вечер". CustomTkinter + заставка.
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image
import pandas as pd


# ============================================================
# НАСТРОЙКИ
# ============================================================

CONTACT_PHONE = "+7 (000) 000-00-00"
CONTACT_EMAIL = "aist.help@example.com"
CONTACT_TELEGRAM = "@aist_support"

# Путь к папке с картинками (рядом с main.py)
ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
SPLASH_PATH = os.path.join(ASSETS_DIR, "splash.png")

# --- Цветовая схема "Тёплый вечер" ---
COLOR_BG       = "#2a1f1a"   # основной фон
COLOR_FRAME    = "#3a2d26"   # панели
COLOR_ACCENT   = "#e5a02a"   # золотой акцент (кнопки)
COLOR_ACCENT_H = "#f0b040"   # акцент при наведении
COLOR_TEXT     = "#f0e5d0"   # основной текст
COLOR_TEXT_DIM = "#a89880"   # приглушённый текст
COLOR_FIELD    = "#1a1210"   # поле ввода/вывода


# ============================================================
# ЗАСТАВКА (splash screen)
# ============================================================

def show_splash():
    """Показывает заставку на 2 секунды, потом открывает главное окно."""
    splash = ctk.CTk()
    splash.title('AIST')
    splash.geometry('736x736')
    splash.resizable(False, False)
    splash.configure(fg_color=COLOR_BG)

    # Размеры и позиция (по центру экрана)
    splash.update_idletasks()
    w = splash.winfo_width()
    h = splash.winfo_height()
    screen_w = splash.winfo_screenwidth()
    screen_h = splash.winfo_screenheight()
    x = (screen_w - w) // 2
    y = (screen_h - h) // 2
    splash.geometry(f"+{x}+{y}")

    # Пытаемся загрузить картинку
    try:
        img = Image.open(SPLASH_PATH)
        img.thumbnail((420, 420), Image.LANCZOS)
        photo = ctk.CTkImage(light_image=img, dark_image=img,
                             size=img.size)

        label_img = ctk.CTkLabel(splash, image=photo, text="")
        label_img.image = photo   # сохраняем ссылку, иначе Python удалит
        label_img.pack(pady=(30, 10), padx=30)
    except Exception as e:
        # Если картинки нет — просто текст
        ctk.CTkLabel(splash, text="(картинка не найдена)",
                     font=("Arial", 12),
                     text_color=COLOR_TEXT_DIM).pack(pady=60)
        print(f"Ошибка загрузки заставки: {e}")

    # Подпись снизу
    ctk.CTkLabel(splash, text="🐦  AIST",
                 font=("Arial", 22, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(5, 0))

    ctk.CTkLabel(splash, text="Ассистент для работы с Excel",
                 font=("Arial", 11),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 20))

    # Через 2 секунды закрываем заставку и открываем главное окно
    def go_to_main():
        splash.destroy()
        main()

    splash.after(2000, go_to_main)
    splash.mainloop()


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
    label_info.configure(text=info)


def show_help():
    """Окно Помощь."""
    if hasattr(show_help, "win") and show_help.win.winfo_exists():
        show_help.win.lift()
        show_help.win.focus_force()
        return
    win = ctk.CTkToplevel(root)
    win.title("Помощь — AIST")
    win.geometry("500x460")
    win.resizable(False, False)
    win.configure(fg_color=COLOR_BG)

    ctk.CTkLabel(win, text="🆘  Помощь",
                 font=("Arial", 20, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(20, 5))

    ctk.CTkLabel(win, text="Если что-то не работает — обращайтесь:",
                 font=("Arial", 12),
                 text_color=COLOR_TEXT).pack(pady=(0, 15))

    contacts_frame = ctk.CTkFrame(win, fg_color=COLOR_FRAME)
    contacts_frame.pack(fill="x", padx=25, pady=5)

    contacts_text = (
        f"📞  Телефон:   {CONTACT_PHONE}\n\n"
        f"📧  Почта:     {CONTACT_EMAIL}\n\n"
        f"💬  Telegram:  {CONTACT_TELEGRAM}"
    )
    ctk.CTkLabel(contacts_frame, text=contacts_text,
                 font=("Consolas", 12),
                 text_color=COLOR_TEXT,
                 justify="left", anchor="w").pack(padx=20, pady=18, fill="x")

    ctk.CTkLabel(win, text="Горячие клавиши:",
                 font=("Arial", 13, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(20, 5),
                                               anchor="w", padx=25)

    hotkeys_text = (
        "Alt + V    — открыть Excel-файл\n"
        "F1         — эта справка\n"
        "F2         — инструкция\n"
        "Ctrl + Q   — закрыть программу\n"
        "Esc        — закрыть это окно"
    )
    ctk.CTkLabel(win, text=hotkeys_text,
                 font=("Consolas", 11),
                 text_color=COLOR_TEXT,
                 justify="left", anchor="w").pack(padx=50, fill="x")

    ctk.CTkButton(win, text="Закрыть", width=140,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=win.destroy).pack(pady=20)

    win.bind("<Escape>", lambda e: win.destroy())
    show_help.win = win


def show_instructions():
    """Окно Инструкция."""
    if hasattr(show_instructions, "win") and show_instructions.win.winfo_exists():
        show_instructions.win.lift()
        show_instructions.win.focus_force()
        return

    win = ctk.CTkToplevel(root)
    win.title("Инструкция — AIST")
    win.geometry("640x620")
    win.resizable(False, False)
    win.configure(fg_color=COLOR_BG)

    ctk.CTkLabel(win, text="📖  Инструкция",
                 font=("Arial", 20, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(20, 10))

    # Текстовое поле с прокруткой
    text_frame = ctk.CTkFrame(win, fg_color=COLOR_BG)
    text_frame.pack(fill="both", expand=True, padx=25, pady=10)

    textbox = tk.Text(
        text_frame, wrap="word", font=("Arial", 12),
        bg=COLOR_FIELD, fg=COLOR_TEXT,
        insertbackground=COLOR_TEXT,
        relief="flat", bd=0, padx=15, pady=15
    )
    scroll = tk.Scrollbar(text_frame, command=textbox.yview)
    textbox.configure(yscrollcommand=scroll.set)

    textbox.pack(side="left", fill="both", expand=True)
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

    Версия: 0.1.0.0
    Тема: «Тёплый вечер»
    Проект развивается постепенно.
    Если есть пожелания — напишите нам
    (контакты в окне «Помощь»).
    """

    textbox.insert("1.0", instructions)
    textbox.config(state="disabled")

    ctk.CTkButton(win, text="Понятно", width=140,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=win.destroy).pack(pady=15)

    win.bind("<Escape>", lambda e: win.destroy())
    show_instructions.win = win


def close_app():
    """Закрывает программу с подтверждением и благодарностью."""
    # Шаг 1. Спрашиваем подтверждение
    if not messagebox.askyesno("Выход", "Закрыть AIST?"):
        return

    # Шаг 2. Показываем окно благодарности
    thanks = ctk.CTkToplevel(root)
    thanks.title("До встречи!")
    thanks.geometry("420x280")
    thanks.resizable(False, False)
    thanks.configure(fg_color=COLOR_BG)

    # Центрируем окно на экране
    thanks.update_idletasks()
    screen_w = thanks.winfo_screenwidth()
    screen_h = thanks.winfo_screenheight()
    x = (screen_w - 420) // 2
    y = (screen_h - 280) // 2
    thanks.geometry(f"+{x}+{y}")

    # Текст благодарности
    ctk.CTkLabel(thanks, text="😊",
                 font=("Arial", 40),
                 text_color=COLOR_ACCENT).pack(pady=(25, 5))

    ctk.CTkLabel(thanks, text="Спасибо!",
                 font=("Arial", 20, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(0, 10))

    ctk.CTkLabel(thanks,
                 text="Спасибо, что пользуетесь\nпрограммой AIST.",
                 font=("Arial", 13),
                 text_color=COLOR_TEXT,
                 justify="center").pack(pady=(0, 5))

    ctk.CTkLabel(thanks, text="Всего хорошего! ✨",
                 font=("Arial", 13, "italic"),
                 text_color=COLOR_TEXT).pack(pady=(0, 5))

    ctk.CTkLabel(thanks, text="Обращайтесь ещё",
                 font=("Arial", 11),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 15))

    # Через 2 секунды закрыть окно и выйти из программы
    thanks.after(4000, lambda: (thanks.destroy(), root.destroy()))

    # Если пользователь закроет окно крестиком — просто выйти
    thanks.protocol("WM_DELETE_WINDOW",
                    lambda: (thanks.destroy(), root.destroy()))


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================

def main():
    """Собирает и показывает главное окно."""
    global root, label_info

    # Базовые настройки CustomTkinter
    ctk.set_appearance_mode("dark")  # тёмная тема
    ctk.set_default_color_theme("dark-blue")

    root = ctk.CTk()
    root.title("🐦 AIST — 0.1.0.0")
    root.geometry("700x520")
    root.configure(fg_color=COLOR_BG)

    # --- Заголовок ---
    ctk.CTkLabel(root, text="AIST",
                 font=("Arial", 26, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(25, 0))

    ctk.CTkLabel(root, text="Ассистент для работы с Excel",
                 font=("Arial", 12),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 20))

    # --- Верхний ряд кнопок ---
    frame_top = ctk.CTkFrame(root, fg_color="transparent")
    frame_top.pack(pady=(5, 3))

    ctk.CTkButton(frame_top, text="📂 Открыть Excel",
                  font=("Arial", 13), width=200, height=40,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=open_file).pack(side="left", padx=6)

    ctk.CTkButton(frame_top, text="❓ Помощь",
                  font=("Arial", 13), width=200, height=40,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_help).pack(side="left", padx=6)

    # --- Нижний ряд кнопок ---
    frame_bottom = ctk.CTkFrame(root, fg_color="transparent")
    frame_bottom.pack(pady=(3, 15))

    ctk.CTkButton(frame_bottom, text="📖 Инструкция",
                  font=("Arial", 13), width=200, height=40,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_instructions).pack(side="left", padx=6)

    ctk.CTkButton(frame_bottom, text="🚪 Закрыть",
                  font=("Arial", 13), width=200, height=40,
                  fg_color=COLOR_FRAME,
                  hover_color="#7a2020",
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color="#a03030",
                  command=close_app).pack(side="left", padx=6)

    # --- Поле с информацией ---
    info_frame = ctk.CTkFrame(root, fg_color=COLOR_FIELD,
                              corner_radius=8)
    info_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    label_info = ctk.CTkLabel(info_frame, text="Файл не выбран",
                              font=("Consolas", 12),
                              text_color=COLOR_TEXT,
                              justify="left", anchor="nw",
                              wraplength=640)
    label_info.pack(padx=20, pady=20, fill="both", expand=True)

    # --- Горячие клавиши ---
    root.bind("<Alt-v>", lambda e: open_file())
    root.bind("<F1>", lambda e: show_help())
    root.bind("<F2>", lambda e: show_instructions())
    root.bind("<Control-q>", lambda e: close_app())

    root.mainloop()
    # ============================================================
    # СТАРТ ПРОГРАММЫ
    # ============================================================

if __name__ == "__main__":
    show_splash()
