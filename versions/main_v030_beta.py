"""
AIST 0.3.0.1 — Ассистент для просмотра Excel-таблиц.
Тема: "Тёплый вечер". CustomTkinter + заставка.
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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
COLOR_BG = "#2a1f1a"  # основной фон
COLOR_FRAME = "#3a2d26"  # панели
COLOR_ACCENT = "#e5a02a"  # золотой акцент (кнопки)
COLOR_ACCENT_H = "#f0b040"  # акцент при наведении
COLOR_TEXT = "#f0e5d0"  # основной текст
COLOR_TEXT_DIM = "#a89880"  # приглушённый текст
COLOR_FIELD = "#1a1210"  # поле ввода/вывода
is_topmost=False


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
        label_img.image = photo  # сохраняем ссылку, иначе Python удалит
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

def clean_dataframe(df):
    """
    Чистит "грязный" DataFrame:
    - удаляет пустые строки сверху и снизу
    - удаляет пустые колонки слева и справа
    - заменяет 'Unnamed: N' на первую строку данных
    - возвращает (очищенный_df, отчёт_очистки)
    """
    report = []
    original_rows = len(df)
    original_cols = len(df.columns)

    # --- 1. Удаляем полностью пустые строки и колонки ---
    df = df.dropna(how="all")  # строки, где все значения NaN
    df = df.dropna(axis=1, how="all")  # колонки, где все значения NaN

    # --- 2. Удаляем пустые строки и колонки в самом начале ---
    # (после dropna они могут остаться "полупустыми" из-за строковых "")
    df = df.loc[~(df == "").all(axis=1)]
    df = df.loc[:, ~(df == "").all(axis=0)]

    # --- 3. Проверяем, что осталось ---
    if df.empty:
        return df, ["Файл полностью пуст"]

    # --- 4. Если все колонки называются 'Unnamed: N' — берём первую строку как заголовки ---
    unnamed_count = sum(1 for c in df.columns if str(c).startswith("Unnamed:"))
    if unnamed_count == len(df.columns):
        # Первая строка станет заголовками
        new_headers = df.iloc[0].tolist()
        df = df.iloc[1:].reset_index(drop=True)
        df.columns = [str(h) for h in new_headers]
        report.append(f"Заголовки взяты из первой строки данных")

    # --- 5. Считаем, что изменилось ---
    removed_rows = original_rows - len(df)
    removed_cols = original_cols - len(df.columns)

    if removed_rows > 0:
        report.append(f"удалено пустых строк: {removed_rows}")
    if removed_cols > 0:
        report.append(f"удалено пустых колонок: {removed_cols}")

    return df, report

def show_functions():
    """Открывает окно со списком функций (меню)."""
    if hasattr(show_functions, "win") and show_functions.win.winfo_exists():
        show_functions.win.lift()
        show_functions.win.focus_force()
        return

    win = ctk.CTkToplevel(root)
    win.title("Функции — AIST")
    win.geometry("420x340")
    win.resizable(False, False)
    win.configure(fg_color=COLOR_BG)

    # Центрируем окно
    win.update_idletasks()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    x = (screen_w - 420) // 2
    y = (screen_h - 340) // 2
    win.geometry(f"+{x}+{y}")

    ctk.CTkLabel(win, text="📋  Функции",
                 font=("Arial", 20, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(25, 5))

    ctk.CTkLabel(win, text="Что сделать с таблицей?",
                 font=("Arial", 11),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 20))

    # --- Кнопка "Открыть Excel" ---
    ctk.CTkButton(win, text="📂  Открыть Excel",
                  font=("Arial", 13), width=300, height=45,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=open_file).pack(pady=8)

    # --- Заглушка для будущих функций ---
    ctk.CTkLabel(win, text="(здесь будут другие функции)",
                 font=("Arial", 10, "italic"),
                 text_color=COLOR_TEXT_DIM).pack(pady=(5, 15))

    ctk.CTkButton(win, text="Закрыть", width=140,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  command=win.destroy).pack(pady=5)

    win.bind("<Escape>", lambda e: win.destroy())
    show_functions.win = win

def toggle_topmost():
    """Переключает режим 'поверх всех окон'."""
    global is_topmost

    is_topmost = not is_topmost
    root.attributes("-topmost", is_topmost)

    if is_topmost:
        btn_topmost.configure(text="📌 Закреплено",
                              fg_color=COLOR_ACCENT,
                              text_color="#1a1210")
    else:
        btn_topmost.configure(text="📌 Закрепить",
                              fg_color=COLOR_FRAME,
                              text_color=COLOR_TEXT)

def open_file():
    """Открывает Excel-файл и показывает данные в таблице."""
    path = filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
    )
    if not path:
        return

    try:
        df = pd.read_excel(path, header=None)
        # print("=== ОТЛАДКА ===")
        # print("Форма до очистки:", df.shape)
        # print("Первые 5 строк:")
        print(df.head())
        # print("---")

        df, report = clean_dataframe(df)
        # print("Форма после очистки:", df.shape)
        # print("Колонки после очистки:", list(df.columns))
        # print("Отчёт:", report)
        # print("=== КОНЕЦ ===")
    except Exception as e:
        messagebox.showerror("Ошибка чтения",
                             f"Не удалось прочитать файл:\n{e}")
        return

    # --- 1. Очищаем таблицу от прошлых данных ---

    print(df.head(3))
    for row in tree.get_children():
        tree.delete(row)

    # --- 2. Настраиваем колонки ---
    columns = [str(c) for c in df.columns]

    tree["columns"] = columns

    # Служебная нулевая колонка (номер строки) — скрываем
    tree["show"] = "headings"

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=140, anchor="w", stretch=False)

    # --- 3. Заполняем данными (первые 500 строк) ---
    limit = 500
    total = len(df)
    show_n = min(limit, total)

    for i in range(show_n):
        row_values = [str(v) for v in df.iloc[i].tolist()]
        tree.insert("", "end", values=row_values)

    # --- 4. Обновляем статус внизу ---
    if total <= limit:
        status_text = f"📂 {path}   |   показано {total} строк"
    else:
        status_text = (f"📂 {path}   |   показано {show_n} из {total} строк "
                       f"(остальные скрыты для скорости)")

    status_label.configure(text=status_text)


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
            📋  МЕНЮ "ФУНКЦИИ"
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Кнопка "📋 Функции" открывает окно
    со списком действий.
    
    Сейчас там одна функция:
    • "📂 Открыть Excel"
    
    Скоро появятся:
    • "🔗 Объединить две таблицы"
    • "📁 Обработать папку"
    • "📊 Свод по контрагенту"

    Если кнопок станет много — они
    не поместятся в главном окне.
    А вменю "Функции" — поместятся.


    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        📌  КНОПКА "ЗАКРЕПИТЬ"
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Кнопка "📌 Закрепить" делает окно AIST
    поверх всех других программ.

    Это удобно, когда вы работаете
    с Excel или 1С одновременно с AIST:
    программа всегда остаётся на виду.

    Как использовать:
       1. Нажмите "📌 Закрепить".
     2. Кнопка изменится на "📌 Закреплено".
       3. AIST теперь поверх всех окон.
       4. Чтобы отменить — нажмите снова.

    Быстрая клавиша: Ctrl + T

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    🛠  ЕСЛИ ФАЙЛ "ГРЯЗНЫЙ" — НЕ СТРАШНО!
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Часто Excel-файлы содержат лишнее:
     • пустая строка сверху (для красоты)
     • пустая колонка слева
     • пустая строка посреди данных

    Раньше это называлось "грязный файл" —
    и с ним было сложно работать.

    ТЕПЕРЬ AIST СПРАВЛЯЕТСЯ С ЭТИМ САМ!
    Он автоматически:
    • убирает пустые строки и колонки
    • берёт заголовки из правильной строки
    • показывает пустые ячейки как пустые

    Вам НЕ нужно чистить файл в Excel —
    просто откройте его в AIST, и всё будет ок!

    В статусной строке внизу будет видно,
    что именно AIST убрал из файла.

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            💡   ПРОСТОЙ СОВЕТ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Если можете — делайте файлы "чистыми":
    • первая строка — заголовки
    • первая колонка — не пустая
     • нет лишних пустых строк

    Тогда AIST откроет файл ещё быстрее.

    Но если файл "грязный" — НЕ ПЕРЕЖИВАЙТЕ.
    0AIST всё равно справится.

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
       F3         функции
       Ctrl+T     Закрепить/открепить роверх окон

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
      🐦  ЧЕСТНОЕ СЛОВО ОТ АВТОРА
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        AIST — молодая программа.
    Она написана человеком, а не корпорацией.

    Иногда она может ошибиться.
    Иногда — не понять файл.
    Иногда — сделать что-то не так, как вы
    ожидали.

    Но она старается. Она улучшается
    с каждой версией. И она очень хочет
    помочь вам в работе.

    Если что-то пошло не так —
    пожалуйста, сообщите нам
    (контакты в окне "Помощь").
    Мы постараемся всё исправить
    в следующей версии.

    Спасибо, что пользуетесь AIST. 🌅

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            🐦  О ПРОГРАММЕ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Версия: 0.3.0.0
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
    thanks.after(3500, lambda: (thanks.destroy(), root.destroy()))

    # Если пользователь закроет окно крестиком — просто выйти
    thanks.protocol("WM_DELETE_WINDOW",
                    lambda: (thanks.destroy(), root.destroy()))


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================

def main():
    """Собирает и показывает главное окно."""
    global root, tree, status_label

    # Базовые настройки CustomTkinter
    ctk.set_appearance_mode("dark")  # тёмная тема
    ctk.set_default_color_theme("dark-blue")

    root = ctk.CTk()
    root.title("🐦 AIST — 0.3.0.0")
    root.geometry("700x520")
    root.configure(fg_color=COLOR_BG)

    # --- Заголовок ---
    ctk.CTkLabel(root, text="AIST",
                 font=("Arial", 26, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(25, 0))

    ctk.CTkLabel(root, text="Ассистент для работы с Excel",
                 font=("Arial", 12),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 20))

    # --- Верхний ряд кнопок ---#
    frame_top = ctk.CTkFrame(root, fg_color="transparent")
    frame_top.pack(pady=(5, 3))

    ctk.CTkButton(frame_top, text="📋  Функции",
                  font=("Arial", 13), width=180, height=42,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=show_functions).pack(side="left", padx=6)

    ctk.CTkButton(frame_top, text="❓  Помощь",
                  font=("Arial", 13), width=180, height=42,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_help).pack(side="left", padx=6)

    # --- Нижний ряд кнопок ---
    frame_bottom = ctk.CTkFrame(root, fg_color="transparent")
    frame_bottom.pack(pady=(3, 15))

    ctk.CTkButton(frame_bottom, text="📖  Инструкция",
                  font=("Arial", 12), width=130, height=38,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_instructions).pack(side="left", padx=5)

    # Кнопка "Закрепить" — меняется при нажатии, поэтому в переменной
    global btn_topmost
    btn_topmost = ctk.CTkButton(frame_bottom, text="📌  Закрепить",
                                font=("Arial", 12), width=130, height=38,
                                fg_color=COLOR_FRAME,
                                hover_color=COLOR_ACCENT,
                                text_color=COLOR_TEXT,
                                border_width=1,
                                border_color=COLOR_ACCENT,
                                command=toggle_topmost)
    btn_topmost.pack(side="left", padx=5)

    ctk.CTkButton(frame_bottom, text="🚪  Закрыть",
                  font=("Arial", 12), width=130, height=38,
                  fg_color=COLOR_FRAME,
                  hover_color="#7a2020",
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color="#a03030",
                  command=close_app).pack(side="left", padx=5)

    # --- Поле с информацией ---
    # --- Таблица с данными ---
    table_frame = ctk.CTkFrame(root, fg_color=COLOR_FIELD,
                               corner_radius=8)
    table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

    # Настраиваем стиль таблицы под нашу тёмную тему
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview",
                    background=COLOR_FIELD,
                    foreground=COLOR_TEXT,
                    fieldbackground=COLOR_FIELD,
                    bordercolor=COLOR_ACCENT,
                    borderwidth=1,
                    font=("Consolas", 10),
                    rowheight=24)
    style.configure("Treeview.Heading",
                    background=COLOR_FRAME,
                    foreground=COLOR_ACCENT,
                    font=("Arial", 11, "bold"))
    style.map("Treeview",
              background=[("selected", COLOR_ACCENT)],
              foreground=[("selected", "#1a1210")])

    # Контейнер для таблицы + прокрутки
    tree_container = tk.Frame(table_frame, bg=COLOR_FIELD)
    tree_container.pack(fill="both", expand=True, padx=10, pady=10)

    # Сама таблица
    tree = ttk.Treeview(tree_container, show="headings")
    tree.grid(row=0, column=0, sticky="nsew")

    # Вертикальная прокрутка (вниз-вверх)
    vsb = ttk.Scrollbar(tree_container, orient="vertical",
                        command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    tree.configure(yscrollcommand=vsb.set)

    # Горизонтальная прокрутка (влево-вправо)
    hsb = ttk.Scrollbar(tree_container, orient="horizontal",
                        command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")
    tree.configure(xscrollcommand=hsb.set)

    # Чтобы таблица растягивалась при увеличении окна
    tree_container.grid_rowconfigure(0, weight=1)
    tree_container.grid_columnconfigure(0, weight=1)

    # --- Строка статуса внизу ---
    status_label = ctk.CTkLabel(root, text="Файл не выбран",
                                font=("Consolas", 10),
                                text_color=COLOR_TEXT_DIM,
                                anchor="w")
    status_label.pack(fill="x", padx=25, pady=(0, 10))

    # --- Горячие клавиши ---
    root.bind("<Alt-v>", lambda e: open_file())
    root.bind("<F1>", lambda e: show_help())
    root.bind("<F2>", lambda e: show_instructions())
    root.bind("<Control-q>", lambda e: close_app())
    root.bind("<F3>", lambda e: show_functions())
    root.bind("<Control-t>", lambda e: toggle_topmost())  # Ctrl+T — закрепить

    root.mainloop()
    # ============================================================
    # СТАРТ ПРОГРАММЫ
    # ============================================================


if __name__ == "__main__":
    show_splash()