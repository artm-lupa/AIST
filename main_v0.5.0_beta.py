"""
AIST 0.5.0.1 — Ассистент для просмотра Excel-таблиц.
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
# --- Настройки программы ---
# Хранятся в памяти. Пока не сохраняются между запусками.
settings = {
    "remove_duplicates": True,  # удалять дубликаты при объединении
}
# --- Пути файлов для объединения ---
merge_files = {
    "file1": None,  # путь к первому файлу
    "file2": None,  # путь ко второму файлу
}

is_topmost = False


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
    # --- 4.1 Если колонки — безымянные (Unnamed: N или 0, 1, 2...) — берём первую строку как заголовки ---

    unnamed_count = sum(1 for c in df.columns if str(c).startswith("Unnamed:"))
    numeric_count = sum(1 for c in df.columns if str(c).isdigit())

    if unnamed_count == len(df.columns) or numeric_count == len(df.columns):
        # Первая строка станет заголовками
        new_headers = df.iloc[0].tolist()
        df = df.iloc[1:].reset_index(drop=True)
        df.columns = [str(h) for h in new_headers]
        report.append("Заголовки взяты из первой строки данных")

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
    win.geometry("420x540")
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

    # --- Кнопка "Объединить два файла" (пока пустая) ---
    ctk.CTkButton(win, text="🔗  Объединить два файла",
                  font=("Arial", 13), width=300, height=45,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_merge).pack(pady=8)

    # --- Кнопка "Сохранить как" ---
    ctk.CTkButton(win, text="💾  Сохранить как",
                  font=("Arial", 13), width=300, height=45,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=save_as).pack(pady=8)

    # --- Кнопка "Выбрать колонки" ---
    ctk.CTkButton(win, text="🔍  Выбрать колонки",
                  font=("Arial", 13), width=300, height=45,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_column_chooser).pack(pady=8)

    # --- Заглушка для будущих функций ---
    ctk.CTkLabel(win, text="(здесь будут другие функции)",
                 font=("Arial", 10, "italic"),
                 text_color=COLOR_TEXT_DIM).pack(pady=(15, 5))

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
    import os
    fleame = os.path.basename(path)

    if total <= limit:
        status_text = f"📂😊 {path}   |   {total} строк | готово!"
    else:
        status_text = (f"📂 {path}   |   {show_n} из {total} строк "
                       f"(большой фаил, показал в первые {limit})")

    status_label.configure(text=status_text)


def position_near_main(win, width, height):
    """
    Позиционирует окно сбоку от главного окна AIST.
    Если места справа нет — ставит слева.
    """
    win.update_idletasks()

    # Получаем позицию и размер главного окна
    root.update_idletasks()
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    root_w = root.winfo_width()

    # Размер экрана
    screen_w = win.winfo_screenwidth()

    # Пробуем поставить справа от главного
    x = root_x + root_w + 10  # 10 пикселей отступ
    y = root_y

    # Если справа не влезает — ставим слева
    if x + width > screen_w:
        x = root_x - width - 10
        # Если и слева не влезает — по центру экрана
        if x < 0:
            x = (screen_w - width) // 2

    win.geometry(f"{width}x{height}+{x}+{y}")
    win.lift()
    win.focus_force()


def show_settings():
    """Открывает окно настроек."""
    if hasattr(show_settings, "win") and show_settings.win.winfo_exists():
        show_settings.win.lift()
        show_settings.win.focus_force()
        return

    win = ctk.CTkToplevel(root)
    win.title("Настройки — AIST")
    win.resizable(False, False)
    win.configure(fg_color=COLOR_BG)
    position_near_main(win, 420, 440)

    # Заголовок
    ctk.CTkLabel(win, text="⚙️  Настройки",
                 font=("Arial", 20, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(25, 5))

    ctk.CTkLabel(win, text="Как AIST должен работать?",
                 font=("Arial", 11),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 20))

    # --- Функция-обёртка для галочки ---
    def on_toggle_duplicates():
        settings["remove_duplicates"] = bool(var_duplicates.get())

    # --- Галочка ---
    var_duplicates = ctk.BooleanVar(value=settings["remove_duplicates"])

    ctk.CTkCheckBox(win,
                    text="Удалять дубликаты при объединении",
                    variable=var_duplicates,
                    command=on_toggle_duplicates,
                    font=("Arial", 13),
                    text_color=COLOR_TEXT,
                    fg_color=COLOR_ACCENT,
                    hover_color=COLOR_ACCENT_H,
                    border_color=COLOR_ACCENT,
                    checkmark_color="#1a1210").pack(pady=10, padx=30,
                                                    anchor="w")

    ctk.CTkLabel(win,
                 text="Когда включено: если в файлах есть\n"
                      "одинаковые строки — останется только одна.",
                 font=("Arial", 10, "italic"),
                 text_color=COLOR_TEXT_DIM,
                 justify="left").pack(pady=(0, 20), padx=50, anchor="w")

    # --- Кнопка Закрыть ---
    ctk.CTkButton(win, text="Готово", width=140,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=win.destroy).pack(pady=15)

    win.bind("<Escape>", lambda e: win.destroy())
    show_settings.win = win


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


def create_table_with_columns(selected_columns):
    """
    Оставляет в таблице только указанные колонки.
    selected_columns — список имён колонок.
    """
    # --- 1. Собираем данные из таблицы ---
    all_columns = list(tree["columns"])
    all_rows = []
    for item in tree.get_children():
        values = tree.item(item)["values"]
        all_rows.append(values)

    # --- 2. Находим индексы нужных колонок ---
    indices = [all_columns.index(col) for col in selected_columns
               if col in all_columns]

    # --- 3. Создаём новые строки с выбранными колонками ---
    new_rows = []
    for row in all_rows:
        new_row = [row[i] for i in indices]
        new_rows.append(new_row)

    # --- 4. Очищаем таблицу ---
    for item in tree.get_children():
        tree.delete(item)

    # --- 5. Устанавливаем новые колонки ---
    tree["columns"] = selected_columns
    tree["show"] = "headings"

    for col in selected_columns:
        tree.heading(col, text=col)
        tree.column(col, width=140, anchor="w", stretch=False)

    # --- 6. Заполняем данными ---
    for row in new_rows:
        tree.insert("", "end", values=row)

    # --- 7. Обновляем статус ---
    status_label.configure(
        text=f"🔍 Оставил: {len(selected_columns)} колонок   |   "
             f"{len(new_rows)} строк"
    )

def show_column_chooser():
    """Открывает окно выбора колонок для создания новой таблицы."""

    # --- 1. Проверяем, что таблица есть ---
    if not tree.get_children():
        messagebox.showwarning(
            "Нет данных",
            "Сначала откройте файл или объедините таблицы."
        )
        return

    # --- 2. Если окно уже открыто — не плодим ---
    if hasattr(show_column_chooser, "win") and show_column_chooser.win.winfo_exists():
        show_column_chooser.win.lift()
        show_column_chooser.win.focus_force()
        return

    # --- 3. Получаем список колонок из таблицы ---
    all_columns = list(tree["columns"])

    # --- 4. Создаём окно ---
    win = ctk.CTkToplevel(root)
    win.title("Выбор колонок — AIST")
    win.resizable(False, False)
    win.configure(fg_color=COLOR_BG)

    # Высота зависит от количества колонок
    position_near_main(win, 520, 560)

    # --- 5. Заголовок ---
    ctk.CTkLabel(win, text="🔍  Выбор колонок",
                 font=("Arial", 18, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(20, 5))

    ctk.CTkLabel(win, text="Отметьте колонки, которые оставить",
                 font=("Arial", 10),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 15))

    # --- 6. Список галочек ---
    vars_dict = {}

    scroll_frame = ctk.CTkScrollableFrame(win, fg_color=COLOR_FRAME,
                                          width=420, height=250)
    scroll_frame.pack(padx=25, pady=5, fill="both", expand=True)

    for col in all_columns:
        var = ctk.BooleanVar(value=True)  # по умолчанию — все выбраны
        vars_dict[col] = var
        ctk.CTkCheckBox(scroll_frame, text=str(col),
                        variable=var,
                        font=("Arial", 12),
                        text_color=COLOR_TEXT,
                        fg_color=COLOR_ACCENT,
                        hover_color=COLOR_ACCENT_H,
                        border_color=COLOR_ACCENT,
                        checkmark_color="#1a1210").pack(anchor="w",
                                                        pady=4, padx=10)

    # --- 7. Кнопки "Все / Ничего" ---
    btn_frame = ctk.CTkFrame(win, fg_color="transparent")
    btn_frame.pack(pady=(5, 10))

    def select_all():
        for var in vars_dict.values():
            var.set(True)

    def deselect_all():
        for var in vars_dict.values():
            var.set(False)

    ctk.CTkButton(btn_frame, text="Все", width=80,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=select_all).pack(side="left", padx=5)

    ctk.CTkButton(btn_frame, text="Ничего", width=80,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=deselect_all).pack(side="left", padx=5)

    # --- 8. Кнопки "Создать / Отмена" ---
    action_frame = ctk.CTkFrame(win, fg_color="transparent")
    action_frame.pack(pady=(0, 20))

    def apply_columns():
        # Собираем только выбранные колонки
        selected = [col for col, var in vars_dict.items() if var.get()]

        if not selected:
            messagebox.showwarning("Ничего не выбрано",
                                   "Выберите хотя бы одну колонку.")
            return

        # Обновляем таблицу
        create_table_with_columns(selected)
        win.destroy()

    ctk.CTkButton(action_frame, text="🔍  Создать таблицу",
                  font=("Arial", 12), width=200, height=40,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=apply_columns).pack(side="left", padx=6)

    ctk.CTkButton(action_frame, text="Отмена", width=120, height=40,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=win.destroy).pack(side="left", padx=6)

    win.bind("<Escape>", lambda e: win.destroy())
    show_column_chooser.win = win

def save_as():
    """Сохраняет текущую таблицу в новый Excel-файл."""

    # --- 1. Проверяем, есть ли что сохранять ---
    if not tree.get_children():
        messagebox.showwarning(
            "Нечего сохранять",
            "Сначала откройте файл или объедините таблицы."
        )
        return

    # --- 2. Спрашиваем, куда сохранить ---
    path = filedialog.asksaveasfilename(
        title="Сохранить как",
        defaultextension=".xlsx",
        filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
        initialfile="AIST_результат.xlsx"
    )

    if not path:
        return  # пользователь нажал «Отмена»

    # --- 3. Собираем данные из таблицы обратно в DataFrame ---
    try:
        # Получаем колонки
        columns = tree["columns"]

        # Собираем все строки
        rows = []
        for item in tree.get_children():
            values = tree.item(item)["values"]
            rows.append(values)

        # Создаём DataFrame
        df = pd.DataFrame(rows, columns=columns)
    except Exception as e:
        messagebox.showerror(
            "Ошибка",
            f"Не удалось собрать данные из таблицы:\n{e}"
        )
        return

    # --- 4. Сохраняем в Excel ---
    try:
        df.to_excel(path, index=False)
    except PermissionError:
        messagebox.showerror(
            "Файл занят",
            "Не удалось сохранить файл.\n\n"
            "Возможно, он открыт в Excel.\n"
            "Закройте Excel и попробуйте снова."
        )
        return
    except Exception as e:
        messagebox.showerror(
            "Ошибка сохранения",
            f"Не удалось сохранить файл:\n{e}"
        )
        return

    # --- 5. Успех — показываем уведомление ---
    import os
    filename = os.path.basename(path)

    messagebox.showinfo(
        "Готово!",
        f"Файл сохранён:\n\n"
        f"  {filename}\n\n"
        f"Строк: {len(df)}\n"
        f"Колонок: {len(df.columns)}"
    )

    # --- 6. Обновляем статус внизу ---
    status_label.configure(
        text=f"💾 Сохранено: {filename}   |   {len(df)} строк"
    )

def do_merge(win):
    """
    Выполняет объединение двух файлов.
    win — окно выбора файлов (его надо закрыть).
    """
    # --- 1. Проверяем, что оба файла выбраны ---
    if not merge_files["file1"] or not merge_files["file2"]:
        messagebox.showwarning(
            "Файлы не выбраны",
            "Пожалуйста, выберите оба файла."
        )
        return

    path1 = merge_files["file1"]
    path2 = merge_files["file2"]

    # --- 1.5. Проверяем, не один ли это и тот же файл ---

    import os
    if os.path.abspath(path1) == os.path.abspath(path2):
        answer = messagebox.askyesno(
            "Один и тот же файл",
            "Вы выбрали один и тот же файл дважды.\n\n"
            "Продолжить? (получится таблица с дубликатами всех строк)"
        )
        if not answer:
            return

    # --- 1.5. Проверяем, не один ли это и тот же файл ---
    import os
    if os.path.abspath(path1) == os.path.abspath(path2):
        answer = messagebox.askyesno(
            "Один и тот же файл",
            "Вы выбрали один и тот же файл дважды.\n\n"
            "Продолжить? (получится таблица с дубликатами всех строк)"
        )
        if not answer:
            return

    # --- 2. Читаем оба файла ---
    try:
        df1 = pd.read_excel(path1, header=None)
        df2 = pd.read_excel(path2, header=None)
    except Exception as e:
        messagebox.showerror(
            "😅 Ошибка чтения",
            f"Не удалось прочитать файл.\n\n"
            f"Возможные причины:\n"
            f"   • Файл открыт в Excel — закройте его\n"
            f"   • Файл повреждён\n"
            f"   • Это не Excel-файл\n\n"
            f"Техническое описание:\n{e}"
        )
        return

    # --- 3. Чистим каждый через существующую функцию ---
    df1, report1 = clean_dataframe(df1)
    df2, report2 = clean_dataframe(df2)

    # --- 4. Проверяем, что оба непустые ---
    if df1.empty or df2.empty:
        messagebox.showerror(
            "Пустой файл",
            "Один из файлов пуст или содержит "
            "только пустые строки."
        )
        return

    # --- 5. Проверяем, что колонки совпадают ---
    cols1 = list(df1.columns)
    cols2 = list(df2.columns)

    if cols1 != cols2:
        # Колонки разные — покажем предупреждение
        answer = messagebox.askyesno(
            "Колонки не совпадают",
            f"Колонки в файлах разные.\n\n"
            f"Файл 1: {cols1}\n\n"
            f"Файл 2: {cols2}\n\n"
            f"Объединить всё равно? "
            f"(недостающие колонки будут пустыми)"
        )
        if not answer:
            return

    # --- 6. Объединяем ---
    rows1 = len(df1)
    rows2 = len(df2)

    combined = pd.concat([df1, df2], ignore_index=True)

    # --- 6.5. Если строк много — предупредим ---
    total_before = len(df1) + len(df2)
    if total_before > 5000:
        answer = messagebox.askyesno(
            "Большой файл",
            f"В двух файлах {total_before} строк.\n\n"
            "Обработка может занять несколько секунд. Продолжить?"
        )
        if not answer:
            return

    # --- 7. Удаляем дубликаты (если включено в настройках) ---
    removed_dups = 0
    if settings["remove_duplicates"]:
        before = len(combined)
        combined = combined.drop_duplicates()
        removed_dups = before - len(combined)

    # --- 8. Закрываем окно выбора файлов ---
    win.destroy()

    # --- 9. Очищаем главную таблицу ---
    for row in tree.get_children():
        tree.delete(row)

    # --- 10. Настраиваем колонки ---
    columns = [str(c) for c in combined.columns]
    tree["columns"] = columns
    tree["show"] = "headings"

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=140, anchor="w", stretch=False)

    # --- 11. Заполняем данными (первые 500 строк) ---
    limit = 500
    total = len(combined)
    show_n = min(limit, total)

    for i in range(show_n):
        row_values = []
        for v in combined.iloc[i].tolist():
            if pd.isna(v):
                row_values.append("")
            else:
                row_values.append(str(v))
        tree.insert("", "end", values=row_values)

    # --- 12. Обновляем статус ---
    parts = [
        f"😁 Объединено: {total} строк",
        f"({rows1} + {rows2})"
    ]
    if removed_dups > 0:
        parts.append(f"🧹 убрал дубликатов: {removed_dups}")
    parts.append(f"показал {show_n}")

    status_label.configure(text="   |   ".join(parts))

    # --- 13. Показываем финальное окно ---
    messagebox.showinfo(
        "🥳Готово!",
        f"Файлы успешно объединены.\n\n"
        f"Было:\n"
        f"📂  Файл 1: {rows1} строк\n"
        f"📂  Файл 2: {rows2} строк\n\n"
        f" Получилось: {total} строк\n"
        + (f"🧹 Удалено дубликатов: {removed_dups}\n"
           if removed_dups > 0 else "")
        + f"\nРезультат показан в главном окне."
    )


def show_merge():
    """Открывает окно выбора двух файлов для объединения."""
    if hasattr(show_merge, "win") and show_merge.win.winfo_exists():
        show_merge.win.lift()
        show_merge.win.focus_force()
        return

    win = ctk.CTkToplevel(root)
    win.title("Объединить два файла — AIST")
    win.resizable(False, False)
    win.configure(fg_color=COLOR_BG)
    position_near_main(win, 560, 420)

    # --- Заголовок ---
    ctk.CTkLabel(win, text="🔗  Объединить два файла",
                 font=("Arial", 18, "bold"),
                 text_color=COLOR_ACCENT).pack(pady=(20, 5))

    ctk.CTkLabel(win, text="Выберите два Excel-файла для объединения",
                 font=("Arial", 10),
                 text_color=COLOR_TEXT_DIM).pack(pady=(0, 20))

    # --- Переменные для отображения имён файлов ---
    var_file1 = ctk.StringVar(value="Не выбран")
    var_file2 = ctk.StringVar(value="Не выбран")

    # --- Функция выбора файла 1 ---
    def choose_file1():
        path = filedialog.askopenfilename(
            title="Выберите первый файл",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if path:
            merge_files["file1"] = path
            # Показываем только имя файла (без длинного пути)
            import os
            var_file1.set(os.path.basename(path))
            print(f">>> Выбран файл 1: {path}")

    # --- Функция выбора файла 2 ---
    def choose_file2():
        path = filedialog.askopenfilename(
            title="Выберите второй файл",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if path:
            merge_files["file2"] = path
            import os
            var_file2.set(os.path.basename(path))
            print(f">>> Выбран файл 2: {path}")

    # --- Блок "Файл 1" ---
    frame1 = ctk.CTkFrame(win, fg_color="transparent")
    frame1.pack(fill="x", padx=30, pady=8)

    ctk.CTkLabel(frame1, text="Файл 1:",
                 font=("Arial", 12),
                 text_color=COLOR_TEXT,
                 width=60, anchor="w").pack(side="left")

    entry1 = ctk.CTkEntry(frame1, textvariable=var_file1,
                          font=("Arial", 11),
                          state="readonly",
                          fg_color=COLOR_FIELD,
                          text_color=COLOR_TEXT,
                          border_color=COLOR_ACCENT)
    entry1.pack(side="left", fill="x", expand=True, padx=8)

    ctk.CTkButton(frame1, text="Выбрать", width=100,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=choose_file1).pack(side="left")

    # --- Блок "Файл 2" ---
    frame2 = ctk.CTkFrame(win, fg_color="transparent")
    frame2.pack(fill="x", padx=30, pady=8)

    ctk.CTkLabel(frame2, text="Файл 2:",
                 font=("Arial", 12),
                 text_color=COLOR_TEXT,
                 width=60, anchor="w").pack(side="left")

    entry2 = ctk.CTkEntry(frame2, textvariable=var_file2,
                          font=("Arial", 11),
                          state="readonly",
                          fg_color=COLOR_FIELD,
                          text_color=COLOR_TEXT,
                          border_color=COLOR_ACCENT)
    entry2.pack(side="left", fill="x", expand=True, padx=8)

    ctk.CTkButton(frame2, text="Выбрать", width=100,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=choose_file2).pack(side="left")

    # --- Кнопки внизу ---
    frame_buttons = ctk.CTkFrame(win, fg_color="transparent")
    frame_buttons.pack(pady=25)
    ctk.CTkButton(frame_buttons, text="🔗  Объединить",
                  font=("Arial", 13), width=180, height=42,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=lambda: do_merge(win)).pack(side="left", padx=8)
    # ↑ Пока пусто. В Шаге 5 дадим логику.

    ctk.CTkButton(frame_buttons, text="Отмена",
                  font=("Arial", 13), width=140, height=42,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=win.destroy).pack(side="left", padx=8)

    win.bind("<Escape>", lambda e: win.destroy())
    show_merge.win = win


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
        🔗  КАК ОБЪЕДИНИТЬ ДВА ФАЙЛА
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Функция "Объединить два файла" позволяет
    сложить две таблицы в одну.

    Когда это нужно:
      • Накладные от двух поставщиков
      • Табели за два месяца
      • Данные из двух садов / бригад

    Как использовать:
       1. Нажмите "📋 Функции".
       2. Нажмите "🔗 Объединить два файла".
      3. Выберите первый файл.
      4. Выберите второй файл.
      5. Нажмите "🔗 Объединить".

    AIST сам:
      • Прочитает оба файла
      • Уберёт пустые строки и колонки
      • Найдёт заголовки
       • Объединит данные
       • Удалит дубликаты (если включено
         в настройках)
       • Покажет результат

    ВАЖНО: AIST НИКОГДА не меняет ваши
    исходные файлы. Только читает их.
    Результат показывается в окне.

    Чтобы сохранить результат — используйте
    функцию "💾 Сохранить как" (появится
    в версии 0.5.0).

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        ⚙️  КАК НАСТРОИТЬ AIST
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    Кнопка "⚙️ Настройки" открывает окно
    с настройками программы.

    Сейчас доступна одна настройка:

      ☑ Удалять дубликаты при объединении

     Если включено: если в двух файлах
     есть одинаковые строки — останется
      только одна.

     По умолчанию: ВКЛЮЧЕНО.

    Скоро появятся другие настройки.

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
       🛡 БЕЗОПАСНОСТЬ ВАШИХ ДАННЫХ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    AIST работает по принципу "не навреди":

      • Никогда не меняет исходные файлы
      • Все результаты — в новых файлах
          • Всё, что делает AIST, можно
        проверить вручную

    Это означает, что вы можете спокойно
    экспериментировать: AIST не испортит
    ваши данные, даже если вы нажмёте
    не ту кнопку.

    Тем не менее, всегда полезно иметь
    резервные копии важных файлов.

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        💾  СОХРАНИТЬ РЕЗУЛЬТАТ
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    После того как AIST прочитал или
    объединил таблицы — результат можно
    сохранить в файл Excel.

    Как сохранить:
      1. Нажмите «📋 Функции».
      2. Нажмите «💾 Сохранить как».
      3. Выберите папку и введите имя файла.
      4. Нажмите «Сохранить».

    Быстрый способ: Ctrl + S.

    Результат сохраняется в формате .xlsx —
    открывается в Excel, как обычный файл.
    Его можно отправить коллеге, распечатать,
    загрузить в 1С.

    ВАЖНО: AIST не меняет ваши исходные
    файлы. Сохранение создаёт НОВЫЙ файл.
    Старые данные остаются на месте.


    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        🔍  ВЫБОР КОЛОНОК
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    
    Если в таблице много колонок, а нужны
    только некоторые — используйте
    функцию «Выбрать колонки».
    
    Когда это нужно:
       • Выгрузка из 1С: 40 колонок,
         а для отчёта нужно только 5
      • Убрать лишние данные перед
        сохранением
       • Оставить только нужные столбцы
    
    Как использовать:
      1. Откройте файл или объедините таблицы.
      2. Нажмите «📋 Функции».
       3. Нажмите «🔍 Выбрать колонки».
      4. Снимите галочки с ненужных колонок.
     5. Нажмите «🔍 Создать таблицу».
    
    Подсказки:
      • Все колонки по умолчанию ОТМЕЧЕНЫ —
        уберите галочки с тех, что не нужны.
      • Кнопка «Все» — отмечает всё.
     • Кнопка «Ничего» — снимает всё.
     • Список можно прокручивать, если
      колонок много.

    Быстрый способ: Ctrl + K.
    
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
       F3         функции (меню)
       Ctrl+T     Закрепить/открепить роверх окон
       F4         Настройки

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

    Версия: 0.5.0.0
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
    root.title("🐦 AIST — 0.5.0.0")
    root.geometry("800x520")
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
                  font=("Arial", 13), width=150, height=42,
                  fg_color=COLOR_ACCENT,
                  hover_color=COLOR_ACCENT_H,
                  text_color="#1a1210",
                  command=show_functions).pack(side="left", padx=6)

    ctk.CTkButton(frame_top, text="❓  Помощь",
                  font=("Arial", 13), width=150, height=42,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_help).pack(side="left", padx=6)

    # --- Нижний ряд кнопок ---
    frame_bottom = ctk.CTkFrame(root, fg_color="transparent")
    frame_bottom.pack(pady=(3, 15))
    # --- Кнопка Настройки (пока пустая) ---
    ctk.CTkButton(frame_bottom, text="⚙️  Настройки",
                  font=("Arial", 12), width=110, height=38,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_settings).pack(side="left", padx=5)

    ctk.CTkButton(frame_bottom, text="📖  Инструкция",
                  font=("Arial", 12), width=110, height=38,
                  fg_color=COLOR_FRAME,
                  hover_color=COLOR_ACCENT,
                  text_color=COLOR_TEXT,
                  border_width=1,
                  border_color=COLOR_ACCENT,
                  command=show_instructions).pack(side="left", padx=5)

    # Кнопка "Закрепить" — меняется при нажатии, поэтому в переменной
    global btn_topmost
    btn_topmost = ctk.CTkButton(frame_bottom, text="📌  Закрепить",
                                font=("Arial", 12), width=110, height=38,
                                fg_color=COLOR_FRAME,
                                hover_color=COLOR_ACCENT,
                                text_color=COLOR_TEXT,
                                border_width=1,
                                border_color=COLOR_ACCENT,
                                command=toggle_topmost)
    btn_topmost.pack(side="left", padx=5)

    ctk.CTkButton(frame_bottom, text="🚪  Закрыть",
                  font=("Arial", 12), width=110, height=38,
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
    root.bind("<F4>", lambda e: show_settings())
    root.bind("<Control-t>", lambda e: toggle_topmost())  # Ctrl+T — закрепить
    root.bind("<Control-s>", lambda e: save_as())

    root.mainloop()
    # ============================================================
    # СТАРТ ПРОГРАММЫ
    # ============================================================


if __name__ == "__main__":
    show_splash()