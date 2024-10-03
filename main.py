import os
import shutil
import sqlite3
import threading
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox

# Константы
EXCEL_FILE = 'sheet1.xlsx'
DB_PATH = 'raffle.db'
BACKUP_DIR = 'backups'

month_translation = {
    "January": "январь", "February": "февраль", "March": "март",
    "April": "апрель", "May": "май", "June": "июнь",
    "July": "июль", "August": "август", "September": "сентябрь",
    "October": "октябрь", "November": "ноябрь", "December": "декабрь"
}

months_russian = list(month_translation.values())


def create_backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    db_backup = f"{BACKUP_DIR}/{timestamp}_raffle.db"
    excel_backup = f"{BACKUP_DIR}/{timestamp}_sheet1.xlsx"

    try:
        shutil.copyfile(DB_PATH, db_backup)
        shutil.copyfile(EXCEL_FILE, excel_backup)
    except FileNotFoundError as e:
        messagebox.showerror("Ошибка", f"Ошибка создания резервной копии: {e}")


def clean_old_backups():
    cutoff_date = datetime.now() - timedelta(days=30)
    for filename in os.listdir(BACKUP_DIR):
        file_date_str = filename.split('_')[0]
        try:
            file_date = datetime.strptime(file_date_str, "%Y%m%d")
            if file_date < cutoff_date:
                os.remove(os.path.join(BACKUP_DIR, filename))
        except (ValueError, OSError):
            continue


def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS raffle (
                date TEXT,
                month TEXT,
                prize_number INTEGER,
                winner_number INTEGER
            )
        ''')
        conn.commit()


def clear_month_data(selected_month):
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM raffle WHERE month = ?', (selected_month,))
        conn.commit()
    messagebox.showinfo("Очистка", f"Данные за {selected_month} успешно очищены.")


def load_data(selected_month):
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=selected_month, header=None, skiprows=1)
    except ValueError:
        messagebox.showerror("Ошибка", "Указанного месяца нет в файле розыгрыша")


def select_winner(data, selected_month):
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT winner_number FROM raffle WHERE month = ?', (selected_month,))
        previous_winners = set(row[0] for row in cursor.fetchall())

        cursor.execute('SELECT prize_number FROM raffle WHERE month = ?', (selected_month,))
        previous_prizes = set(row[0] for row in cursor.fetchall())

        available_contestants = data[data.iloc[:, 2].isna()]
        available_prizes = data[~data[3].isin(previous_prizes) & data[3].notna() & data[4].notna()]

        if available_prizes.empty:
            return None, "В этом месяце все призы уже разыграны между победителями."

        if available_contestants.empty:
            return None, "Все участники уже выиграли. Розыгрыш завершён."

        prize_row = available_prizes.iloc[0]
        prize_number = prize_row[3]
        prize_name = prize_row[4]
        prize_provider = prize_row[6]

        winner = None
        while winner is None:
            potentialwinner = available_contestants.sample(1).iloc[0]
            if pd.isna(potentialwinner[2]):
                winner = potentialwinner

        return (winner, prize_number, prize_name, prize_provider), None


def update_winner(data, winnerindex, prize_number, selected_month):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('INSERT INTO raffle (date, month, prize_number, winner_number) VALUES (?, ?, ?, ?)',
                     (datetime.now().strftime("%Y-%m-%d"), selected_month, prize_number, winnerindex))
        conn.commit()

    data.at[winnerindex, 2] = prize_number

    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            data.to_excel(writer, sheet_name=selected_month, index=False, header=False, startrow=1)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка обновления Excel: {e}")


def clear_winners(selected_month):
    if messagebox.askyesno("Подтверждение", "Вы действительно хотите очистить результаты розыгрыша?"):
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE  FROM raffle WHERE month=?', (selected_month,))
            conn.commit()
        data = load_data(selected_month)
        data[2] = ""

        try:
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name=selected_month, index=False, header=False, startrow=1)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обновления Excel: {e}")
        for item in winners_tree.get_children():
            winners_tree.delete(item)


def pick_winner():
    selected_month = month_var.get()
    data = load_data(selected_month)
    winnerinfo, message = select_winner(data, selected_month)
    if winnerinfo:
        winner, prize_number, prize_name, prize_provider = winnerinfo
        info = (winner[1], prize_name, prize_provider)
        update_winner(data, winner.name, prize_number, selected_month)
        display_winner(info)
    else:
        messagebox.showinfo("Информация", message)


def display_winner(info):
    winners_tree.insert("", 0, values=info)


def simulate_calculations():
    progress_bar['value'] = 0

    def updateprogress(i):
        if i <= 100:
            progress_bar['value'] = i
            root.after(100, updateprogress, i+1)
        else:
            pick_winner()

    updateprogress(0)


def start_raffle():
    threading.Thread(target=simulate_calculations).start()


def show_about():
    messagebox.showinfo("О программе",
                        "Рандомайзер для выбора победителей в ежемесячном розыгрыше призов среди донов группы "
                        "\nМАЯК - Максимально Адекватно и Ясно о Коже"
                        " \n\n ©seligor")


def center_window(root, width=600, height=410):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    x = (screenwidth // 2) - (width // 2)
    y = (screenheight // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')


def update_title():
    selected_month = month_var.get()
    root.title(f"МАЯК - Розыгрыш призов среди донов за {selected_month}")


def copy_to_clipboard():
    # Извлечение данных из Treeview
    data = []
    for child in winners_tree.get_children():
        row = winners_tree.item(child)['values']
        data.append('\t'.join(map(str, row)))

    # Форматирование данных в виде таблицы
    table = '\n'.join(data)

    # Копирование в буфер обмена
    root.clipboard_clear()
    root.clipboard_append(table)
    root.update()  # Это необходимо для сохранения данных в буфер обмена
    messagebox.showinfo("Успех", "Данные скопированы в буфер обмена.")

def clear_list():
    if messagebox.askyesno("Подтверждение", "Вы действительно хотите очистить список?"):
        for item in winners_tree.get_children():
            winners_tree.delete(item)

def show_context_menu(event):
    context_menu.post(event.x_root, event.y_root)

# Создание GUI
root = tk.Tk()
root.title("Розыгрыш призов")
center_window(root)

# Меню
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Файл", menu=file_menu)
file_menu.add_command(label=f"Очистить результаты розыгрыша", command=lambda: clear_winners(month_var.get()))
file_menu.add_separator()
file_menu.add_command(label="Выход", command=root.quit)

help_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Помощь", menu=help_menu)
help_menu.add_command(label="О программе", command=show_about)

# Победители
winners_frame = ttk.Frame(root)
winners_frame.pack(pady=0)
winners_tree = ttk.Treeview(winners_frame, columns=("name", "prize", "provider"), show="headings", height=15)
winners_tree.heading("name", text="Имя победителя")
winners_tree.heading("prize", text="Приз")
winners_tree.heading("provider", text="Спонсор подарка")
winners_tree.pack(side=tk.LEFT)

winners_scrollbar = ttk.Scrollbar(winners_frame, orient="vertical", command=winners_tree.yview)
winners_tree.configure(yscroll=winners_scrollbar.set)
winners_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Рамка для выбора месяца и прогрессбара
controls_frame = ttk.Frame(root)
controls_frame.pack(pady=10)

# Выпадающий список для выбора месяца
month_var = tk.StringVar()
month_var.set(month_translation[datetime.now().strftime("%B")])  # Текущий месяц по умолчанию

update_title()
month_menu = ttk.Combobox(controls_frame, textvariable=month_var, values=months_russian, state="readonly")
month_menu.bind("<<ComboboxSelected>>", lambda event: update_title())
month_menu.pack(side=tk.LEFT, padx=5)

# Прогрессбар
progress_bar = ttk.Progressbar(controls_frame, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(side=tk.LEFT, padx=5)

# Кнопка
button = tk.Button(root, text="Выбрать победителя", command=start_raffle)
button.pack(pady=10)

# Создание контекстного меню
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Скопировать в буфер обмена", command=copy_to_clipboard)
context_menu.add_command(label="Очистить список", command=clear_list)

# Привязка контекстного меню к щелчку правой кнопкой мыши
winners_tree.bind("<Button-3>", show_context_menu)

# Инициализация базы данных
init_db()
create_backup()
clean_old_backups()

root.mainloop()
