import os
import tkinter as tk
import win32com.client
from tkinter import ttk, filedialog

data = []
exe_file = ""
table = None

def open_file():
    global exe_file, table
    # Определение допустимых типов файлов
    filetypes = [("EXE Files", "*.exe"), ("Shortcut Files", "*.lnk")]

    # Отображение окна диалога для выбора файла exe
    exe_file = filedialog.askopenfilename(title="Выберите файл exe", filetypes=filetypes)

    # Получение только имени файла
    exe_file = os.path.basename(exe_file)

    if exe_file:
        # Заменяем текст в кнопке "Обзор" на имя выбранного файла
        browse_button.config(text=exe_file)
        data = []

        filename = 'data.txt'

        # Проверяем существование файла
        if os.path.isfile(filename):
            # Чтение данных из файла
            with open('data.txt', 'r') as f:
                for line in f:
                    row = line.strip().split(',')
                    data.append(row)

            # Обновляем таблицу
            update_table()

def add_row():
    global table
    hotkey = ""
    if ctrl_var.get():
        hotkey += "ctrl+"
    if alt_var.get():
        hotkey += "alt+"
    if shift_var.get():
        hotkey += "shift+"
    hotkey += hotkey_entry.get()

    interval = interval_entry.get()

    # Проверяем, что все поля заполнены
    if exe_file and hotkey and interval:
        # Создаем новую строку данных
        new_row = [exe_file, hotkey, interval]

        # Добавляем новую строку в список данных
        data.append(new_row)

        # Обновляем таблицу
        update_table()

def remove_row():
    global table, data
    # Получаем выделенную строку
    selected_row = table.selection()
    if selected_row:
        # Удаляем выбранную строку из таблицы
        table.delete(selected_row)

        # Удаляем выбранную строку из массива данных
        index = int(selected_row[0][1:]) - 1
        if index >= 0 and index < len(data):
            del data[index]

        # Сохраняем изменения
        save()

def update_table():
    global table,data,dd,i
    # Очищаем таблицу
    i=0
    table.delete(*table.get_children())

    # Отображаем данные
    for values in data:
        # Вставляем данные в таблицу, начиная со столбца "Путь к exe"
        item = table.insert("", "end", text=values[0], values=(values[1], values[2]))
        # print(value[0], value[1], value[2])
        dd.append(values[0])
        i=i+1
    table.column('#0',width=180)
    table.column('#1',width=120)
    table.column('#2',width=100)
    # table.column('#3',width=10)
    if os.path.exists('0.txt'):
        autorun_var=1
 
def save():
    global data
    filename = 'data.txt'
    with open(filename, 'w', encoding='utf-8') as f:
        for row in data:
            f.write(','.join(str(value) for value in row) + '\n')

    autorun=autorun_var.get()
    if autorun==True:
        with open('0.txt', 'w', encoding='utf-8') as f:f.write('0')
    else:os.remove('0.txt')
    if os.path.exists('0.txt'):
        current_directory = os.getcwd()

        # Path to the file for which you want to create the shortcut
        source_file = os.path.join(current_directory, 'start.exe')

        # Shortcut name
        shortcut_name = 'start'

        # Path to the startup folder
        startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')

        # Create the shortcut object
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut_file = os.path.join(startup_folder, f'{shortcut_name}.lnk')
        shortcut = shell.CreateShortcut(shortcut_file)
        shortcut.TargetPath = source_file
        shortcut.WorkingDirectory = os.path.dirname(source_file)
        shortcut.WindowStyle = 7  # Hidden window style (7)
        shortcut.save()
    else:
        # Путь к папке с ярлыком
        shortcut_folder = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')

        # Имя ярлыка для удаления
        shortcut_name = 'start.lnk'

        # Полный путь к ярлыку
        shortcut_path = os.path.join(shortcut_folder, shortcut_name)

        # Проверяем, существует ли ярлык
        if os.path.exists(shortcut_path):
            # Удаляем ярлык
            os.remove(shortcut_path)
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'auto save.exe':
            # Завершение процесса
            proc.kill()
    
def check_data_file():
    global data
    filename = 'data.txt'
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as f:
            for line in f:
                row = line.strip().split(',')
                data.append(row)
        # Обновляем таблицу
        update_table()

def validate_input(*args):
    # Получаем введенное значение
    value = hotkey_entry.get()
    
    # Проверяем, что это английский символ
    if len(value) > 1 or not value.isalpha() or not value.isascii():
        # Оставляем только первый символ, если он английский
        value = value[0] if value and value.isalpha() and value.isascii() else hotkey_entry.set("")
        hotkey_entry.set(value)

def validate_input1(*args):
    value = interval_entry.get()
    if not value.isdigit():
        interval_entry.set("")
dd=[]
root = tk.Tk()
root.resizable(False,False)
root.title('AutoSave')
root.iconbitmap('i.ico')
hk=''
# Создаем фрейм для блока с прокруткой
block_frame = tk.Frame(root)
block_frame.grid(row=0, column=0, padx=5, pady=5)

# Создаем поле для прокрутки
scrollbar = tk.Scrollbar(block_frame)
scrollbar.pack(side="right", fill="y")

# Создаем таблицу
table = ttk.Treeview(block_frame, columns=("exe","Hotkeys"), yscrollcommand=scrollbar.set, padding=0)
table.heading("#0", text="Путь к exe")
table.heading("#1", text="Горячие клавиши")
table.heading("#2", text="Интервал (сек)")
table.pack(padx=0, pady=0)
table.configure(height=5)
# Привязываем прокрутку к таблице
scrollbar.config(command=table.yview)

frame = tk.LabelFrame(root)
frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
frame.config(borderwidth=2, relief="groove")

exe_frame = tk.LabelFrame(frame, text="Настройки exe", padx=5, pady=5)
exe_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

exe_frame.config(borderwidth=2, relief="groove")

browse_button = tk.Button(exe_frame, text="Обзор", command=open_file)
browse_button.grid(row=0, column=1, sticky="e", padx=5, pady=5)
exe_file_label = tk.Label(exe_frame, text="Путь к exe:")
exe_file_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)
hotkey_frame = tk.LabelFrame(frame, text="Горячие клавиши", padx=5, pady=5,width=round(root.winfo_width()/2))
hotkey_frame.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

hotkey_entry = tk.StringVar()
hotkey_entry.set('s')
hotkey_entry.trace("w", validate_input)

hotkey_entry_widget = tk.Entry(hotkey_frame, textvariable=hotkey_entry,width=5)
hotkey_entry_widget.grid(row=0, column=4, sticky="w", padx=5, pady=5)

ctrl_var = tk.IntVar()
ctrl_var.set(1)
ctrl_checkbox = tk.Checkbutton(hotkey_frame, text="ctrl+", variable=ctrl_var)
ctrl_checkbox.grid(row=0, column=1, sticky="w", pady=5)

alt_var = tk.IntVar()
alt_checkbox = tk.Checkbutton(hotkey_frame, text="alt+", variable=alt_var)
alt_checkbox.grid(row=0, column=2, sticky="w", pady=5)

shift_var = tk.IntVar()
shift_checkbox = tk.Checkbutton(hotkey_frame, text="shift+", variable=shift_var)
shift_checkbox.grid(row=0, column=3, sticky="w", pady=5)

time_frame = tk.LabelFrame(frame, text="Задержка", padx=5, pady=5)
time_frame.grid(row=1, column=1, padx=5, sticky="ew")

interval_label = tk.Label(time_frame, text="Интервал (сек):")
interval_label.grid(row=5, column=1, sticky="e", padx=5, pady=5)

autorun_frame = tk.LabelFrame(frame, text="Автозапуск", padx=5, pady=5)
autorun_frame.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
autorun_var = tk.IntVar()
autorun_checkbox = tk.Checkbutton(autorun_frame, text="Автозапуск с системой", variable=autorun_var)
autorun_checkbox.grid(row=4, column=3, sticky="w", pady=5)

interval_label = tk.Label(time_frame, text="Интервал (сек):")

# Создание переменной StringVar для поля ввода
interval_entry = tk.StringVar()
interval_entry.set('60')
interval_entry.trace("w", validate_input1)

# Создание виджета Entry с привязкой к переменной interval_entry
interval_entry_widget = tk.Entry(time_frame, textvariable=interval_entry,width=5)
interval_entry_widget.grid(row=5, column=2, sticky="w",padx=5 )

but=tk.LabelFrame(root)
but.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
remove_button = tk.Button(but, text="Удалить", command=remove_row)
remove_button.grid(row=5, column=0, sticky="ew", padx=5, pady=5)
add_button = tk.Button(but, text="Добавить", command=add_row)
add_button.grid(row=5, column=1, sticky="ew", padx=5, pady=5)

save_button = tk.Button(but, text="Сохранить", command=save)
save_button.grid(row=5, column=2, sticky="ew", padx=5, pady=5)
check_data_file()

root.mainloop()