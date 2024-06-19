from pystray import Icon, MenuItem
from PIL import Image
import subprocess
import psutil

# Функции, вызываемые при выборе пунктов меню
def on_show(icon, item):
    subprocess.Popen(['settings.exe'])

def on_exit(icon, item):
    # Получение списка всех процессов с именем 'settings.exe'
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'auto save.exe':
            # Завершение процесса
            proc.kill()
    icon.stop()

# Создание иконки
image = Image.open("i.ico")  # Путь к иконке
menu = (
    MenuItem('Настройки', on_show),
    MenuItem('Выход', on_exit)
)

icon = Icon("App Name", image, "Auto Save", menu)

# Запуск иконки в трее
icon.run()
