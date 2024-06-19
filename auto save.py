import itertools
import psutil
import pygetwindow as gw
import ctypes
import win32gui
import win32api
import win32process
import win32con
import pyautogui 
import time
import os
import keyboard
from icecream import ic
ic.configureOutput(includeContext=True,prefix='(ic):')

ts=0.5
filename = 'data.txt'

def is_window_title(hwnd):
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    return bool(style & win32con.WS_CAPTION)

def get_window_title_by_pid(pid):
    def enum_windows_callback(hwnd, lparam):
        window_pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        if window_pid == pid and is_window_title(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            window_titles.append(window_title)
    
    window_titles = []
    win32gui.EnumWindows(enum_windows_callback, None)
    window_titles = list(filter(None, map(str.strip, window_titles)))
    time.sleep(ts) 
    window_titlesa = [title for title in gw.getAllTitles() if title.strip()]
    time.sleep(ts) 
    common_titles = set(window_titlesa) & set(window_titles)
    common_titles = [title for title in common_titles if title]
    common_titles = list(filter(bool, common_titles))

    return common_titles if common_titles else None

def get_pids_by_exe(config):
    pids = [proc.info['pid'] for proc in psutil.process_iter(['pid', 'name']) if proc.info['name'].lower() == config.lower()]
    time.sleep(ts) 
    return pids

def activ_win_for_save(exe_file):
    window_titles = []

    for exe in exe_file:
        pids = get_pids_by_exe(exe)
        time.sleep(ts) 
        window_titles.extend(get_window_title_by_pid(pid) for pid in pids if pid is not None)

    return window_titles

def check_window_exists(window_id):
    try:
        window = gw.Window(window_id)
        return True
    except gw.PyGetWindowException:
        return False


def titleto(title1,idw=None):
    # Если title1 пустое или не является словарем, создаем пустой словарь
    if not title1 or not isinstance(title1, dict):
        title1_dict = {}
    else:
        title1_dict = dict(title1)

    pid = idw  
 
    if not pid in title1_dict:
        title1_dict[pid] = 0
    time.sleep(ts) 
    
    ic(title1_dict,pid)

    # Создаем новый словарь title1 только с существующими окнами
    title1 = {}
    for key, value in title1_dict.items():
        if check_window_exists(key):
            title1[key] = value
        time.sleep(ts)  
    return title1

def title_pid_exe(window_title):
    def callback(hwnd, hwnds):
        length = ctypes.windll.user32.GetWindowTextLengthA(hwnd)
        buffer = ctypes.create_string_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextA(hwnd, buffer, length + 1)
        if buffer.value.decode('cp1251') == window_title and ctypes.windll.user32.IsWindowVisible(hwnd):
            hwnds[0] = hwnd
            return False
        return True
    
    hwnds = (ctypes.c_ulong * 1)()
    ctypes.windll.user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_ulong, ctypes.POINTER(ctypes.c_ulong))(callback), ctypes.byref(hwnds))
    return hwnds[0] if hwnds[0] else None

def pid_to_title(hwnd):
    return win32gui.GetWindowText(hwnd)

def key(keys):
    key_dict = {
        'ctrl': keyboard.is_pressed('ctrl'),
        'shift': keyboard.is_pressed('shift'),
        'alt': keyboard.is_pressed('alt'),
        'win': keyboard.is_pressed('win')
    }

    if not any(key_dict.values()):
        keys_list = [key.strip().lower() for key in keys.split('+')]
        modifiers = [key for key in keys_list if key in key_dict]
        non_modifiers = [key for key in keys_list if key not in key_dict]

        time.sleep(ts)

        for modifier in reversed(modifiers):
            keyboard.press(modifier)

        for key in non_modifiers:
            keyboard.press(key)
            keyboard.release(key)

        for modifier in modifiers:
            keyboard.release(modifier)
     

if os.path.exists(filename):
    b=0
    p=0
    active_window=gw.getActiveWindow()
    win=''
    target_value = 2
    title1=[]
    config = {'exe': [], 'time': [], 'key': [],'log':[]} 
    with open(filename, 'r', encoding='utf-8') as f:
        for line in f:
            row = line.strip().split(',')
            config['exe'].append(row[0])
            config['key'].append(row[1])
            config['time'].append(float(row[2]))
            config['log'].append(False)
    
    title=activ_win_for_save(config['exe'])
    active_window=gw.getActiveWindow()
    if title!=[]:
        while False in config['log']:
            if len(config['exe'])>0: 
                if title!=[]:
                    if p<len(config['log'])+1:
                        if config['log'][p]==False:
                            try:
                                win = gw.getWindowsWithTitle(title[p][b])[0]
                            except:
                                break
                            win0 = win.title
                            win1=gw.getActiveWindow().title
                            if win0 != win1:
                                win0=win1
                            if win0 == win1:
                                title1=titleto(title1)
                                while len(title[p])!=b:
                                    time.sleep(3)
                                    win = gw.getWindowsWithTitle(title[p][b])[0]
                                    win0 = win.title
                                    win = gw.getWindowsWithTitle(win0)[0]
                                    win.activate()
                                    if win!=None:
                                        w1=win0
                                        exe=title_pid_exe(w1)
                                        if exe:
                                            pid = ctypes.c_ulong()
                                            ctypes.windll.user32.GetWindowThreadProcessId(exe, ctypes.byref(pid))
                                            exe_path = psutil.Process(pid.value).exe()
                                            exe = os.path.basename(exe_path)
                                            index = config['exe'].index(exe)
                                            hk = config['key'][index]
                                            time.sleep(0.5) 
                                            key(hk)
                                            time.sleep(0.5) 
                                            config['log'][p]=True
                                            b = b + 1
                p=p+1
                b=0
    active_window.activate()
    while True: 
        time.sleep(5)
        try:
            title=activ_win_for_save(config['exe'])
            current_active_window=gw.getActiveWindowTitle()
            title = list(itertools.chain.from_iterable(title))
        except:
            time.sleep(1)
        if current_active_window in title:
            idw=title_pid_exe(current_active_window)
            title1=titleto(title1,idw)
            if idw in title1:
                title1[idw]=1
                exe=title_pid_exe(current_active_window)
                active_window=gw.getActiveWindow() 
                if exe:
                    active_window=gw.getActiveWindow() 
                    pid = ctypes.c_ulong()
                    ctypes.windll.user32.GetWindowThreadProcessId(exe, ctypes.byref(pid))
                    exe_path = psutil.Process(pid.value).exe()
                    exe = os.path.basename(exe_path)
                    if exe in config['exe']:
                        active_window=gw.getActiveWindow() 
                        index = config['exe'].index(exe)
                        time_value = config['time'][index]
                        time_value = float(time_value)
                        hk = str(config['key'][index])
                        time.sleep(time_value/3*2)
                        if win32api.GetKeyState(0x01) >= 0 and win32api.GetKeyState(0x02) >= 0:
                            key(hk)
                        title1[idw]=2
                        time.sleep(time_value/3)
        current_active_window = gw.getActiveWindow()
        target_keys=[]
        if title1!=[] and active_window!=current_active_window:
            target_keys = [key for key, value in title1.items() if value == target_value]
            if target_keys!=[]:
                for index1 in target_keys:
                    exe1 = pid_to_title(index1)
                    if exe1:
                        exe = title_pid_exe(exe1)
                        pid = ctypes.c_int()
                        tw = gw.getActiveWindow()
                        ctypes.windll.user32.GetWindowThreadProcessId(exe, ctypes.byref(pid))
                        exe1=gw.getWindowsWithTitle(exe1)[0]
                        exe_path = psutil.Process(pid.value).exe()
                        exe = os.path.basename(exe_path)
                        index = config['exe'].index(exe)  
                        hk = config['key'][index]
                        exe1.activate()
                        if win32api.GetKeyState(0x01) >= 0 and win32api.GetKeyState(0x02) >= 0:
                            time.sleep(0.2)
                            key(hk)
                            time.sleep(0.2)
                            title1[index1]=1
                current_active_window.activate()
else:
    time.sleep(5)