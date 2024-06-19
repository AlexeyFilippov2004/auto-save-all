import subprocess
import os
import time
script_dir = os.path.dirname(os.path.abspath(__file__))

subprocess.Popen([os.path.join(script_dir, 'auto save.exe')])
time.sleep(0.2)
subprocess.Popen([os.path.join(script_dir, 'tray.exe')])
