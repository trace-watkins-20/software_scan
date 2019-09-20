import os
import sys
import subprocess
import pywinauto
import time

attempts = 0

cmd = subprocess.Popen('cmd.exe /K powershell -command "Start-Process cmd -Verb RunAs')
while attempts < 3:
    try:
        app = pywinauto.Application().connect(title_re='Admin.*')
        dialogs = app.window(title_re='Admin.*')
        time.sleep(4)    
    except:
        attempts+=1
if attempts <= 3:
    dialogs.type_keys('wmic{ENTER}')
    dialogs.type_keys("/output:U:\Python\Python37-32\office_scan\'computername'.txt /node:'computername' list brief")
