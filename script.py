# Program to send bulk customized messages through Telegram Desktop application

import subprocess
import time
import os

try:
    import pandas
except (ImportError, NameError, RuntimeError, ModuleNotFoundError):
    subprocess.run(["pip", "install", "pandas"])


try:
    import openpyxl
except (ImportError, NameError, RuntimeError, ModuleNotFoundError):
    subprocess.run(["pip", "install", "openpyxl"])


try:
    import pyautogui
except (ImportError, NameError, RuntimeError, ModuleNotFoundError):
    subprocess.run(["pip", "install", "pyautogui"])



excel_data = pandas.read_excel('data.xlsx')

count = 0

time.sleep(3)

os.system('/opt/Telegram/./Telegram')


pyautogui.click(900, 900)
time.sleep(15)
for column in excel_data['usernames'].tolist():
      
      pyautogui.press('esc')
      pyautogui.hotkey('ctrl', 'f')
      time.sleep(1)
      pyautogui.write(str("rao Ali Nawaz"));
      pyautogui.press('enter')
      time.sleep(2)
      pyautogui.press('down')
      pyautogui.press('enter')
      pyautogui.write(str("ok"));
      pyautogui.press('enter')
      pyautogui.press('esc')
      count = count + 1

print('The script executed successfully.')