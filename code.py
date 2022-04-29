# Program to send bulk customized messages through Telegram Desktop application

from encodings import utf_8
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

#get message file
message_file_path = str(input("message file path: "))
# if message_file_path == ".":
#     message_file_path = os.getcwd
with open(message_file_path) as f:
    text_message = f.read()
    print(text_message)


#get the usernames of telegram to send message to that users

excel_sheets = []
excel_sheets += [each for each in os.listdir(os.getcwd()) if each.endswith('.xlsx')]
print(excel_sheets)

for excel_sheet in excel_sheets:

    excel_data = pandas.read_excel(str(excel_sheet))

    count = 0

    time.sleep(3)

    os.system('telegram-desktop')

    pyautogui.click(900, 900)

    for column in excel_data['usernames'].tolist():
        
        pyautogui.press('esc')
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        pyautogui.write(str(excel_data['usernames'][count]));
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('down')
        pyautogui.press('enter')
        pyautogui.write(str(excel_data['usernames'][count]));
        pyautogui.press('enter')
        pyautogui.press('esc')
        count = count + 1
results =[]
print('The script executed successfully.')
