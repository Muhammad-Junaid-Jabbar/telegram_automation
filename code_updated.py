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

# get message file
message_file_path = str(input("message file path: "))

with open(message_file_path) as f:
    text_message = f.read()
    print(text_message)
a = input("telegram folders loctaion: ")
os.chdir(a)
print(os.getcwd())
p = os.listdir(os.getcwd())
directories = []
for i in p:
    if os.path.isdir(i):
        directories.append(os.path.abspath(i))
        # print(i)
    print(directories)
for directory in directories:
    os.chdir(directory)
    print(os.listdir(os.getcwd()))

    # get the usernames of telegram to send message to that users
    excel_sheets = []
    excel_sheets += [each for each in os.listdir(os.getcwd()) if each.endswith(".xlsx")]
    if excel_sheets == []:
        continue
    print(excel_sheets)

    for excel_sheet in excel_sheets:

        excel_data = pandas.read_excel(str(excel_sheet))

        count = 0

        time.sleep(3)
        # tele = os.path.join(
        #     "C:",
        #     os.sep,
        #     "Users",
        #     "rao",
        #     "AppData",
        #     "Roaming",
        #     "Telegram Desktop",
        #     "Telegram.exe",
        # )
        # os.system(tele)
        # os.system('telegram-desktop')
        subprocess.Popen(
            "C:\\Users\\rao\\AppData\\Roaming\\Telegram Desktop\\Telegram.exe"
        )
        time.sleep(3)
        pyautogui.click(900, 900)

        for column in excel_data["usernames"].tolist():

            pyautogui.press("esc")
            pyautogui.hotkey("ctrl", "f")
            time.sleep(1)
            pyautogui.write(str(excel_data["usernames"][count]))
            pyautogui.press("enter")
            time.sleep(2)
            pyautogui.press("down")
            pyautogui.press("enter")
            pyautogui.write(str(text_message))
            pyautogui.press("enter")
            pyautogui.press("esc")
            count = count + 1

    excel_sheets = []
pyautogui.hotkey("ctrl", "q")
print("The script executed successfully.")
