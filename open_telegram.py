import os

# os.system('telegram-desktop')
# import pyautogui

# #pyautogui.moveTo(X, Y, Seconds)
# pyautogui.moveTo(900, 900, 2) #Move to X=100, Y=100 over a 2 seconds period
# pyautogui.click(900, 900)
# os.listdir(".")
results = []
results += [each for each in os.listdir(os.getcwd()) if each.endswith('.xlsx')]

print(results)
results =[]
print(results)

