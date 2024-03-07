import pandas as pd
import numpy as np
import time
import pyautogui
import pyperclip
import platform


message = "ENTER YOUR MESSAGE HERE..."
excel_file_name = "Book1.xlsx"

# START
# Do Not Edit these values
current_platform = platform.system()
darwin = "Darwin"
ctrl = ""
start = ""
enter = ""
# END

sent_array = np.array("")


def assign_keys():
    global ctrl, start, enter, darwin

    if (current_platform == darwin):
        ctrl = "command"
        start = "space"
        enter = "enter"

    else:
        ctrl = "ctrl"
        start = "win"
        enter = "enter"


def open_whatsapp():

    time.sleep(1)

    pyperclip.copy('whatsapp')

    # # Open WA
    if current_platform == darwin:
        pyautogui.keyDown(ctrl)
        pyautogui.press(start)
        pyautogui.keyUp(ctrl)
    else:
        pyautogui.press(start)
    pyautogui.hotkey(ctrl, 'v')
    pyautogui.sleep(2)
    pyautogui.press(enter)

    time.sleep(3)


def iterate_number():
    df = pd.read_excel(excel_file_name, sheet_name='Sheet1')

    for index, row in df.iterrows():
        phone_number = str(row[1])
        if " ::: " in phone_number:
            # Found the delimiter then split the merged contacts and send message
            p = phone_number.split(" ::: ")

            for number in p:
                send_whatsapp_messages(number)

        else:
            send_whatsapp_messages(phone_number)


def send_whatsapp_messages(number):
    global sent_array
    # Check for duplicate value

    phone_number = str(number).strip().rstrip(".00").strip()

    if phone_number.startswith("0"):
        phone_number = phone_number[1:]
    if phone_number.startswith("+"):
        phone_number = phone_number[1:]
    if phone_number.startswith("+92"):
        phone_number = phone_number[3:]
    if phone_number.startswith("92"):
        phone_number = phone_number[2:]

    else:
        phone_number = phone_number

    if phone_number not in sent_array:

        pyperclip.copy(phone_number)

        time.sleep(1)

        pyautogui.click(x=340, y=60)
        time.sleep(1)
        pyautogui.click(x=515, y=170)
        pyautogui.hotkey(ctrl, 'v')

        time.sleep(1)

        pyautogui.click(x=450, y=250)

        time.sleep(1)

        pyperclip.copy(message)

        pyautogui.hotkey(ctrl, 'v')

        pyautogui.press(enter)

        # Update the global sent_array
        sent_array = np.append(sent_array, phone_number)

        time.sleep(1)


 
assign_keys()

open_whatsapp()
iterate_number()