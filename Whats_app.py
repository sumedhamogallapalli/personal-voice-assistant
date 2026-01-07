from pyautogui import click, write, sleep
import keyboard
from keyboard import press
from keyboard import write
from time import sleep

def Whatsapp_msg(name, msg):
    sleep(3)

    # write name
    click(x=271, y=137)
    sleep(1)
    write(name)

    # click on name
    sleep(2)
    click(x=248, y=233)
    sleep(2)

    # click on message option
    click(x=837, y=964)
    sleep(2)

    # to send message
    write(msg)
    sleep(1)
    keyboard.press("enter")
    sleep(3)

    # closing whatsapp
    click(x=1890, y=2)
    sleep(5)

# Example usage:
# Whatsapp_msg("amma", "hello")

def Whatsapp_call(name):
    # opening app
    #startfile("C:\\Users\\rocky\OneDrive\Desktop\WhatsApp.lnk")

    sleep(3)

    # write name
    click(x=271, y=137)
    sleep(1)
    write(name)

    # click on name
    sleep(2)
    click(x=248, y=233)
    sleep(2)

    # for call
    click(x=1781, y=82)

    # for ending call (x=1808, y=82)
    keyboard.wait("ctrl")
    click(x=1085, y=830)
    sleep(3)

    # close app
    click(x=1890, y=2)
    sleep(5)

# Example usage:
# Whatsapp_call("yashik")

def Whatsapp_video_call(name):
    # opening app
    #startfile("C:\\Users\\rocky\OneDrive\Desktop\WhatsApp.lnk")

    sleep(3)

    # write name
    click(x=271, y=137)
    sleep(1)
    write(name)

    # click on name
    sleep(2)
    click(x=248, y=233)
    sleep(2)

    # for video call
    click(x=1729, y=91)

    # waiting for Ctrl key
    keyboard.wait("ctrl")
    click(x=1103, y=936)
    sleep(3)

    # close app
    click(x=1890, y=2)
    sleep(5)

# Example usage:
# Whatsapp_video_call("yashik")
