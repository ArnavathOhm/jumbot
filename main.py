from pyautogui import *
import pyautogui as pyg
import os
import sys
import keyboard
import numpy as np
import win32api
import win32con
from datetime import datetime

'''
==========================
= Non-Comercial Use Only =
==========================
created by Arnavath
bulid v.0.0.2

note: DON'T CHANGE ANY SINGLE CHARACTER IF UNNECESSARY

*editable if neccesary

© 2020 Arnavath
'''

bot_version = "0.0.2 by: Arnavath"

#X: 1333 Y:  112 RGB: (248, 248, 248)
#X: 1334 Y:  147 RGB: (248, 248, 248)
#>210, <190, <190


# define version
def version():
    print("=" * 50)
    lib_list = [[pyg, "pyautoGUI"]]
    print("bot verison : " + bot_version)
    for i in range(len(lib_list)):
        print(lib_list[i][1], "version : ", lib_list[i][0].__version__)
    print("=" * 50)


def zoom_isrun():
    # X:    0 Y:  727 RGB: ( 17,  17,  17)
    # X:    0 Y:  767 RGB: ( 17,  17,  17)
    # X: 1365 Y:  729 RGB: ( 38,  73,  93)
    # width = 1365 height = 40

    if pyg.locateOnScreen("zoom_icon.png", grayscale=False, confidence=0.9) != None:
        return None
    else:
        print("zoom is not running\nplease run the zoom application and run the script again..!")
        print("press esc to exit.......")
        keyboard.wait("esc")
        sys.exit()

# class
class dispcor:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.color_Backgroud = [255, 255, 255]  # white
        self.color_text = [0, 0, 0]  # black
        self.cam_is_on_color = [0, 0, 150]  # blue

    def click(self):
        win32api.SetCursorPos((self.x, self.y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(0.01)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    def veriv(self):
        for i in range(34):
            if pixel(1197 + i, 130)[0] < 150:
                return "Hadir"
        return "Tidak Hadir"


def database():
    global list_nim
    with open("list_nim_20.txt", "r") as file:
        list_nim = (file.read()).split(",")

    file_kelompok = open(f"kelompok\data_kelompok.txt", "r")

    data_kelompok = file_kelompok.readlines()
    for i in range(len(data_kelompok)):
        data_kelompok[i] = (data_kelompok[i][:-1]).split(",")

    file_kelompok.close()

    # devisi
    global kelompok_1
    kelompok_1 = []
    global kelompok_2
    kelompok_2 = []
    global kelompok_3
    kelompok_3 = []
    global kelompok_4
    kelompok_4 = []
    global kelompok_5
    kelompok_5 = []
    global kelompok_6
    kelompok_6 = []
    global kelompok_7
    kelompok_7 = []
    global kelompok_8
    kelompok_8 = []
    global kelompok_9
    kelompok_9 = []
    global kelompok_10
    kelompok_10 = []
    global kelompok_11
    kelompok_11 = []
    global kelompok_12
    kelompok_12 = []
    global list_kelompok
    list_kelompok = [kelompok_1,
                     kelompok_2,
                     kelompok_3,
                     kelompok_4,
                     kelompok_5,
                     kelompok_6,
                     kelompok_7,
                     kelompok_8,
                     kelompok_9,
                     kelompok_10,
                     kelompok_11,
                     kelompok_12]

    for i in range(len(list_kelompok)):
        list_kelompok[i].append(data_kelompok[0][i])
    for i in range(len(list_kelompok)):
        for k in range(len(data_kelompok[i + 1])):
            list_kelompok[i].append(data_kelompok[i + 1][k])

    with open("berhalangan.txt", "r") as file:
        global list_berhalangan
        list_berhalangan = []
        file = file.read()
        if len(file) >= 1:
            list_berhalangan = file.split(",")

    for i in list_berhalangan:
        if i in list_nim:
            list_nim.remove(i)


def filename():
    global report_path
    global date
    date = str(datetime.now()).split(" ")[0]
    report_name = date
    report_path = os.path.join('report', report_name)


def date():
    global time_now
    global date_now
    date_now, time_now = str(datetime.now()).split(" ")


def process():
    absen_kelompok_1 = ["kelompok 1\n"]
    absen_kelompok_2 = ["kelompok 2\n"]
    absen_kelompok_3 = ["kelompok 3\n"]
    absen_kelompok_4 = ["kelompok 4\n"]
    absen_kelompok_5 = ["kelompok 5\n"]
    absen_kelompok_6 = ["kelompok 6\n"]
    absen_kelompok_7 = ["kelompok 7\n"]
    absen_kelompok_8 = ["kelompok 8\n"]
    absen_kelompok_9 = ["kelompok 9\n"]
    absen_kelompok_10 = ["kelompok 10\n"]
    absen_kelompok_11 = ["kelompok 11\n"]
    absen_kelompok_12 = ["kelompok 12\n"]
    absen_kelompok = [absen_kelompok_1,
                     absen_kelompok_2,
                     absen_kelompok_3,
                     absen_kelompok_4,
                     absen_kelompok_5,
                     absen_kelompok_6,
                     absen_kelompok_7,
                     absen_kelompok_8,
                     absen_kelompok_9,
                     absen_kelompok_10,
                     absen_kelompok_11,
                     absen_kelompok_12]

    filename()
    try:
        file = open(f"{report_path}.txt", "x")
    except:
        file = open(f"{report_path}.txt", "w+")
    file.mode = "w"
    file.write('')
    file.mode = "a"
    file.write(f"tanggal : {date_now}\n")
    file.write(f"waktu : {time_now}\n\n")
    file.write("list kehadiran :\n")
    absent = 0
    list_absent = []
    dispcor(1236, 96).click()
    for i in list_nim:
        keyboard.write(str(i))
        time.sleep(0.25)

        # verivy id
        desc = dispcor(1323, 100).veriv()
        time.sleep(0.1)
        

        # write on file
        if desc == "Tidak Hadir":
            absent += 1
            list_absent.append(i)
        for l in range(3):
            keyboard.write("\b")
        file.write(str(i) + ":" + desc + "\n")
        for l in range(len(list_kelompok)):
            if i in list_kelompok[l]:
                absen_kelompok[l].append(str(i) + ":" + desc + "\n")
        dispcor(1236, 96).click()

    file.write("\n==Statistik==\n\nTidak Hadir : {}\nHadir : {}\nSudah Keluar : 6\nBerhalangan : {}\n".format(absent, (len(list_nim)- absent), len(list_berhalangan)))

    file.write("\n")
    
    for i in absen_kelompok:
        for k in range(len(i)):
            file.write(i[k])
        file.write("\n")
    file.write("\n")

    file.write("\n\nList yang Tidak Hadir atau Berhalangan : \n\n")

    if len(list_absent) > 0:
        file.write("\n\nTidak Hadir :\n\n")
        for i in list_absent:
            file.write(str(i) + "\n")

    if len(list_berhalangan) > 0:
        file.write("\n\nBerhalangan Hadir :\n\n")
        for i in list_berhalangan:
            file.write(str(i)+"\n")

    file.write("\nSudah Keluar :\n")

    with open("keluar.txt", "rt") as file_keluar:
        list_keluar = (file_keluar.read()).split(",")
    for i in list_keluar:
        file.write(i + "\n")
    file.close()

def main():
    date()
    version()
    zoom_isrun()
    database()
    filename()
    process()

if __name__ == "__main__":
    main()
