#/For importing webinar attendance data from cVent into GEAATTD. Java window must
#be in taskbark slot immediately below a screen-right aligned task bark for mouse
#to click into window. Banner should be full screen, with event already created
#and correct data entry block selected.
import pyautogui
import os
import pandas
import pywinauto
import numpy
import win32gui
import datetime

os.chdir('C:\\Users\\tenman\\Desktop\\R')

#PROMT USER FOR SHEET NUMBER
sheet_counter = (input("Sheet Name: "))

#MOVE mouse to position just below Task View on the Task bannerID
pyautogui.moveTo(1861,191, duration = 1 )
pyautogui.click()
pyautogui.moveTo(217,293, duration = 1 )
pyautogui.click()


data = pandas.read_excel('cVent Data - Webinars.xls', sheetname=sheet_counter)

date= data.get_value(1, 'Start Date')
date = str(date)
date = date[:10]
date = datetime.datetime.strptime(date, "%Y-%m-%d")
date = date.strftime("%d-%b-%Y")

max_rows = len(data)
row_counter = 0
while row_counter <= max_rows:
    bannerID = data.get_value(row_counter, 'Banner ID')
    bannerID = int(bannerID)
    bannerID = str(bannerID)
    pyautogui.typewrite(bannerID)
    pyautogui.press('space')
    pyautogui.press('tab')
    pyautogui.typewrite('ATTEND')
    pyautogui.press('tab')
    pyautogui.typewrite(date)
    pyautogui.typewrite('NOFEE')
    pyautogui.press('tab')
    pyautogui.typewrite(date)
    pyautogui.keyDown('ctrl')
    pyautogui.press('s')
    pyautogui.keyUp('ctrl')
    if row_counter != max_rows:
        pyautogui.press('f6') #MIGHT NEED TO MOVE THIS TO THE TOP OF THE LOOP
    row_counter = row_counter + 1
