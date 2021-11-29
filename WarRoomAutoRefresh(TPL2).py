# -*- coding: utf-8 -*-
"""
Created on Fri Aug 14 16:48:02 2020

@author: cckuo
"""

from tkinter import Tk, Text, Button, Label
from datetime import datetime
import win32gui
import win32con
import win32com.client
import time
# import pyautogui
from selenium import webdriver


def open_tableau_site(url):
    driver = webdriver.Edge(r'C:\Users\PL2_WarRoom\Downloads\msedgedriver.exe')
    driver.maximize_window()
    driver.implicitly_wait(8)

    driver.get("http://172.16.1.22/#/signin")
    driver.find_element_by_name('username').send_keys('pl_warroom')
    driver.find_element_by_name('password').send_keys('2021_Q2')
    driver.find_element_by_css_selector(
        '[class="tb-orange-button tb-button-login"]').click()
    time.sleep(5)
    driver.get(url)


def close_web(web_title):
    hwnd = win32gui.FindWindow(None, web_title)
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)


def update_web(web_title):
    print(web_title)
    hwnd = win32gui.FindWindow(None, web_title)
    shell = win32com.client.Dispatch("WScript.Shell")  # 加上這兩句
    shell.SendKeys('%')  # 就可以正常切换窗口
    win32gui.SetForegroundWindow(hwnd)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # 放大web
    win32gui.PostMessage(exhwnd, win32con.WM_KEYDOWN, win32con.VK_F5, 0)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_F5, 0)
    # 模擬滑鼠移動
    # x, y = pyautogui.position()
    # pyautogui.moveRel(100, 100, duration=0.25)
    # pyautogui.moveTo(x, y, duration=0.25)


def update_all_excel(xlpath, xlfilename, macro_name):

    xlApp = win32com.client.Dispatch("Excel.Application")  # 開啟設定
    try:
        # xlApp.Windows(xlfilename).Activate  # 選擇excel視窗
        xlApp.Workbooks(xlfilename).Activate  # 選擇excel
        # print("now excel")
    except:
        xlApp.Workbooks.Open(Filename=xlpath+xlfilename)  # 開啟excel
        xlApp.Visible = True
        # print("new excel")
        flash_list()

    shell = win32com.client.Dispatch("WScript.Shell")  # 加上這兩句
    shell.SendKeys('%')  # 就可以正常切换窗口了 3q
    hwnd = list(dic_hwnd_title.keys())[
        list(dic_hwnd_title.values()).index(xlfilename+' - Excel')]
    time.sleep(1)  # 等待1秒

    win32gui.SetForegroundWindow(hwnd)
    # maximize window
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    xlApp.Application.Run(xlfilename+"!"+macro_name)  # 執行巨集程式


def flash_list():

    global web_list, all_win_list

    # list all window
    win32gui.EnumWindows(get_all_hwnd, 0)
    # 列印出全部的web name#for h, t in dic_hwnd_title.items():
    all_win_list = list(set(list(dic_hwnd_title.values())))
    # list all website
    web_list = [i for i in all_win_list if i.find("Edge") > 0]


def get_all_hwnd(hwnd, mouse):
    # list all window
    if (win32gui.IsWindow(hwnd)
            and win32gui.IsWindowEnabled(hwnd)
            and win32gui.IsWindowVisible(hwnd)):
        dic_hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})


def countdown(t):

    while t >= 0 and running:
        time.sleep(1)
        t -= 1
        win.update()


def description():

    handbook = Tk()
    handbook.geometry('320x110')
    handbook.title('程式更新說明')
    text = Text(handbook, height=11)
    text.insert('insert', "\n----------------程式更新說明----------------\n\n")
    text.insert('insert', "點按「自動更新」按鈕，程式將每小時自動更新。\n\n")
    text.insert('insert', "更新的項目為Excel檔及所有開啟的網頁。\n")
    text.pack()
    text.config(state='disabled')
    handbook.mainloop()


def auto_refresh():

    global running, web_list
    btn_AutoRefresh.configure(state='disabled')
    running = True
    flash_list()
    while running:
        # update_all_excel('\\\\172.16.2.16\\pl\\電視牆資料\\','TV_KPI日報表.xlsm','UpDate_Workbook')
        # update_all_excel('\\\\172.16.2.16\\pl\\',
                        #  'Yield_Info_稼動率分析.xlsm', 'main_LOSS')

        
        flash_list()
        for i in web_list:
            close_web(i)
        for i in url_list:
            open_tableau_site(i)
        
        countdown(3600)


def stop():
    global running
    btn_AutoRefresh.configure(state='normal')
    running = False


if __name__ == '__main__':

    dic_hwnd_title = {}
    url_list = ['http://172.16.1.22/#/views/PL2Warroom/sheet0']
    running = False
    # xl_list = [['\\\\172.16.2.16\\pl\\電視牆資料\\','TV_KPI日報表.xlsm'],['\\\\172.16.2.16\\pl\\電視牆資料\\','Yield_Info_稼動率分析.xlsm]]

    win = Tk()
    win.geometry('200x50')
    # win.title('AutoRefresh')
    btn_AutoRefresh = Button(win, text='自動更新', command=auto_refresh)
    btn_AutoRefresh.grid(row=0, column=0, sticky='w', padx=5)
    btn_Stop = Button(win, text='停止更新', command=stop)
    btn_Stop.grid(row=0, column=1, padx=2)
    btn_Description = Button(win, text='說明', command=description)
    btn_Description.grid(row=0, column=2, sticky='e', padx=2)
    Label(win, text='© 2021 Sumika Designed by cckuo', font=("Arial", 9)).grid(
        row=1, column=0, stick='w', columnspan=3, pady=2)
    win.mainloop()
