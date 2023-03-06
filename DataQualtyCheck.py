# -*- coding: utf-8 -*-
"""
Created on Mar 2023
@author: SEA3523
"""
import tkinter as tk
import tkinter.ttk as ttk 

from tkinter import *
from tkinter import messagebox as msb

import getpass
import pandas as pd
import xlrd

import datetime

root = tk.Tk() #создание окна
root.geometry("800x600+400+100") #установка размеров окна
root.resizable(False, False) #запрет на ресайз
root.title('Data Quality Check Potential Plant')
root.iconbitmap("se.ico") #установка иконки 

label = ttk.Label(text="Data Check Potential",
                  font = 'Calibri 20',
                  ) # создаем текстовую метку
label.grid(row=0, column=0)   # размещаем метку в окне

style = ttk.Style(root)
root.tk.call('source', 'azure/azure.tcl')
style.theme_use('azure')

style.configure('.',font="Calibri 14", foreground="black", padding=8)
style.configure("Accentbutton", foreground='white', tk_setPalette="black")
style.configure("Togglebutton", foreground='white')

for c in range(3): root.columnconfigure(index=c, weight=300)
for r in range(10): root.rowconfigure(index=r, weight=60)

def storeStrategyOfWarehousesCheck():
   #получение имени пользователя компьютера getpass.getuser()
   path = "C:/Users/" + getpass.getuser() + "/Nextcloud2/Master Data Potential/Data Quality Check/DataStore/"
   #перечень файлов
   nameRefer = 'Номенклатура.xls'
   nameStratege = 'стратегия.xls'
   now = datetime.datetime.now()
   # df = xlrd.open_workbook(path + nameRefer)
   dfRefer = pd.read_excel(io = path + nameRefer)
   dfStratege = pd.read_excel(io = path + nameStratege)
   nameFile = 'check_storage_strategy_'+ getpass.getuser() +'_' + now.strftime("%d-%m-%Y %H-%M")+ '.xlsx'
   writer = pd.ExcelWriter(nameFile)
   # write dataframe to excel
   dfStratege.merge(dfRefer.rename({'Код ТМЦ': 'Код'},axis=1), left_on='Код ТМЦ', right_on='Код', how='left').to_excel(writer)
   # save the excel
   writer.save()
   msb.showinfo("Информационное сообщение", "Результаты проверки записаны в файл " + nameFile) 

# frame = Frame(
#    root,
#    padx=7,
#    pady=7
# )
# frame.grid(column=0, row=1)

buttonCheckStorageStrategy = ttk.Button(
    text="Проверить данные справочника SS",
    style="Accentbutton",
    command = storeStrategyOfWarehousesCheck
 )
buttonCheckStorageStrategy.grid(row=1, column=1, ipadx=6, ipady=6, padx=4, pady=4, sticky=NSEW)
#buttonCheckStorageStrategy.pack(anchor=NE, padx=[20, 60], pady=10, ipadx=10, ipady=10)

buttonCheckStorageStrategy = ttk.Button(
    text="Проверить справочник МТМ",
    style="Accentbutton",
    #command = создать функцию
 )
buttonCheckStorageStrategy.grid(row=2, column=1, ipadx=6, ipady=6, padx=4, pady=4, sticky=NSEW)
#buttonCheckStorageStrategy.pack(side=LEFT, fill=Y)

# var = tk.StringVar()

# togglebutton = ttk.Checkbutton(
# 	root,
# 	text='Toggle button',
# 	style='Togglebutton',
# 	variable=var,
# 	onvalue=1)
# togglebutton.pack()

# var2 = tk.StringVar()
# togglebutton2 = ttk.Checkbutton(
# 	root,
# 	text='Switch button',
# 	variable=var2,
# 	onvalue=1,
# 	style="Switch")
# togglebutton2.pack()
 
# cal_btn = Button(
#    frame,
#    text='Проверка заполненности кода ТН ВЭД',
# )
# cal_btn.grid(row=5, column=2)
 
# import pandas as pd
# my_file = open("C:/Users/sea3523/Desktop/BabyFile.xlsx", "w+",  errors='ignore')
# StrategyFilePath = "C:/Users/sea3523/Desktop/Strategy.xlsx"
# xlStrategy = pd.ExcelFile(StrategyFilePath)
# print(xlStrategy.sheet_names)

# my_file = open("C:/Users/sea3523/Desktop/BabyFile.xlsx", "w+",  errors='ignore')

# with open('C:/Users/sea3523/Desktop/Strategy.xls', errors='ignore') as f:
#     text = f.read()
#     for row in text:        
#         print(row)
#         # my_file.write(row)
# f.close()
# my_file.close()

root.mainloop()