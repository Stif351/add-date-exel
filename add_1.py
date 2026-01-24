import tkinter as tk
from tkinter import *
from tkinter import font
from tkinter import ttk

import openpyxl
from openpyxl import load_workbook


def add_date():
    pass
    wb = load_workbook("K_8.xlsx")
    ws = wb.active
    data_gas_num = data_gas.get()
    ws['Z57'] = "Нові дані"
    ws['Z58'] = data_gas_num
    wb.save("K_8.xlsx")

    print(data_gas_num)

# /////////////////////////////////  ///////////////////////////////////////////////////////////////////////////
root = tk.Tk()
root.title("MAIN WINDOW")
root.geometry("900x800")
root.resizable(False, False)

courier_10 = font.Font(family="Courier", size=10, weight=font.BOLD)
courier_14 = font.Font(family="Courier", size=14, weight=font.BOLD)
courier_18 = font.Font(family="Courier", size=18, weight=font.BOLD)
width_frame = 800

label = tk.Label(root, text="ДОДАВАННЯ ПОКАЗНИКІВ ЛІЧИЛЬНІКІВ У ФАЙЛ", fg="BLACK", font=courier_18)
label.grid(row=0, column=0, columnspan=3, ipadx=6, ipady=6, padx=5, pady=15)

# ========================  main frame  =======================================

lf_MF = ttk.Frame(root, borderwidth=10, relief=SUNKEN)
lf_MF.config(width=850, height=600)
lf_MF.grid_propagate(False)

label = tk.Label(lf_MF, text="ДОДАВАННЯ ПОКАЗНИКІВ КВІТНЕВА-8", fg="BLACK", font=courier_18)
label.grid(row=0, column=0, columnspan=3, ipadx=6, ipady=6, padx=5, pady=15)

# ========================  end main frame  =======================================

# ***************************** frame date count 8  **********************************

lf_H8 = ttk.Frame(lf_MF, borderwidth=10, relief=SUNKEN)
lf_H8.config(width=width_frame, height=300)
lf_H8.grid_propagate(False)

# ************************************  ГАЗ  ****************************************************************

label_month = tk.Label(lf_H8, text="Введить показники лічильника газу: ", font=courier_14, foreground='red')
label_month.grid(row=0, column=0, ipadx=6, ipady=6, padx=5, pady=5)

data_gas = ttk.Entry(lf_H8)
data_gas.grid(row=0, column=1, ipadx=6, ipady=6)




# *************************************  СВІТЛО  **********************************************************************

label_year = tk.Label(lf_H8, text="Введить показники лічильника електроенергії: ", font=courier_14, foreground='red')
label_year.grid(row=1, column=0, ipadx=6, ipady=6, padx=5, pady=5)

data_light = ttk.Entry(lf_H8)
data_light.grid(row=1, column=1, ipadx=6, ipady=6)


# ************************************ ВОДА ***************************************************************************

label_year = tk.Label(lf_H8, text="Введить показники лічильника води: ", font=courier_14, foreground='red')
label_year.grid(row=2, column=0, ipadx=6, ipady=6, padx=5, pady=5)

data_water = ttk.Entry(lf_H8)
data_water.grid(row=2, column=1, ipadx=6, ipady=6)


# ************************************ КНОПКА ДОБАВИТИ  ***********************************************************

save_btn = tk.Button(lf_H8, text="ЗБЕРЕГТИ ПОКАЗНИКИ", font=courier_10, state='normal', command=add_date)
save_btn.grid(row=3, column=0, ipadx=6, ipady=6, padx=50, pady=30)


lf_H8.grid(column=0, row=1, padx=20, pady=10, sticky=W)


lf_MF.grid(column=0, row=1, ipadx=6, ipady=6, padx=20, pady=20)

root.mainloop()

'''
from openpyxl import load_workbook

wb = load_workbook("example.xlsx")
ws = wb.active

ws['C3'] = "Нові дані"
wb.save("example.xlsx")

'''