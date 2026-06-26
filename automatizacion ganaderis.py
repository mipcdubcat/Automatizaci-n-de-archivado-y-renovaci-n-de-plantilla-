import pandas as pd
import openpyxl as opx
import time
import os
from datetime import datetime as dt
DIARIO_G = "C:/Users/DUBAN/documents/ganaderia/DIARIO GANADERIA.xlsx"
Hora_actual = dt.now().strftime("%H:%M")
Horas = ["01:00","02:00", "03:00", "04:00", "05:00", "06:00", "07:00", "08:00", "09:00", "10:00", "11:00", "12:00", "13:00"]
while True:
    Hora_actual = dt.now().strftime("%H:%M")
    HORA = dt.now().strftime("%H")
    mes = dt.now().strftime("%d")
    if Hora_actual in Horas:
        print("guardando informe diario...")
        oldbook = pd.read_excel("C:/Users/DUBAN/OneDrive/DIARIO GANADERIA.xlsx")
        newbook = f"C:/Users/DUBAN/documents/ganaderia/DIA_{mes}/{Hora_actual.replace(":", "-")}.xlsx"
        os.remove("C:/Users/DUBAN/OneDrive/DIARIO GANADERIA.xlsx")
        onedrive = pd.read_excel("C:/Users/DUBAN/documents/ganaderia/DIARIO GANADERIA.xlsx")
        carpeta = "C:/Users/DUBAN/OneDrive/DIARIO GANADERIA.xlsx"
        onedrive.to_excel(carpeta)
        oldbook.to_excel(newbook)
    if Hora_actual == "00:01":
        HOY = dt.now().strftime("%d-%m-%Y")
        print("guardando informe mensual...")
        oldbook_m = pd.read_excel("C:/Users/DUBAN/OneDrive/MENSUAL GANADERIA.xlsx")
        newbook_m = f"C:/Users/DUBAN/documents/ganaderia/24_HORAS/{HOY}.xlsx"
        oldbook_m.to_excel(newbook_m)
    time.sleep(60)
