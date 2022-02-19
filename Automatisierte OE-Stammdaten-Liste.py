#Automatisierte OE-Stammdaten-Liste
from time import time, sleep
import os, shutil
from shutil import move
import glob
import webbrowser
from datetime import datetime
import win32com.client
from win32com.client import Dispatch, constants
import pandas as pd
import re

datum_now = datetime.now().strftime("%Y_%m_%d")
datum2 = datetime.now().strftime("%d.%m.%Y")

def moven():
    for file in glob.glob("C:\\Users\\Maximilian.Rasch\\Downloads/"+ "*ZENSIERT*" + ".xls"):
            global d
            u2 = "C:\\Users\\Maximilian.Rasch\\Downloads\\" + os.path.basename(file)
            d = "C:\\Users\\Maximilian.Rasch\\Desktop\\Arbeitsordner\\" + os.path.basename(file)
            shutil.move(u2, d)
            
webbrowser.open('LINK ZUR WEBSITE ZENSIERT')

def convert(d):
    fname = d
    excel = win32com.client.dynamic.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)


    global excel_file_1
    excel_file_1 = "C:\\Users\\Maximilian.Rasch\\Desktop\\Arbeitsordner\\" + datum_now + "_EXCEL DATEI ZENSIERT.xlsx"

    wb.SaveAs(excel_file_1, FileFormat = 51)    #FileFormat = 51 ->  .xlsx
    wb.Close()                               #FileFormat = 56 -> .xls
    excel.Application.Quit()


def filter(excel_file_1):

    # Einlesen der Excel Datei in ein Dataframe
    df1 = pd.read_excel(excel_file_1)

    filtered_df_1 = df1.loc[
        # Filter nach: Ohne VWT, Alles über fünfstellig, Alle RL, Alle Standorte in Neu-Isenburg
        (df1["Region"] != "VWT")
        & (df1["OENr"] < 9999) #OE mit 4 Stelliger Nummer raus
        & (df1["t_DistrictmanagerName"] != "RL Ost") 
        & (df1["t_DistrictmanagerName"] != "RL Nord")
        & (df1["t_DistrictmanagerName"] != "RL West")
        & (df1["t_DistrictmanagerName"] != "RL Mitte")
        & (df1["t_DistrictmanagerName"] != "RL Süd")
        & (df1["t_DistrictmanagerName"] != "RL Süd-Ost") # ALLE Regionalleitungen raus
        & (df1["Ort"] != "Neu-Isenburg") #Alle Kostenstellen Neu Isenburg raus
        & (~df1["OEName1"].str.contains("Budget"))
    ]
    # Hinzufügen von ZENSIERT -> weil vorher mit Filer Neu-Isenburg rausgenommen
    filtered_df_2 = df1.loc[(df1["ZENSIERT] == ZENSIERT)]

    # Zusammenfügen von beiden Dataframes
    vor_ergebnis = filtered_df_1.append(filtered_df_2)

    # Sortieren nach Numemer -> Aufsteigend
    df = vor_ergebnis.sort_values(by=["ZENSIERT"])

    # Schreiben der DF in neue Excel Datei (Mit Blattname & ohne Index in erster Zeile)
    global excel_filtered
    excel_filtered = "C:\\Users\Maximilian.Rasch\\Desktop\\Arbeitsordner\\" + str(datum_now) + "_ZENSIERT_bereinigt.xlsx"

    writer = pd.ExcelWriter(
        excel_filtered
    )
    df.to_excel(writer, sheet_name="Sheet1", index=False)

    #Formatieren der Kopfzeilen
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets["Sheet1"].set_column(col_idx, col_idx, column_width)


    writer.save()

ol = win32com.client.dynamic.Dispatch('Outlook.Application')
msg = ol.CreateItem(0)

def mailen():
    msg.GetInspector
    bodystart = re.search("<body.*?>", msg.HTMLBody)
    msg.HTMLBody = re.sub(bodystart.group(), bodystart.group()+"Sehr geehrte Damen,<br /><br />anbei erhalten Sie die Liste mit Stand vom " + datum2 + ".<br /><br />Bei Anmerkungen oder Rückfragen stehe ich Ihnen gerne zur Verfügung.", 
    msg.HTMLBody)

    msg.Subject = "ZENSIERT " +  datum2
    msg.To = "EMAIL 1 ZENSIERT; EMAIL 2 ZENSIERT"
    msg.CC = "EMAIL 3ZENSIERT;EMAIL 4ZENSIERT; E-Mail 5ZENSIERT"

    attachment1 = excel_filtered
    msg.Attachments.Add(Source=attachment1)

    msg.display()
    #Mail senden
    #newMail.send()

def moven2():

            u2 = excel_filtered
            d = "PFAD zensiert" + os.path.basename(excel_filtered)
            shutil.move(u2, d)
            u3 = excel_file_1
            d1 = "Pfad zensiert" + os.path.basename(excel_file_1)
            shutil.move(u3, d1)

sleep(2)
moven()
sleep(2)
convert(d)
sleep(1)
filter(excel_file_1)
mailen()
moven2()