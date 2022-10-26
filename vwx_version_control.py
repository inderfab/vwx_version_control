#!/usr/bin/env python
# coding: utf-8

# # Vectorworks Versionen Übersicht
# ### Skript erstellt eine Tabelle mit allen Vectorworksdateien im selektierten Pfad
# Version 1.0, Fabio Indergand

# 
# #### Bitte den Suchpfad eingeben:
# Zum Beispiel Rechtsklick auf gewünschten Ordner, Alt Taste drücken und "Pfadname kopieren", innerhalb der Anführungszeichen einfügen

# In[1]:

print('Welcher Ordner soll nach VWX Dateien durchsucht werden?\nBeispiel: /Users/fabioindergand')
searchpath = input()


# #### Bitte den Exportpfad inklusive Dateinamen eingeben:

# In[2]:

print('Wo soll das Excel gespeichert werden (inkl Dateinamen angeben?\nBeispiel: /Users/fabioindergand/Desktop/Dateiname')
exportpath = input()


# #### Jetzt nur noch im Menü unter Kernel-Restart and Run All anwählen, die Datei wird dann am gewünschten Ort abgelegt.
# Wenn oben Rechts unter Logout der Kreis schwarz ist, läuft das Programm, sobald er wieder leer ist sollte der Prozess abgeschlossen sein.
# Je nach gewähltem Ordner braucht der Terminal noch Zugangsrechte, diese werden dann laufend abgefragt.
# Als Grundlage müssen noch einige Module installiert werden

# In[17]:


import glob,os
import regex as re
import pandas as pd
import xlsxwriter

data = []
conversionDate = {1:1, 2:2, 3:3, 4:4, 5:5, 6:6, 7:7, 8:8, 9:9, 10:10, 11:11, 12:12, 13:2008, 14:2009, 15:2010, 16:2011, 17:2012, 18:2013, 19:2014, 20:2015, 21:2016, 22:2017, 23:2018, 24:2019, 25:2020, 26:2021, 27:2022, 28:2023, 29:2024, 30:2025}


def findVersion(file):
    with open(file, 'rb') as content:
        bcontent = content.readlines(1)
        content = str(bcontent)
        versions = re.findall(r"VW\d+\.?\d.?\d?", content)
        
        startend=[]
        for version in versions:
            nbr = re.findall("(?<=VW)\d*",version)[0]
            version = re.sub("(?<=VW)\d*",str(conversionDate[int(nbr)]),version)
            startend.append(version)
        return startend

for root, dirs, files in os.walk(searchpath):
    for file in files:
        if file.endswith(".vwx"):
            path = os.path.join(root, file)
            startend = findVersion(path)
            data.append([file, startend[0], startend[1],''])

df = pd.DataFrame(data,columns =['File',"Start","Ende","Link"])
#df.to_excel('{}/Dateiversionen.xlsx'.format(searchpath), sheet_name='Versionen', index=False)

writer = pd.ExcelWriter('{}.xlsx'.format(exportpath), engine='xlsxwriter')
df.to_excel(writer, sheet_name='Versionen', index=False)
workbook  = writer.book
worksheet = writer.sheets['Versionen']

for index, row in df.iterrows():
    rownr = index + 2
    worksheet.write_url('D{}'.format(rownr),'file://{}'.format(os.path.join(searchpath, row["File"]) ))

worksheet.set_column(0, 0, 40)
worksheet.set_column(1, 2, 12)
worksheet.set_column(3, 3, 70)

writer.save()
print('Tabelle gespeichert unter{}'.format(exportpath))
