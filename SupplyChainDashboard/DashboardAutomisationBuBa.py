#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Feb 19 21:39:25 2022

@author: Robert_Hennings
"""

#Packages importieren
import numpy as np
import pandas as pd
import xlwings as xw



#Dateipfad festlegen mit der Datei die die Urpsungsbloombergschnittstelle darstellt

path = "path/SupplyChainDashboard_DatenTableau.xlsx"
path_to_save = "path_save/Downloads/"
name_new_wb = "Daten_Export.xlsx"
data_after_sheet = "BloombergDatenImport-->"
wb = xw.Book(path)

#Reading in Data from the sheets
#Getting an overview of all existing sheets
print(wb.sheets.count) #number of sheets
print(wb.sheets) #names of sheets


#Finding out which Index the sheet BloombergDatenImport has
#Reading in the data per sheet from there on

sheet_list = list(wb.sheets)
#List and Index of every sheet in the Workbook
sheet_list_name = [i.name for i in sheet_list]

data_after_sheet in sheet_list_name


np.where(np.array(sheet_list_name) == data_after_sheet)[0][0]
pos_data_after_sheet = np.where(np.array(sheet_list_name) == data_after_sheet)[0][0]
#2 position
#ab dieser Position sollen nun alle sheets als eigener dataframe eingelesen werden
del wb.sheets.name[0]
del wb.sheets.name[1]
del wb.sheets.name[2]

for i in wb.sheets:
    #Dynamically create Data frames
    vars()[i.name] = pd.DataFrame(i.range("A1:E10").value)
    

#die eingelesenen neuen Daten (nach Aktualisierung) sollen nun in eine neue Exceldatei sheet für sheet geschrieben werden

#opening new excel workbook
wb_new = xw.Book()

#Adding all new created sheets to dfs as a new sheet in the new workbook
for i in wb.sheets:
    #print(i)
    #Dynamically create new sheets in new workbook
    wb_new.sheets.add(i.name, before="Tabelle 1")
    wb_new.sheets[i.name].range("A1:E10").value = wb.sheets[i.name].range("A1:E10").value
  
#trying the same but with using copy 
for i in wb.sheets:
    #print(i.name)
    i.copy(before=wb_new.sheets[0])





wb_new.sheets[0].delete()
wb_new.sheets["HowTo"].delete()
#Move the copied sheets in the correct order


#ws1.api.Move(None, After=ws3.api)

wb_new.sheets[0].name ="Test1"
wb_new.sheets.add("Test2")
wb_new.sheets.add("Test3")



help(wb.sheets[0])



wb_new.save(path_to_save+name_new_wb)







    
    



