# -*- coding: utf-8 -*-
"""
Created on Fri Mar 25 10:42:46 2022

@author: up17292
"""
#This Python Script maintains the functionality of the Supply Chain Dashboard and produces the "Export" file that Tableau reads in 
#to produce the Graphs
#Author: Intern Robert Hennings
#Importing all the necessary Packages that will be used within this cript for the functionality
import pandas as pd
import numpy as np
import xlwings as xw
import openpyxl
import matplotlib.pyplot as plt
import seaborn as sns
import time
#Setting the currently open book as the active one in use to work with it, Idea: try not to work with fixed file paths as these might change over time
#so it is more convenient to only edit the paths to this .py file and the path where the Export file should be saved
#book = xw.books.active
book = xw.Book("path\SupplyChainDashboard.xlsm")
#Reading in and saving all the single excel sheets of the excel file in a dataframe for loopin over them to collect the data
df = pd.DataFrame(index = range(0,book.sheets.count), columns = range(0,2))
for i in range(0,book.sheets.count):
#    print(i, book.sheets[i].name)
    df.iloc[i,0] = i
    df.iloc[i,1] = book.sheets[i].name
df.columns =["Sheet_Number","Sheet_Name"]
#The print method in line 24 can be activated to see all the existing sheets that will be printed ot when the for loop is running

#Now creating a function that performs all the necessary operations to clean the data and cut it off at the end so that there will be even lenghts for all data columns in each sheet
def ReadDataTable(table_name, where):
    globals()[f"{table_name}"] = pd.DataFrame(book.sheets[table_name].range(where).options(expand="table").value) #saving each excel sheet as a DataFrame to work with the pandas library in python
    globals()[f"{table_name}"].columns = globals()[f"{table_name}"].loc[0] #setting the first row as the columns of the frame, can also be done via header attribute in the options method one line above
    globals()[f"{table_name}"].drop([0], inplace =True) #then dropping this row as it is no longer needed
    globals()[f"{table_name}"].reset_index(drop=True, inplace=True) #resetting the index to start again at 0 like a normal one would do in python
    globals()[f"{table_name}"].set_index( globals()[f"{table_name}"].iloc[:,0], inplace=True) #setting the first column (the dates) as the new index
    globals()[f"{table_name}"].drop(labels=globals()[f"{table_name}"].columns[0],axis=1, inplace =True) #then dropping the first column as we no longer need it (it is the new index now)
    globals()[f"{table_name}"].dropna(inplace=True) #dropping all missing values to cut the frame evenly off at the end so that Tableau cant produce Data Spikes in its Graphs wehen there is one column longer than the others (uneven dimensions)
#Some Sheets need a special treatment where the Na values shouldnt be dropped so wehen we compare the both functions ReadDataTable and ReadDtaTableSpecial the only difference is the last row where NAs arent dropped in ReadDataTableSpecial
def ReadDataTableSpecial(table_name, where):
    globals()[f"{table_name}"] = pd.DataFrame(book.sheets[table_name].range(where).options(expand="table").value) #saving each excel sheet as a DataFrame
    globals()[f"{table_name}"].columns = globals()[f"{table_name}"].loc[0] #setting the first row as the columns of the frame
    globals()[f"{table_name}"].drop([0], inplace =True) #then dropping this row
    globals()[f"{table_name}"].reset_index(drop=True, inplace=True) #resetting the index to start again at 0 like normal
    globals()[f"{table_name}"].set_index( globals()[f"{table_name}"].iloc[:,0], inplace=True) #setting the first column (the dates) as the new index
    globals()[f"{table_name}"].drop(labels=globals()[f"{table_name}"].columns[0],axis=1, inplace =True) #then dropping the first column
   
#Reading in the path where to save the generated Export File and its name which are given by the user 
path_save = book.sheets[0].range("B19").value+"\\"+book.sheets[0].range("B20").value+".xlsx" 

#Now we use the previsously defined functions ReadDataTable and ist Special Variant to read in the Data from the single Sheets
#Reading in the Data for the Totalscore that begins in X1 in the sheet HeatmapGesamtScore
ReadDataTable(df.iloc[6,1],"X1")
#After reading in the Data all rows that contain #NV should be dropped to leave only the relevant values
#Rows with #NV are read in as the value -2146826246 so these should be eliminated
HeatmapGesamtScore.drop((HeatmapGesamtScore[HeatmapGesamtScore.Gesamtscore ==-2146826246]).index, inplace=True)
#Now reading in the Data with the normal Function as there is no need for special treatment

for i in df.iloc[9:18,1]:
    ReadDataTable(i,"A1")
#Special Handling for this sheet: IfoMaterialMangelGewerbe
ReadDataTableSpecial(df.iloc[18,1],"A1")
#Normal Handling again here
for i in df.iloc[19:25,1]:
    ReadDataTable(i,"A1")

#Special Handling for this sheet: Kiel_Indicator_Waren_Hafenstau
ReadDataTableSpecial(df.iloc[25,1],"A1")

for i in df.iloc[27:30,1]:
    ReadDataTable(i,"A1")
#Special Handling for this sheet: Diesel
ReadDataTableSpecial(df.iloc[30,1],"A1")

for i in df.iloc[31:39,1]:
    ReadDataTable(i,"A1")
#Single sheets with normal treatment but the first one will be read in from the cell A8 on, as the third one too
ReadDataTable(df.iloc[39,1],"A8")
ReadDataTable(df.iloc[42,1],"A1")
ReadDataTable(df.iloc[43,1],"A8")
ReadDataTable(df.iloc[44,1],"A1")


for i in df.iloc[46:49,1]:
    ReadDataTable(i,"A1")
#Single sheet with normal Treatment
ReadDataTable(df.iloc[50,1],"A1")


#In case at a later point in time sheets are added to the file these should of course be processed too
#Integrating the reading and writing in of additional sheets, the standard file has 52 sheets so if ones are added the loop will be executed
if book.sheets.count > 52:
    for i in df.iloc[52:df.shape[0],1]:
        ReadDataTable(i,"A1") #only restriction is that the data has to begin in cell A1 without any breaks as here the normal ReadDataTable fucntion is used
else:
    pass

#Write the actual Time and Date into the main sheet to let the User know and see when the data was last updated into an Export file
book.sheets[0].range("B21").value  = time.asctime(time.localtime(time.time()))
#Now after all the updated data is read in and cleaned/cut off at NaN or #NV values
#the data has to be dropped into a  new file to be read in by Tableau

#The part in the lines 145 - 148 serves as BackUp because the best version currently operates with the lines 151 and 152 where there is everything
#running in an invisible excel instance so that the user sees only a python prompt and nothing else working in the foreground
#but if necessary the lines 146 and 148 can be uncommented and used as well (but then comment lines 151 and 152)
#Opening a new book and dropping everything in step by step as the user can see
#book_new = xw.Book()
#Setting the new opened file as the active one in utilization
#book_new = xw.books.active

#Better later updated version here in the lines 151,152: Generate an invsisble Excel instance and add a new book wehere the data will be placed
excel_app =xw.App(visible=False)
book_new= excel_app.books.add()
#Writing a small function to dump each of the generated DataFrames into a new excel sheet in the openend workbook with the same names as in the original sheet

def WriteDataTable(table_name, where):
    book_new.sheets.add(table_name) #Set the Sheet name to exactly the original ones from the main file
    book_new.sheets[table_name].range(where).value = globals()[f"{table_name}"] #say wehere to drop the data in the sheet
#Just like in the first case wehen reading in the data now writing it to the new file
#changing the order in which the sheets will appear to be in the correct order
if book.sheets.count > 52:
    for i in df.iloc[52:df.shape[0],1]:
        WriteDataTable(i,"A1")
else:
    pass

WriteDataTable(df.iloc[50,1],"A1")
for i in df.iloc[46:49,1]:
    WriteDataTable(i,"A1")
WriteDataTable(df.iloc[44,1],"A1")
WriteDataTable(df.iloc[43,1],"A1")
WriteDataTable(df.iloc[42,1],"A1")
WriteDataTable(df.iloc[39,1],"A1")
for i in df.iloc[31:39,1]:
    WriteDataTable(i,"A1")
WriteDataTable(df.iloc[30,1],"A1")
for i in df.iloc[27:30,1]:
    WriteDataTable(i,"A1")
WriteDataTable(df.iloc[25,1],"A1")
for i in df.iloc[19:25,1]:
    WriteDataTable(i,"A1")
WriteDataTable(df.iloc[18,1],"A1")
for i in df.iloc[9:18,1]:
    WriteDataTable(i,"A1")
WriteDataTable(df.iloc[6,1],"A1")

#At the end saving the file with the tableau-ready data and closing the new workbokk that contains the sheets
book_new.save(path_save) #saving the Export File at the desired file path with the desired name
book_new.close() #closing the file 