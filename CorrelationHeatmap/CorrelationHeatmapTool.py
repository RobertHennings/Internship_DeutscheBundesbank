#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Mar  1 17:11:25 2022

@author: Robert_Hennings
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Feb 26 20:30:52 2022

@author: Robert_Hennings
"""

"""Generate a correlation heatmap from custom loaded securities
Steps:
    1) Load the Ticker data into the excel file in the specified areas
    2) Click the generate Heatmap Button to generate a Correlation Heatmap into the next sheet
"""

#Script with chnaged Input Parameters

import xlwings as xw
import pandas as pd
import seaborn as sns
import numpy as np
import matplotlib.pyplot as plt
from pandas import ExcelWriter
import time
#def main():

wb = xw.books.active

path = wb.sheets[0].range("B16").value
file_name_out = wb.sheets[0].range("B17").value
path_save_out = path+"\\"+file_name_out+".xlsx" 
wb.sheets[0].range("B18").value  = time.asctime(time.localtime(time.time()))
title_heatmap = wb.sheets[1].range("B2").value
font_size_title = wb.sheets[1].range("B3").value
xaxis = wb.sheets[1].range("B4").value
yaxis = wb.sheets[1].range("B5").value
include_input_data = wb.sheets[1].range("B6").value
color_heatmap = wb.sheets[1].range("B7").value
rotation_axis = wb.sheets[1].range("B8").value
font_size = wb.sheets[1].range("B9").value
fig_dpi = wb.sheets[1].range("B10").value
   

#Read the data in the data input sheet
df = pd.DataFrame(wb.sheets[2].range("A1").options(expand="table").value)
#Setting the Columnnames
df.columns = df.loc[0]
df.drop([0], inplace=True)
df.reset_index(drop=True, inplace=True)
df.head()


df['Datum'] = pd.to_datetime(df['Datum'])
df.set_index('Datum', inplace=True)

#Deleting the Column with the dates
#df.drop(df.columns[0], axis=1, inplace=True)
#Adjusting the data types from object to int64
df = pd.DataFrame(data=df, dtype=np.float64)
df.dtypes
#generating the correlation heatmap
df_corr = df.corr()

#preparing for plotting
axis_labels = df.columns
heat = sns.heatmap(df_corr,annot=True,xticklabels=axis_labels, yticklabels=axis_labels, cmap=color_heatmap)
plt.xticks(rotation=rotation_axis)
sns.set(font_scale=font_size)
heat.set_title(title_heatmap,fontsize=font_size_title)
heat.set(xlabel=xaxis, ylabel=yaxis)
fig = heat.get_figure()
fig.set_dpi(fig_dpi)

#fig.savefig(path+"Correlation_Heatmap_Tool_Output.png",dpi=1200,bbox_inches="tight")

#Saving the picture to the new generated output file


   # with pd.ExcelWriter(path+filename_output) as writer:
           #df.to_excel(writer, sheet_name="BloombergInputData")
           #df_corr.style.background_gradient(cmap="hot").to_excel(writer,sheet_name="Correlation_Heatmap_Excel")

excel_app =xw.App(visible=False)

wb_out = excel_app.books.add()
#wb_out = xw.books.active
wb_out.sheets.add("Heatmap.png")
wb_out.sheets["Heatmap.png"].pictures.add(fig,name="Correlation Heatmap",update=True)
wb_out.sheets.add("Correlation_Heatmap_Excel")
wb_out.sheets["Correlation_Heatmap_Excel"].range("A1").value = df.corr()
#    wb_out.sheets["Correlation_Heatmap_Excel"].range("A1").value = df_corr.style.background_gradient(cmap="hot")
wb_out.sheets[2].delete()

if include_input_data == "Yes":
    wb_out.sheets.add("Bloomberg Data Input")
    wb_out.sheets["Bloomberg Data Input"].range("A1").value = df
else:
     
     pass

wb_out.save(path_save_out)
wb_out.close()


wb = xw.books.active
wb.save()
#wb.close()
#if __name__ == "__main__":
#    main()

#with ExcelWriter("R:\Zentrale\ZB-M\ZB-M-NEU\Daten\CEPH\Market Intelligence\Praktikanten\Robert Hennings/Test.xlsx") as writer:
#    df_corr.style.background_gradient(cmap="hot").to_excel(writer, sheet_name="Test")
    


