#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 19:22:46 2022

@author: Robert_Hennings
"""

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt




sns
dir(sns)

df.corr()

sns.heatmap(df,linewidths=2,linecolor="black",square=True,cmap="RdYlGn_r", xticklabels=False,annot=True)
plt.title("Test")
# plt.ylabel("Hallo")
plt.text(-1.2,2,"T")
plt.savefig("/Users/Robert_Hennings/Downloads/Fig.png", dpi=400)

t = np.random.randint(-1,17,10)
t2 = np.random.randint(-1,17,10)
t3 = np.random.randint(-1,17,10)
t4 = np.random.randint(-1,17,10)
t5 = np.random.randint(-1,17,10)

df = pd.DataFrame([t,t2,t3,t4,t5])


help(sns.heatmap)


help(plt.text)
rows = 8
columns = 10

numb = rows*columns
z = np.random.randint(-2,25,numb)
z.ndim 

z = z.reshape((rows,columns))

df = pd.DataFrame(z)

df.index = ["Var1", "Var2", "Var3", "Var4", "Var5", "Var6", "Var7", "Var8"]



