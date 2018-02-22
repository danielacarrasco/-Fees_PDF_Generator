# -*- coding: utf-8 -*-
"""
Created on Fri Nov 17 10:33:49 2017

@author: carrascod
"""

import wx 
import xlwings as xlw
from datetime import datetime
import glob
import os
import fnmatch
import numpy as np

FeesTemplate_main = wx.App(False) 
wb = xlw.Book('new_template.xltm')
sh_data = wb.sheets['Data'] 
#frame = wx.Frame(None, wx.ID_ANY, "Hello World") 


#frame.Show(True) 


#FeesTemplate_main.MainLoop()

source_path = 'A:\\Planning and Performance\\Student Compliance and Reporting\\Publishing Requirements\\Publications 2018\\'
list_of_files = sorted(glob.iglob(source_path+'*'), key=os.path.getctime, reverse = True) # * means all if need specific format then *.csv



latest_folder = list_of_files[0]
second_to_last = list_of_files[1]
last_date = datetime.fromtimestamp(os.path.getctime(second_to_last)).strftime('%m/%d/%Y')
sh_data.range('A1'.format(1)).value = last_date

file_names = ['*AU_AND_INT_CO*','*AU_FEES*','*INT_FEES*','*CSP_FEES*','*CAP_ASSES_FEES*']

for file in os.listdir(latest_folder):
    if fnmatch.fnmatch(file, file_names[0]) is True:
        sh_data.range('B1').value = latest_folder+'\\'+file #AU_AND_INT
    elif fnmatch.fnmatch(file, file_names[1]) is True:
        sh_data.range('B2').value = latest_folder+'\\'+file # AU_FEES
    elif fnmatch.fnmatch(file, file_names[2]) is True:
        sh_data.range('B3').value = latest_folder+'\\'+file # INT_FEES
    elif fnmatch.fnmatch(file, file_names[3]) is True:
        sh_data.range('B4').value = latest_folder+'\\'+file # CSP_FEES
    elif fnmatch.fnmatch(file, file_names[4]) is True:
        sh_data.range('B5').value = latest_folder+'\\'+file # CAP_ASSES

my_macro = wb.macro('Get_Copy_Data_ALL_FEES')
my_macro()
FeesTemplate_main.MainLoop()