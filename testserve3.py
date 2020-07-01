# -*- coding: utf-8 -*-
'''
Name file : testserver3.py
Lasted Created : 31 MAR 2020
For : Final Project last semester in Mechatronic Engineering King Mongkut of technology thonburi
Created by: Boonnuea Boonmana 59070502208
Order : 2/4
'''

import sys
sys.path.append('C:\\Users\\super\\Desktop\\4\\Final_Project\\used\\AGV_new\\goto')
import request3 as req3
#import requests as req
from flask import Flask
import pandas as pd
import xlwings as xw
import os
app = Flask(__name__)
def lastRow(idx, workbook, col=1):
    """ Find the last row in the worksheet that contains data.

    idx: Specifies the worksheet to select. Starts counting from zero.

    workbook: Specifies the workbook

    col: The column in which to look for the last cell containing data.
    """

    ws = workbook.sheets[idx]
 
    lwr_r_cell = ws.cells.last_cell                                             #  lower right cell
    lwr_row = lwr_r_cell.row                                                    #  row of the lower right cell
    lwr_cell = ws.range((lwr_row, col))                                         #  change to your specified column

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')                                           #  go up untill you hit a non-empty cell

    return lwr_cell.row


gotx = []
goty = []
numb = []
j = 0
chk = 0
counter = 0
@app.route("/")                                                                 #  To creat spread site
def home():
    global chkx,chky
    # os.system('AGV_GUI_20.py')
    req3.start('http://192.168.1.44:5000',r'C:/Users/super/Desktop/4/Final_Project/used/AGV_new/request.xlsx')
    wb = xw.Book(r'C:/Users/super/Desktop/4/Final_Project/used/AGV_new/goto/requestout.xlsx') 
    #lastRow('Sheet1',wb)
    row = lastRow('Sheet1',wb)
    sht = wb.sheets['Sheet1']
    for i in range(2,row):
        gotx.append(int(sht.range('B{0}'.format(i)).value))
        goty.append(int(sht.range('C{0}'.format(i)).value))
        numb.append(sht.range('D{0}'.format(i)).value)
        j = 0
    print(chkx,chky)
    
    if chkx == str(gotx[j]) and  chky == str(goty[j]) :
        overall = "001${}/{}${}".format(gotx[j+1],goty[j+1],numb[j+1])
    elif j > len(gotx):
        overall = "001${}/{}${}".format(gotx[0],goty[0],numb[0])
    else:
        j = j+1
    
    #############################################################################    
    """
    This condition used for take data and send back to server
    this condition given to change for support data
    """
    #if out == 'done':    
    #    overall = '001$3/0$up'
    #elif out == 'tuning':
    #overall = '001$3/0$up'
        #data = numb.split("/")
        #overall = int(data[0])+int(data[1])
    #############################################################################
    # req.start()
    print(overall)
    print('send')
    return str(overall)
    
if __name__ == "__main__":
    """  
    use to initial file for running. Will change to main file if AGV_GUI.py runnign perfectly. 
    These used for run server.
    """
    # app.run(debug=True)
    # req.start('http://192.168.1.44:5000',r'C:/Users/super/Desktop/4/Final_Project/data/request.xlsx')
    app.run(host='0.0.0.0',debug=True, threaded=True,  use_reloader=False)