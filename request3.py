# -*- coding: utf-8 -*-
'''
Name file : requst3.py
Lasted Created : 31 MAR 2020
For : Final Project last semester in Mechatronic Engineering King Mongkut of technology thonburi
Created by: Boonnuea Boonmana 59070502208
Order : 1/4
'''

#!/usr/bin/env python3

import requests as req
import http.server
import socketserver
import time
import pandas as pd
from datetime import date
import xlwings as xw

timedelay = time.time()

def lastRow(idx, workbook, col=1):
    """ Find the last row in the worksheet that contains data.

    idx: Specifies the worksheet to select. Starts counting from zero.

    workbook: Specifies the workbook

    col: The column in which to look for the last cell containing data.
    """

    ws = workbook.sheets[idx]

    lwr_r_cell = ws.cells.last_cell      # lower right cell
    lwr_row = lwr_r_cell.row             # row of the lower right cell
    lwr_cell = ws.range((lwr_row, col))  # change to your specified column

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')    # go up untill you hit a non-empty cell

    return lwr_cell.row

def start(url,file):
        """ This def we used for running keep data from server and save to
        
        excel in real time but it use to open excel all the time

        """
    
    
        wb = xw.Book(file)   #  Open file excel
        # lastRow('Sheet1',wb) 
        row = lastRow('Sheet1',wb)                                                  #  Find last row of data
        sht = wb.sheets['Sheet1']                                                   #  Select sheet to save data
        
        # d1 = date.today()
        # day = d1.day
        # month = d1.month
        # year = d1.year
        
        # url = 'http://192.168.1.44:5000'                                          #  IP address for my apartment wifi
        # url = 'http://192.168.43.102:5000'                                          #  IP addres for my phone wifi
        x = req.get(url, params={'<p>':'</p>'})                                     #  Get data from IP address
        r = x.text                                                                  #  Change data from IP into string
        r = r.replace('\n','')                                                      #  Change text
        s = r.split("$")                                                            #  Spilt text by used '$'
        a = s[1]                                                                    #  Data 1 from text
        b = s[2]                                                                    #  Data 2 from text
        c = s[3]                                                                    #  Date 3 from text
        d = s[4]                                                                    #  Data 4 from text
        sht.range('A{0}'.format(row+1)).value = row+1                               #  First column save data
        sht.range('B{0}'.format(row+1)).value = a                                   #  Second  column save data 1
        sht.range('C{0}'.format(row+1)).value = b                                   #  Third column save data 2
        sht.range('D{0}'.format(row+1)).value = c                                   #  Third column save data 3
        sht.range('E{0}'.format(row+1)).value = d                                   #  Third column save data 4

        ''' # keep data into array
        # point.append(a)
        # center.append(b)
        # angle.append(c)
        # status.append(d)
        '''

        ''' # print to check data
        print(a)
        print(b)
        print(c)
        print(d)
        '''

        # time.sleep(1)                                                             #  For keep delay to check data
        # print('save')
        # if count == 10:                                                           #  Use to check data 10 row
        #     break                                                                 #  Use in whlie loop
        return s[-1]
def stop():
    print('stop reading!')