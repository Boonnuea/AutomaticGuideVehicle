import sys
sys.path.append('C:\\Users\\super\\Desktop\\4\\Final_Project\\used\\pi')
import Astar_algorithm as ast
from tkinter import *
from tkinter import ttk
import tkinter.messagebox
import tkinter as tk
import time
from PIL import Image,ImageTk                                     # pip install pillow
import numpy as np

import pandas as pd
import xlwings as xw
from threading import Thread


'''variable'''
###############################################################################
i = 1                                   # For number of task

startAGV = [[4,0],[4,8]]

AGVstart = []
AGVstop = []
costAGV = []

inputrequest = []
outputrequest = []
costrequest = []

listpoint = []
listhitch = []

listnumber = []                         # List number

maxnumlist = []                         # Max number start to end (All)
maxnumlistA1 = []                       # Max number start to end (AGV1)
maxnumlistA2 = []                       # Max number start to end (AGV2)

Line = []                               # Line start to end (All)
LineA1 = []                             # Line start to end (AGV1)
LineA2 = []                             # Line start to end (AGV2)

LineAGV = []                            # Line AGV to start (All)
LineAGVA1 = []                          # Line AGV1 to start
LineAGVA2 = []                          # Line AGV2 to start

startcircleA1 = []                      # start circle (AGV1)
startcircleA2 = []                      # start circle (AGV2)

endcircleA1 = []                        # end circle (AGV1)
endcircleA2 = []                        # end circle (AGV2)

AGVnow1 = []                            # AGV1 circle (Simulation)
AGVnow2 = []                            # AGV2 circle (Simulation)

AGVtostart = []                         # Max number AGV to start (All)
AGVtostartA1 = []                       # Max number AGV1 to start
AGVtostartA2 = []                       # Max number AGV2 to start

listNum1 = []                           # Number in Num1 for simulation start to end (AGV1)
listNum2 = []                           # Number in Num2 for simulation start to end (AGV2)

listnum1 = []                           # Number in num1 for simulation AGV1 to end
listnum2 = []                           # Number in num1 for simulation AGV2 to end

finishAGV1 = []                         # Finish simlation AGV1 to start
finishAGV2 = []                         # Finish simlation AGV2 to start

finishA1 = []                           # Finish simlation start to end (AGV1)
finishA2 = []                           # Finish simlation start to end (AGV2)

listdatay1 = []                         # y for simlation AGV1 to start
listdatax1 = []                         # x for simlation AGV1 to start

listDatay1 = []                         # y for simlation start to end (AGV1)
listDatax1 = []                         # x for simlation start to end (AGV1)

listdatay2 = []                         # y for simlation AGV2 to start
listdatax2 = []                         # x for simlation AGV2 to start

listDatay2 = []                         # y for simlation start to end (AGV2)
listDatax2 = []                         # x for simlation start to end (AGV2)

sim1 = []                               # Simulation (AGV1)
sim2 = []                               # Simulation (AGV2)

AGV1order = []                          # AGV1 Order before
AGV2order = []                          # AGV2 Order before

area = [[0, 0, 0, 0, 0, 0, 1, 0, 0],
        [0, 0, 0, 0, 0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0, 0, 0, 0, 0],
        [0, 0, 1, 0, 0, 0, 0, 0, 0]]  # area used

###############################################################################
'''open excel'''
file = r'C:/Users/super/Desktop/4/Final_Project/used/AGV_condition/data.xlsx'

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


###############################################################################

###############################################################################
'''loop for cal Astar'''
''' 1 '''
def calrequest(file,areaused,pick,send):
    '''
    this sequence used for calculate distance for pickup and send of request
    and keep it into excel
    :param areaused: area that we used to calculate
    :param pick: point of picking request [x,y]
    :param send: point of sending request [x,y]
    '''
    ''' main calculate'''
    wb = xw.Book(file) 
    
    Num = []
    
    rowrequest = lastRow('Request',wb)                                         #  Find last row of data
    shtrequest = wb.sheets['Request']                                           #  Select sheet to save data
    
    '''calculate request'''
    maze = areaused
    start = [pick[0], pick[1]]                                                  # starting position
    end = [send[0], send[1]]                                                    # ending position
    cost = 1                                                                    # cost per movement
    pathout = ast.search(maze,cost, start, end)
    
    direction = []
    redirection = []
    Num = []
    prey = []
    costout = 0
    
    for y in pathout:
        for x in y:                                                             # outx = [Each row] in array
            if x != -1:                                                         # outy = Row in array
                direction.append([pathout.index(y),y.index(x)]) #[x,y]
                Num.append(int(x))                                                  # Num = Number 0,1,2,... of path to move in array
    for direct in range(len(Num)):
        point = Num.index(direct)
        redirection.append(direction[point])
    prey = redirection[0][1]
    #print(prey)
    for get in range(len(redirection)):
        if redirection[get][1] != prey:
            costout = costout + 2
            prey = redirection[get][1]
            #print(costout,prey)
        else:
            costout = costout + 1
            #print(costout,prey)
    a = rowrequest
    b = start[0]                                                                #  Data 1 = pick x coordinates
    c = start[1]                                                                #  Data 2 = pick y coordinates
    d = end[0]                                                                  #  Data 3 = send x coordinates      
    e = end[1]                                                                  #  Data 4 = send y coordinates
    f = costout                                                                #  Date 5 = cost
    shtrequest.range('A{0}'.format(rowrequest+1)).value = a                     #  First column save order list
    shtrequest.range('B{0}'.format(rowrequest+1)).value = b                     #  Second  column save Data 1
    shtrequest.range('C{0}'.format(rowrequest+1)).value = c                     #  Third column save Data 2
    shtrequest.range('D{0}'.format(rowrequest+1)).value = d                     #  Fourth column save Data 3
    shtrequest.range('E{0}'.format(rowrequest+1)).value = e                     #  Third column save Data 4
    shtrequest.range('F{0}'.format(rowrequest+1)).value = f                     #  Fourth column save Data 5
    shtrequest.range('G{0}'.format(rowrequest+1)).value = 1                     #  Fifth column save Data 6

    return start,end,max(Num),pathout
###############################################################################
'''loop for cal and choose AGV'''
def calAGV(area,startAGV,file,inputsheet,outputsheet):
    '''
    this sequence used for calculate all of AGV car that can calculate
    :param numbcar: number of cars and all of car position [x,y]
    :param areaused: area for testing
    '''
    wb = xw.Book(file) 
    
    pick = []
    send = []
    cost = []
    orderin = []
    status = []
    picknum = []
    listAGV1 = [[],[],[],[],[],[],[],[]]          # order \ startx \ starty \ stopx \ stopy \ cost \ order
    listAGV2 = [[],[],[],[],[],[],[],[]]         # order \ startx \ starty \ stopx \ stopy \ cost \ order
    
    '''get main data'''
    numberquest = lastRow(inputsheet,wb)                                        # Find last row of data                                                                                                 #  Select sheet to save data
    getsheet = wb.sheets[inputsheet]                                            # Open sheet of excel    
    
    getlist = int(getsheet.range('A{0}'.format(numberquest)).value)
    # print('{},{}'.format(getlist,numberquest))
    
    if getlist == 1:
        pick.append([int(getsheet.range('B{0}'.format(numberquest)).value)
                ,int(getsheet.range('C{0}'.format(numberquest)).value)])
        send.append([int(getsheet.range('D{0}'.format(numberquest)).value)
                ,int(getsheet.range('E{0}'.format(numberquest)).value)])
        cost.append(int(getsheet.range('F{0}'.format(numberquest)).value))
        status.append(getsheet.range('G{0}'.format(numberquest)).value)
        orderin.append(int(getsheet.range('H{0}'.format(numberquest)).value))
        picknum.append(int(getsheet.range('I{0}'.format(numberquest)).value))
        
    else:
        for i in range (numberquest+1-getlist,numberquest+1):
            pick.append([int(getsheet.range('B{0}'.format(i)).value)
                    ,int(getsheet.range('C{0}'.format(i)).value)])
            send.append([int(getsheet.range('D{0}'.format(i)).value)
                    ,int(getsheet.range('E{0}'.format(i)).value)])
            cost.append(int(getsheet.range('F{0}'.format(i)).value))
            status.append(getsheet.range('G{0}'.format(i)).value)
            orderin.append(int(getsheet.range('H{0}'.format(i)).value))
            picknum.append(int(getsheet.range('I{0}'.format(i)).value))
   
    '''calculate request'''                                                       # Choose car for once
    
    rowAGV1 = lastRow(outputsheet[0],wb)                        # Find last row of data
    shtAGV1 = wb.sheets[outputsheet[0]]   
    rowAGV2 = lastRow(outputsheet[1],wb)                        # Find last row of data
    shtAGV2 = wb.sheets[outputsheet[1]]   
    
    order = 0
    preorder = [[],[]]
    passcost = [[],[]]
    #prevalcost = [[],[]]
    Pickin = []                                                          
    Sendin = []    
    #Costout = []
    Costout = [cost.copy(),cost.copy()]                                      
    Pickin = [pick.copy(),pick.copy()]
    Sendin = [send.copy(),send.copy()] 
    lowCost = [0,0]
    preCost = [0,0]
    calcost = 0
    
    status.insert(0,'non')
    picknum.insert(0,0)
    
    Pickin[0].insert(0,[0,0])
    Sendin[0].insert(0,startAGV[0])
    Costout[0].insert(0,0)
    Pickin[1].insert(0,[0,0])
    Sendin[1].insert(0,startAGV[1])
    Costout[1].insert(0,0)
    
    print('data AGV1:')
    print(Pickin[0])
    print(Sendin[0])
    print('data AGV2:')
    print(Pickin[1])
    print(Sendin[1])
    print(orderin)
    for a in range(len(Pickin[0])): 
        Cost1 = []
        Cost2 = []
        Costx = []
        lowCostx = []
        for id in range(len(startAGV)): 
            for b in range(len(Sendin[id])):
              
                '''
                this path used for calculate all of the posible way form send point
                to recieve point all of the way that possible
                '''
                ### reset data ###
                if lowCost[0] == b or lowCost[1] == b:
                    continue
                elif lowCost[id] == []:
                    calcost = b
                else:
                    calcost = lowCost[id]
                ### cal the way###
                #print(Sendin[id][lowCost[id]][0], Sendin[id][lowCost[id]][1])
                maze = area
                end = [Pickin[id][b][0], Pickin[id][b][1]]                                # starting position
                start = [Sendin[id][calcost][0], Sendin[id][calcost][1]]                                  # ending position
                # print('start:')
                [cost,direct] = cal(start,end,maze)                                      # Num = Number 0,1,2,... of path to move in array
                costost = Costout[id][b]
                if id == 0:
                    Cost1.append(max(cost)+costost)
                elif id == 1:
                    Cost2.append(max(cost)+costost)
                '''
                # print('1 : {}'.format(Cost1))
                # print('2 : {}'.format(Cost2))
                # print('start number: {}'.format(a))
                # print('car: {}'.format(id))
                # print('cycle times: {}'.format(b))
                '''
         
        print('order : {}'.format(order))    
        print('Cost AGV1: {}'.format(Cost1))
        print('Cost AGV2: {}'.format(Cost2))
        #print(status)
        #print(picknum)
        #print('pre cost1: {}'.format(lowCost[0]))
        #print('pre cost2: {}'.format(lowCost[1]))
        preCost[0] = lowCost[0]
        preCost[1] = lowCost[1]
        #print('past cost1: {}'.format(preCost[0]))
        #print('past cost2: {}'.format(preCost[1]))
        
        if Cost1 == [] and Cost2 == []:
            '''No point to calculate'''
            break
        elif len(Cost1) == 1 or len(Cost2) == 1:
            '''If it has last 1 point'''
            print('case0')
            if order == 0:
                if Cost1 != []:
                    del Pickin[0][0]
                    del Sendin[0][0]
                    del Costout[0][0]
                if Cost2 != []:
                    del Pickin[1][0]
                    del Sendin[1][0]
                    del Costout[1][0]
                del status[0]
                del picknum[0]
                del orderin[0]
                print('type0')
                    
            elif order > 0:
                if preCost[0] != [] and preCost[1] != []:
                    if preCost[0] > preCost[1]:
                    
                        del Pickin[0][preCost[0]]
                        del Sendin[0][preCost[0]]
                        del Pickin[1][preCost[0]]
                        del Sendin[1][preCost[0]]
                        del Costout[0][preCost[0]]
                        del Costout[1][preCost[0]]
                        del orderin[preCost[0]]
                        del status[preCost[0]]
                        del picknum[preCost[0]]
                    
                        del Pickin[0][preCost[1]]
                        del Sendin[0][preCost[1]]
                        del Pickin[1][preCost[1]]
                        del Sendin[1][preCost[1]]
                        del Costout[0][preCost[1]]
                        del Costout[1][preCost[1]]
                        del orderin[preCost[1]]
                        del status[preCost[1]]
                        del picknum[preCost[1]]
                    
                        print('type1')
                    
                    elif preCost[0] < preCost[1]:
                    
                        del Pickin[0][preCost[1]]
                        del Sendin[0][preCost[1]]
                        del Pickin[1][preCost[1]]
                        del Sendin[1][preCost[1]]
                        del Costout[0][preCost[1]]
                        del Costout[1][preCost[1]]
                        del orderin[preCost[1]]
                        del status[preCost[1]]
                        del picknum[preCost[1]]
                    
                        del Pickin[0][preCost[0]]
                        del Sendin[0][preCost[0]]
                        del Pickin[1][preCost[0]]
                        del Sendin[1][preCost[0]]
                        del Costout[0][preCost[0]]
                        del Costout[1][preCost[0]]
                        del orderin[preCost[0]]
                        del status[preCost[0]]
                        del picknum[preCost[0]]
            
                        print('type2')            
                elif preCost[0] != [] and preCost[1] == []: 
                    del Pickin[0][preCost[0]]
                    del Sendin[0][preCost[0]]
                    del Pickin[1][preCost[0]]
                    del Sendin[1][preCost[0]]
                    del Costout[0][preCost[0]]
                    del Costout[1][preCost[0]]
                    del orderin[preCost[0]]
                    del status[preCost[0]]
                    del picknum[preCost[0]]
                    
                elif preCost[0] == [] and preCost[1] != []: 
                    del Pickin[0][preCost[1]]
                    del Sendin[0][preCost[1]]
                    del Pickin[1][preCost[1]]
                    del Sendin[1][preCost[1]]
                    del Costout[0][preCost[1]]
                    del Costout[1][preCost[1]]
                    del orderin[preCost[1]]
                    del status[preCost[1]]
                    del picknum[preCost[1]]
                    
            if Cost1 != [] and Cost2 != []:

                if  order > 0 and listAGV1[5] != [] and listAGV1[5][-1] > min(Cost2):
                    
                    lowCost[1] = Cost2.index(min(Cost2))
                        
                    listAGV2[0].append(order)
                    listAGV2[1].append(Pickin[1][lowCost[1]][0])
                    listAGV2[2].append(Pickin[1][lowCost[1]][1])
                    listAGV2[3].append(Sendin[1][lowCost[1]][0])
                    listAGV2[4].append(Sendin[1][lowCost[1]][1])
                    listAGV2[5].append(Costout[1][lowCost[1]])
                    listAGV2[6].append(status[lowCost[1]])
                    listAGV2[7].append(picknum[lowCost[1]])
                    
                elif order > 0 and listAGV2[5] != [] and listAGV2[5][-1] > min(Cost1):
                    
                    lowCost[0] = Cost1.index(min(Cost1))
                       
                    listAGV1[0].append(order)
                    listAGV1[1].append(Pickin[0][lowCost[0]][0])
                    listAGV1[2].append(Pickin[0][lowCost[0]][1])
                    listAGV1[3].append(Sendin[0][lowCost[0]][0])
                    listAGV1[4].append(Sendin[0][lowCost[0]][1])
                    listAGV1[5].append(Costout[0][lowCost[0]])
                    listAGV1[6].append(status[lowCost[0]])
                    listAGV1[7].append(picknum[lowCost[0]])
                    
                else:
                    if Cost1 > Cost2:
                            
                        lowCost[1] = Cost2.index(min(Cost2))
                            
                        listAGV2[0].append(order)
                        listAGV2[1].append(Pickin[1][lowCost[1]][0])
                        listAGV2[2].append(Pickin[1][lowCost[1]][1])
                        listAGV2[3].append(Sendin[1][lowCost[1]][0])
                        listAGV2[4].append(Sendin[1][lowCost[1]][1])
                        listAGV2[5].append(Costout[1][lowCost[1]])
                        listAGV2[6].append(status[lowCost[1]])
                        listAGV2[7].append(picknum[lowCost[1]])
                    
                    elif Cost1 <= Cost2:
                    
                        lowCost[0] = Cost1.index(min(Cost1))
                        
                        listAGV1[0].append(order)
                        listAGV1[1].append(Pickin[0][lowCost[0]][0])
                        listAGV1[2].append(Pickin[0][lowCost[0]][1])
                        listAGV1[3].append(Sendin[0][lowCost[0]][0])
                        listAGV1[4].append(Sendin[0][lowCost[0]][1])
                        listAGV1[5].append(Costout[0][lowCost[0]])
                        listAGV1[6].append(status[lowCost[0]])
                        listAGV1[7].append(picknum[lowCost[0]])
                    
            elif Cost2 != [] and Cost1 == []:
                lowCost[0] = Cost1.index(min(Cost1))
                        
                listAGV1[0].append(order)
                listAGV1[1].append(Pickin[0][lowCost[0]][0])
                listAGV1[2].append(Pickin[0][lowCost[0]][1])
                listAGV1[3].append(Sendin[0][lowCost[0]][0])
                listAGV1[4].append(Sendin[0][lowCost[0]][1])
                listAGV1[5].append(Costout[0][lowCost[0]])
                listAGV1[6].append(status[lowCost[0]])
                listAGV1[7].append(picknum[lowCost[0]])
                
            elif Cost2 == [] and Cost1 != []:
                lowCost[1] = Cost2.index(min(Cost2))
                        
                listAGV2[0].append(order)
                listAGV2[1].append(Pickin[1][lowCost[1]][0])
                listAGV2[2].append(Pickin[1][lowCost[1]][1])
                listAGV2[3].append(Sendin[1][lowCost[1]][0])
                listAGV2[4].append(Sendin[1][lowCost[1]][1])
                listAGV2[5].append(Costout[1][lowCost[1]])
                listAGV2[6].append(status[lowCost[1]])
                listAGV2[7].append(picknum[lowCost[1]])

            #if lowCost[0] != []:
             #   print('output sort AGV1:({},{}) to ({},{})'.format(Pickin[0][lowCost[0]][0],Pickin[0][lowCost[0]][1],
             #         Sendin[0][lowCost[0]][0],Sendin[0][lowCost[0]][1]))
            print('remain AGV1:')
            print(Pickin[0])
            print(Sendin[0])
            #if lowCost[1] != []:
            #    print('output sort AGV2:({},{}) to ({},{})'.format(Pickin[1][lowCost[1]][0],Pickin[1][lowCost[1]][1],
            #          Sendin[1][lowCost[1]][0],Sendin[1][lowCost[1]][1]))
            print('remain AGV2:')
            print(Pickin[1])
            print(Sendin[1])
            order = order + 1
            print('____________________________________________________')

        else:
            '''If it have more than 2 point to calculate'''
            print('case1')
            if len(Sendin[0]) == 2 or len(Sendin[1]) == 2:
                if preCost[0] != []:
                    lowCost[0] = Cost1.index(min(Cost1))
                    del Pickin[0][preCost[0]]
                    del Sendin[0][preCost[0]]
                    del Pickin[1][preCost[0]]
                    del Sendin[1][preCost[0]]
                    del status[preCost[0]]
                    del picknum[preCost[0]]
                    print('end AGV1:({},{}) to ({},{})'.format(Pickin[0][preCost[0]][0],Pickin[0][preCost[0]][1],
                          Sendin[0][preCost[0]][0],Sendin[0][preCost[0]][1]))
                    listAGV1[0].append(order)
                    listAGV1[1].append(Pickin[0][lowCost[0]][0])
                    listAGV1[2].append(Pickin[0][lowCost[0]][1])
                    listAGV1[3].append(Sendin[0][lowCost[0]][0])
                    listAGV1[4].append(Sendin[0][lowCost[0]][1])
                    listAGV1[5].append(Costout[0][lowCost[0]])
                    listAGV1[6].append(status[lowCost[0]])
                    listAGV1[7].append(picknum[lowCost[0]])
                if preCost[1] != []:
                    lowCost[1] = Cost2.index(min(Cost2))
                    del Pickin[0][preCost[1]]
                    del Sendin[0][preCost[1]]
                    del Pickin[1][preCost[1]]
                    del Sendin[1][preCost[1]]
                    del status[preCost[1]]
                    del picknum[preCost[1]]
                    print('end AGV2:({},{}) to ({},{})'.format(Pickin[1][preCost[1]][0],Pickin[1][preCost[1]][1],
                          Sendin[1][preCost[1]][0],Sendin[1][preCost[1]][1]))
                    listAGV1[0].append(order)
                    listAGV1[1].append(Pickin[0][lowCost[1]][0])
                    listAGV1[2].append(Pickin[0][lowCost[1]][1])
                    listAGV1[3].append(Sendin[0][lowCost[1]][0])
                    listAGV1[4].append(Sendin[0][lowCost[1]][1])
                    listAGV1[5].append(Costout[1][lowCost[1]])
                    listAGV1[6].append(status[lowCost[1]])
                    listAGV1[7].append(picknum[lowCost[1]])
                    
                print('__________________________________________________________')
                break
            else:
                print('input AGV1:({},{}) to ({},{})'.format(Pickin[0][preCost[0]][0]
                    ,Pickin[0][preCost[0]][1],Sendin[0][preCost[0]][0],Sendin[0][preCost[0]][1]))
                print('input AGV2:({},{}) to ({},{})'.format(Pickin[1][preCost[1]][0]
                    ,Pickin[1][preCost[1]][1],Sendin[1][preCost[1]][0],Sendin[1][preCost[1]][1]))
                #print(orderin[preCost[0]],orderin[preCost[0]+1],orderin[preCost[0]-1])
                #print(orderin[preCost[1]],orderin[preCost[1]+1],orderin[preCost[1]-1])
                print(orderin)
                lowCost[0] = Cost1.index(min(Cost1))
                lowCost[1] = Cost2.index(min(Cost2))
                
                print('prelow1 : {}'.format(lowCost[0]))
                print('prelow2 : {}'.format(lowCost[1]))
            
                if lowCost[0] == lowCost[1] :
                    if min(Cost1) <= min(Cost2):
                    
                        Costx = Cost2.copy()
                        Costx.sort()
                        if len(Costx) != 1 :
                            if Costx[0] != Costx[1]:
                                lowCost[1] = Cost2.index(Costx[1])
                                #print('here1{}'.format(order))
                            elif Costx[0] == Costx[1]:
                                Cost2.pop(lowCost[1])
                                lowCostx = Cost2.index(min(Cost2))
                                lowCostx = lowCostx + 1
                                lowCost[1] = lowCostx
                                #print('here2{}'.format(order))                    
                    elif min(Cost1) > min(Cost2): 
                    
                        Costx = Cost1.copy()
                        Costx.sort()
                        if len(Costx) != 1:
                            if Costx[0] != Costx[1]:
                                lowCost[0] = Cost1.index(Costx[1])
                                #print('here3{}'.format(order))
                            elif Costx[0] == Costx[1]:
                                Cost1.pop(lowCost[0])
                                lowCostx = Cost1.index(min(Cost1))
                                lowCostx = lowCostx + 1
                                lowCost[0] = lowCostx
                                #print('here4{}'.format(order))  
                        
                        
                print('low1 : {}'.format(lowCost[0]))
                print('low2 : {}'.format(lowCost[1]))
                if order == 0:  
                    #print('order 1st : {},{}'.format(lowCost[0],lowCost[1]))
                    if orderin[lowCost[0]] == orderin[lowCost[1]] and max(orderin) == orderin[lowCost[0]]:
                        if Cost1[lowCost[0]] < Cost2[lowCost[1]]:
                            if lowCost[0] != lowCost[1] and Cost1[lowCost[1]] > Cost2[lowCost[1]]:
                                lowCost[0] = lowCost[0]
                                lowCost[1] = lowCost[1]
                            else:
                                lowCost[0] = lowCost[0]
                                if orderin[lowCost[0]] == orderin[lowCost[0]-1] and lowCost[0] != 0:
                                    if orderin[lowCost[0]-1] == orderin[lowCost[0]-2] and lowCost[0]-1 != 0:
                                        if orderin[lowCost[0]-2] == orderin[lowCost[0]-3] and lowCost[0]-2 != 0: 
                                            print('back3')
                                            lowCost[0] = lowCost[0] - 3
                                            preorder[0] = orderin[lowCost[0]]
                                        else:
                                            print('back2')
                                            lowCost[0] = lowCost[0] - 2
                                            preorder[0] = orderin[lowCost[0]]
                                    else:
                                        print('back1')
                                        lowCost[0] = lowCost[0] - 1
                                        preorder[0] = orderin[lowCost[0]]
                                elif orderin[lowCost[0]] != orderin[lowCost[0]-1]:
                                    print('no back')
                                    lowCost[0] = lowCost[0]
                                    preorder[0] = []
                            
                                lowCost[1] = []
                            
                        elif Cost1[lowCost[0]] > Cost2[lowCost[1]]:
                            if lowCost[0] != lowCost[1] and Cost2[lowCost[0]] > Cost1[lowCost[0]]:
                                lowCost[0] = lowCost[0]
                                lowCost[1] = lowCost[1]
                            else:
                                lowCost[1] = lowCost[1]
                                if orderin[lowCost[1]] == orderin[lowCost[1]-1] and lowCost[1] != 0:
                                    if orderin[lowCost[1]-1] == orderin[lowCost[1]-2] and lowCost[1]-1 != 0:
                                        if orderin[lowCost[1]-2] == orderin[lowCost[1]-3] and lowCost[1]-2 != 0: 
                                            print('back3')
                                            lowCost[1] = lowCost[1] - 3
                                            preorder[1] = orderin[lowCost[1]]
                                        else:
                                            print('back2')
                                            lowCost[1] = lowCost[1] - 2
                                            preorder[1] = orderin[lowCost[1]]
                                    else:
                                        print('back1')
                                        lowCost[1] = lowCost[1] - 1
                                        preorder[1] = orderin[lowCost[1]]
                                elif orderin[lowCost[1]] != orderin[lowCost[1]-1]:
                                    print('no back')
                                    lowCost[1] = lowCost[1]
                                    preorder[1] = []
                            
                                lowCost[0] = []
                            
                    elif orderin[lowCost[0]] == orderin[lowCost[1]] and max(orderin) != orderin[lowCost[0]]:
                        if Cost1[lowCost[0]] < Cost2[lowCost[1]]:
                            
                            lowCost[0] = lowCost[0]
                            if orderin[lowCost[0]] == orderin[lowCost[0]-1] and lowCost[0] != 0:
                                if orderin[lowCost[0]-1] == orderin[lowCost[0]-2] and lowCost[0]-1 != 0:
                                    if orderin[lowCost[0]-2] == orderin[lowCost[0]-3] and lowCost[0]-2 != 0: 
                                        print('back3')
                                        lowCost[0] = lowCost[0] - 3
                                        preorder[0] = orderin[lowCost[0]]
                                    else:
                                        print('back2')
                                        lowCost[0] = lowCost[0] - 2
                                        preorder[0] = orderin[lowCost[0]]
                                else:
                                    print('back1')
                                    lowCost[0] = lowCost[0] - 1
                                    preorder[0] = orderin[lowCost[0]]
                            elif orderin[lowCost[0]] != orderin[lowCost[0]-1]:
                                print('no back')
                                lowCost[0] = lowCost[0]
                                preorder[0] = []
                            
                            lowCost[1] = orderin.index(orderin[lowCost[1]]+1)
                            
                        elif Cost1[lowCost[0]] > Cost2[lowCost[1]]:
                            
                            lowCost[1] = lowCost[1]
                            if orderin[lowCost[1]] == orderin[lowCost[1]-1] and lowCost[1] != 0:
                                if orderin[lowCost[1]-1] == orderin[lowCost[1]-2] and lowCost[1]-1 != 0:
                                    if orderin[lowCost[1]-2] == orderin[lowCost[1]-3] and lowCost[1]-2 != 0: 
                                        print('back3')
                                        lowCost[1] = lowCost[1] - 3
                                        preorder[1] = orderin[lowCost[1]]
                                    else:
                                        print('back2')
                                        lowCost[1] = lowCost[1] - 2
                                        preorder[1] = orderin[lowCost[1]]
                                else:
                                    print('back1')
                                    lowCost[1] = lowCost[1] - 1
                                    preorder[1] = orderin[lowCost[1]]
                            elif orderin[lowCost[1]] != orderin[lowCost[1]-1]:
                                print('no back')
                                lowCost[1] = lowCost[1]
                                preorder[1] = []
                            
                            lowCost[0] = orderin.index(orderin[lowCost[0]]+1)
                            
                    elif orderin[lowCost[0]] != orderin[lowCost[0]]:
                        
                        if orderin[lowCost[0]] == orderin[lowCost[0]-1] and lowCost[0] != 0:
                            if orderin[lowCost[0]-1] == orderin[lowCost[0]-2] and lowCost[0]-1 != 0:
                                if orderin[lowCost[0]-2] == orderin[lowCost[0]-3] and lowCost[0]-2 != 0: 
                                    print('back3')
                                    lowCost[0] = lowCost[0] - 3
                                    preorder[0] = orderin[lowCost[0]]
                                else:
                                    print('back2')
                                    lowCost[0] = lowCost[0] - 2
                                    preorder[0] = orderin[lowCost[0]]
                            else:
                                print('back1')
                                lowCost[0] = lowCost[0] - 1
                                preorder[0] = orderin[lowCost[0]]
                        elif orderin[lowCost[0]] != orderin[lowCost[0]-1]:
                            print('no back')
                            lowCost[0] = lowCost[0]
                            preorder[0] = []
                    
                        if orderin[lowCost[1]] == orderin[lowCost[1]-1] and lowCost[1] != 0:
                            if orderin[lowCost[1]-1] == orderin[lowCost[1]-2] and lowCost[1]-1 != 0:
                                if orderin[lowCost[1]-2] == orderin[lowCost[1]-3] and lowCost[1]-2 != 0: 
                                    print('back3')
                                    lowCost[1] = lowCost[1] - 3
                                    preorder[1] = orderin[lowCost[1]]
                                else:
                                    print('back2')
                                    lowCost[1] = lowCost[1] - 2
                                    preorder[1] = orderin[lowCost[1]]
                            else:
                                print('back1')
                                lowCost[1] = lowCost[1] - 1
                                preorder[1] = orderin[lowCost[1]]
                        elif orderin[lowCost[1]] != orderin[lowCost[1]-1]:
                            print('no back')
                            lowCost[1] = lowCost[1]
                            preorder[1] = []
                else:
                    #print('order : {},{}'.format(lowCost[0],lowCost[1]))
                    if orderin[lowCost[0]] == orderin[lowCost[1]] and max(orderin) == orderin[lowCost[0]]:
                        if Cost1[lowCost[0]] < Cost2[lowCost[1]]:
                            
                            lowCost[0] = lowCost[0]
                            if preorder[0] != []:
                                del orderin[preCost[0]]
                                if preCost[0] > preCost[1]:
                                    passcost[0] = orderin.index(preorder[0])-1
                                elif preCost[0] < preCost[1]:
                                    passcost[0] = orderin.index(preorder[0])
                                if orderin[passcost[0]] == orderin[passcost[0]+1] and passcost[0] < len(orderin)-1:
                                    lowCost[0] = passcost[0]
                                    preorder[0] = orderin[lowCost[0]]
                                else:
                                    lowCost[0] = passcost[0]
                                    preorder[0] = []
                            else:
                                del orderin[preCost[0]]
                                if orderin[lowCost[0]] == orderin[lowCost[0]+1] and passcost[0] < len(orderin)-1:
                                    lowCost[0] = lowCost[0]
                                    preorder[0] = orderin[lowCost[0]]
                                else:
                                    lowCost[0] = lowCost[0]
                                    preorder[0] = []
                                    
                            lowCost[1] = []
                            
                        elif Cost1[lowCost[0]] > Cost2[lowCost[1]]:
                            
                            lowCost[1] = lowCost[1]
                            if preorder[1] != []:
                                del orderin[preCost[1]]
                                if preCost[1] > preCost[0]:
                                    passcost[1] = orderin.index(preorder[1])-1
                                elif preCost[1] < preCost[0]:
                                    passcost[1] = orderin.index(preorder[1])
                                if orderin[passcost[1]] == orderin[passcost[1]+1] and passcost[1] < len(orderin)-1:
                                    lowCost[1] = passcost[1]
                                    preorder[1] = orderin[lowCost[1]]
                                else:
                                    lowCost[1] = passcost[1]
                                    preorder[1] = []
                            else:
                                del orderin[preCost[1]]
                                if orderin[lowCost[1]] == orderin[lowCost[1]+1] and passcost[1] < len(orderin)-1:
                                    lowCost[1] = passcost[1]
                                    preorder[1] = orderin[lowCost[1]]
                                else:
                                    lowCost[1] = lowCost[1]
                                    preorder[1] = []
                                    
                            lowCost[0] = []
                            
                    elif orderin[lowCost[0]] == orderin[lowCost[1]] and max(orderin) != orderin[lowCost[0]]:
                        if Cost1[lowCost[0]] < Cost2[lowCost[1]]:
                            
                            lowCost[0] = lowCost[0]
                            if preorder[0] != []:
                                del orderin[preCost[0]]
                                if preCost[0] > preCost[1]:
                                    passcost[0] = orderin.index(preorder[0])-1
                                elif preCost[0] < preCost[1]:
                                    passcost[0] = orderin.index(preorder[0])
                                if orderin[passcost[0]] == orderin[passcost[0]+1] and passcost[0] < len(orderin)-1:
                                    lowCost[0] = passcost[0]
                                    preorder[0] = orderin[lowCost[0]]
                                else:
                                    lowCost[0] = passcost[0]
                                    preorder[0] = []
                            else:
                                del orderin[preCost[0]]
                                if orderin[lowCost[0]] == orderin[lowCost[0]+1] and passcost[0] < len(orderin)-1:
                                    lowCost[0] = lowCost[0]
                                    preorder[0] = orderin[lowCost[0]]
                                else:
                                    lowCost[0] = lowCost[0]
                                    preorder[0] = []
                                
                            lowCost[1] = orderin.index(lowCost[1]+1)
                            
                        elif Cost1[lowCost[0]] > Cost2[lowCost[1]]:
                            lowCost[1] = lowCost[1]
                            if preorder[1] != []:
                                del orderin[preCost[1]]
                                if preCost[1] > preCost[0]:
                                    passcost[1] = orderin.index(preorder[1])-1
                                elif preCost[1] < preCost[0]:
                                    passcost[1] = orderin.index(preorder[1])
                                if orderin[passcost[1]] == orderin[passcost[1]+1] and passcost[1] < len(orderin)-1:
                                    lowCost[1] = passcost[1]
                                    preorder[1] = orderin[lowCost[1]]
                                else:
                                    lowCost[1] = passcost[1]
                                    preorder[1] = []
                            else:
                                del orderin[preCost[1]]
                                if orderin[lowCost[1]] == orderin[lowCost[1]+1] and passcost[1] < len(orderin)-1:
                                    lowCost[1] = passcost[1]
                                    preorder[1] = orderin[lowCost[1]]
                                else:
                                    lowCost[1] = lowCost[1]
                                    preorder[1] = []
                            
                            lowCost[0] = orderin.index(lowCost[0]+1)
                            
                    elif orderin[lowCost[0]] != orderin[lowCost[1]]:
                        
                        if preorder[0] != []:
                            del orderin[preCost[0]]
                            if preCost[0] > preCost[1]:
                                passcost[0] = orderin.index(preorder[0])-1
                            elif preCost[0] < preCost[1]:
                                passcost[0] = orderin.index(preorder[0])
                            if orderin[passcost[0]] == orderin[passcost[0]+1] and passcost[0] < len(orderin)-1:
                                lowCost[0] = passcost[0]
                                preorder[0] = orderin[lowCost[0]]
                            else:
                                lowCost[0] = passcost[0]
                                preorder[0] = []
                        else:
                            del orderin[preCost[0]]
                            if orderin[lowCost[0]] == orderin[lowCost[0]+1] and passcost[0] < len(orderin)-1:
                                lowCost[0] = lowCost[0]
                                preorder[0] = orderin[lowCost[0]]
                            else:
                                lowCost[0] = lowCost[0]
                                preorder[0] = []
                    
                        if preorder[1] != []:
                            del orderin[preCost[1]]
                            if preCost[1] > preCost[0]:
                                passcost[1] = orderin.index(preorder[1])-1
                            elif preCost[1] < preCost[0]:
                                passcost[1] = orderin.index(preorder[1])
                            if orderin[passcost[1]] == orderin[passcost[1]+1] and passcost[1] < len(orderin)-1:
                                lowCost[1] = passcost[1]
                                preorder[1] = orderin[lowCost[1]]
                            else:
                                lowCost[1] = passcost[1]
                                preorder[1] = []
                        else:
                            del orderin[preCost[1]]
                            if orderin[lowCost[1]] == orderin[lowCost[1]+1] and passcost[1] < len(orderin)-1:
                                lowCost[1] = passcost[1]
                                preorder[1] = orderin[lowCost[1]]
                            else:
                                lowCost[1] = lowCost[1]
                                preorder[1] = []
                    
                    
                print('firstlow1 : {}'.format(lowCost[0]))
                print('firstlow2 : {}'.format(lowCost[1]))
                
                if order == 0:
                    del Pickin[0][0]
                    del Sendin[0][0]
                    del Pickin[1][0]
                    del Sendin[1][0]
                    del status[0]
                    del picknum[0]
                    #del orderin[0]
                    print('type0')
                elif order > 0:
                    if preCost[0] != [] and preCost[1] != []:
                        if preCost[0] > preCost[1]:
                            
                            del Pickin[0][preCost[0]]
                            del Sendin[0][preCost[0]]
                            del Pickin[1][preCost[0]]
                            del Sendin[1][preCost[0]]
                            del status[preCost[0]]
                            del picknum[preCost[0]]
                        
                            del Pickin[0][preCost[1]]
                            del Sendin[0][preCost[1]]
                            del Pickin[1][preCost[1]]
                            del Sendin[1][preCost[1]]
                            del status[preCost[1]]
                            del picknum[preCost[1]]
                        
                            print('type1')
                    
                        elif preCost[0] < preCost[1]:
                    
                            del Pickin[0][preCost[1]]
                            del Sendin[0][preCost[1]]
                            del Pickin[1][preCost[1]]
                            del Sendin[1][preCost[1]]
                            del status[preCost[1]]
                            del picknum[preCost[1]]
                        
                            del Pickin[0][preCost[0]]
                            del Sendin[0][preCost[0]]
                            del Pickin[1][preCost[0]]
                            del Sendin[1][preCost[0]]
                            del status[preCost[0]]
                            del picknum[preCost[0]]
            
                            print('type2')
                    elif preCost[0] != [] and preCost[1] == []:
                        
                        del Pickin[0][preCost[0]]
                        del Sendin[0][preCost[0]]
                        del Pickin[1][preCost[0]]
                        del Sendin[1][preCost[0]]
                        del status[preCost[0]]
                        del picknum[preCost[0]]
                        
                    elif preCost[0] == [] and preCost[1] != []: 
                        
                        del Pickin[0][preCost[1]]
                        del Sendin[0][preCost[1]]
                        del Pickin[1][preCost[1]]
                        del Sendin[1][preCost[1]]
                        del status[preCost[1]]
                        del picknum[preCost[1]]
                        
                    #lowCost[0] = lowCost[0]
                    #lowCost[1] = lowCost[1]
            
                # print(lowCost)
                if lowCost[0] != []:
                    listAGV1[0].append(order)
                    listAGV1[1].append(Pickin[0][lowCost[0]][0])
                    listAGV1[2].append(Pickin[0][lowCost[0]][1])
                    listAGV1[3].append(Sendin[0][lowCost[0]][0])
                    listAGV1[4].append(Sendin[0][lowCost[0]][1])
                    listAGV1[5].append(Costout[0][lowCost[0]])
                    listAGV1[6].append(status[lowCost[0]])
                    listAGV1[7].append(picknum[lowCost[0]])
                    
                if lowCost[1] != []:
                    listAGV2[0].append(order)
                    listAGV2[1].append(Pickin[1][lowCost[1]][0])
                    listAGV2[2].append(Pickin[1][lowCost[1]][1])
                    listAGV2[3].append(Sendin[1][lowCost[1]][0])
                    listAGV2[4].append(Sendin[1][lowCost[1]][1])
                    listAGV2[5].append(Costout[1][lowCost[1]])
                    listAGV2[6].append(status[lowCost[1]])
                    listAGV2[7].append(picknum[lowCost[1]])
            
                print(orderin)
                if lowCost[0] != []:
                    print('output sort AGV1:({},{}) to ({},{})'.format(Pickin[0][lowCost[0]][0],Pickin[0][lowCost[0]][1],
                          Sendin[0][lowCost[0]][0],Sendin[0][lowCost[0]][1]))
                print('remain AGV1:')
                print(Pickin[0])
                print(Sendin[0])
                if lowCost[1] != []:
                    print('output sort AGV2:({},{}) to ({},{})'.format(Pickin[1][lowCost[1]][0],Pickin[1][lowCost[1]][1],
                          Sendin[1][lowCost[1]][0],Sendin[1][lowCost[1]][1]))
                print('remain AGV2:')
                print(Pickin[1])
                print(Sendin[1])
                order = order + 1
                print('__________________________________________________________________________')
                                #  Fourth column save data 3
    
    
    for x in range(len(listAGV1[0])):
        shtAGV1.range('A{0}'.format(rowAGV1+x+2)).value = listAGV1[0][x]       #  First column save order list
        shtAGV1.range('B{0}'.format(rowAGV1+x+2)).value = listAGV1[1][x]       #  Second  column save Data 1
        shtAGV1.range('C{0}'.format(rowAGV1+x+2)).value = listAGV1[2][x]       #  Third column save Data 2
        shtAGV1.range('D{0}'.format(rowAGV1+x+2)).value = listAGV1[3][x]       #  Fourth column save Data 3
        shtAGV1.range('E{0}'.format(rowAGV1+x+2)).value = listAGV1[4][x]       #  Third column save Data 4
        shtAGV1.range('F{0}'.format(rowAGV1+x+2)).value = listAGV1[5][x]       #  Fourth column save Data 5
        shtAGV1.range('G{0}'.format(rowAGV1+x+2)).value = listAGV1[6][x]       #  Third column save Data 4
        shtAGV1.range('H{0}'.format(rowAGV1+x+2)).value = listAGV1[7][x]       #  Fourth column save Data 5
    
    for x in range(len(listAGV2[0])):
        shtAGV2.range('A{0}'.format(rowAGV2+x+2)).value = listAGV2[0][x]       #  First column save order list
        shtAGV2.range('B{0}'.format(rowAGV2+x+2)).value = listAGV2[1][x]       #  Second  column save Data 1
        shtAGV2.range('C{0}'.format(rowAGV2+x+2)).value = listAGV2[2][x]       #  Third column save Data 2
        shtAGV2.range('D{0}'.format(rowAGV2+x+2)).value = listAGV2[3][x]       #  Fourth column save Data 3
        shtAGV2.range('E{0}'.format(rowAGV2+x+2)).value = listAGV2[4][x]       #  Third column save Data 4
        shtAGV2.range('F{0}'.format(rowAGV2+x+2)).value = listAGV2[5][x]       #  Fourth column save Data 5
        shtAGV2.range('G{0}'.format(rowAGV2+x+2)).value = listAGV2[6][x]       #  Third column save Data 4
        shtAGV2.range('H{0}'.format(rowAGV2+x+2)).value = listAGV2[7][x]       #  Fourth column save Data 5

    print('end')
    return listAGV1,listAGV2
###############################################################################
''' 3 '''   
def pointcon(area,file,inputsheet,outputsheet):
    '''Condition to find same pick point and send point'''
    wb = xw.Book(file)
    
    order = []
    pick = []
    send = []
    cost = []
    picknum = []    
    numberquest = lastRow(inputsheet,wb)  # Find last row of data                                                                                                 #  Select sheet to save data
    getsheet = wb.sheets[inputsheet]        # Open sheet of excel
    for i in range (2,numberquest+1):
        order.append(int(getsheet.range('A{0}'.format(i)).value))
        pick.append([int(getsheet.range('B{0}'.format(i)).value)
            ,int(getsheet.range('C{0}'.format(i)).value)])
        send.append([int(getsheet.range('D{0}'.format(i)).value)
            ,int(getsheet.range('E{0}'.format(i)).value)])
        cost.append(int(getsheet.range('F{0}'.format(i)).value))    
        picknum.append(int(getsheet.range('G{0}'.format(i)).value))    
    '''for find the same point and connection'''    
    copypick = pick.copy()
    copysend = send.copy()
    copycost = cost.copy()
    
    used = []
    count = 0
    ordercount = 1
    
    listpoint = [[],[],[],[],[],[],[],[]]          # order \ startx \ starty \ stopx \ stopy \ cost \ order
    
    print('pick :')
    print(pick)
    print('send :')
    print(send)
    if len(order) != 1:
        for i in range(len(pick)):
            for j in range(len(copypick)):
                if i < j and j not in used and i not in used:
                    # print(used)
                    # print(len(copypick)-len(used))
                    print('i:{} ({},{}),j:{} ({},{})'.format(i,pick[i],send[i],j,copypick[j],copysend[j]))
                    if send[i] == copypick[j] or pick[i] == copysend[j]:
                        '''find quest that are connected'''
                        print('quest connected')
                        if send[i] == copypick[j] and pick[i] != copysend[j]:
                            print('i quest connected j quest')
                            '''First quest'''
                            listpoint[0].append(ordercount)
                            listpoint[1].append(pick[i])
                            listpoint[2].append(send[i])
                            listpoint[3].append(cost[i])
                            listpoint[4].append('pick2send')
                            listpoint[5].append(count)
                            listpoint[6].append(picknum[i])
                            ordercount = ordercount + 1
                            '''Second quest'''
                            listpoint[0].append(ordercount)
                            listpoint[1].append(copypick[j])
                            listpoint[2].append(copysend[j])
                            listpoint[3].append(copycost[j])
                            listpoint[4].append('pick2send')
                            listpoint[5].append(count)
                            listpoint[6].append(picknum[i])
                            ordercount = ordercount + 1
                            
                        elif pick[i] == copysend[j] and send[i] != copypick[j]:
                            print('j quest connected i quest')
                            '''First quest'''
                            listpoint[0].append(ordercount)
                            listpoint[1].append(copypick[j])
                            listpoint[2].append(copysend[j])
                            listpoint[3].append(copycost[j])
                            listpoint[4].append('pick2send')
                            listpoint[5].append(count)
                            listpoint[6].append(picknum[i])
                            ordercount = ordercount + 1
                            '''Second quest'''
                            listpoint[0].append(ordercount)
                            listpoint[1].append(pick[i])
                            listpoint[2].append(send[i])
                            listpoint[3].append(cost[i])
                            listpoint[4].append('pick2send')
                            listpoint[5].append(count)
                            listpoint[6].append(picknum[i])
                            ordercount = ordercount + 1
                        elif pick[i] == copysend[j] and send[i] == copypick[j]:
                            print('two quest reverse')
                            '''First quest'''
                            listpoint[0].append(ordercount)
                            listpoint[1].append(pick[i])
                            listpoint[2].append(send[i])
                            listpoint[3].append(cost[i])
                            listpoint[4].append('connected')
                            listpoint[5].append(count)
                            listpoint[6].append(picknum[i])
                            ordercount = ordercount + 1
                            '''Second quest'''
                            listpoint[0].append(ordercount)
                            listpoint[1].append(copypick[j])
                            listpoint[2].append(copysend[j])
                            listpoint[3].append(copycost[j])
                            listpoint[4].append('connected')
                            listpoint[5].append(count)
                            listpoint[6].append(picknum[i])
                            ordercount = ordercount + 1
                            
                        count = count + 1
                        used.append(i)
                        used.append(j)
                    else:    
                        if send[i] == copysend[j] and pick[i] != copypick[j]:
                            '''when have same send point in different quest'''
                            [cost,directcon]=cal(pick[i],copypick[j],area)
                            costcon = max(cost)
                            print('have point same send point')
                            #print('on point : ({},{}) to ({},{})'.format(pick[i][0],pick[i][1]
                            #    ,send[i][0],send[i][0]))
                            #print('and point : ({},{}) to ({},{})'.format(copypick[j][0],copypick[j][1]
                            #    ,copysend[j][0],copysend[j][1]))
                            if costcon < (cost[i]+copycost[j]):
                                print('change positon to best quality')
                                if cost[i] <= copycost[j]:
                                    '''First quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(copypick[j])
                                    listpoint[2].append(pick[i])
                                    listpoint[3].append(costcon)
                                    listpoint[4].append('pick2pick')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                    '''Second quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(pick[i])
                                    listpoint[2].append(send[i])
                                    listpoint[3].append(cost[i])
                                    listpoint[4].append('pick2send')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                    
                                elif cost[i] > copycost[j]:
                                    '''First quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(pick[i])
                                    listpoint[2].append(copypick[j])
                                    listpoint[3].append(costcon)
                                    listpoint[4].append('pick2pick')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                    '''Second quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(copypick[j])
                                    listpoint[2].append(copysend[j])
                                    listpoint[3].append(copycost[j])
                                    listpoint[4].append('pick2send')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                    
                                count = count + 1
                                used.append(i)
                                used.append(j)
                                
                            elif costcon >= (cost[i]+copycost[j]):
                                print('no change point')
                                '''Not use another quest'''
                                listpoint[0].append(ordercount)
                                listpoint[1].append(pick[i])
                                listpoint[2].append(send[i])
                                listpoint[3].append(cost[i])
                                listpoint[4].append('pick2send')
                                listpoint[5].append(count)
                                listpoint[6].append(picknum[i])
                                ordercount = ordercount + 1
                                count = count + 1
                                used.append(i)

                        elif pick[i] == copypick[j] and send[i] != copysend[j]:
                            '''when have same pick point in different quest'''
                            [cost,directcon]=cal(send[i],copysend[j],area)
                            costcon = max(cost)
                            print('have point same pick point')
                            #print('on point : ({},{}) to ({},{})'.format(pick[i][0],pick[i][1]
                            #    ,send[i][0],send[i][0]))
                            #print('and point : ({},{}) to ({},{})'.format(copypick[j][0],copypick[j][1]
                            #    ,copysend[j][0],copysend[j][1]))
                            if costcon < (cost[i]+copycost[j]):
                                print('change positon to best quality')
                                if cost[i] <= copycost[j]:
                                    '''First quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(pick[i])
                                    listpoint[2].append(send[i])
                                    listpoint[3].append(cost[i])
                                    listpoint[4].append('pick2send')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i]+1)
                                    ordercount = ordercount + 1
                                    '''Second quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(send[i])
                                    listpoint[2].append(copysend[j])
                                    listpoint[3].append(costcon)
                                    listpoint[4].append('send2send')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                    
                                elif cost[i] > copycost[j]:
                                    '''First quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(copypick[j])
                                    listpoint[2].append(copysend[j])
                                    listpoint[3].append(cost[j])
                                    listpoint[4].append('pick2send')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i]+1)
                                    ordercount = ordercount + 1
                                    '''Second quest'''
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(copysend[j])
                                    listpoint[2].append(send[i])
                                    listpoint[3].append(costcon)
                                    listpoint[4].append('send2send')
                                    listpoint[5].append(count)
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                
                                count = count + 1
                                used.append(i)
                                used.append(j)                    
                            
                            elif costcon >= (cost[i]+copycost[j]):
                                print('no change point') 
                                '''Not use another quest'''
                                listpoint[0].append(ordercount)
                                listpoint[1].append(pick[i])
                                listpoint[2].append(send[i])
                                listpoint[3].append(cost[i])
                                listpoint[4].append('pick2send')
                                listpoint[5].append(count)
                                listpoint[6].append(picknum[i])
                                ordercount = ordercount + 1
                                count = count + 1
                                used.append(i)
                        
                        elif send[i] == copysend[j] and pick[i] == copypick[j]:
                            print('have same point same direction')
                            '''Have same either pick point and send point'''
                            listpoint[0].append(ordercount)
                            listpoint[1].append(pick[i])
                            listpoint[2].append(send[i])
                            listpoint[3].append(cost[i])
                            listpoint[4].append('pick2send')
                            listpoint[5].append(count)
                            listpoint[6].append(picknum[i])
                            ordercount = ordercount + 1
                            count = count + 1
                            used.append(i)
                            used.append(j)
                    
                        else:
                            if j+1 >= len(copysend)-len(used) :
                                if i not in used :
                                    '''Not have same neither pick point or send point'''
                                    print('no change for i')
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(pick[i])
                                    listpoint[2].append(send[i])
                                    listpoint[3].append(cost[i])
                                    listpoint[4].append('pick2send')
                                    listpoint[5].append(count) 
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                    count = count + 1
                                    used.append(i)
                                if j not in used :
                                    '''Not have same neither pick point or send point'''
                                    print('no change for j')
                                    listpoint[0].append(ordercount)
                                    listpoint[1].append(pick[j])
                                    listpoint[2].append(send[j])
                                    listpoint[3].append(cost[j])
                                    listpoint[4].append('pick2send')
                                    listpoint[5].append(count) 
                                    listpoint[6].append(picknum[i])
                                    ordercount = ordercount + 1
                                    count = count + 1
                                    used.append(j)
            if i+1 > len(send)-len(used) and i not in used:
                print('last one')
                listpoint[0].append(ordercount)
                listpoint[1].append(pick[i])
                listpoint[2].append(send[i])
                listpoint[3].append(cost[i])
                listpoint[4].append('pick2send')
                listpoint[5].append(count) 
                listpoint[6].append(picknum[i])
                ordercount = ordercount + 1
                count = count + 1
                used.append(i)
    else:  
        print('one quest')
        listpoint[0]=ordercount
        listpoint[1]=pick
        listpoint[2]=send
        listpoint[3]=cost
        listpoint[4]='non'
        listpoint[5]=count
        listpoint[6]=1
        ordercount = ordercount + 1
        count = count + 1

    print(listpoint[0])
    print('new pick :')
    print(listpoint[1][0])
    print('new pick :')
    print(listpoint[2][0])
    print('new cost :')
    print(listpoint[3])
    print(listpoint[4])
    print('__________________________________________________________________')
    
   
    rowout = lastRow(outputsheet,wb)                        # Find last row of data
    shtout = wb.sheets[outputsheet]           
    if len(listpoint[1]) == 1:
        if rowout >= 2:
            shtout.range('A{0}'.format(rowout+2)).value = listpoint[0]                #  First column save order list
            shtout.range('B{0}'.format(rowout+2)).value = listpoint[1][0][0]           #  Second  column save Data 1
            shtout.range('C{0}'.format(rowout+2)).value = listpoint[1][0][1]           #  Third column save Data 2
            shtout.range('D{0}'.format(rowout+2)).value = listpoint[2][0][0]           #  Fourth column save Data 3
            shtout.range('E{0}'.format(rowout+2)).value = listpoint[2][0][1]           #  Third column save Data 4
            shtout.range('F{0}'.format(rowout+2)).value = listpoint[3]              #  Fourth column save Data 5
            shtout.range('G{0}'.format(rowout+2)).value = listpoint[4]                #  Fifth column save Data 5
            shtout.range('H{0}'.format(rowout+2)).value = listpoint[5]            #  Sixth column save Data 5
            shtout.range('I{0}'.format(rowout+2)).value = listpoint[6]            #  Sixth column save Data 5
        else:
            shtout.range('A{0}'.format(rowout+1)).value = listpoint[0]                #  First column save order list
            shtout.range('B{0}'.format(rowout+1)).value = listpoint[1][0][0]           #  Second  column save Data 1
            shtout.range('C{0}'.format(rowout+1)).value = listpoint[1][0][1]           #  Third column save Data 2
            shtout.range('D{0}'.format(rowout+1)).value = listpoint[2][0][0]           #  Fourth column save Data 3
            shtout.range('E{0}'.format(rowout+1)).value = listpoint[2][0][1]           #  Third column save Data 4
            shtout.range('F{0}'.format(rowout+1)).value = listpoint[3]              #  Fourth column save Data 5
            shtout.range('G{0}'.format(rowout+1)).value = listpoint[4]            #  Fifth column save Data 5
            shtout.range('H{0}'.format(rowout+1)).value = listpoint[5]            #  Sixth column save Data 5
            shtout.range('I{0}'.format(rowout+1)).value = listpoint[6]            #  Sixth column save Data 5
    else:
        for x in range(len(listpoint[3])):
            if rowout >= 2:
                shtout.range('A{0}'.format(rowout+x+2)).value = listpoint[0][x]                #  First column save order list
                shtout.range('B{0}'.format(rowout+x+2)).value = listpoint[1][x][0]           #  Second  column save Data 1
                shtout.range('C{0}'.format(rowout+x+2)).value = listpoint[1][x][1]           #  Third column save Data 2
                shtout.range('D{0}'.format(rowout+x+2)).value = listpoint[2][x][0]           #  Fourth column save Data 3
                shtout.range('E{0}'.format(rowout+x+2)).value = listpoint[2][x][1]           #  Third column save Data 4
                shtout.range('F{0}'.format(rowout+x+2)).value = listpoint[3][x]              #  Fourth column save Data 5
                shtout.range('G{0}'.format(rowout+x+2)).value = listpoint[4][x]            #  Fifth column save Data 5
                shtout.range('H{0}'.format(rowout+x+2)).value = listpoint[5][x]            #  Sixth column save Data 5
                shtout.range('I{0}'.format(rowout+x+2)).value = listpoint[6][x]            #  Sixth column save Data 5
            else:
                shtout.range('A{0}'.format(rowout+x+1)).value = listpoint[0][x]                #  First column save order list
                shtout.range('B{0}'.format(rowout+x+1)).value = listpoint[1][x][0]           #  Second  column save Data 1
                shtout.range('C{0}'.format(rowout+x+1)).value = listpoint[1][x][1]           #  Third column save Data 2
                shtout.range('D{0}'.format(rowout+x+1)).value = listpoint[2][x][0]           #  Fourth column save Data 3
                shtout.range('E{0}'.format(rowout+x+1)).value = listpoint[2][x][1]           #  Third column save Data 4
                shtout.range('F{0}'.format(rowout+x+1)).value = listpoint[3][x]              #  Fourth column save Data 5
                shtout.range('G{0}'.format(rowout+x+1)).value = listpoint[4][x]            #  Fifth column save Data 5
                shtout.range('H{0}'.format(rowout+x+1)).value = listpoint[5][x]            #  Sixth column save Data 5
                shtout.range('I{0}'.format(rowout+x+1)).value = listpoint[6][x]            #  Sixth column save Data 5
    return listpoint

###############################################################################
''' 4 '''
def hitchhikingcon(area,file,inputsheet,outputsheet):
    
    '''Condition to find same pick point and send point'''
    wb = xw.Book(file)
    
    order = []
    pick = []
    send = []
    cost = []  
    status = []
    number = []
    picknum = []  
    
    numberquest = lastRow(inputsheet,wb)  # Find last row of data                                                                                                 #  Select sheet to save data
    getsheet = wb.sheets[inputsheet]        # Open sheet of excel
    
    getlist = int(getsheet.range('A{0}'.format(numberquest)).value)
    # print('{},{}'.format(getlist,numberquest))
    
    if getlist == 1:
        order.append(int(getsheet.range('A{0}'.format(numberquest)).value))
        pick.append([int(getsheet.range('B{0}'.format(numberquest)).value)
                ,int(getsheet.range('C{0}'.format(numberquest)).value)])
        send.append([int(getsheet.range('D{0}'.format(numberquest)).value)
                ,int(getsheet.range('E{0}'.format(numberquest)).value)])
        cost.append(int(getsheet.range('F{0}'.format(numberquest)).value))
        status.append(getsheet.range('G{0}'.format(numberquest)).value)
        number.append(int(getsheet.range('H{0}'.format(numberquest)).value))
        picknum.append(int(getsheet.range('I{0}'.format(numberquest)).value))
    else:
        for i in range (numberquest+1-getlist,numberquest+1):
            order.append(int(getsheet.range('A{0}'.format(i)).value))
            pick.append([int(getsheet.range('B{0}'.format(i)).value)
                ,int(getsheet.range('C{0}'.format(i)).value)])
            send.append([int(getsheet.range('D{0}'.format(i)).value)
                ,int(getsheet.range('E{0}'.format(i)).value)])
            cost.append(int(getsheet.range('F{0}'.format(i)).value))    
            status.append(getsheet.range('G{0}'.format(i)).value)
            number.append(int(getsheet.range('H{0}'.format(i)).value)) 
            picknum.append(int(getsheet.range('I{0}'.format(i)).value))
    
    keepdirect = []
    usedalready = []
    ordercounter = 1
    
    listhitch = [[],[],[],[],[],[],[],[]]          # order \ startx \ starty \ stopx \ stopy \ cost \ order
    print(pick,send)
    for i in range(len(pick)):
        #print(i)
        [costt,direction]=cal(pick[i],send[i],area)
        keepdirect.append(direction)
        print(direction)
    copydirect = keepdirect.copy()
    if len(keepdirect) != 1:
        for j in range(len(keepdirect)):
            data = keepdirect[j]
            for l in range(len(copydirect)):
                if j < l and l not in usedalready and j not in usedalready:
                    for k in range(len(data)):
                        outcost = []
                        if copydirect[l][0] == keepdirect[j][k] and k+1 != len(data) and k != 0:
                            print('have pick point on the moving direction')
                            print('{},{}'.format(keepdirect[j][k],copydirect[l][0]))
                            print('{},{},{}'.format(j,l,k))
                            [cost,direction]=cal(keepdirect[j][0],copydirect[l][0],area)
                            costcon = max(cost)
                            outcost.append(costcon) # cost[0] from 1st pick point to 2nd pick point
                            [cost,direction]=cal(copydirect[l][0],keepdirect[j][-1],area)
                            costcon = max(cost)
                            outcost.append(costcon) # cost[1] from 2nd pick point to 1st send point
                            [cost,direction]=cal(keepdirect[j][-1],copydirect[l][-1],area)
                            costcon = max(cost)
                            outcost.append(costcon) # cost[2] from 1st send point to 2nd send point
                            [cost,direction]=cal(copydirect[l][0],copydirect[l][-1],area)
                            costcon = max(cost)
                            outcost.append(costcon) # cost[3] from 2nd pick point to 2nd send point
                            
                            if outcost[0]+outcost[1] >= outcost[2]:
                                if outcost[0]+outcost[1]+outcost[2] <= outcost[0]+outcost[2]+outcost[3]:
                                    #if send 1st quest first shorter than 2nd quest first
                                    print('send first quest first')
                                    '''from pick 1st quest to pick 2nd quest'''
                                    listhitch[0].append(ordercounter)
                                    listhitch[1].append(keepdirect[j][0])
                                    listhitch[2].append(copydirect[l][0])
                                    listhitch[3].append(outcost[0])
                                    listhitch[5].append(number[j])
                                    listhitch[6].append(picknum[j]+2)
                                    ordercounter = ordercounter + 1
                                    '''from pick 2nd quest to send 1st quest'''
                                    listhitch[0].append(ordercounter)
                                    listhitch[1].append(copydirect[l][0])
                                    listhitch[2].append(keepdirect[j][-1])
                                    listhitch[3].append(outcost[1])
                                    listhitch[5].append(number[j])
                                    listhitch[6].append(picknum[j]+1)
                                    ordercounter = ordercounter + 1
                                    '''from send 1st quest to send 2nd quest'''
                                    listhitch[0].append(ordercounter)
                                    listhitch[1].append(keepdirect[j][-1])
                                    listhitch[2].append(copydirect[l][-1])
                                    listhitch[3].append(outcost[2])
                                    listhitch[5].append(number[j])
                                    listhitch[6].append(picknum[j]+2)
                                    ordercounter = ordercounter + 1
                                    
                                    listhitch[4].append('pick2pick')
                                    listhitch[4].append('pick2send')
                                    listhitch[4].append('send2send')
                                    '''
                                    if status[j] == 'non' :
                                        listhitch[4].append('2 quest')
                                        listhitch[4].append('send 1st quest')
                                        listhitch[4].append('2 contains')
                                    elif status[j] == 'connected' :
                                        listhitch[4].append('3 quest')
                                        listhitch[4].append('send 1st quest')
                                        listhitch[4].append('3 contains')
                                    elif status[j] == 'same' :
                                        listhitch[4].append('2 quest')
                                        listhitch[4].append('send 1st quest')
                                        listhitch[4].append('3 contains')
                                    '''
                                        
                                elif outcost[0]+outcost[1]+outcost[2] > outcost[0]+outcost[2]+outcost[3]:
                                    #if send 2nd quest first shorter than 1st quest first
                                    print('send next quest first')
                                    '''from pick 1st quest to pick 2nd quest'''
                                    listhitch[0].append(ordercounter)
                                    listhitch[1].append(keepdirect[j][0])
                                    listhitch[2].append(copydirect[l][0])
                                    listhitch[3].append(outcost[0])
                                    listhitch[5].append(number[j])
                                    listhitch[6].append(picknum[j]+2)
                                    ordercounter = ordercounter + 1
                                    '''from pick 2nd quest to send 2nd quest'''
                                    listhitch[0].append(ordercounter)
                                    listhitch[1].append(copydirect[l][0])
                                    listhitch[2].append(copydirect[l][-1])
                                    listhitch[3].append(outcost[3])
                                    listhitch[5].append(number[j])
                                    listhitch[6].append(picknum[j]+1)
                                    ordercounter = ordercounter + 1
                                    '''from send 2nd quest to send 1st quest'''
                                    listhitch[0].append(ordercounter)
                                    listhitch[1].append(copydirect[l][-1])
                                    listhitch[2].append(keepdirect[j][-1])
                                    listhitch[3].append(outcost[2])
                                    listhitch[5].append(number[j])
                                    listhitch[6].append(picknum[j])
                                    ordercounter = ordercounter + 1
                                    
                                    listhitch[4].append('pick2pick')
                                    listhitch[4].append('pick2send')
                                    listhitch[4].append('send2send')
                                    '''
                                    if status[j] == 'non' :
                                        listhitch[4].append('2 quest')
                                        listhitch[4].append('send 1st quest')
                                        listhitch[4].append('2 contains')
                                    elif status[j] == 'connected' :
                                        listhitch[4].append('3 quest')
                                        listhitch[4].append('send 1st quest')
                                        listhitch[4].append('3 contains')
                                    elif status[j] == 'connected' :
                                        listhitch[4].append('2 quest')
                                        listhitch[4].append('send 1st quest')
                                        listhitch[4].append('3 contains')
                                    '''    
                                usedalready.append(l)
                                usedalready.append(j)
                                        
                            elif outcost[0] < outcost[2]:
                                print('keep more quest is waste more time')
                        
                                listhitch[0].append(ordercounter)
                                listhitch[1].append(keepdirect[j][0])
                                listhitch[2].append(keepdirect[j][-1])
                                listhitch[3].append(cost[j])
                                listhitch[4].append('pick2send')
                                listhitch[5].append(number[j])
                                listhitch[6].append(picknum[j])
                                ordercounter = ordercounter + 1
                                usedalready.append(j)
                                
                        elif copydirect[l][-1] == keepdirect[j][k] and k+1 != len(data) and k != 0:
                            #if send is on the moving direction
                            print('have send point on the moving direction')
                            print('{},{}'.format(keepdirect[j][k],copydirect[l][-1]))
                            print('{},{},{}'.format(j,l,k))
                            [cost,direction]=cal(keepdirect[j][0],copydirect[l][0],area)
                            costcon = max(cost)
                            outcost.append(costcon) # cost[0] from 1st pick point to 2nd pick point
                            [cost,direction]=cal(copydirect[l][0],copydirect[l][-1],area)
                            costcon = max(cost)
                            outcost.append(costcon) # cost[1] from 2nd pick point to 1st send point
                            [cost,direction]=cal(copydirect[l][-1],keepdirect[j][-1],area)
                            costcon = max(cost)
                            outcost.append(costcon) # cost[2] from 2nd pick point to 2nd send point
                            if cost[j] >= outcost[0]+outcost[1]+outcost[2]:
                                # if send 1st is more than or equal send 1st and 2nd together 
                                print('it is possible to keep this quest')
                                '''from pick 1st quest to pick 2nd quest'''
                                listhitch[0].append(ordercounter)
                                listhitch[1].append(keepdirect[j][0])
                                listhitch[2].append(copydirect[l][0])
                                listhitch[3].append(outcost[0])
                                listhitch[5].append(number[j])
                                listhitch[6].append(picknum[j]+2)
                                ordercounter = ordercounter + 1
                                print('send first quest first')
                                listhitch[0].append(ordercounter)
                                listhitch[1].append(copydirect[l][0])
                                listhitch[2].append(copydirect[l][-1])
                                listhitch[3].append(outcost[1])
                                listhitch[5].append(number[j])
                                listhitch[6].append(picknum[j]+1)
                                ordercounter = ordercounter + 1
                                '''from send 1st quest to send 2nd quest'''
                                listhitch[0].append(ordercounter)
                                listhitch[1].append(copydirect[l][-1])
                                listhitch[2].append(keepdirect[j][-1])
                                listhitch[3].append(outcost[2])
                                listhitch[5].append(number[j])
                                listhitch[6].append(picknum[j])
                                ordercounter = ordercounter + 1
                                
                                listhitch[4].append('pick2pick')
                                listhitch[4].append('pick2send')
                                listhitch[4].append('send2send')
                                '''
                                if status[j] == 'non' :
                                    listhitch[4].append('2 quest')
                                    listhitch[4].append('send 1st quest')
                                    listhitch[4].append('2 contains')
                                elif status[j] == 'connected' :
                                    listhitch[4].append('3 quest')
                                    listhitch[4].append('send 1st quest')
                                    listhitch[4].append('3 contains')
                                elif status[j] == 'same' :
                                    listhitch[4].append('2 quest')
                                    listhitch[4].append('send 1st quest')
                                    listhitch[4].append('3 contains') 
                                '''   
                                usedalready.append(l)
                                usedalready.append(j)
                                    
                            elif cost[j] < outcost[0]+outcost[1]+outcost[2]:
                                print('keep more quest is waste the time')
                                listhitch[0].append(ordercounter)
                                listhitch[1].append(keepdirect[j][0])
                                listhitch[2].append(keepdirect[j][-1])
                                listhitch[3].append(cost[j])
                                listhitch[4].append('pick2send')
                                listhitch[5].append(number[j])
                                listhitch[6].append(picknum[j])
                                ordercounter = ordercounter + 1
                                usedalready.append(j)
                                
                    if l+1 == len(copydirect) and j not in usedalready:
                        print('there is no hitchhiking')
                        print('{},{},{}'.format(j,l,k))
                        listhitch[0].append(ordercounter)
                        listhitch[1].append(keepdirect[j][0])
                        listhitch[2].append(keepdirect[j][-1])
                        listhitch[3].append(cost[j])
                        listhitch[4].append(status[j])
                        listhitch[5].append(number[j])
                        listhitch[6].append(picknum[j])
                        ordercounter = ordercounter + 1
                        
                        usedalready.append(j)
            if j+1 == len(keepdirect) and j not in usedalready:
                #ordercounter = order[j]
                print('there is no hitchhiking')
                print('{},{},{}'.format(j,l,k))
                listhitch[0].append(ordercounter)
                listhitch[1].append(keepdirect[j][0])
                listhitch[2].append(keepdirect[j][-1])
                listhitch[3].append(cost[j])
                listhitch[4].append(status[j])
                listhitch[5].append(number[j])
                listhitch[6].append(picknum[j])
                ordercounter = ordercounter + 1
                usedalready.append(j)
                        
                # 'hitchhiking'
    else:
        listhitch[0].append(ordercounter)
        listhitch[1].append(keepdirect[0][0])
        listhitch[2].append(keepdirect[0][-1])
        listhitch[3].append(cost)
        listhitch[4].append('pick2send')
        listhitch[5].append(number[0])
        listhitch[6].append(picknum[0])
        ordercounter = ordercounter + 1       
        
    print(listhitch[0])
    print('new pick :')
    print(listhitch[1])
    print('new pick :')
    print(listhitch[2])
    print('new cost :')
    print(listhitch[3])
    print(listhitch[4])
    print(listhitch[5])
    print(listhitch[6])
    print('__________________________________________________________________')
      
    rowused = lastRow(outputsheet,wb)                        # Find last row of data
    shtused = wb.sheets[outputsheet]                        
    if len(listhitch[0]) == 1:
        if rowused >= 2:
            shtused.range('A{0}'.format(rowused+2)).value = listhitch[0]                #  First column save order list
            shtused.range('B{0}'.format(rowused+2)).value = listhitch[1][0][0]           #  Second  column save Data 1
            shtused.range('C{0}'.format(rowused+2)).value = listhitch[1][0][1]         #  Third column save Data 2
            shtused.range('D{0}'.format(rowused+2)).value = listhitch[2][0][0]           #  Fourth column save Data 3
            shtused.range('E{0}'.format(rowused+2)).value = listhitch[2][0][1]           #  Third column save Data 4
            shtused.range('F{0}'.format(rowused+2)).value = listhitch[3]             #  Fourth column save Data 5
            shtused.range('G{0}'.format(rowused+2)).value = listhitch[4]            #  Fifth column save Data 5
            shtused.range('H{0}'.format(rowused+2)).value = listhitch[5]            #  Sixth column save Data 5
            shtused.range('I{0}'.format(rowused+2)).value = listhitch[6]            #  Sixth column save Data 5
        else:
            shtused.range('A{0}'.format(rowused+1)).value = listhitch[0]                #  First column save order list
            shtused.range('B{0}'.format(rowused+1)).value = listhitch[1][0][0]           #  Second  column save Data 1
            shtused.range('C{0}'.format(rowused+1)).value = listhitch[1][0][1]           #  Third column save Data 2
            shtused.range('D{0}'.format(rowused+1)).value = listhitch[2][0][0]           #  Fourth column save Data 3
            shtused.range('E{0}'.format(rowused+1)).value = listhitch[2][0][1]           #  Third column save Data 4
            shtused.range('F{0}'.format(rowused+1)).value = listhitch[3]              #  Fourth column save Data 5
            shtused.range('G{0}'.format(rowused+1)).value = listhitch[4]            #  Fifth column save Data 5
            shtused.range('H{0}'.format(rowused+1)).value = listhitch[5]            #  Sixth column save Data 5
            shtused.range('I{0}'.format(rowused+1)).value = listhitch[6]            #  Sixth column save Data 5
        
    else:
        for x in range(len(listhitch[1])):
            if rowused >= 2:
                shtused.range('A{0}'.format(rowused+x+2)).value = listhitch[0][x]                #  First column save order list
                shtused.range('B{0}'.format(rowused+x+2)).value = listhitch[1][x][0]           #  Second  column save Data 1
                shtused.range('C{0}'.format(rowused+x+2)).value = listhitch[1][x][1]         #  Third column save Data 2
                shtused.range('D{0}'.format(rowused+x+2)).value = listhitch[2][x][0]           #  Fourth column save Data 3
                shtused.range('E{0}'.format(rowused+x+2)).value = listhitch[2][x][1]           #  Third column save Data 4
                shtused.range('F{0}'.format(rowused+x+2)).value = listhitch[3][x]              #  Fourth column save Data 5
                shtused.range('G{0}'.format(rowused+x+2)).value = listhitch[4][x]            #  Fifth column save Data 5
                shtused.range('H{0}'.format(rowused+x+2)).value = listhitch[5][x]            #  Sixth column save Data 5
                shtused.range('I{0}'.format(rowused+x+2)).value = listhitch[6][x]            #  Sixth column save Data 5
            else:
                shtused.range('A{0}'.format(rowused+x+1)).value = listhitch[0][x]                #  First column save order list
                shtused.range('B{0}'.format(rowused+x+1)).value = listhitch[1][x][0]           #  Second  column save Data 1
                shtused.range('C{0}'.format(rowused+x+1)).value = listhitch[1][x][1]           #  Third column save Data 2
                shtused.range('D{0}'.format(rowused+x+1)).value = listhitch[2][x][0]           #  Fourth column save Data 3
                shtused.range('E{0}'.format(rowused+x+1)).value = listhitch[2][x][1]           #  Third column save Data 4
                shtused.range('F{0}'.format(rowused+x+1)).value = listhitch[3][x]              #  Fourth column save Data 5
                shtused.range('G{0}'.format(rowused+x+1)).value = listhitch[4][x]            #  Fifth column save Data 5
                shtused.range('H{0}'.format(rowused+x+1)).value = listhitch[5][x]            #  Sixth column save Data 5
                shtused.range('I{0}'.format(rowused+x+2)).value = listhitch[6][x]            #  Sixth column save Data 5

    return listhitch
###############################################################################
''' 7 '''
def takeAGV(area,file,inputsheet,outputsheet):
    '''Condition to find same pick point and send point'''
    wb = xw.Book(file)
    
    order = []
    pick = []
    send = []
    cost = []  
    status = []
    number = []
    picknum = []
    
    numberquest = lastRow(inputsheet,wb)  # Find last row of data                                                                                                 #  Select sheet to save data
    getsheet = wb.sheets[inputsheet]        # Open sheet of excel
    
    getlist = int(getsheet.range('A{0}'.format(numberquest)).value)
    # print('{},{}'.format(getlist,numberquest))
    
    if getlist == 1:
        order.append(int(getsheet.range('A{0}'.format(numberquest)).value))
        pick.append([int(getsheet.range('B{0}'.format(numberquest)).value)
                ,int(getsheet.range('C{0}'.format(numberquest)).value)])
        send.append([int(getsheet.range('D{0}'.format(numberquest)).value)
                ,int(getsheet.range('E{0}'.format(numberquest)).value)])
        cost.append(int(getsheet.range('F{0}'.format(numberquest)).value))
        status.append(getsheet.range('G{0}'.format(numberquest)).value)
        number.append(int(getsheet.range('H{0}'.format(numberquest)).value))
        picknum.append(int(getsheet.range('I{0}'.format(numberquest)).value))
    else:
        for i in range (numberquest+1-getlist,numberquest+1):
            order.append(int(getsheet.range('A{0}'.format(i)).value))
            pick.append([int(getsheet.range('B{0}'.format(i)).value)
                ,int(getsheet.range('C{0}'.format(i)).value)])
            send.append([int(getsheet.range('D{0}'.format(i)).value)
                ,int(getsheet.range('E{0}'.format(i)).value)])
            cost.append(int(getsheet.range('F{0}'.format(i)).value))    
            status.append(getsheet.range('G{0}'.format(i)).value)
            number.append(int(getsheet.range('H{0}'.format(i)).value)) 
            picknum.append(int(getsheet.range('I{0}'.format(i)).value)) 
    
    
    AGV1 = [[],[],[],[],[],[],[],[]]          # order \ startx \ starty \ stopx \ stopy \ cost \ order
    AGV2 = [[],[],[],[],[],[],[],[]]          # order \ startx \ starty \ stopx \ stopy \ cost \ order
    count = 0
    for j in range(len(pick)):
        if count % 2 == 1:
            AGV1[0].append(order)
            AGV1[1].append(pick[j])
            AGV1[2].append(send[j])
            AGV1[3].append(cost[j])
            AGV1[4].append(status[j])
            AGV1[5].append(number[j])
            AGV1[6].append(picknum[j])
            count = count + 1
        elif count % 2 == 0:
            AGV2[0].append(order)
            AGV2[1].append(pick[j])
            AGV2[2].append(send[j])
            AGV2[3].append(cost[j])
            AGV2[4].append(status[j])
            AGV2[5].append(number[j])
            AGV1[6].append(picknum[j])
            count = count + 1
            
    rowAGV1 = lastRow(outputsheet[0],wb)                        # Find last row of data
    shtAGV1 = wb.sheets[outputsheet[0]]   
    rowAGV2 = lastRow(outputsheet[1],wb)                        # Find last row of data
    shtAGV2 = wb.sheets[outputsheet[1]]   
    
    for x in range(len(AGV1[0])):
        shtAGV1.range('A{0}'.format(rowAGV1+x+2)).value = AGV1[0][x]       #  First column save order list
        shtAGV1.range('B{0}'.format(rowAGV1+x+2)).value = AGV1[1][x]       #  Second  column save Data 1
        shtAGV1.range('C{0}'.format(rowAGV1+x+2)).value = AGV1[2][x]       #  Third column save Data 2
        shtAGV1.range('D{0}'.format(rowAGV1+x+2)).value = AGV1[3][x]       #  Fourth column save Data 3
        shtAGV1.range('E{0}'.format(rowAGV1+x+2)).value = AGV1[4][x]       #  Third column save Data 4
        shtAGV1.range('F{0}'.format(rowAGV1+x+2)).value = AGV1[5][x]       #  Fourth column save Data 5
        shtAGV1.range('G{0}'.format(rowAGV1+x+2)).value = AGV1[6][x]       #  Fourth column save Data 5
    for x in range(len(AGV2[0])):
        shtAGV2.range('A{0}'.format(rowAGV2+x+2)).value = AGV2[0][x]       #  First column save order list
        shtAGV2.range('B{0}'.format(rowAGV2+x+2)).value = AGV2[1][x]       #  Second  column save Data 1
        shtAGV2.range('C{0}'.format(rowAGV2+x+2)).value = AGV2[2][x]       #  Third column save Data 2
        shtAGV2.range('D{0}'.format(rowAGV2+x+2)).value = AGV2[3][x]       #  Fourth column save Data 3
        shtAGV2.range('E{0}'.format(rowAGV2+x+2)).value = AGV2[4][x]       #  Third column save Data 4
        shtAGV2.range('F{0}'.format(rowAGV2+x+2)).value = AGV2[5][x]       #  Fourth column save Data 5
        shtAGV2.range('G{0}'.format(rowAGV2+x+2)).value = AGV2[6][x]       #  Fourth column save Data 5
    return AGV1,AGV2

###############################################################################
''' 5 '''
def cal(start,stop,areaused):
    '''Cal distance'''
    direction = []
    redirection = []
    Num = []
    prey = []
    ### cal the way###
    maze = areaused
    cost = 1                                                                     # cost per movement
    costout = 0
    costall = []
    pathout = ast.search(maze,cost, start, stop)
    for y in pathout:
        for x in y:                                                             # outx = [Each row] in array
            if x != -1:                                                         # outy = Row in array
                direction.append([pathout.index(y),y.index(x)]) #[x,y]
                Num.append(int(x))
    for direct in range(len(Num)):
        point = Num.index(direct)
        redirection.append(direction[point])
    prey = redirection[0][1]
    #print(prey)
    for get in range(len(redirection)):
        if redirection[get][1] != prey:
            costout = costout + 2
            prey = redirection[get][1]
            #print(costout,prey)
        else:
            costout = costout + 1
            #print(costout,prey)
        costall.append(costout)

    
    #print('pathout :')
    #print(pathout)
    #print('x :{}'.format(redirection[0]))
    #print('y :{}'.format(redirection[1]))
    
    directionout = redirection.copy()
    
    return costall,directionout
###############################################################################
''' 6 '''
def samepointcost(area,file,inputsheet,outputsheet):
    '''Condition to find same pick point and send point'''
    wb = xw.Book(file)
    
    order0 = []
    pick0 = []
    send0 = []
    cost0 = []  
    status0 = []
    number0 = []
    
    order1 = []
    pick1 = []
    send1 = []
    cost1 = []  
    status1 = []
    number1 = []
    
    AGVlist = [[],[]]
    getsheet = [[],[]]
    getlist = [[],[]]

    AGVlist[0] = lastRow(inputsheet[0],wb)  # Find last row of data                                                                                                 #  Select sheet to save data
    getsheet[0] = wb.sheets[inputsheet[0]]  # Open sheet of excel
    getlist[0] = int(getsheet[0].range('A{0}'.format(AGVlist[0])).value)
    # print('{},{}'.format(getlist,numberquest))
    
    AGVlist[1] = lastRow(inputsheet[1],wb)  # Find last row of data                                                                                                 #  Select sheet to save data
    getsheet[1] = wb.sheets[inputsheet[1]]  # Open sheet of excel
    getlist[1] = int(getsheet[1].range('A{0}'.format(AGVlist[1])).value)
    # print('{},{}'.format(getlist,numberquest))
    
    if getlist[0] == 1:
        order0.append(int(getsheet[0].range('A{0}'.format(AGVlist[0])).value))
        pick0.append([int(getsheet[0].range('B{0}'.format(AGVlist[0])).value)
                ,int(getsheet[0].range('C{0}'.format(AGVlist[0])).value)])
        send0.append([int(getsheet[0].range('D{0}'.format(AGVlist[0])).value)
                ,int(getsheet[0].range('E{0}'.format(AGVlist[0])).value)])
        cost0.append(int(getsheet[0].range('F{0}'.format(AGVlist[0])).value))
        status0.append(getsheet[0].range('G{0}'.format(AGVlist[0])).value)
        number0.append(int(getsheet[0].range('H{0}'.format(AGVlist[0])).value))
    else:
        for i in range (AGVlist[0]+1-getlist[0],AGVlist[0]+1):
            order0.append(int(getsheet[0].range('A{0}'.format(i)).value))
            pick0.append([int(getsheet[0].range('B{0}'.format(i)).value)
                ,int(getsheet[0].range('C{0}'.format(i)).value)])
            send0.append([int(getsheet[0].range('D{0}'.format(i)).value)
                ,int(getsheet[0].range('E{0}'.format(i)).value)])
            cost0.append(int(getsheet[0].range('F{0}'.format(i)).value))    
            status0.append(getsheet[0].range('G{0}'.format(i)).value)
            number0.append(int(getsheet[0].range('H{0}'.format(i)).value))  
            
    if getlist[1] == 1:
        order1.append(int(getsheet[1].range('A{0}'.format(AGVlist[1])).value))
        pick1.append([int(getsheet[1].range('B{0}'.format(AGVlist[1])).value)
                ,int(getsheet[1].range('C{0}'.format(AGVlist[1])).value)])
        send1.append([int(getsheet[1].range('D{0}'.format(AGVlist[1])).value)
                ,int(getsheet[1].range('E{0}'.format(AGVlist[1])).value)])
        cost1.append(int(getsheet[1].range('F{0}'.format(AGVlist[1])).value))
        status1.append(getsheet[1].range('G{0}'.format(AGVlist[1])).value)
        number1.append(int(getsheet[1].range('H{0}'.format(AGVlist[1])).value))
    else:
        for i in range (AGVlist[1]+1-getlist[1],AGVlist[1]+1):
            order1.append(int(getsheet[1].range('A{0}'.format(i)).value))
            pick1.append([int(getsheet[1].range('B{0}'.format(i)).value)
                ,int(getsheet[1].range('C{0}'.format(i)).value)])
            send1.append([int(getsheet[1].range('D{0}'.format(i)).value)
                ,int(getsheet[1].range('E{0}'.format(i)).value)])
            cost1.append(int(getsheet[1].range('F{0}'.format(i)).value))    
            status1.append(getsheet[1].range('G{0}'.format(i)).value)
            number1.append(int(getsheet[1].range('H{0}'.format(i)).value))
            
    directionall = [[],[]]
    costall = [[],[]]
    recentcost = 0
    directionend = [[],[]]
    for j in range(len(pick0)):
        [costcon,direction]=cal(pick0,send0,area)
        directionall[0].append(direction)
        costall[0].append(costcon+recentcost)
        recentcost = recentcost + max(costcon)
        print (costall[0])
        print (directionall[0])
    for j in range(len(pick1)):
        [costcon,direction]=cal(pick1,send1,area)
        directionall[1].append(direction)
        costall[1].append(costcon+recentcost)
        recentcost = recentcost + max(costcon)
        print (costall[1])
        print (directionall[1])
    
    if len(costall[0]) >= len(costall[1]):
        for k in range(len(directionall[1])):
            for l in range(len(costall[1])):
                if directionall[0][l] == directionall[1][k] and costall[0][l] == costall[1][k] and l+1 == len(costall[1]):
                    print('have same point on 2 car')
                    
                else:
                    directionend[0].append(directionall[0][k])
                    directionend[1].append(directionall[1][k])
                
    elif len(costall[0]) < len(costall[1]):
        for k in range(len(directionall[0])):
            for l in range(len(costall[0])):
                if directionall[0][l] == directionall[1][k] and costall[0][l] == costall[1][k] and l+1 == len(costall[0]):
                    print('have same point on 2 car')
                    
                else:
                    directionend[0].append(directionall[0][k])
                    directionend[1].append(directionall[1][k])
                
    return
###############################################################################
def writeline(AGV1,AGV2):
    #######################################################################
    #Line
    #######################################################################
    #Delete Gui
    # Delete task in AGV1 table
    X = AGVTable1.get_children()
    for item in X:
        AGVTable1.delete(item)

    # Delete task in AGV2 table
    Y = AGVTable2.get_children()
    for item in Y:
        AGVTable2.delete(item)

    if maxnumlist != []:

        # Delete line

        # Delete line from start to end (A1)
        for R in range(len(maxnumlistA1)):
            for r in range(maxnumlistA1[R]):
                myCanvas.delete(LineA1[R]-r)
        maxnumlistA1.clear()
        LineA1.clear()
        
        # Delete line from start to end (A2)
        for S in range(len(maxnumlistA2)):
            for s in range(maxnumlistA2[S]):
                myCanvas.delete(LineA2[S]-s)    
        maxnumlistA2.clear()
        LineA2.clear()

        # Delete line from AGV to start A1
        for T in range(len(AGVtostartA1)):
            for t in range(AGVtostartA1[T]):
                myCanvas.delete(LineAGVA1[T]-t)
        AGVtostartA1.clear()
        LineAGVA1.clear()

        # Delete line from AGV to start A2
        for U in range(len(AGVtostartA2)):
            for u in range(AGVtostartA2[U]):
                myCanvas.delete(LineAGVA2[U]-u)
        AGVtostartA2.clear()
        LineAGVA2.clear()

        # Delete circle

        # Delete start circle A1
        for V in range(len(startcircleA1)):
            myCanvas.delete(startcircleA1[V])
        startcircleA1.clear()
        
        # Delete end circle A1
        for W in range(len(endcircleA1)):
            myCanvas.delete(endcircleA1[W])
        endcircleA1.clear()

        # Delete start circle A2
        for v in range(len(startcircleA2)):
            myCanvas.delete(startcircleA2[v])
        startcircleA2.clear()

        # Delete end circle A2
        for w in range(len(endcircleA2)):
            myCanvas.delete(endcircleA2[w])
        endcircleA2.clear()
    
###############################################################################
    
    # Clear list

    listnum1.clear()
    listnum2.clear()

    listNum1.clear()
    listNum2.clear()

    listDatay1.clear()
    listDatax1.clear()

    listDatay2.clear()
    listDatax2.clear()

    listdatay1.clear()
    listdatax1.clear()

    listdatay2.clear()
    listdatax2.clear()

    # AGV1
    for xx in range(len(AGV1[0])):
        # print('xx',xx)

        # print("Shortest path from start i to end i task :"+str(i))
        maze = [[0, 0, 0, 0, 0, 0, 1, 0, 0],
                [0, 0, 0, 0, 0, 0, 1, 0, 0],
                [0, 0, 1, 0, 0, 0, 1, 0, 0],
                [0, 0, 1, 0, 0, 0, 1, 0, 0],
                [0, 0, 1, 0, 0, 0, 1, 0, 0],
                [0, 0, 1, 0, 0, 0, 1, 0, 0],
                [0, 0, 1, 0, 0, 0, 1, 0, 0],
                [0, 0, 1, 0, 0, 0, 0, 0, 0],
                [0, 0, 1, 0, 0, 0, 0, 0, 0]]

        # Simulation start to end "not finish"
        if finishA1 == []:
            # Draw line start to end
            start = [AGV1[1][xx], AGV1[2][xx]] # starting position

            ######################################################################################
            ''''How many paper AGV1'''
            ######################################################################################

            '''PaperAGV1 = Label(root,text=(f"[{AGV1[0][xx]+1}]"),font=("bold",14))
            PaperAGV1.place(x=840,y=59)'''
        
        else:
            if xx in range(len(finishA1)):
                # Simulation start to end "finish"
                if finishA1[xx] == 1: 
                    # Not draw line start to end
                    start = [AGV1[3][xx], AGV1[4][xx]] # ending position

            # Simulation start to end "not finish"
            else:
                # Draw line start to end
                start = [AGV1[1][xx], AGV1[2][xx]] # starting position

                ######################################################################################
                ''''How many paper AGV1'''
                ######################################################################################

                '''PaperAGV1 = Label(root,text=(f"[{AGV1[0][xx]+1}]"),font=("bold",14))
                PaperAGV1.place(x=840,y=59)'''
        
        end = [AGV1[3][xx], AGV1[4][xx]] # ending position
        cost = 1 # cost per movement

        path = ast.search(maze,cost, start, end)

        # print('\n'.join([''.join(["{:" ">3d}".format(item) for item in row]) 
        # for row in path]))

        '''print("start: " + str(start))
        print("end: " + str(end))'''

        outy = []
        outx = []
        Datay = []
        Datax = []
        Num = []
        Befy = 10 
        Befx = 10
        Afty = 0
        Aftx = 0

        for x in path:
            outx.append(x)                              # outx = [Each row] in array
            for z in x:
                outy.append(z)                          # outy = Row in array
                if z != -1:
                    Datay.append((len(outy)-1)%9)       # len(outy) = 81
                    Datax.append(len(outx)-1)           # len(outx) = 9
                    Num.append(z)

        maxnumlist.append(max(Num)) 
        
        if xx in range(len(finishA1)):
            # Simulation start to end "finish"
            if finishA1[xx] == 1: 
                print('no listNum1')
                if xx in range(len(listNum1)):
                    listNum1[xx] = [0]
                    listDatax1[xx] = [0]
                    listDatay1[xx] = [0]
                else:
                    listNum1.append([0])
                    listDatax1.append([0])
                    listDatay1.append([0])
        else:
            # Simulation start to end "not finish"
            listNum1.append(Num)
            listDatax1.append(Datax)
            listDatay1.append(Datay)

        # Show line from start i to end i
        for out in range(max(Num)+1):
            pos = Num.index(out)                        # pos = Number of path to move 
            if Befy != 10 or Befx != 10:
                Afty = Datay[pos]                       # Afty = Number 0,1,2,... of y
                Aftx = Datax[pos]                       # Aftx = Number 0,1,2,... of x
                line_SE = myCanvas.create_line(30+(30*Befy), 30+(30*Befx),30+(30*Afty), 30+(30*Aftx), fill="blue", width=10)
                Befy = Afty
                Befx = Aftx
                pass
            else:
                Befy = Datay[pos]                       # Befy = Last number of Datay
                Befx = Datax[pos]                       # Befx = Last number of Datax
                pass
                line_SE = 0
        
        Line.append(line_SE)
        LineA1.append(line_SE)
        maxnumlistA1.append(max(Num))

        # Calculate distance between AGV1 to start
        # print("Calculate distance between AGV1 to start")
        maze = area

        # First task
        if xx == 0:
            # Simulation AGV1 to start not "finish"
            if finishAGV1 == []:
                # Draw line AGV1 to start
                start = [4, 0] # AGV1 home position

            # Simulation AGV1 to start "finish"
            else: 
                # Not draw line AGV1 to start
                start = [AGV1[1][xx], AGV1[2][xx]] # starting position

        # 2nd task ++
        else :
            # Simulation AGV1 to start not "finish"
            if finishAGV1 == []:
                # Draw line previous end to start
                start = [AGV1[3][xx-1], AGV1[4][xx-1]] # previous ending position
            else:
                # Simulation AGV1 to start "finish"
                if xx in range(len(finishAGV1)):
                    if finishAGV1[xx] == 1:
                        # Not draw line AGV1 to start
                        start = [AGV1[1][xx], AGV1[2][xx]] # starting position
                else:
                    # Draw line previous end to start
                    start = [AGV1[3][xx-1], AGV1[4][xx-1]] # previous ending position
        
        end = [AGV1[1][xx], AGV1[2][xx]]
        cost = 1 # cost per movement

        path = ast.search(maze,cost, start, end)

        #print('\n'.join([''.join(["{:" ">3d}".format(item) for item in row]) 
        #for row in path]))
        print("start: " + str(start))
        print("end: " + str(end))

        outy = []
        outx = []
        datay1 = []
        datax1 = []
        num1 = []
        befy1 = 10 
        befx1 = 10
        afty1 = 0
        aftx1 = 0

        for x in path:
            outx.append(x)                              # outx = [Each row] in array
            for z in x:
                outy.append(z)                          # outy = Row in array
                if z != -1:
                    datay1.append((len(outy)-1)%9)       # len(outy) = 81
                    datax1.append(len(outx)-1)           # len(outx) = 9
                    num1.append(z)

        AGVtostart.append(max(num1))
        AGVtostartA1.append(max(num1))

        # Show line from AGV1 to start i        
        for out in range(max(num1)+1):
            pos = num1.index(out)                        # pos = Number of path to move 
            if befy1 != 10 or befx1 != 10:
                afty1 = datay1[pos]                       # Afty = Number 0,1,2,... of y
                aftx1 = datax1[pos]                       # Aftx = Number 0,1,2,... of x
                line_AS = myCanvas.create_line(30+(30*befy1), 30+(30*befx1),30+(30*afty1), 30+(30*aftx1), fill="cyan", width=5)
                befy1 = afty1
                befx1 = aftx1
                pass
            else:
                befy1 = datay1[pos]                       # Befy = Last number of Datay
                befx1 = datax1[pos]                       # Befx = Last number of Datax
                pass
                line_AS = 0

        LineAGV.append(line_AS)
        LineAGVA1.append(line_AS)

        # Simulation AGV1 to start "finish"
        if xx in range(len(finishAGV1)):
            if finishAGV1[xx] == 1: 
                print('no listnum1')
                if xx in range(len(listnum1)):
                    listnum1[xx] = [0]
                    listdatay1[xx] = [0]
                    listdatax1[xx] = [0]
                else:
                    listnum1.append([0])
                    listdatay1.append([0])
                    listdatax1.append([0])
        # Simulation AGV1 to start "not finish"
        else:
            if xx in range(len(finishA1)):
                if finishA1[xx] == 1 :
                    print('no listnum1')
                    #finishAGV1[xx] = 1
                    #print('finishAGV1',finishAGV1)
                    if xx in range(len(listnum1)):
                        listnum1[xx] = [0]
                        listdatay1[xx] = [0]
                        listdatax1[xx] = [0]
                    else:
                        listnum1.append([0])
                        listdatay1.append([0])
                        listdatax1.append([0])

            else:
                # Simulation AGV1 to start "not finish"
                listnum1.append(num1)
                listdatay1.append(datay1)
                listdatax1.append(datax1)

        # Insert task in AGV1 table
        AGVTable1.insert('', 'end', 
                                values=(xx+1,f"({AGV1[1][xx]},{AGV1[2][xx]})" ,
                                        f"({AGV1[3][xx]},{AGV1[4][xx]})",AGV1[6][xx],AGV1[7][xx]))
        
        # Simulation AGV1 to start "finish"    
        if xx in range(len(finishAGV1)):
            if finishAGV1[xx] == 1:
                print("No startcircle 1")
                if xx in range(len(startcircleA1)):
                    startcircleA1[xy] = 0
                else:
                    startcircleA1.append(0)

        # Simulation AGV1 to start "not finish"
        else:
            if xx in range(len(finishA1)):
                if finishA1[xx] == 1 :
                    print("No startcircle 1")
                    if xx in range(len(startcircleA1)):
                        startcircleA1[xy] = 0
                    else:
                        startcircleA1.append(0)
            else:
                # Create start position
                startcircle = create_circle(30+30*(AGV1[2][xx]), 30+30*(AGV1[1][xx]), 8, myCanvas,fill='green')
                startcircleA1.append(startcircle)
        
        # Simulation start to end "finish"
        if xx in range(len(finishA1)):
            if finishA1[xx] == 1:
                print("No endcircle 1")
                if xx in range(len(endcircleA1)):
                    endcircleA1[xy] = 0
                else:
                    endcircleA1.append(0)

        # Simulation start to end "not finish"
        else:
            # Create end position
            endcircle = create_circle(30+30*(AGV1[4][xx]), 30+30*(AGV1[3][xx]), 8, myCanvas,fill='brown')
            endcircleA1.append(endcircle)

        # Display of AGV from AGV_now to end for AGV1
        def AGV_movealong_AE1():
            print('AE1')
            #print('listnum1',listnum1)
            for L in range(len(AGVnow1)):
                #print('L',L)

                # Last AGV1 position
                if L == len(AGVnow1)-1:
                    # Delete last AGV1 position
                    myCanvas.delete(AGVnow1[L])

            for sth in range(len(finishA1),len(listnum1)):
                #print('sth',sth)

                # AGV_now to starti
                for out in range(max(listnum1[sth])+1):
                    pos1 = listnum1[sth].index(out)
                    if out < max(listnum1[sth]):
                        afty1 = listdatay1[sth][pos1]
                        aftx1 = listdatax1[sth][pos1]
                        befy1_1 = afty1
                        befx1_1 = aftx1
                        # make AGV move in each point
                        agv_by_path1 = create_circle(30+30*(befy1_1), 30+30*(befx1_1), 14, myCanvas,fill='blue')

                        time.sleep(1)
                        myCanvas.delete(agv_by_path1)
                        out += 1
                        print(0)
                    
                    # last point of AGV_now path = starti
                    elif out == max(listnum1[sth]):

                        befy1_1 = listdatay1[sth][pos1]
                        befx1_1 = listdatax1[sth][pos1]

                        print(1)
                        finishAGV1.append(1)
                        
                        # delete line from AGVstart to starti
                        #print('delete line from AGVstart to starti')
                        for AS in range(max(listnum1[sth])):
                            myCanvas.delete(LineAGVA1[sth]-AS)

                        if sth in range(len(startcircleA1)):
                            # Delete start circle
                            myCanvas.delete(startcircleA1[sth])

                        # Simulation AGV1 to start "finish"
                        listnum1[sth] = [0]
                
                # starti to end i
                #print('listNum1',listNum1)
                for out in range(max(listNum1[sth])+1):

                    pos = listNum1[sth].index(out)

                    if out < max(listNum1[sth]):

                        Afty = listDatay1[sth][pos]
                        Aftx = listDatax1[sth][pos]
                        Befy_1 = Afty
                        Befx_1 = Aftx

                        if out == 0:
                            print('stop')
                            stopforpick = create_circle(30+30*(Befy_1), 30+30*(Befx_1), 14, myCanvas,fill='blue')
                            time.sleep(1)
                            myCanvas.delete(stopforpick)
                        else:
                            pass
                        
                        agv_by_path2 = create_circle(30+30*(Befy_1), 30+30*(Befx_1), 14, myCanvas,fill='blue')

                        time.sleep(1)
                        myCanvas.delete(agv_by_path2)

                        out += 1
                        #print(2)
                
                    # last point of start to end path = endi
                    elif out == max(listNum1[sth]):
                        
                        Befy_1 = listDatay1[sth][pos]
                        Befx_1 = listDatax1[sth][pos]
                        
                        print(3)

                        # delete line from starti to endi
                        #print('delete line from starti to endi')
                        for SE1 in range(maxnumlistA1[sth]):
                            myCanvas.delete(LineA1[sth]-SE1)

                        if sth in range(len(endcircleA1)):
                            # Delete end circle
                            myCanvas.delete(endcircleA1[sth])

                        # Simulation start to end "finish"
                        listNum1[sth] = [0]

                        finishA1.append(1)

                        sim1[sth] = 0

                        #print('finishA1',finishA1)
                        #print('listNum1',listNum1)
                        if (len(finishA1) < len(listNum1)):
                            # Start first simulation
                            threadAE1 = Thread(target = AGV_movealong_AE1)
                            threadAE1.start()

                            sim1.append(1)

                        if (len(finishA1) == len(listNum1)):    
                            # create circle at last point
                            AGV1_now = create_circle(30+30*(listDatay1[sth][pos]), 30+30*(listDatax1[sth][pos]), 14, myCanvas,fill='blue')
                            AGVnow1.append(AGV1_now)

                        for L in range(len(AGVnow1)):
                                    
                            # Last AGV1 position
                            if L == len(AGVnow1)-1:
                                # Have next task to do so it's not last AGV1 position anymore
                                if len(finishA1) < len(maxnumlistA1) :
                                    # Delete AGV1 position
                                    myCanvas.delete(AGVnow1[L])
                                       
                                # Last AGV1 position
                                else:
                                    print('not delete AGV1 now')
                                    
                            # Not the last AGV1 position 
                            else:
                                # Delete AGV1 position
                                myCanvas.delete(AGVnow1[L])

                        '''# AGV1 "not finish" sending paper 
                        if sth+1 < len(listNum1):
                            # Show number of paper left
                            PaperAGV1 = Label(root,text=(f"[{len(listNum1)-(sth+1)}]"),font=("bold",14))
                            PaperAGV1.place(x=840,y=59)

                        # AGV1 "finish" sending paper 
                        else:
                            PaperAGV1 = Label(root,text=("[0]"),font=("bold",14))
                            PaperAGV1.place(x=840,y=59)'''
        
        #print('AGV1order',AGV1order)
        #print('len(AGV1[1]',len(AGV1[1]))
        
        # Haven't simulation yet
        if (all(d == 0 for d in sim1)) and picklist[-1] == AGV1[1][-1] and picklist2[-1] == AGV1[2][-1] and sendlist[-1] == AGV1[3][-1] and sendlist2[-1] == AGV1[4][-1]:
            if AGV1order == []:
                print('second sim1')
                # Start first simulation
                threadAE1 = Thread(target = AGV_movealong_AE1)
                threadAE1.start()

                sim1.append(1)

            else:
                if AGV1order[-1] < len(AGV1[1]):
                    print('sim1')
                    # Start first simulation
                    threadAE1 = Thread(target = AGV_movealong_AE1)
                    threadAE1.start()

                    sim1.append(1)
            
        elif sim1 == []:
            print('first sim1')
            # Start first simulation
            threadAE1 = Thread(target = AGV_movealong_AE1)
            threadAE1.start()

            sim1.append(1)

    # AGV1 order before        
    AGV1order.append(len(AGV1[1]))

    # AGV2
    for xy in range(len(AGV2[0])):

        #print('xy',xy)
        #print("Shortest path from start i to end i task :"+str(i))
        maze = area

        # Simulation start to end "not finish"
        if finishA2 == []:
            # Draw line start to end
            start = [AGV2[1][xy], AGV2[2][xy]] # starting position

            ######################################################################################
            ''''How many paper AGV2'''
            ######################################################################################

            '''PaperAGV2 = Label(root,text=(f"[{AGV2[0][xy]+1}]"),font=("bold",14))
            PaperAGV2.place(x=840,y=259)'''
        
        else:
            if xy in range(len(finishA2)): 
                # Simulation start to end "finish"
                if finishA2[xy] == 1:
                    # Not draw line start to end
                    start = [AGV2[3][xy], AGV2[4][xy]] # ending position
            
            # Simulation start to end "not finish"
            else:
                # Draw line start to end
                start = [AGV2[1][xy], AGV2[2][xy]] # starting position

                ######################################################################################
                ''''How many paper AGV2'''
                ######################################################################################

                '''PaperAGV2 = Label(root,text=(f"[{AGV2[0][xy]+1}]"),font=("bold",14))
                PaperAGV2.place(x=840,y=259)'''
        
        end = [AGV2[3][xy], AGV2[4][xy]] # ending position
        cost = 1 # cost per movement

        path = ast.search(maze,cost, start, end)

        # print('\n'.join([''.join(["{:" ">3d}".format(item) for item in row]) 
        # for row in path]))

        print("start: " + str(start))
        print("end: " + str(end))

        outy = []
        outx = []
        Datay = []
        Datax = []
        Num = []
        Befy = 10 
        Befx = 10
        Afty = 0
        Aftx = 0

        for x in path:
            outx.append(x)                              # outx = [Each row] in array
            for z in x:
                outy.append(z)                          # outy = Row in array
                if z != -1:
                    Datay.append((len(outy)-1)%9)       # len(outy) = 81
                    Datax.append(len(outx)-1)           # len(outx) = 9
                    Num.append(z)

        maxnumlist.append(max(Num))

        if xy in range(len(finishA2)):
            # Simulation start to end "finish"
            if finishA2[xy] == 1: 
                print('no listNum2')
                if xy in range(len(listNum2)):
                    listNum2[xy] = [0]
                    listDatax2[xy] = [0]
                    listDatay2[xy] = [0]
                else:
                    listNum2.append([0])
                    listDatax2.append([0])
                    listDatay2.append([0])
        else:
            # Simulation start to end "not finish"
            listNum2.append(Num)
            listDatax2.append(Datax)
            listDatay2.append(Datay)

        # Show line from start i to end i
        for out in range(max(Num)+1):
            pos = Num.index(out)                        # pos = Number of path to move 
            if Befy != 10 or Befx != 10:
                Afty = Datay[pos]                       # Afty = Number 0,1,2,... of y
                Aftx = Datax[pos]                       # Aftx = Number 0,1,2,... of x
                line_SE = myCanvas.create_line(30+(30*Befy), 30+(30*Befx),30+(30*Afty), 30+(30*Aftx), fill="purple", width=10)
                Befy = Afty
                Befx = Aftx
                pass
            else:
                Befy = Datay[pos]                       # Befy = Last number of Datay
                Befx = Datax[pos]                       # Befx = Last number of Datax
                pass
                line_SE = 0

        Line.append(line_SE)
        LineA2.append(line_SE)
        maxnumlistA2.append(max(Num))

        # Calculate distance between AGV2 to start
        # print("Calculate distance between AGV2 to start")
        maze = area

        # First task
        if xy == 0:
            # Simulation AGV2 to start not "finish"
            if finishAGV2 == []:
                # Draw line AGV2 to start
                start = [4, 8] # AGV2 home position

            # Simulation AGV2 to start "finish"    
            else:
                # Not draw line AGV2 to start
                start = [AGV2[1][xy], AGV2[2][xy]] # starting position
        
        # 2nd task ++
        else:
            # Simulation AGV2 to start not "finish"
            if finishAGV2 == []:
                # Draw line previous end to start
                start = [AGV2[3][xy-1], AGV2[4][xy-1]] # previous ending position
            else:
                # Simulation AGV2 to start "finish"
                if xy in range(len(finishAGV2)):
                    if finishAGV2[xy] == 1:
                        # Not draw line AGV2 to start
                        start = [AGV2[1][xy], AGV2[2][xy]] # starting position
                else:    
                    # Draw line previous end to start
                    start = [AGV2[3][xy-1], AGV2[4][xy-1]] # previous ending position

        end = [AGV2[1][xy], AGV2[2][xy]]
        cost = 1 # cost per movement

        path = ast.search(maze,cost, start, end)

        # print('\n'.join([''.join(["{:" ">3d}".format(item) for item in row]) 
        # for row in path]))
        print("start: " + str(start))
        print("end: " + str(end))

        outy = []
        outx = []
        datay2 = []
        datax2 = []
        num2 = []
        befy2 = 10 
        befx2 = 10
        afty2 = 0
        aftx2 = 0

        for x in path:
            outx.append(x)                              # outx = [Each row] in array
            for z in x:
                outy.append(z)                          # outy = Row in array
                if z != -1:
                    datay2.append((len(outy)-1)%9)       # len(outy) = 81
                    datax2.append(len(outx)-1)           # len(outx) = 9
                    num2.append(z)

        AGVtostart.append(max(num2))
        AGVtostartA2.append(max(num2))

        # Show line from AGV2 to start i        
        for out in range(max(num2)+1):
            pos = num2.index(out)                        # pos = Number of path to move 
            if befy2 != 10 or befx2 != 10:
                afty2 = datay2[pos]                       # Afty = Number 0,1,2,... of y
                aftx2 = datax2[pos]                       # Aftx = Number 0,1,2,... of x
                line_AS = myCanvas.create_line(30+(30*befy2), 30+(30*befx2),30+(30*afty2), 30+(30*aftx2), fill="magenta", width=5)
                befy2 = afty2
                befx2 = aftx2
                pass
            else:
                befy2 = datay2[pos]                       # Befy = Last number of Datay
                befx2 = datax2[pos]                       # Befx = Last number of Datax
                pass
                line_AS = 0

        LineAGV.append(line_AS)
        LineAGVA2.append(line_AS)

        # Simulation AGV2 to start "finish"
        if xy in range(len(finishAGV2)):
            if finishAGV2[xy] == 1: 
                print('no listnum2')
                if xy in range(len(listnum2)):
                    listnum2[xy] = [0]
                    listdatay2[xy] = [0]
                    listdatax2[xy] = [0]
                else:
                    listnum2.append([0])
                    listdatay2.append([0])
                    listdatax2.append([0])

        else:
            if xy in range(len(finishA2)):
                if finishA2[xy] == 1 :
                    print('no listnum2')
                    #finishAGV2[xy] = 1
                    #print('finishAGV1',finishAGV2)
                    #listnum2[xy] = [0]
                    if xy in range(len(listnum2)):
                        listnum2[xy] = [0]
                        listdatay2[xy] = [0]
                        listdatax2[xy] = [0]
                    else:
                        listnum2.append([0])
                        listdatay2.append([0])
                        listdatax2.append([0])

            else:
                # Simulation AGV2 to start "not finish"
                listnum2.append(num2)
                listdatay2.append(datay2)
                listdatax2.append(datax2)

        # Insert task in AGV2 table
        AGVTable2.insert('', 'end', 
                                values=(xy+1,f"({AGV2[1][xy]},{AGV2[2][xy]})" ,
                                        f"({AGV2[3][xy]},{AGV2[4][xy]})",AGV2[6][xy],AGV2[7][xy]))

        # Simulation AGV2 to start "finish"    
        if xy in range(len(finishAGV2)):
            if finishAGV2[xy] == 1:
                print("No startcircle 2")
                if xy in range(len(startcircleA2)):
                    startcircleA2[xy] = 0
                else:
                    startcircleA2.append(0)

        # Simulation AGV2 to start "not finish"
        else:
            if xy in range(len(finishA2)):
                if finishA2[xy] == 1 :
                    print("No startcircle 2")
                    if xy in range(len(startcircleA2)):
                        startcircleA2[xy] = 0
                    else:
                        startcircleA2.append(0)
            else:
                # Create start position
                startcircle = create_circle(30+30*(AGV2[2][xy]), 30+30*(AGV2[1][xy]), 8, myCanvas,fill='red')
                startcircleA2.append(startcircle)

        # Simulation start to end "finish"
        if xy in range(len(finishA2)):
            if finishA2[xy] == 1:
                print("No endcircle 2")
                if xy in range(len(endcircleA2)):
                    endcircleA2[xy] = 0
                else:
                    endcircleA2.append(0)

        # Simulation start to end "not finish"
        else:
            # Create end position
            endcircle = create_circle(30+30*(AGV2[4][xy]), 30+30*(AGV2[3][xy]), 8, myCanvas,fill='orange')
            endcircleA2.append(endcircle)

        def AGV_movealong_AE2():
            print('AE2')

            for L in range(len(AGVnow2)):
                #print('L',L)

                # Last AGV2 position
                if L == len(AGVnow2)-1:
                    # Delete last AGV2 position
                    myCanvas.delete(AGVnow2[L])
                    #print('delete AGV2 now')

            #print('len(finishA2)',len(finishA2))
            #print('len(listnum2)',len(listnum2))
            for sth in range(len(finishA2),len(listnum2)):
                #print('sth',sth)
                #print('listnum2',listnum2)

                # AGV_now to starti
                for out in range(max(listnum2[sth])+1):
                    pos2 = listnum2[sth].index(out)
                    if out < max(listnum2[sth]):
                        afty2 = listdatay2[sth][pos2]                      
                        aftx2 = listdatax2[sth][pos2]
                        befy2_2 = afty2
                        befx2_2 = aftx2
                        # make AGV move in each point
                        agv_by_path11 = create_circle(30+30*(befy2_2), 30+30*(befx2_2), 14, myCanvas,fill='magenta')

                        time.sleep(1)
                        myCanvas.delete(agv_by_path11)
                        out += 1
                    
                    # last point of AGV_now path = starti
                    elif out == max(listnum2[sth]):

                        befy2_2 = listdatay2[sth][pos2]
                        befx2_2 = listdatax2[sth][pos2]
                        
                        finishAGV2.append(1)
                    
                        # delete line from AGVstart to starti
                        #print('delete line from AGVstart to starti')
                        for AS in range(max(listnum2[sth])):
                            myCanvas.delete(LineAGVA2[sth]-AS)

                        if sth in range(len(startcircleA2)):
                            # Delete start circle
                            myCanvas.delete(startcircleA2[sth])

                        # Simulation AGV2 to start "finish"
                        listnum2[sth] = [0]
                
                # starti to end i
                for out in range(max(listNum2[sth])+1):

                    pos = listNum2[sth].index(out)

                    if out < max(listNum2[sth]):
                        
                        Afty = listDatay2[sth][pos]
                        Aftx = listDatax2[sth][pos]
                        Befy_2 = Afty
                        Befx_2 = Aftx

                        if out == 0:
                            print('stop')
                            stopforpick2 = create_circle(30+30*(Befy_2), 30+30*(Befx_2), 14, myCanvas,fill='magenta')
                            time.sleep(1)
                            myCanvas.delete(stopforpick2)
                        else:
                            pass
                        
                        agv_by_path22 = create_circle(30+30*(Befy_2), 30+30*(Befx_2), 14, myCanvas,fill='magenta')
                        
                        time.sleep(1)
                        myCanvas.delete(agv_by_path22)
                        
                        out += 1
                        
                    # last point of start to end path = endi
                    elif out == max(listNum2[sth]):

                        Befy_2 = listDatay2[sth][pos]
                        Befx_2 = listDatax2[sth][pos]

                        print(33)

                        '''# AGV2 "not finish" sending paper 
                        if sth+1 < len(listNum2):
                            # Show number of paper left
                            PaperAGV2 = Label(root,text=(f"[{len(listNum2)-(sth+1)}]"),font=("bold",14))
                            PaperAGV2.place(x=840,y=259)

                        # AGV2 "finish" sending paper 
                        else:
                            PaperAGV2 = Label(root,text=("[0]"),font=("bold",14))
                            PaperAGV2.place(x=840,y=259)'''

                        # delete line from starti to endi
                        #print('delete line from starti to endi')
                        for SE2 in range(maxnumlistA2[sth]):
                            myCanvas.delete(LineA2[sth]-SE2)

                        if sth in range(len(endcircleA2)):
                            # Delete end circle
                            myCanvas.delete(endcircleA2[sth])

                        # Simulation start to end "finish"
                        listNum2[sth] = [0]

                        finishA2.append(1)

                        sim2[sth] = 0

                        if (len(finishA2) < len(listNum2)):
                            # Start first simulation  
                            threadAE2 = Thread(target = AGV_movealong_AE2)
                            threadAE2.start()

                            sim2.append(1)

                        if (len(finishA2) == len(listNum2)):    
                            # create circle at last point
                            AGV2_now = create_circle(30+30*(listDatay2[sth][pos]), 30+30*(listDatax2[sth][pos]), 14, myCanvas,fill='magenta')     
                            AGVnow2.append(AGV2_now)

                        for L in range(len(AGVnow2)):
                                    
                            # Last AGV2 position
                            if L == len(AGVnow2)-1:
                                # Have next task to do so it's not last AGV2 position anymore
                                if len(finishA2) < len(maxnumlistA2):
                                    # Delete AGV2 position
                                    myCanvas.delete(AGVnow2[L])
                                    #print('delete AGV2 now')
                                        
                                # Last AGV2 position
                                else:
                                    print('not delete AGV2 now')
                                    
                            # Not the last AGV2 position 
                            else:
                                # Delete AGV2 position
                                myCanvas.delete(AGVnow2[L])
                                #print('delete AGV2 now')
        
        #print('len(AGV2[0])',len(AGV2[0]))
        #print('AGV2order',AGV2order)
        #print('len(AGV2[1]',len(AGV2[1]))
        # Haven't simulation yet
        if (all(v == 0 for v in sim2)) and picklist[-1] == AGV2[1][-1] and picklist2[-1] == AGV2[2][-1] and sendlist[-1] == AGV2[3][-1] and sendlist2[-1] == AGV2[4][-1]:
            if AGV2order == []:
            
                print('second sim2')
                # Start first simulation  
                threadAE2 = Thread(target = AGV_movealong_AE2)
                threadAE2.start()

                sim2.append(1)
            else:
                if AGV2order[-1] < len(AGV2[1]):
                    print('sim2')
                    # Start first simulation  
                    threadAE2 = Thread(target = AGV_movealong_AE2)
                    threadAE2.start()

                    sim2.append(1)

        elif sim2 == []:
            print('first sim2')
            # Start first simulation
            threadAE2 = Thread(target = AGV_movealong_AE2)
            threadAE2.start()

            sim2.append(1)

    # AGV2 order before
    AGV2order.append(len(AGV2[1]))

    print('end')
    return

###############################################################################
'''GUI that make for control AGV'''
########################################################################################
'''Create Window object'''
########################################################################################
root = Tk()
root.geometry('1100x540')                                         # Window size
root.title("AGV")                                                 # Name of window             

########################################################################################
'''Menu bar'''
########################################################################################
menu=Menu(root)
root.config(menu=menu)

'''About'''
def abt():
    tkinter.messagebox.showinfo("Welcome to authors","This is demo for menu fields")

'''Exit'''
def exitt():
    exit()
    
'''File Menu'''
subm1=Menu(menu)
menu.add_cascade(label="File",menu=subm1)
subm1.add_command(label="Exit",command=exitt)

'''Option Menu'''
subm2=Menu(menu)
menu.add_cascade(label="Option",menu=subm2)
subm2.add_command(label="About",command=abt)

########################################################################################
'''Create Object'''
########################################################################################

'''Connect Button for IP AGV1'''
connect1=Label(root, text="connecting",font=("bold",10))
disconnect1=Label(root, text="   cancel  ",font=("bold",10))
connect2=Label(root, text="connecting",font=("bold",10))
disconnect2=Label(root, text="   cancel  ",font=("bold",10))
strreading=Label(root, text="start reading",font=("bold",10))
stpreading=Label(root, text="stop reading",font=("bold",10))

def conip1():
    ipa11=ipa1.get()
    print(f"IP AGV1 {ipa11}")
    connect1.place(x=470,y=95)
    disconnect1.place_forget()

but_connect1 = Button(root,text='Connect',width=12,bg='brown',fg='white',command=conip1).place(x=270, y=95)

ipa1=IntVar()
ipa2=IntVar()

'''Cancel Button for IP AGV1'''
def canip1():
    ipa11=ipa1.get()
    print(f"IP AGV1 Cancel")
    disconnect1.place(x=470,y=95)
    connect1.place_forget()
    
but_cancel1 = Button(root,text='Cancel',width=12,bg='brown',fg='white',command=canip1).place(x=360, y=95)

'''Confirm Button for IP AGV2'''
def conip2():
    ipa21=ipa2.get()
    print(f"IP AGV2 {ipa21}")
    connect2.place(x=470,y=135)
    disconnect2.place_forget()
    
but_connect2 = Button(root,text='Connect',width=12,bg='brown',fg='white',command=conip2).place(x=270, y=133)

'''Confirm Button for IP AGV2'''
def canip2():
    ipa21=ipa2.get()
    print(f"IP AGV2 Cancel")
    disconnect2.place(x=470,y=135)
    connect2.place_forget()

but_cancel2 = Button(root,text='Cancel',width=12,bg='brown',fg='white',command=canip2).place(x=360, y=133)

def start():
    strreading.place(x=430,y=290)
    stpreading.place_forget()

but_start = Button(root,text='Start',width=12,bg='brown',fg='white',command=start).place(x=420, y=200)

def stop():
    stpreading.place(x=430,y=290)
    strreading.place_forget()

but_stop = Button(root,text='Stop',width=12,bg='brown',fg='white',command=stop).place(x=420, y=260)

#######################################################################################
'''AGV display'''
#######################################################################################
myCanvas = Canvas(root, width=300, height=300, borderwidth=0, highlightthickness=0, bg="white")
'''myCanvas.pack(side="left",anchor=tk.SE)'''
myCanvas.place(x=80,y=180)

'''Create AGV on grid'''
def create_circle(x, y, r, canvasName, **kwargs): #center coordinates, radius
    x0 = x - r
    y0 = y - r
    x1 = x + r
    y1 = y + r
    return canvasName.create_oval(x0, y0, x1, y1, **kwargs)

'''Create line'''
for x in range(0,9):
    myCanvas.create_line(30+(30*x), 0, 30+(30*x), 300)
    myCanvas.create_line(0, 30+(30*x), 300, 30+(30*x))
myCanvas.create_line(30+(30*2), (30*3),30+(30*2), (30*10), fill="black", width=5)
myCanvas.create_line(30+(30*6), (30*0),30+(30*6), (30*7), fill="black", width=5)
myCanvas.create_rectangle(0, 0,300, 300, outline = 'black', width=5)

# Position 1
myCanvas.create_rectangle(255, 20, 285, 40, fill = "cyan")
# Position 2
myCanvas.create_rectangle(165, 20, 195, 40, fill = "cyan")
# Position 3
myCanvas.create_rectangle(105, 260, 135, 280, fill = "cyan")
# Position 4
myCanvas.create_rectangle(15, 260, 45, 280, fill = "cyan")

#######################################################################################
'''AGV Tabel'''
#######################################################################################

AGVTable1 = ttk.Treeview(height=6,columns=('No.','Pick','Quantity','Status','No. of paper'),show="headings")    
AGVTable1.place(x=550,y=100)

AGVLabel1=Label(root, text="AGV1", font=("bold",15))
AGVLabel1.place(x=780,y=60)

AGVTable2 = ttk.Treeview(height=6,columns=('No.','Pick','Quantity','Status','No. of paper'),show="headings")    
AGVTable2.place(x=550,y=300)

AGVLabel2=Label(root, text="AGV2", font=("bold",15))
AGVLabel2.place(x=780,y=260)

AGVPosition=Label(root, text="Position and status of AGV", font=("bold",13))
AGVPosition.place(x=130,y=490)

'''Heading of AGV1 table'''
AGVTable1.heading('No.',text="No.")
AGVTable1.heading('Pick',text="Pick")
AGVTable1.heading('Quantity',text="Send")
AGVTable1.heading('Status',text="Status")
AGVTable1.heading('No. of paper',text="No. of paper")

'''Column of AGV1 table'''
AGVTable1.column('No.', width=100, minwidth=100, stretch=tk.NO,anchor='center')
AGVTable1.column('Pick', width=100, minwidth=70, stretch=tk.NO,anchor='center')
AGVTable1.column('Quantity', width=100, minwidth=70, stretch=tk.NO,anchor='center')
AGVTable1.column('Status', width=100, minwidth=100, stretch=tk.NO,anchor='center')
AGVTable1.column('No. of paper', width=100, minwidth=100, stretch=tk.NO,anchor='center')

'''Scroll Bar for AGV1 table'''
scrollbar1 = Scrollbar(root, orient="vertical",command=AGVTable1.yview)
scrollbar1.place(x=1050,y=100)
AGVTable1.configure(yscrollcommand=scrollbar1.set)

'''Heading of AGV2 table'''
AGVTable2.heading('No.',text="No.")
AGVTable2.heading('Pick',text="Pick")
AGVTable2.heading('Quantity',text="Send")
AGVTable2.heading('Status',text="Status")
AGVTable2.heading('No. of paper',text="No. of paper")

'''Column of AGV2 table'''
AGVTable2.column('No.', width=100, minwidth=100, stretch=tk.NO,anchor='center')
AGVTable2.column('Pick', width=100, minwidth=70, stretch=tk.NO,anchor='center')
AGVTable2.column('Quantity', width=100, minwidth=70, stretch=tk.NO,anchor='center')
AGVTable2.column('Status', width=100, minwidth=100, stretch=tk.NO,anchor='center')
AGVTable2.column('No. of paper', width=100, minwidth=100, stretch=tk.NO,anchor='center')

'''Scroll Bar for AGV2 table'''
scrollbar2 = Scrollbar(root, orient="vertical",command=AGVTable1.yview)
scrollbar2.place(x=1050,y=300)
AGVTable2.configure(yscrollcommand=scrollbar2.set)

'''Automated Guided Vehicle'''
label_0 = Label(root, text="Automated Guided Vehicle",relief="solid",width=25,font=("arial",19,"bold"))
label_0.place(x=350,y=10)

######################################################################################
''''X,Y'''
######################################################################################

'''X'''
Y = Label(root,text="Y = 0     1      2      3      4      5      6     7      8",font=("bold",10))
Y.place(x=80,y=158)
'''Y'''
X = Label(root,text="X = 0",font=("bold",10))
X.place(x=43,y=200)
X1 = Label(root,text="1",font=("bold",10))
X1.place(x=65,y=230)
X2 = Label(root,text="2",font=("bold",10))
X2.place(x=65,y=260)
X3 = Label(root,text="3",font=("bold",10))
X3.place(x=65,y=290)
X4 = Label(root,text="4",font=("bold",10))
X4.place(x=65,y=320)
X5 = Label(root,text="5",font=("bold",10))
X5.place(x=65,y=350)
X6 = Label(root,text="6",font=("bold",10))
X6.place(x=65,y=380)
X7 = Label(root,text="7",font=("bold",10))
X7.place(x=65,y=410)
X8 = Label(root,text="8",font=("bold",10))
X8.place(x=65,y=440)

###########################

#######################################################################################
'''Point data getting'''
#######################################################################################


puf = IntVar()          # Put data (Pick up from) in table
puf2 = IntVar()
st = IntVar()           # Put data (Send to) in table
st2 = IntVar()

picklist = []           # List of pick x
picklist2 = []          # List of pick y
sendlist = []           # List of send x
sendlist2 = []          # List of send y

##############################################################################
'''Press confirm button'''
'''Main'''
###############################################################################
def get():

    global i

###############################################################################
    '''Get data form label'''

    pick = puf.get()           # Get pick input x
    pick2 = puf2.get()         # Get pick input y
    send = st.get()            # Get send input x
    send2 = st2.get()          # Get send input y
    picklist.append(pick)      # Put in picklist x
    picklist2.append(pick2)    # Put in picklist y
    sendlist.append(send)      # Put in sendlist x
    sendlist2.append(send2)    # Put in sendlist y

    ''' Calculate shortest path from start i to end i ''' 
    print("###############################################################################")
    print("Shortest path from start i to end i task "+str(i))

###############################################################################
    '''for calculate request'''
    pickquest = [pick,pick2]
    sendquest = [send,send2]
    [startquest,stopquest,costquest,pathquest] = calrequest(file,area,pickquest,sendquest)
    
    inputrequest.append(startquest)
    outputrequest.append(stopquest)
    costrequest.append(costquest)
    
    listpoint = pointcon(area,file,'Request','condition')
    listhitch = hitchhikingcon(area,file,'condition','hitchhiking')
    #[AGV1,AGV2] = takeAGV(area,file,inputsheet,outputsheet)
    [AGV1,AGV2] = calAGV(area,startAGV,file,'hitchhiking',['AGV1take','AGV2take'])
    
    writeline(AGV1,AGV2)
    '''calculate agv to request + cost'''
    
###############################################################################

    
    '''calculate AGV to request'''
    # maze = area
    # start = startAGV1                                     # starting position
    # end = [pick, pick2]                                   # ending position
    # cost = 1                                              # cost per movement
    # path = ast.search(maze,cost, start, end)
    
###############################################################################

#    print('\n'.join([''.join(["{:" ">3d}".format(item) for item in row]) 
#    for row in pathquest]))

#    print("start: " + str(inputrequest))
#    print("end: " + str(outputrequest))
#    print("Num :" +str(costrequest))
#    print("total cost: " +str(costrequest))
#    print(type(pickquest))
#    print(type(sendquest))
#    print(type(startquest))
#    print(type(stopquest))
    # maxnumlist.append(max(Num))
    # print("Maxnumlist: " +str(maxnumlist))


#######################################################################################
'''Label point'''
#######################################################################################
        
'''Pick up from'''
label_1 = Label(root, text="Pick up from :",width=20, font=("bold",10))
label_1.place(x=10,y=60)

'''Input Text Pick up from'''
tk.Entry(textvariable=puf,width=3).place(x=140,y=60)
tk.Entry(textvariable=puf2,width=3).place(x=165,y=60)

'''Confirm Button (Pick up from)'''
tk.Button(root, text="Confirm",width=12,bg='brown',fg='white', command=get).place(x=305,y=55)

''''Send to'''
label_1 = Label(root, text="Send to :", font=("bold",10))
label_1.place(x=190,y=60)

'''Input Text Send to'''
tk.Entry(textvariable=st,width=3).place(x=250,y=60)
tk.Entry(textvariable=st2,width=3).place(x=275,y=60)

#######################################################################################
'''IP text'''
#######################################################################################

'''IP address AGV1'''
label_1 = Label(root, text="IP AGV1 :",width=20, font=("bold",10))
label_1.place(x=20,y=100)

'''Input Text IP address AGV1'''
'''entry_1= Entry(root,textvar=ipa1)
entry_1.place(x=140,y=100)'''
tk.Entry(textvariable=ipa1).place(x=140,y=100)

'''IP address AGV2'''
label_1 = Label(root, text="IP AGV2 :",width=20, font=("bold",10))
label_1.place(x=20,y=138)

'''Input Text IP address AGV2'''
tk.Entry(textvariable=ipa2).place(x=140,y=138)    

#######################################################################################

root.iconify()
root.update()
root.deiconify()
root.mainloop()