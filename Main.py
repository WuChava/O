# -*- coding: utf-8 -*-
import Performance
import Division
import Inbound
import os
import Library

rawdata_path=Library.getSetting('source', 'rawdata', 1, 0, 0)
rawdata_date=Library.getSetting('source', 'reportdate', 0, 1, 0)
utrate=Library.getSetting('source', 'utrate', 0, 0, 1)
if utrate==None: utrate=0
wordtimemarkup=Library.getSetting('source', 'wordtimemarkup', 0, 0, 1)
if wordtimemarkup==None: wordtimemarkup=0

inboundlist_path=Library.getSetting('source', 'inboundlist', 0, 0, 0)

isexist=0
while(not isexist):
    inputString='0'
    itmes=['1','2', '3', '4', '9']
    while (inputString not in itmes):
        inputString = input('1. Execute the Performance Report\n2. Execute the Performance Report with UT rate markup\n3. Execute the 分時表 Report\n4. Execute the Inboundlist Report\n9. Exit!!!\nWhat is your choise (input 1 or 2 or 3 or 4 or 9): ')

    if inputString=='1':
        Performance.MainPerformance(rawdata_path, rawdata_date, 0, utrate, wordtimemarkup, debug=1)
    elif inputString=='2':
        Performance.MainPerformance(rawdata_path, rawdata_date, 1, utrate, wordtimemarkup, debug=1) 
    elif inputString=='3':
        Division.MainDivision(rawdata_path, rawdata_date)
    elif inputString=='4':
        Inbound.MainInbound(inboundlist_path)
    elif inputString=='9':
        isexist=1

os.system("pause")