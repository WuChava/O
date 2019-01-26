# -*- coding: utf-8 -*-
import Performance
import Division
import Inbound
import os
import Library

rawdata_path=Library.getSetting('source', 'rawdata', 1, 0)
rawdata_date=Library.getSetting('source', 'reportdate', 0, 1)
inboundlist_path=Library.getSetting('source', 'inboundlist', 0, 0)

inputString='0'
itmes=['1','2']
while (inputString not in itmes):
    inputString = input('1. Execute the Performance Report\n2. Execute the Inboundlist Report\nWhat is your choise (input 1 or 2): ')

if inputString=='1':
    Performance.MainPerformance(rawdata_path, rawdata_date, debug=1)
    Division.MainDivision(rawdata_path, rawdata_date)
elif inputString=='2':
    Inbound.MainInbound(inboundlist_path)

os.system("pause")