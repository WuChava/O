# -*- coding: utf-8 -*-
import Inbound
import os
import Library

rawdata_path=Library.getSetting('source', 'rawdata', 1, 0)
rawdata_date=Library.getSetting('source', 'reportdate', 0, 1)
inboundlist_path=Library.getSetting('source', 'inboundlist', 0, 0)

Inbound.MainInbound(inboundlist_path)

os.system("pause")