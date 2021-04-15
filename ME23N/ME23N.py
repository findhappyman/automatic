#!/usr/bin/python3
# -*- coding: utf-8 -*-

import subprocess as subp
import win32con
import win32clipboard as w

def overwrite_SAP(PO='',item='',**kwargs):
    content_var ='''
session.findById("wnd[0]").resizeWorkingPane 116,43,false
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = "{0}"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").key = "  {1}"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1106/ssubSUB2:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").setCurrentCell 2, "BELNR"

    '''.format(PO,item)

    return content_var

def getClipboard():
    w.OpenClipboard()
    t = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    return t

if __name__ == '__main__':
    content_fix = '''If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If
    '''

    Clipboard = getClipboard().split('\t')
    content_PO = Clipboard[0]
    content_item = Clipboard[1].split('\r')[0]
    content_SAP = overwrite_SAP(content_PO,content_item)
    content = content_fix + content_SAP
    Me23n_file = open('me23n.vbs', 'w')
    Me23n_file.write(content)
    Me23n_file.close()
    subp.call('cscript me23n.vbs')
