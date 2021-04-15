#!/usr/bin/python3
# -*- coding: utf-8 -*-

import subprocess as subp
import pandas as pd
import win32con
import win32clipboard as w

def overwrite_SAP(PN='',Qty='',**kwargs):
    content_var ='''
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:9304/subSUBSCREEN_TC:SAPMV45A:9305/tblSAPMV45AZZ_TCTRL_U_ERF_REKLA/ctxtRV45A-MABNR[1,0]").text = "{0}"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:9304/subSUBSCREEN_TC:SAPMV45A:9305/tblSAPMV45AZZ_TCTRL_U_ERF_REKLA/txtRV45A-KWMENG[2,0]").text = "{1}"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:9304/subSUBSCREEN_TC:SAPMV45A:9305/tblSAPMV45AZZ_TCTRL_U_ERF_REKLA/ctxtVBAP-PSTYV[3,0]").text = "ZKü1"
session.findById("wnd[0]").sendVKey 0
on error resume next
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:9304/subSUBSCREEN_TC:SAPMV45A:9305/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POLO").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
err.Clear
err.raise 6
'''.format(PN,Qty)

    return content_var

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
    pd.
    Clipboard = getClipboard().split('\t')
    content_PO = Clipboard[0]
    content_item = Clipboard[1].split('\r')[0]
    content_SAP = overwrite_SAP(content_PO,content_item)
    content = content_fix + content_SAP
    Me23n_file = open('../ME23N/me23n.vbs', 'w')
    Me23n_file.write(content)
    Me23n_file.close()
    subp.call('cscript me23n.vbs')