#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------
import sys, win32com.client
import time

#-Sub Main--------------------------------------------------------------
def Main():

  try:

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
      return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
      SapGuiAuto = None
      return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
      application = None
      SapGuiAuto = None
      return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
      connection = None
      application = None
      SapGuiAuto = None
      return

    #session.findById("wnd[0]").resizeWorkingPane(65, 19, 0)
    session.findById("wnd[0]/tbar[0]/okcd").text = "mm03"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text="01L-919-821-B"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP02").select()

  except:
    print(sys.exc_info()[0])

  finally:
    session = None
    connection = None
    application = None
    SapGuiAuto = None

#-Main------------------------------------------------------------------
if __name__ == "__main__":
  Main()

#-End-------------------------------------------------------------------