'''
This script uses Python's `contextlib.contextmanager`
to allow for easier readability of code.

Credits to Stefan Schnell for setting up the baseline code.
https://blogs.sap.com/2017/09/19/how-to-use-sap-gui-scripting-inside-python-programming-language/
'''

from contextlib import contextmanager
import win32com.client

@contextmanager
def SAPSession(SAP_GUI_AUTO, conn=0, sess=0):
    if isinstance(SAP_GUI_AUTO, win32com.client.CDispatch):
        application = SAP_GUI_AUTO.GetScriptingEngine
        connection = application.Children(conn)
        session = connection.Children(sess)

    yield session

    session = None
    connection = None
    application = None
    SAP_GUI_AUTO = None

if __name__ == '__main__':
    
    SAPGUI = win32com.client.GetObject('SAPGUI')

    with SAPSession(SAPGUI) as session:
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = '/nSE16'
        session.findById("wnd[0]").sendVKey(0)