'''
This script utilizes `multiprocessing` library
for SAP GUI automation.
Credits to Stefan Schnell for setting up the baseline code.
https://blogs.sap.com/2017/09/19/how-to-use-sap-gui-scripting-inside-python-programming-language/
'''

from multiprocessing import Process

import win32com.client

def SAP_Init():
    SAP_GUI_AUTO = win32com.client.GetObject('SAPGUI')

    if isinstance(SAP_GUI_AUTO, win32com.client.CDispatch):
        application = SAP_GUI_AUTO.GetScriptingEngine
        connection = application.Children(0)
    return connection

def TASK_I():
    ''' Enter T-code SE16 '''

    print('TASK I started.')
    connection = SAP_Init()
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = '/nSE16'
    session.findById("wnd[0]").sendVKey(0)

    print('TASK I complete.')

def TASK_II():
    ''' Enter T-code FBL3N '''

    print('TASK II started.')
    connection = SAP_Init()
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = '/nFBL3N'
    session.findById("wnd[0]").sendVKey(0)

    print('TASK II complete.')

if __name__ == '__main__':
    try:
        window1 = Process(target = TASK_I)
        window2 = Process(target = TASK_II)

        window1.start()
        window2.start()

    except Exception as e:
        print(e)
