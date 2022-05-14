import win32com.client
import win32.lib.win32con as win32con
import win32con
import sys,os
import time
from time import sleep
import subprocess
import subprocess
import time
import openpyxl
from openpyxl import *
import re
import psycopg2
import os
import win32com.client as win32
from datetime import datetime
import collections
from PyQt5.QtWidgets import QApplication

###########################################################

if os.path.exists(r"E:\CWIP_Report\Dummy_folder\second_demo.xls"):
    os.remove(r"E:\CWIP_Report\Dummy_folder\second_demo.xls")
if os.path.exists(r"E:\CWIP_Report\Dummy_folder\second_demo.xlsx"):
    os.remove(r"E:\CWIP_Report\Dummy_folder\second_demo.xlsx")
    
###########################################################
    
def sap_connection(connection1):
    global session,session1
    error = 'SAP Connection Error'
    
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(10)
    SapGuiAuto = win32com.client.GetObject("SAPGUI")

    if not type(SapGuiAuto) == win32com.client.CDispatch:
            return error
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return error
    connection = application.OpenConnection(connection1, True)
    if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return error
    if connection.DisabledByServer == True:
            application = None
            SapGuiAuto = None
            return error
    session = connection.Children(0)
    session1 = session
    if not type(session) == win32com.client.CDispatch:
            connection = None

            application = None
            SapGuiAuto = None
            return error
    if session.Info.IsLowSpeedConnection == True:
            connection = None
            application = None
            SapGuiAuto = None
            return error
    return 'Success'

for i in range(5):
	sap_conn = sap_connection(connection1 = 'YOUR SAP SERVER HERE !')
	print(sap_conn)
	if sap_conn == 'Success':
		break
	else:
		pass

time.sleep(2)
print("SAP connection established..")

####################################################################
# Extracting data from excel
path=r"YOUR EXCEL FILE PATH HERE !"
wb_s=load_workbook(path)
ws=wb_s.active
unic = []
row = 0
for i in range(2,3412):
               
    code_li=ws['E'+str(i)].value
    #print(code_li)
    if code_li not in unic:
        unic.append(code_li)
print(unic)

for r in range(len(unic)):
    #print(r)
    r = str(r)
####################################################################

#For login process
def first():
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "YOUR USER ID HERE !"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "YOUR PASSWORD HERE !"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[0]/okcd").text = "YOUR T-CODE HERE !"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "YOUR VALUE HERE !"
    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").caretPosition = 12
    session.findById("wnd[1]").sendVKey(0)
##first()

# for putting rows from excel
def second():
##    session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = "CL/M0000/00029"
##    session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").caretPosition = 14
##    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = j
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").caretPosition = 19
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    sleep(3)
    
    r=j
    print(j,r)
    
    path_d = r"E:\CWIP_Report\Dummy_folder"   # FOLDER FOR DOWNLOADING FILE !
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = path_d
    r=r.replace('/','_')
    r=r.replace('.','_')
    file_name="tummy_%s.xls"%r
    print(file_name)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
    rep = session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").text
    err = "File E:\CWIP_Report\Dummy_folder\tummy_CL_M0000_00021.x already exists" # TO REPLACE THE FILE NAME SIMULTANEOUSLY !
    if rep == err:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        
    session.findById("wnd[1]").sendVKey(0)
    sleep(5) 
    ##########converting#############
    fname = r"E:\CWIP_Report\Dummy_folder\tummy_%s.xls"%r
    destin_file=r"E:\CWIP_Report\Dummy_folder\bummy\tummy_%s.xls"%r

    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(destin_file+"x", FileFormat = 51)
    wb.Close()
    excel.Application.Quit()
    #session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncn43n"
    session.findById("wnd[0]").sendVKey(0)

    
##second()

# for saving the file 
##def third():
##    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "E:\CWIP_Report\Dummy_folder"
##    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Second_demo.xls"
##    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
##    session.findById("wnd[1]").sendVKey(0)
##    sleep(5)
####third()
##
### for coming back after download
##def back():
##    session.findById("wnd[1]/tbar[0]/btn[0]").press()
##    session.findById("wnd[1]").close()
####    session.findById("wnd[0]/tbar[0]/btn[3]").press()
####    session.findById("wnd[0]/tbar[0]/btn[3]").press()
##    session.findById("wnd[0]").sendVKey(3)
####back()

###################################################################################
#for closing the t-code window
##def exit():
##    session.findById("wnd[0]").close()
##    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
##exit()
###################################################################################

first()
try:
    for j in unic:
        
        print(j)
        second()

except:
    print("Kuch to gadbad h dayaa ")
