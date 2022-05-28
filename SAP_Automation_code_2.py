from pywinauto.application import Application
import win32gui
import win32con
import sys,os, win32com.client
import time
import subprocess
import subprocess
from openpyxl import *
import psycopg2 as ps
import pandas as pd

def sap_connection(connection1):
    print("xxSAP connection")
    global session,session1
    error = 'SAP Connection Error'
    
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"  # YOUR SAP PATH HERE
    subprocess.Popen(path)
    time.sleep(10)
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    print('YES')
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

for i in range(5):   # for establising the connection of SAP
	sap_conn = sap_connection(connection1 = ' YOUR SAP SERVER HERE ! ')
	print(sap_conn)
	if sap_conn == 'Success':
            
		break
	else:
		pass



#database
con = ps.connect(database="DATABASE_NAME_HERE", user="postgres", password="PASSWORD_HERE", host="localhost", port="5432")
cur = con.cursor()

cur.execute(" select invoice_number, bill_from, totalunitsbilled, vc, fixedcost, fixedcostpeak, carryingcost, others_, incometax, incentive, rrascharges, fixedcostoffpeak, fixedcostoffset, incentiveoffpeak, incentiveoffset, remarks2 from tata_power where mappingsheet='APCPL' ")
lis = cur.fetchall()

#print(lis)
con.commit()
cur.close()
con.close()

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "SAP_ID_HERE"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "SAP_PASSWORD_HERE"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/okcd").text = "SAP_T-CODE_HERE"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/btnNEW_BILL_ENTRY").press()
session.findById("wnd[0]/usr/cmbTYPE_TRANSACTION").key = "P"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/cmbWA_ZMMTR_PMG_HEADER-BILL_TYPE").key = "6"
session.findById("wnd[0]/usr/cmbWA_ZMMTR_PMG_HEADER-BILL_TYPE").setFocus()
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").text = "some_code_here"
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").setFocus()
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").caretPosition = 7
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-INV_DATE").text = "09.02.22"
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-INV_DATE").caretPosition = 8
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").text = "28.02.22"
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").setFocus()
session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").caretPosition = 8
session.findById("wnd[0]").sendVKey(0)


count=1
tcspurchase = ""
def function_1(j, invoice_number, bill_from, totalunitsbilled, vc, fixedcost,fixedcostpeak, carryingcost, others_, incometax, incentive, rrascharges, fixedcostoffpeak, fixedcostoffset, incentiveoffpeak, incentiveoffset, tcspurchase, remarks2):
    fixedcostpeak = fixedcost + fixedcostpeak

    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/cmbWA_ZMMTR_PMG_ITEM-PLANT_VEN_CODE[1,{}]".format(j)).key = "some_code_here"
    #InvoiceNumber
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-INVOICE_NUM[2,{}]".format(j)).text = invoice_number
    #BillFrom
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).text = bill_from
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)
    #UnitsBilled
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-UNIT_BILLED[8,{}]".format(j)).text = totalunitsbilled
    #VariableCost
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0002[9,{}]".format(j)).text = vc
    #FixedCostPeak
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0003[10,{}]".format(j)).text = fixedcostpeak
    #CarryingCost
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0004[11,{}]".format(j)).text = carryingcost
    #Others
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0027[12,{}]".format(j)).text = others_
    #IncomeTax
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0028[13,{}]".format(j)).text = incometax
    #Incentive
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0029[14,{}]".format(j)).text = incentive
    #RRASCharges
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0031[15,{}]".format(j)).text = rrascharges
    #FixedCOstOffPeak
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0033[16,{}]".format(j)).text = fixedcostoffpeak
    #FixedCostOffSet
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0034[17,{}]".format(j)).text = fixedcostoffset
    #IncentiveOffPeak
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0036[18,{}]".format(j)).text = incentiveoffpeak
    #IncentiveOffSet
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0037[19,{}]".format(j)).text = incentiveoffset
    #TCSPurchase
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0038[20,{}]".format(j)).text = tcspurchase
    #Remarks
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-REMARKS[23,{}]".format(j)).text = remarks2

    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-REMARKS[23,0]").setFocus()
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-REMARKS[23,0]").caretPosition = 6
    session.findById("wnd[0]").sendVKey(0)


for j in range (0,len(lis)):
    x = lis[j]
    invoice_number = x[0]
    bill_from = x[1].replace("-",".")
    totalunitsbilled = x[2]
    vc = x[3]
    fixedcost = x[4]
    fixedcostpeak = x[5]
    carryingcost = x[6]
    others_ = x[7]
    incometax = x[8]
    incentive = x[9]
    rrascharges = x[10]
    fixedcostoffpeak = x[11]
    fixedcostoffset = x[12]
    incentiveoffpeak = x[13]
    incentiveoffset = x[14]
    tcspurchase = ""
    remarks2 = x[15]
    
    #print(fixedcostpeak)
    if j>=9:
        j = 8
        time.sleep(3)
        session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM").verticalScrollbar.position = count
        count= count+1
        
    function_1(j,invoice_number, bill_from, totalunitsbilled, vc, fixedcost,fixedcostpeak, carryingcost, others_, incometax, incentive, rrascharges, fixedcostoffpeak, fixedcostoffset, incentiveoffpeak, incentiveoffset, tcspurchase, remarks2)



print (Mission Accomplished :))



















    
