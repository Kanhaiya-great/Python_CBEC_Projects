import sys
import pyodbc
import xlsxwriter

import subprocess as sp
sp.call('cls',shell=True)

print("                                         ------------------------------------------------------ ")
print("                                         |                                                     |")
print("                                         |       Monthly TRU Data - BED Specific               |")
print("                                         |                                                     |")
print("                                         -------------------------------------------------------")                                 

From_dt=input('Please enter From Date YYYYMMDD:')
To_dt=input('Please enter To Date YYYYMMDD:')

def getMonth(x):
        switcher={
                
                '01':'January',
                '02':'February',
                '03':'March',
                '04':'April ',
                '05':'May',
                '06':'June',
                '07':'July',
                '08':'August',
                '09':'September', 
                '10':'October',
                '11':'November' ,
                '12':'December'
                }     
        return switcher.get(x)
    
FromYear=From_dt[0:][:4]
FromMonth=getMonth(From_dt[4:][:2])
FromDay=From_dt[7:][:2]
ToYear=To_dt[0:][:4]
ToMonth=getMonth(To_dt[4:][:2])
ToDay=To_dt[6:][:2]

if (FromYear==ToYear and FromMonth==ToMonth):
    FileNm='Data_BED_Specific_'+FromMonth+FromYear+'.xlsx'
else:
    FileNm='Data_BED_Specific_'+FromMonth+FromYear+ '-'+ ToMonth+ToYear +'.xlsx'
    
#FileNm='Data_BED_Specific_'+FromMonth+FromYear+'.xlsx'
Path="C:\Python_TRU_Reports\Monthly\\"

print("Output FileName: "+Path+FileNm)

wb = xlsxwriter.Workbook(Path+FileNm)

################## Create "Note" Worksheet
#sheet1 = wb.add_worksheet('Note')
cell_format = wb.add_format()

cell_format = wb.add_format()

note_format = wb.add_format({
    'bold': 11,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
   'fg_color': '#B3B6B7',
    'font_size':'11'})
title_format = wb.add_format({
    'bold': 11,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#B3B6B7',
    'font_size':'11'})
ColumnHead_format = wb.add_format({
    'bold': 10,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#B3B6B7',
    'font_size':'10',
    'text_wrap':1})
Num_format = wb.add_format({'border': 1,'num_format':'0.00'})
date_format = wb.add_format({'border': 1,'num_format': 'dd-mm-yyyy'})
cell_format = wb.add_format({'border': 1,})

#sheet1.merge_range('B2:K2','Data has been extracted for the period mentioned in format sheet',note_format)

################## Create Data "Summary" Worksheet for CTH2
sheet2 = wb.add_worksheet('CTH2')

if (FromYear==ToYear and FromMonth==ToMonth):
    sheet2.merge_range('A1:BP1', 'Central Excise - CETH (2 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear   , title_format)
else:
    sheet2.merge_range('A1:BP1', 'Central Excise - CETH (2 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear + ' to ' +  ToMonth + '-' + ToYear , title_format)
    
    
sheet2.merge_range('A2:D2', '', title_format)
sheet2.merge_range('E2:BP2', '(In Rs. Crore)', title_format)
sheet2.merge_range('A3:C3', '', title_format)
sheet2.write('D3:D3', '(In Rs. Crore)', title_format)
sheet2.merge_range('E3:BK3', 'Total Duty Payable', title_format)
sheet2.merge_range('BL3:BM3', '', title_format)
sheet2.merge_range('BN3:BP3', 'Total Duty Foregone (In Rs. Crore)', title_format)

sheet2.write('A4:A4','S.No.',ColumnHead_format)
sheet2.set_column(0, 0, 5)
sheet2.write('B4:B4','CETH [2]',ColumnHead_format)
sheet2.write('C4:C4','Assessable Value',ColumnHead_format)
sheet2.write('D4:D4','DUTY_SPECFC',ColumnHead_format)
sheet2.write('E4:E4','Total Duty Payable',ColumnHead_format)
sheet2.write('F4:F4','ADC_LVD_CT_75',ColumnHead_format)
sheet2.write('G4:G4','ADC_LVD_CT_75_31',ColumnHead_format)
sheet2.write('H4:H4','ADC_LVD_CT_75_35',ColumnHead_format)
sheet2.write('I4:I4','ADE',ColumnHead_format)
sheet2.write('J4:J4','ADE_LVD_CL_85',ColumnHead_format)
sheet2.write('K4:K4','AED_GSI',ColumnHead_format)
sheet2.write('L4:L4','AED_TTA',ColumnHead_format)
sheet2.write('M4:M4','BCD',ColumnHead_format)
sheet2.write('N4:M4','CENVAT_D',ColumnHead_format)
sheet2.write('O4:O4','CESS_AUTOMBL',ColumnHead_format)
sheet2.write('P4:P4','CESS_BEEDI',ColumnHead_format)
sheet2.write('Q4:Q4','CESS_CASHEW',ColumnHead_format)
sheet2.write('R4:R4','CESS_CHROME',ColumnHead_format)
sheet2.write('S4:S4','CESS_COAL_COKE',ColumnHead_format)
sheet2.write('T4:T4','CESS_COFFEE',ColumnHead_format)
sheet2.write('U4:U4','CESS_COPRA',ColumnHead_format)
sheet2.write('V4:V4','CESS_COTN_FBRC',ColumnHead_format)
sheet2.write('W4:W4','CESS_COTTON',ColumnHead_format)
sheet2.write('X4:X4','CESS_CRUDEOIL',ColumnHead_format)
sheet2.write('Y4:Y4','CESS_FIBER',ColumnHead_format)
sheet2.write('Z4:Z4','CESS_FILMS',ColumnHead_format)
sheet2.write('AA4:AA4','CESS_IRON_ORE',ColumnHead_format)
sheet2.write('AB4:AB4','CESS_JUTE',ColumnHead_format)
sheet2.write('AC4:AC4','CESS_LAC',ColumnHead_format)
sheet2.write('AD4:AD4','CESS_LIME_DLMT',ColumnHead_format)
sheet2.write('AE4:AE4','CESS_MAGNSE',ColumnHead_format)
sheet2.write('AF4:AF4','CESS_MARINE',ColumnHead_format)
sheet2.write('AG4:AG4','CESS_MATCHES',ColumnHead_format)
sheet2.write('AH4:AH4','CESS_MEDICINAL',ColumnHead_format)
sheet2.write('AI4:AI4','CESS_MNMD_FBRC',ColumnHead_format)
sheet2.write('AJ4:AJ4','CESS_NATRL_GAS',ColumnHead_format)
sheet2.write('AK4:AK4','CESS_OIL',ColumnHead_format)
sheet2.write('AL4:AL4','CESS_OTHR_CMDT',ColumnHead_format)
sheet2.write('AM4:AM4','CESS_PAPER',ColumnHead_format)
sheet2.write('AN4:AM4','CESS_RAYON',ColumnHead_format)
sheet2.write('AO4:AO4','CESS_RUBBER',ColumnHead_format)
sheet2.write('AP4:AP4','CESS_SALT',ColumnHead_format)
sheet2.write('AQ4:AQ4','CESS_STRW_BRD',ColumnHead_format)
sheet2.write('AR4:AR4','CESS_SUGAR',ColumnHead_format)
sheet2.write('AS4:AS4','CESS_TEA',ColumnHead_format)
sheet2.write('AT4:AT4','CESS_TEXTILE',ColumnHead_format)
sheet2.write('AU4:AU4','CESS_TOBCCO',ColumnHead_format)
sheet2.write('AV4:AV4','CESS_VEG_OIL',ColumnHead_format)
sheet2.write('AW4:AW4','CESS_WOOLEN',ColumnHead_format)
sheet2.write('AX4:AX4','CLEAN_ENVIRONMENT_CESS',ColumnHead_format)
sheet2.write('AY4:AY4','CVD',ColumnHead_format)
sheet2.write('AZ4:AZ4','EDU_CESS',ColumnHead_format)
sheet2.write('BA4:BA4','EDU_CESS_ST',ColumnHead_format)
sheet2.write('BB4:BB4','EXPORT_DUTY',ColumnHead_format)
sheet2.write('BC4:BC4','INFRASTRUCTURE CESS',ColumnHead_format)
sheet2.write('BD4:BD4','NCCD',ColumnHead_format)
sheet2.write('BE4:BE4','OTHERS',ColumnHead_format)
sheet2.write('BF4:BF4','SAD',ColumnHead_format)
sheet2.write('BG4:BG4','SAED',ColumnHead_format)
sheet2.write('BH4:BH4','SEC_EDU_CESS',ColumnHead_format)
sheet2.write('BI4:BI4','SEC_EDU_CESS_ST',ColumnHead_format)
sheet2.write('BJ4:BJ4','SED',ColumnHead_format)
sheet2.write('BK4:BK4','SERVICE_TAX',ColumnHead_format)
sheet2.write('BL4:BL4','PLA',ColumnHead_format)
sheet2.write('BM4:BM4','CENVAT',ColumnHead_format)
sheet2.write('BN4:BM4','Total Duty Foregone',ColumnHead_format)
sheet2.write('BO4:BO4','Basic Excise Foregone',ColumnHead_format)
sheet2.write('BP4:BP4','Other Components of Duty Foregone',ColumnHead_format)							
sheet2.set_column(1, 67, 9)

											
print("Connecting to SmartView for CTH2")
cnxn = pyodbc.connect("DSN=smartview_w1")
cursor1 = cnxn.cursor()
print("Running query for CTH2")
cursor1.execute("SELECT C.CETH_NO_2 CETH_NO, SUM(B.ASBL_VALUE)/10000000 AS ASBL_VALUE, B.DUTY_SPECFC AS DUTY_SPECFC, SUM(B.DUTY_PAYBL)/10000000 AS DUTY_PAYBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 54 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 58 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_31, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 57 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_35, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 13 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  6 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE_LVD_CL_85, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  9 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_GSI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 11 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_TTA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  1 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS BCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  7 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CENVAT_D, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 22 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_AUTOMBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 17 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_BEEDI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 41 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CASHEW, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 42 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CHROME, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 34 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COAL_COKE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 31 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COFFEE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 27 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COPRA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 49 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTN_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 20 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTTON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 18 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CRUDEOIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 43 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FIBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 40 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FILMS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 35 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_IRON_ORE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 21 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_JUTE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 45 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LAC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 36 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LIME_DLMT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 38 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MAGNSE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 46 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MARINE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 51 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MATCHES, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 44 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MEDICINAL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 50 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MNMD_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 53 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_NATRL_GAS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 29 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 52 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OTHR_CMDT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 19 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_PAPER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 47 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RAYON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 30 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RUBBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 37 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SALT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 32 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_STRW_BRD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 16 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SUGAR, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 15 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 39 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEXTILE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 28 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TOBCCO, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 33 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_VEG_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 48 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_WOOLEN, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 56 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CLEAN_ENVIRONMENT_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  2 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CVD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 14 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 25 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  4 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EXPORT_DUTY, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 55 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [INFRASTRUCTURE CESS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 12 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS NCCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  5 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [OTHERS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  3 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 23 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 24 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 26 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  8 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 10 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SERVICE_TAX, SUM(ISNULL(B.DUTY_PAID_ACNT_CURNT,0))/10000000 AS PLA , SUM(ISNULL(B.DUTY_PAID_ACNT_CREDT,0))/10000000 AS CENVAT, SUM(ISNULL(B.DUTY_FOREGONE,0))/10000000 AS DUTY_FOREGONE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 7 THEN B.DUTY_FOREGONE ELSE 0 END)/10000000 AS Basic_Excise_Foregone, DUTY_FOREGONE - Basic_Excise_Foregone AS Other_Components_of_Duty_Foregone FROM DWUSER.DIM_COM_DUTY_HEAD_T A,DWUSER.FACT_CE_ITEM_DUTY_PAYBALE_T B,DWUSER.DIM_COM_CETSH_T C WHERE A.DW_DUTY_HEAD_KEY = B.DW_DUTY_HEAD_KEY AND B.RETRN_MONTH_YEAR BETWEEN " +From_dt+  " AND "+To_dt+ " AND C.DW_CETSH_KEY = B.DW_CETSH_KEY AND C.CETH_NO_2 <> 'XX' GROUP BY CETH_NO,DUTY_SPECFC ")
print("Writing data to output file for CTH2")

cnt=0
rno=4
for r in cursor1:
    cnt = cnt + 1
    sheet2.write(rno, 0, cnt,cell_format)
    for c in range(67):
        if (c >= 0 and c <= 67):
            
            if(r[c]==None):
                sheet2.write(rno, c+1, '--',Num_format)
            else:
                sheet2.write(rno, c+1, r[c],Num_format)
        else:
            sheet2.write(rno, c+1, r[c],cell_format)
    rno = rno + 1

cursor1.close()
cnxn.close()

################## Create Data "Summary" Worksheet for CTH4
sheet3 = wb.add_worksheet('CTH4')

if (FromYear==ToYear and FromMonth==ToMonth):
    sheet3.merge_range('A1:BP1', 'Central Excise - CETH (4 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear   , title_format)
else:
    sheet3.merge_range('A1:BP1', 'Central Excise - CETH (4 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear + ' to ' +  ToMonth + '-' + ToYear  , title_format)
    
sheet3.merge_range('A2:D2', '', title_format)
sheet3.merge_range('E2:BP2', '(In Rs. Crore)', title_format)
sheet3.merge_range('A3:C3', '', title_format)
sheet3.write('D3:D3', '(In Rs. Crore)', title_format)
sheet3.merge_range('E3:BK3', 'Total Duty Payable', title_format)
sheet3.merge_range('BL3:BM3', '', title_format)
sheet3.merge_range('BN3:BP3', 'Total Duty Foregone (In Rs. Crore)', title_format)

sheet3.write('A4:A4','S.No.',ColumnHead_format)
sheet3.set_column(0, 0, 5)
sheet3.write('B4:B4','CETH [4]',ColumnHead_format)
sheet3.write('C4:C4','Assessable Value',ColumnHead_format)
sheet3.write('D4:D4','DUTY_SPECFC',ColumnHead_format)
sheet3.write('E4:E4','Total Duty Payable',ColumnHead_format)
sheet3.write('F4:F4','ADC_LVD_CT_75',ColumnHead_format)
sheet3.write('G4:G4','ADC_LVD_CT_75_31',ColumnHead_format)
sheet3.write('H4:H4','ADC_LVD_CT_75_35',ColumnHead_format)
sheet3.write('I4:I4','ADE',ColumnHead_format)
sheet3.write('J4:J4','ADE_LVD_CL_85',ColumnHead_format)
sheet3.write('K4:K4','AED_GSI',ColumnHead_format)
sheet3.write('L4:L4','AED_TTA',ColumnHead_format)
sheet3.write('M4:M4','BCD',ColumnHead_format)
sheet3.write('N4:M4','CENVAT_D',ColumnHead_format)
sheet3.write('O4:O4','CESS_AUTOMBL',ColumnHead_format)
sheet3.write('P4:P4','CESS_BEEDI',ColumnHead_format)
sheet3.write('Q4:Q4','CESS_CASHEW',ColumnHead_format)
sheet3.write('R4:R4','CESS_CHROME',ColumnHead_format)
sheet3.write('S4:S4','CESS_COAL_COKE',ColumnHead_format)
sheet3.write('T4:T4','CESS_COFFEE',ColumnHead_format)
sheet3.write('U4:U4','CESS_COPRA',ColumnHead_format)
sheet3.write('V4:V4','CESS_COTN_FBRC',ColumnHead_format)
sheet3.write('W4:W4','CESS_COTTON',ColumnHead_format)
sheet3.write('X4:X4','CESS_CRUDEOIL',ColumnHead_format)
sheet3.write('Y4:Y4','CESS_FIBER',ColumnHead_format)
sheet3.write('Z4:Z4','CESS_FILMS',ColumnHead_format)
sheet3.write('AA4:AA4','CESS_IRON_ORE',ColumnHead_format)
sheet3.write('AB4:AB4','CESS_JUTE',ColumnHead_format)
sheet3.write('AC4:AC4','CESS_LAC',ColumnHead_format)
sheet3.write('AD4:AD4','CESS_LIME_DLMT',ColumnHead_format)
sheet3.write('AE4:AE4','CESS_MAGNSE',ColumnHead_format)
sheet3.write('AF4:AF4','CESS_MARINE',ColumnHead_format)
sheet3.write('AG4:AG4','CESS_MATCHES',ColumnHead_format)
sheet3.write('AH4:AH4','CESS_MEDICINAL',ColumnHead_format)
sheet3.write('AI4:AI4','CESS_MNMD_FBRC',ColumnHead_format)
sheet3.write('AJ4:AJ4','CESS_NATRL_GAS',ColumnHead_format)
sheet3.write('AK4:AK4','CESS_OIL',ColumnHead_format)
sheet3.write('AL4:AL4','CESS_OTHR_CMDT',ColumnHead_format)
sheet3.write('AM4:AM4','CESS_PAPER',ColumnHead_format)
sheet3.write('AN4:AM4','CESS_RAYON',ColumnHead_format)
sheet3.write('AO4:AO4','CESS_RUBBER',ColumnHead_format)
sheet3.write('AP4:AP4','CESS_SALT',ColumnHead_format)
sheet3.write('AQ4:AQ4','CESS_STRW_BRD',ColumnHead_format)
sheet3.write('AR4:AR4','CESS_SUGAR',ColumnHead_format)
sheet3.write('AS4:AS4','CESS_TEA',ColumnHead_format)
sheet3.write('AT4:AT4','CESS_TEXTILE',ColumnHead_format)
sheet3.write('AU4:AU4','CESS_TOBCCO',ColumnHead_format)
sheet3.write('AV4:AV4','CESS_VEG_OIL',ColumnHead_format)
sheet3.write('AW4:AW4','CESS_WOOLEN',ColumnHead_format)
sheet3.write('AX4:AX4','CLEAN_ENVIRONMENT_CESS',ColumnHead_format)
sheet3.write('AY4:AY4','CVD',ColumnHead_format)
sheet3.write('AZ4:AZ4','EDU_CESS',ColumnHead_format)
sheet3.write('BA4:BA4','EDU_CESS_ST',ColumnHead_format)
sheet3.write('BB4:BB4','EXPORT_DUTY',ColumnHead_format)
sheet3.write('BC4:BC4','INFRASTRUCTURE CESS',ColumnHead_format)
sheet3.write('BD4:BD4','NCCD',ColumnHead_format)
sheet3.write('BE4:BE4','OTHERS',ColumnHead_format)
sheet3.write('BF4:BF4','SAD',ColumnHead_format)
sheet3.write('BG4:BG4','SAED',ColumnHead_format)
sheet3.write('BH4:BH4','SEC_EDU_CESS',ColumnHead_format)
sheet3.write('BI4:BI4','SEC_EDU_CESS_ST',ColumnHead_format)
sheet3.write('BJ4:BJ4','SED',ColumnHead_format)
sheet3.write('BK4:BK4','SERVICE_TAX',ColumnHead_format)
sheet3.write('BL4:BL4','PLA',ColumnHead_format)
sheet3.write('BM4:BM4','CENVAT',ColumnHead_format)
sheet3.write('BN4:BM4','Total Duty Foregone',ColumnHead_format)
sheet3.write('BO4:BO4','Basic Excise Foregone',ColumnHead_format)
sheet3.write('BP4:BP4','Other Components of Duty Foregone',ColumnHead_format)							
sheet3.set_column(1, 67, 9)

print("Connecting to SmartView for CTH4")
cnxn = pyodbc.connect("DSN=smartview_w1")
cursor1 = cnxn.cursor()
print("Running query for CTH4")
cursor1.execute("SELECT C.CETH_NO_4 CETH_NO, SUM(B.ASBL_VALUE)/10000000 AS ASBL_VALUE, B.DUTY_SPECFC AS DUTY_SPECFC, SUM(B.DUTY_PAYBL)/10000000 AS DUTY_PAYBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 54 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 58 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_31, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 57 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_35, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 13 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  6 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE_LVD_CL_85, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  9 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_GSI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 11 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_TTA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  1 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS BCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  7 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CENVAT_D, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 22 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_AUTOMBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 17 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_BEEDI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 41 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CASHEW, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 42 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CHROME, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 34 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COAL_COKE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 31 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COFFEE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 27 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COPRA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 49 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTN_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 20 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTTON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 18 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CRUDEOIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 43 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FIBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 40 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FILMS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 35 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_IRON_ORE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 21 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_JUTE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 45 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LAC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 36 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LIME_DLMT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 38 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MAGNSE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 46 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MARINE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 51 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MATCHES, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 44 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MEDICINAL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 50 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MNMD_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 53 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_NATRL_GAS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 29 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 52 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OTHR_CMDT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 19 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_PAPER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 47 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RAYON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 30 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RUBBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 37 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SALT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 32 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_STRW_BRD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 16 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SUGAR, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 15 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 39 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEXTILE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 28 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TOBCCO, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 33 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_VEG_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 48 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_WOOLEN, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 56 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CLEAN_ENVIRONMENT_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  2 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CVD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 14 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 25 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  4 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EXPORT_DUTY, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 55 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [INFRASTRUCTURE CESS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 12 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS NCCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  5 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [OTHERS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  3 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 23 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 24 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 26 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  8 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 10 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SERVICE_TAX, SUM(ISNULL(B.DUTY_PAID_ACNT_CURNT,0))/10000000 AS PLA , SUM(ISNULL(B.DUTY_PAID_ACNT_CREDT,0))/10000000 AS CENVAT, SUM(ISNULL(B.DUTY_FOREGONE,0))/10000000 AS DUTY_FOREGONE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 7 THEN B.DUTY_FOREGONE ELSE 0 END)/10000000 AS Basic_Excise_Foregone, DUTY_FOREGONE - Basic_Excise_Foregone AS Other_Components_of_Duty_Foregone FROM DWUSER.DIM_COM_DUTY_HEAD_T A,DWUSER.FACT_CE_ITEM_DUTY_PAYBALE_T B,DWUSER.DIM_COM_CETSH_T C WHERE A.DW_DUTY_HEAD_KEY = B.DW_DUTY_HEAD_KEY AND B.RETRN_MONTH_YEAR BETWEEN " +From_dt+  " AND "+To_dt+ " AND C.DW_CETSH_KEY = B.DW_CETSH_KEY AND C.CETH_NO_2 <> 'XX' GROUP BY CETH_NO,DUTY_SPECFC ")
print("Writing data to output file for CTH4")

cnt=0
rno=4
for r in cursor1:
    cnt = cnt + 1
    sheet3.write(rno, 0, cnt,cell_format)
    for c in range(67):
        if (c >= 0 and c <= 67):
            
            if(r[c]==None):
                sheet3.write(rno, c+1, '--',Num_format)
            else:
                sheet3.write(rno, c+1, r[c],Num_format)
        else:
            sheet3.write(rno, c+1, r[c],cell_format)
    rno = rno + 1

cursor1.close()
cnxn.close()

################## Create Data "Summary" Worksheet for CTH6
sheet4 = wb.add_worksheet('CTH6')

if (FromYear==ToYear and FromMonth==ToMonth):
    sheet4.merge_range('A1:BP1', 'Central Excise - CETH (6 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear   , title_format)
else:
    sheet4.merge_range('A1:BP1', 'Central Excise - CETH (6 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear + ' to ' +  ToMonth + '-' + ToYear  , title_format)
    
sheet4.merge_range('A2:D2', '', title_format)
sheet4.merge_range('E2:BP2', '(In Rs. Crore)', title_format)
sheet4.merge_range('A3:C3', '', title_format)
sheet4.write('D3:D3', '(In Rs. Crore)', title_format)
sheet4.merge_range('E3:BK3', 'Total Duty Payable', title_format)
sheet4.merge_range('BL3:BM3', '', title_format)
sheet4.merge_range('BN3:BP3', 'Total Duty Foregone (In Rs. Crore)', title_format)

sheet4.write('A4:A4','S.No.',ColumnHead_format)
sheet4.set_column(0, 0, 5)
sheet4.write('B4:B4','CETH [6]',ColumnHead_format)
sheet4.write('C4:C4','Assessable Value',ColumnHead_format)
sheet4.write('D4:D4','DUTY_SPECFC',ColumnHead_format)
sheet4.write('E4:E4','Total Duty Payable',ColumnHead_format)
sheet4.write('F4:F4','ADC_LVD_CT_75',ColumnHead_format)
sheet4.write('G4:G4','ADC_LVD_CT_75_31',ColumnHead_format)
sheet4.write('H4:H4','ADC_LVD_CT_75_35',ColumnHead_format)
sheet4.write('I4:I4','ADE',ColumnHead_format)
sheet4.write('J4:J4','ADE_LVD_CL_85',ColumnHead_format)
sheet4.write('K4:K4','AED_GSI',ColumnHead_format)
sheet4.write('L4:L4','AED_TTA',ColumnHead_format)
sheet4.write('M4:M4','BCD',ColumnHead_format)
sheet4.write('N4:M4','CENVAT_D',ColumnHead_format)
sheet4.write('O4:O4','CESS_AUTOMBL',ColumnHead_format)
sheet4.write('P4:P4','CESS_BEEDI',ColumnHead_format)
sheet4.write('Q4:Q4','CESS_CASHEW',ColumnHead_format)
sheet4.write('R4:R4','CESS_CHROME',ColumnHead_format)
sheet4.write('S4:S4','CESS_COAL_COKE',ColumnHead_format)
sheet4.write('T4:T4','CESS_COFFEE',ColumnHead_format)
sheet4.write('U4:U4','CESS_COPRA',ColumnHead_format)
sheet4.write('V4:V4','CESS_COTN_FBRC',ColumnHead_format)
sheet4.write('W4:W4','CESS_COTTON',ColumnHead_format)
sheet4.write('X4:X4','CESS_CRUDEOIL',ColumnHead_format)
sheet4.write('Y4:Y4','CESS_FIBER',ColumnHead_format)
sheet4.write('Z4:Z4','CESS_FILMS',ColumnHead_format)
sheet4.write('AA4:AA4','CESS_IRON_ORE',ColumnHead_format)
sheet4.write('AB4:AB4','CESS_JUTE',ColumnHead_format)
sheet4.write('AC4:AC4','CESS_LAC',ColumnHead_format)
sheet4.write('AD4:AD4','CESS_LIME_DLMT',ColumnHead_format)
sheet4.write('AE4:AE4','CESS_MAGNSE',ColumnHead_format)
sheet4.write('AF4:AF4','CESS_MARINE',ColumnHead_format)
sheet4.write('AG4:AG4','CESS_MATCHES',ColumnHead_format)
sheet4.write('AH4:AH4','CESS_MEDICINAL',ColumnHead_format)
sheet4.write('AI4:AI4','CESS_MNMD_FBRC',ColumnHead_format)
sheet4.write('AJ4:AJ4','CESS_NATRL_GAS',ColumnHead_format)
sheet4.write('AK4:AK4','CESS_OIL',ColumnHead_format)
sheet4.write('AL4:AL4','CESS_OTHR_CMDT',ColumnHead_format)
sheet4.write('AM4:AM4','CESS_PAPER',ColumnHead_format)
sheet4.write('AN4:AM4','CESS_RAYON',ColumnHead_format)
sheet4.write('AO4:AO4','CESS_RUBBER',ColumnHead_format)
sheet4.write('AP4:AP4','CESS_SALT',ColumnHead_format)
sheet4.write('AQ4:AQ4','CESS_STRW_BRD',ColumnHead_format)
sheet4.write('AR4:AR4','CESS_SUGAR',ColumnHead_format)
sheet4.write('AS4:AS4','CESS_TEA',ColumnHead_format)
sheet4.write('AT4:AT4','CESS_TEXTILE',ColumnHead_format)
sheet4.write('AU4:AU4','CESS_TOBCCO',ColumnHead_format)
sheet4.write('AV4:AV4','CESS_VEG_OIL',ColumnHead_format)
sheet4.write('AW4:AW4','CESS_WOOLEN',ColumnHead_format)
sheet4.write('AX4:AX4','CLEAN_ENVIRONMENT_CESS',ColumnHead_format)
sheet4.write('AY4:AY4','CVD',ColumnHead_format)
sheet4.write('AZ4:AZ4','EDU_CESS',ColumnHead_format)
sheet4.write('BA4:BA4','EDU_CESS_ST',ColumnHead_format)
sheet4.write('BB4:BB4','EXPORT_DUTY',ColumnHead_format)
sheet4.write('BC4:BC4','INFRASTRUCTURE CESS',ColumnHead_format)
sheet4.write('BD4:BD4','NCCD',ColumnHead_format)
sheet4.write('BE4:BE4','OTHERS',ColumnHead_format)
sheet4.write('BF4:BF4','SAD',ColumnHead_format)
sheet4.write('BG4:BG4','SAED',ColumnHead_format)
sheet4.write('BH4:BH4','SEC_EDU_CESS',ColumnHead_format)
sheet4.write('BI4:BI4','SEC_EDU_CESS_ST',ColumnHead_format)
sheet4.write('BJ4:BJ4','SED',ColumnHead_format)
sheet4.write('BK4:BK4','SERVICE_TAX',ColumnHead_format)
sheet4.write('BL4:BL4','PLA',ColumnHead_format)
sheet4.write('BM4:BM4','CENVAT',ColumnHead_format)
sheet4.write('BN4:BM4','Total Duty Foregone',ColumnHead_format)
sheet4.write('BO4:BO4','Basic Excise Foregone',ColumnHead_format)
sheet4.write('BP4:BP4','Other Components of Duty Foregone',ColumnHead_format)
sheet4.set_column(1, 67, 9)


print("Connecting to SmartView for CTH6")
cnxn = pyodbc.connect("DSN=smartview_w1")
cursor1 = cnxn.cursor()
print("Running query for CTH6")
cursor1.execute("SELECT C.CETH_NO_6 CETH_NO, SUM(B.ASBL_VALUE)/10000000 AS ASBL_VALUE, B.DUTY_SPECFC AS DUTY_SPECFC, SUM(B.DUTY_PAYBL)/10000000 AS DUTY_PAYBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 54 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 58 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_31, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 57 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_35, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 13 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  6 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE_LVD_CL_85, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  9 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_GSI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 11 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_TTA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  1 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS BCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  7 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CENVAT_D, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 22 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_AUTOMBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 17 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_BEEDI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 41 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CASHEW, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 42 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CHROME, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 34 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COAL_COKE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 31 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COFFEE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 27 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COPRA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 49 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTN_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 20 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTTON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 18 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CRUDEOIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 43 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FIBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 40 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FILMS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 35 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_IRON_ORE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 21 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_JUTE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 45 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LAC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 36 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LIME_DLMT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 38 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MAGNSE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 46 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MARINE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 51 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MATCHES, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 44 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MEDICINAL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 50 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MNMD_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 53 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_NATRL_GAS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 29 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 52 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OTHR_CMDT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 19 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_PAPER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 47 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RAYON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 30 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RUBBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 37 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SALT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 32 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_STRW_BRD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 16 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SUGAR, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 15 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 39 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEXTILE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 28 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TOBCCO, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 33 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_VEG_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 48 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_WOOLEN, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 56 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CLEAN_ENVIRONMENT_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  2 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CVD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 14 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 25 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  4 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EXPORT_DUTY, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 55 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [INFRASTRUCTURE CESS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 12 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS NCCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  5 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [OTHERS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  3 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 23 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 24 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 26 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  8 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 10 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SERVICE_TAX, SUM(ISNULL(B.DUTY_PAID_ACNT_CURNT,0))/10000000 AS PLA , SUM(ISNULL(B.DUTY_PAID_ACNT_CREDT,0))/10000000 AS CENVAT, SUM(ISNULL(B.DUTY_FOREGONE,0))/10000000 AS DUTY_FOREGONE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 7 THEN B.DUTY_FOREGONE ELSE 0 END)/10000000 AS Basic_Excise_Foregone, DUTY_FOREGONE - Basic_Excise_Foregone AS Other_Components_of_Duty_Foregone FROM DWUSER.DIM_COM_DUTY_HEAD_T A,DWUSER.FACT_CE_ITEM_DUTY_PAYBALE_T B,DWUSER.DIM_COM_CETSH_T C WHERE A.DW_DUTY_HEAD_KEY = B.DW_DUTY_HEAD_KEY AND B.RETRN_MONTH_YEAR BETWEEN " +From_dt+  " AND "+To_dt+ " AND C.DW_CETSH_KEY = B.DW_CETSH_KEY AND C.CETH_NO_2 <> 'XX' GROUP BY CETH_NO,DUTY_SPECFC ")
print("Writing data to output file for CTH6")

cnt=0
rno=4
for r in cursor1:
    cnt = cnt + 1
    sheet4.write(rno, 0, cnt,cell_format)
    for c in range(67):
        if (c >= 0 and c <= 67):
            
            if(r[c]==None):
                sheet4.write(rno, c+1, '--',Num_format)
            else:
                sheet4.write(rno, c+1, r[c],Num_format)
        else:
            sheet4.write(rno, c+1, r[c],cell_format)
    rno = rno + 1

cursor1.close()
cnxn.close()


################## Create Data "Summary" Worksheet for CTH8
sheet5 = wb.add_worksheet('CTH8')

if (FromYear==ToYear and FromMonth==ToMonth):
    sheet5.merge_range('A1:BP1', 'Central Excise - CETH (8 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear   , title_format)
else:
    sheet5.merge_range('A1:BP1', 'Central Excise - CETH (8 Digit) wise Duty Details for ' + FromMonth + '-' + FromYear + ' to ' +  ToMonth + '-' + ToYear  , title_format)

sheet5.merge_range('A2:D2', '', title_format)
sheet5.merge_range('E2:BP2', '(In Rs. Crore)', title_format)
sheet5.merge_range('A3:C3', '', title_format)
sheet5.write('D3:D3', '(In Rs. Crore)', title_format)
sheet5.merge_range('E3:BK3', 'Total Duty Payable', title_format)
sheet5.merge_range('BL3:BM3', '', title_format)
sheet5.merge_range('BN3:BP3', 'Total Duty Foregone (In Rs. Crore)', title_format)

sheet5.write('A4:A4','S.No.',ColumnHead_format)
sheet5.set_column(0, 0, 5)
sheet5.write('B4:B4','CETH [8]',ColumnHead_format)
sheet5.write('C4:C4','Assessable Value',ColumnHead_format)
sheet5.write('D4:D4','DUTY_SPECFC',ColumnHead_format)
sheet5.write('E4:E4','Total Duty Payable',ColumnHead_format)
sheet5.write('F4:F4','ADC_LVD_CT_75',ColumnHead_format)
sheet5.write('G4:G4','ADC_LVD_CT_75_31',ColumnHead_format)
sheet5.write('H4:H4','ADC_LVD_CT_75_35',ColumnHead_format)
sheet5.write('I4:I4','ADE',ColumnHead_format)
sheet5.write('J4:J4','ADE_LVD_CL_85',ColumnHead_format)
sheet5.write('K4:K4','AED_GSI',ColumnHead_format)
sheet5.write('L4:L4','AED_TTA',ColumnHead_format)
sheet5.write('M4:M4','BCD',ColumnHead_format)
sheet5.write('N4:M4','CENVAT_D',ColumnHead_format)
sheet5.write('O4:O4','CESS_AUTOMBL',ColumnHead_format)
sheet5.write('P4:P4','CESS_BEEDI',ColumnHead_format)
sheet5.write('Q4:Q4','CESS_CASHEW',ColumnHead_format)
sheet5.write('R4:R4','CESS_CHROME',ColumnHead_format)
sheet5.write('S4:S4','CESS_COAL_COKE',ColumnHead_format)
sheet5.write('T4:T4','CESS_COFFEE',ColumnHead_format)
sheet5.write('U4:U4','CESS_COPRA',ColumnHead_format)
sheet5.write('V4:V4','CESS_COTN_FBRC',ColumnHead_format)
sheet5.write('W4:W4','CESS_COTTON',ColumnHead_format)
sheet5.write('X4:X4','CESS_CRUDEOIL',ColumnHead_format)
sheet5.write('Y4:Y4','CESS_FIBER',ColumnHead_format)
sheet5.write('Z4:Z4','CESS_FILMS',ColumnHead_format)
sheet5.write('AA4:AA4','CESS_IRON_ORE',ColumnHead_format)
sheet5.write('AB4:AB4','CESS_JUTE',ColumnHead_format)
sheet5.write('AC4:AC4','CESS_LAC',ColumnHead_format)
sheet5.write('AD4:AD4','CESS_LIME_DLMT',ColumnHead_format)
sheet5.write('AE4:AE4','CESS_MAGNSE',ColumnHead_format)
sheet5.write('AF4:AF4','CESS_MARINE',ColumnHead_format)
sheet5.write('AG4:AG4','CESS_MATCHES',ColumnHead_format)
sheet5.write('AH4:AH4','CESS_MEDICINAL',ColumnHead_format)
sheet5.write('AI4:AI4','CESS_MNMD_FBRC',ColumnHead_format)
sheet5.write('AJ4:AJ4','CESS_NATRL_GAS',ColumnHead_format)
sheet5.write('AK4:AK4','CESS_OIL',ColumnHead_format)
sheet5.write('AL4:AL4','CESS_OTHR_CMDT',ColumnHead_format)
sheet5.write('AM4:AM4','CESS_PAPER',ColumnHead_format)
sheet5.write('AN4:AM4','CESS_RAYON',ColumnHead_format)
sheet5.write('AO4:AO4','CESS_RUBBER',ColumnHead_format)
sheet5.write('AP4:AP4','CESS_SALT',ColumnHead_format)
sheet5.write('AQ4:AQ4','CESS_STRW_BRD',ColumnHead_format)
sheet5.write('AR4:AR4','CESS_SUGAR',ColumnHead_format)
sheet5.write('AS4:AS4','CESS_TEA',ColumnHead_format)
sheet5.write('AT4:AT4','CESS_TEXTILE',ColumnHead_format)
sheet5.write('AU4:AU4','CESS_TOBCCO',ColumnHead_format)
sheet5.write('AV4:AV4','CESS_VEG_OIL',ColumnHead_format)
sheet5.write('AW4:AW4','CESS_WOOLEN',ColumnHead_format)
sheet5.write('AX4:AX4','CLEAN_ENVIRONMENT_CESS',ColumnHead_format)
sheet5.write('AY4:AY4','CVD',ColumnHead_format)
sheet5.write('AZ4:AZ4','EDU_CESS',ColumnHead_format)
sheet5.write('BA4:BA4','EDU_CESS_ST',ColumnHead_format)
sheet5.write('BB4:BB4','EXPORT_DUTY',ColumnHead_format)
sheet5.write('BC4:BC4','INFRASTRUCTURE CESS',ColumnHead_format)
sheet5.write('BD4:BD4','NCCD',ColumnHead_format)
sheet5.write('BE4:BE4','OTHERS',ColumnHead_format)
sheet5.write('BF4:BF4','SAD',ColumnHead_format)
sheet5.write('BG4:BG4','SAED',ColumnHead_format)
sheet5.write('BH4:BH4','SEC_EDU_CESS',ColumnHead_format)
sheet5.write('BI4:BI4','SEC_EDU_CESS_ST',ColumnHead_format)
sheet5.write('BJ4:BJ4','SED',ColumnHead_format)
sheet5.write('BK4:BK4','SERVICE_TAX',ColumnHead_format)
sheet5.write('BL4:BL4','PLA',ColumnHead_format)
sheet5.write('BM4:BM4','CENVAT',ColumnHead_format)
sheet5.write('BN4:BM4','Total Duty Foregone',ColumnHead_format)
sheet5.write('BO4:BO4','Basic Excise Foregone',ColumnHead_format)
sheet5.write('BP4:BP4','Other Components of Duty Foregone',ColumnHead_format)
sheet5.set_column(1, 67, 9)

print("Connecting to SmartView for CTH8")
cnxn = pyodbc.connect("DSN=smartview_w1")
cursor1 = cnxn.cursor()
print("Running query for CTH8")
cursor1.execute("SELECT C.CETH_NO_8 CETH_NO, SUM(B.ASBL_VALUE)/10000000 AS ASBL_VALUE, B.DUTY_SPECFC AS DUTY_SPECFC, SUM(B.DUTY_PAYBL)/10000000 AS DUTY_PAYBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 54 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 58 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_31, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 57 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADC_LVD_CT_75_35, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 13 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  6 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS ADE_LVD_CL_85, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  9 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_GSI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 11 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS AED_TTA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  1 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS BCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  7 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CENVAT_D, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 22 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_AUTOMBL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 17 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_BEEDI, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 41 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CASHEW, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 42 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CHROME, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 34 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COAL_COKE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 31 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COFFEE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 27 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COPRA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 49 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTN_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 20 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_COTTON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 18 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_CRUDEOIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 43 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FIBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 40 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_FILMS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 35 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_IRON_ORE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 21 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_JUTE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 45 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LAC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 36 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_LIME_DLMT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 38 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MAGNSE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 46 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MARINE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 51 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MATCHES, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 44 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MEDICINAL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 50 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_MNMD_FBRC, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 53 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_NATRL_GAS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 29 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 52 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_OTHR_CMDT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 19 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_PAPER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 47 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RAYON, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 30 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_RUBBER, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 37 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SALT, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 32 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_STRW_BRD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 16 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_SUGAR, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 15 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEA, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 39 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TEXTILE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 28 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_TOBCCO, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 33 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_VEG_OIL, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 48 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CESS_WOOLEN, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 56 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CLEAN_ENVIRONMENT_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  2 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS CVD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 14 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 25 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  4 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS EXPORT_DUTY, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 55 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [INFRASTRUCTURE CESS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 12 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS NCCD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  5 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS [OTHERS], SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  3 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAD, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 23 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SAED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 24 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 26 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SEC_EDU_CESS_ST, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY =  8 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SED, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 10 THEN B.DUTY_PAYBL ELSE 0 END)/10000000 AS SERVICE_TAX, SUM(ISNULL(B.DUTY_PAID_ACNT_CURNT,0))/10000000 AS PLA , SUM(ISNULL(B.DUTY_PAID_ACNT_CREDT,0))/10000000 AS CENVAT, SUM(ISNULL(B.DUTY_FOREGONE,0))/10000000 AS DUTY_FOREGONE, SUM(CASE WHEN A.DW_DUTY_HEAD_KEY = 7 THEN B.DUTY_FOREGONE ELSE 0 END)/10000000 AS Basic_Excise_Foregone, DUTY_FOREGONE - Basic_Excise_Foregone AS Other_Components_of_Duty_Foregone FROM DWUSER.DIM_COM_DUTY_HEAD_T A,DWUSER.FACT_CE_ITEM_DUTY_PAYBALE_T B,DWUSER.DIM_COM_CETSH_T C WHERE A.DW_DUTY_HEAD_KEY = B.DW_DUTY_HEAD_KEY AND B.RETRN_MONTH_YEAR BETWEEN " +From_dt+  " AND "+To_dt+ " AND C.DW_CETSH_KEY = B.DW_CETSH_KEY AND C.CETH_NO_2 <> 'XX' GROUP BY CETH_NO,DUTY_SPECFC ")
print("Writing data to output file for CTH8")

cnt=0
rno=4
for r in cursor1:
    cnt = cnt + 1
    sheet5.write(rno, 0, cnt,cell_format)
    for c in range(67):
        if (c >= 0 and c <= 67):
            
            if(r[c]==None):
                sheet5.write(rno, c+1, '--',Num_format)
            else:
                sheet5.write(rno, c+1, r[c],Num_format)
        else:
            sheet5.write(rno, c+1, r[c],cell_format)
    rno = rno + 1

cursor1.close()
cnxn.close()
wb.close()

print("End of Run")
