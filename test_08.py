#!/usr/bin/env python
# usage: python test_07.py -p /home/user/Desktop/r230d_0608/R230D_5G/
# plan: 1. Use dict replace to 'if else'. 
#       2. add function to creat new excel file used for temporary information and data , remove the template of excle files.
#       3. extract subfunction from the main functions. increase the flexibility.
#       4. dict
import os 
import subprocess 
import re 
import hashlib 
import getopt
from getopt import GetoptError
from bs4 import BeautifulSoup
import sys
from xlrd import open_workbook
from xlrd import open_workbook,cellname
from xlutils.copy import copy
import xlwt 
from xlwt import Workbook, easyxf
# local sub functions import
from sub_functions import os_info
from sub_functions import fail_report
from sub_functions import excel_utils

#
def log_statistic(i_path):
    temp=''
    temp_1 = ""
    temp_100 = ""
    temp_110 = ""
    temp_111 = ""
    temp_112 = ""
    temp_2 = {}
    temp_3 = {}
    temp_10 = {}
    temp_11 = {}
    i =1 # counter for "Log Quantity"
    k =1 # counter for "Actual SN Quantity"
    l = 0 # counter for "Repeat Times"
    m = 0 # counter for "cross border detection". the number of parts from filename split should no less than 3.
    n = 0 # counter for "Restart Times"
    p = 1 # abstract flag for "parse the node from soup.desendants", delete one string from dobule info string.
    p_1 = 1
    sum_actual_SN = 0 # sumary of actual SN
    invalid_file_name = 0 # The number of invalid file name in input directory
    #rate_final_result_P = 0 # count of repeat test success 
    #rate_final_result_F = 0
    # column define
    COL_Actual_SN_Quantity = 0
    COL_SN_Num = 1
    COL_Test_Date = 2
    COL_Test_Time = 3
    COL_Test_Result = 4
    COL_Log_Quantity = 5
    COL_First_Pass = 5
    COL_Repeat_Times = 6
    COL_Station_ID = 7
    COL_Final_Result = 8
    COL_Restart_Times = 9
    COL_Report_Hyperlink = 10
    COL_Fail_Comments = 11
    COL_Error_Types = 12
    COL_remarks = 13
    COL_Error_416 = 14
    COL_Error_LP_VsaDataCapture = 15
    
    sheet_Log_Statistic = 3    
    # the display style of excel file.
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')
    style1 = xlwt.easyxf('pattern:pattern solid, fore_colour yellow;''align: vertical center, horizontal left;''font: bold true, colour black;',num_format_str='#0')
    style2 = xlwt.easyxf('align: vertical center, horizontal center;''font: bold true, colour black;',num_format_str='#0')
    # keywords for pick the Test Station ID from the .html log.
    w1 = 'Station ID:'
    w2 = ', Product Code:'
    # open an excel template file , copy and add some contents to generate final report
    rb = open_workbook('/home/user/0_Daily_work/git_python/0_200pcs.xls',formatting_info=True)
    wb = copy(rb)
    # select a sheet that will be addtion statistic data.
    ws = wb.get_sheet(3)
    ws.write(0, COL_Actual_SN_Quantity, "Actual Quantity of SN",style1)	
    ws.write(0, COL_SN_Num, "SN Number",style1)	
    ws.write(0, COL_Test_Date, "Test Date",style1)	
    ws.write(0, COL_Test_Time, "Test Time",style1)	
    ws.write(0, COL_Test_Result, "Test Result",style1)	
    ws.write(0, COL_First_Pass, "Next or Previous Station Result",style1)	
    ws.write(0, COL_Repeat_Times, "Repeat Times",style1)	
    ws.write(0, COL_Station_ID, "Station ID",style1)
    ws.write(0, COL_Final_Result, "Final Results of Repeat Test",style1) 
    ws.write(0, COL_Restart_Times, "Restart Times",style1)
    ws.write(0, COL_Report_Hyperlink, "Hyperlink of Fail log",style1)
    ws.write(0, COL_Fail_Comments, "Comments of Fail",style1)
    ws.write(0, COL_Error_Types, "COL_Error_Types",style1)
    ws.write(0, COL_Error_416, "COL_Error_416",style1)
    ws.write(0, COL_Error_LP_VsaDataCapture, "COL_Error_LP_VsaDataCapture",style1)
    # algorithm : Remove duplicate SN number, create a set, if element is NOT in set, add it, else the element is in set, don't add.
    lines_seen = set() 
    filenames = os.listdir(os.path.dirname(i_path))
    filenames.sort()
    for filename in filenames:
            print i
            n = 0
            p = 1
            p_1 = 1
            m = 0
            temp_1 = ""
            temp_100 = ""
            temp_110 = ""
            temp_111 = ""
            temp_112 = ""
            temp_3 = {}
            temp_11 = {}
            if filename[0:2] =="HW":
                continue
            #ws.write(i, COL_Log_Quantity, i)	
            temp=filename.split("_")
            # cross border detection.
            for val in temp:
                m = m +1
            if m < 3 :
                invalid_file_name += 1            	
                continue # if filename have no correct format , jump out of this cycle.
            print temp[0]
            # algorithm : Remove duplicate SN number
            if temp[0] not in lines_seen:
                ws.write(i, COL_SN_Num, temp[0])
                ws.write(i, COL_Actual_SN_Quantity, k,style2)
                lines_seen.add(temp[0])
                sum_actual_SN += 1
                temp_2 = {}
                temp_10 = {}
                k = k +1
                l = 1
            else:
                l = l + 1
                ws.write(i, COL_SN_Num, temp[0])
                if len(temp_2) <> 0 and temp_2.has_key('\n'):
                    if temp_2.values()[0] == temp[0]:
                        del temp_2['\n']
                        temp_10 = temp_2.copy()
                        print temp_10

            ws.write(i, COL_Test_Date, temp[1])
            ws.write(i, COL_Test_Time, temp[2])
            if os.path.splitext(temp[3])[0] == "F":
                ws.write(i,  COL_Test_Result, os.path.splitext(temp[3])[0],style0)
            else:
                ws.write(i,  COL_Test_Result, os.path.splitext(temp[3])[0])
            # line number (SN counter) updated at here, Caution!    
            i = i +1
            if os.path.splitext(temp[3])[0] == "F":
                ws.write(i-1, COL_Repeat_Times, l,style2)
            if os.path.splitext(temp[3])[0] == "P":
                if l <> 1 and l <> "":
                    ws.write(i-1,COL_Repeat_Times,l,style2)  
                    ws.write(i-1,COL_Final_Result ,"OK")
                    #rate_final_result_P += 1
            # extract Station ID from .html log.        
            f_fullname = os.path.join(i_path, filename)
            f = open(f_fullname,'r')
            buff = f.read()
            buff = buff.replace('\n','')
            pat = re.compile(w1+'(.*?)'+w2,re.S)
            result = pat.findall(buff) 
            ws.write(i-1, COL_Station_ID, result)
            print temp[1]
            print temp[2]
            print os.path.splitext(temp[3])[0]  
            
            for iii in re.finditer("RESULT: FAIL",buff):
                print iii.group(),iii.span()
                print buff[iii.span()[0]-100 : iii.span()[1]]
                temp_110 = (buff[iii.span()[0]-100 : iii.span()[1]])
                ws.write(i-1, COL_Error_Types, "%s"%temp_110[:],style2)    	

            for iiiiii in re.finditer("txgain=416, power=",buff):
                print iiiiii.group(),iiiiii.span()
                print buff[iiiiii.span()[0]-100 : iiiiii.span()[1]]
                temp_111 = (buff[iiiiii.span()[0]-100 : iiiiii.span()[1]])
                ws.write(i-1, COL_Error_416, "%s"%temp_111[:],style2)
        	
            for ii in re.finditer("LP_VsaDataCapture returned error: Data capture failed",buff):
                print ii.group(),ii.span()
                print buff[ii.span()[0]-100 : ii.span()[1]]
                temp_112 = (buff[ii.span()[0]-100 : ii.span()[1]])
                ws.write(i-1, COL_Error_LP_VsaDataCapture, "%s"%temp_112[:],style2)
                
            # write the hyperlink of .html log to excel
        #if os.path.splitext(temp[3])[0] == "F":    
            f_httpname_i = os.path.join(i_path, filename)
            f_httpname = 'file://' + f_httpname_i
            print (f_httpname)
            ws.write(i-1,  COL_Report_Hyperlink, xlwt.Formula('Hyperlink("%s")'%f_httpname))
            ws.write(i-1,  COL_remarks, f_httpname)
                # estimate the times that the "ATE restart button" have be clicked.
            soup=BeautifulSoup(open(f_fullname),"lxml")
            for child in soup.descendants:
                if child.string ==("\n"+"art_dn.sh"):
                    if n > 0:
                        ws.write(i-1, COL_Restart_Times, n,style2)
                    n = n + 1
                if child.string == "FAIL":
                    temp_1 += (child.previous.previous.string )
                    temp_3[child.previous.previous.string] = temp[0] #dict
                    if p%2 == 1:
                        print(child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string)	
                        #if (child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string) <> "-":
                        temp_1 += (" : " + child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string )
                        ws.write(i-1,COL_Fail_Comments,temp_1[:-1])
                        temp_2.update(temp_3)
                    p += 1
                    print temp_2
                #if len(temp_10) <> 0 and os.path.splitext(temp[3])[0] == "P" and (child.string in temp_10.keys()):
                if len(temp_10.keys()) <> 0 and (child.string in temp_10.keys()):
                    if p_1%2 == 1:
                        temp_100 += child.string
                        temp_100 += (" : " + child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string + "\n")
                        ws.write(i-1,COL_Fail_Comments,temp_100[:-1])
                    p_1 += 1
                    #ws.write(i-1,COL_Fail_Comments,child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string)
    # caculate the number of illegal .html log
    ws.write(i+1, COL_First_Pass, "invalid_file_name = ")  
    ws.write(i+2, COL_First_Pass, invalid_file_name)
    ws.write(i+2,COL_SN_Num, xlwt.Formula("subtotal(3,B2:B%d)"%i))
    #ws.write(i+1, COL_Final_Result,"rate_final_result_P = ")
    #ws.write(i+2, COL_Final_Result,rate_final_result_P)
    wb.save('/home/user/0_Daily_work/git_python/0_200pcs_1.xls')
    excel_utils.excel_rd_md_wr('/home/user/0_Daily_work/git_python','0_200pcs_1.xls','/home/user/0_Daily_work/git_python/0_200pcs_2.xls',sum_actual_SN)
    return log_statistic

# -p []: input path, must end with "/",example: -p /home/user/Desktop/r230d_0608/R230D_5G/      
# -i []: input file , example:
# -o []: output file, example:
def main():
    d = ""
    a = ""
    try:
        opts,args=getopt.getopt(sys.argv[1:], 'p:i:o:d:', ['path=','inputfile=','outputfile=','debug='])
    except GetoptError:
        sys.exit()
    for key,values in opts:
        if key in ('-p',''):
            a=values
            print a
        if key in ('-i',''):
            b=values
            print b      
        if key in ('-o',''):
            c=values
            print c 
        if key in ('-d',''):
            d=values
            print d
    if d =="debug":
        fail_report.node_abstract('/home/user/Desktop/R230D_2G/081638000016_2016-04-12_14-52-40_F.html','/home/user/Desktop/1.txt')  
    if a <> "":           
        #log_statistic('/home/user/Desktop/r230d_0608/R230D_5G/')
        log_statistic(a)
    elif a == "":
        print "example: python test08.py -p /home/user/Desktop/r230d_0608/R230D_5G/"+"\n"


if __name__ == "__main__": 
    print (os_info.get_osinfo())
    fail_report.fail_comments()
    if False:
        excel_utils.excel_styles()
        fail_report.fail_log_print('/home/user/Desktop/R230D_2G/081638000016_2016-04-12_14-52-40_F.html')
    main()     

 
 
 




    
