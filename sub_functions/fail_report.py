import os
import sys
from bs4 import BeautifulSoup
from xlrd import *
from xlwt import *
from xlutils import *
from xlutils.copy import copy

def fail_log_print(i_filename):
    soup=BeautifulSoup(open('%s'%i_filename),"lxml")
    for child in soup.descendants:
        if child.string == "FAIL":
    	      print(child.previous.previous.string)	
    	      print(child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string)	
    return fail_log_print

def node_abstract(i_filename,o_filename):    
    soup=BeautifulSoup(open("%s"%i_filename),'lxml')
    output_file = open("%s"%o_filename,"w")
    output_file.write("thank you very much")
    len(list(soup.children))
    l = len(list(soup.descendants))
    print (list(soup.descendants)[20])
    for i in range(l):
        output_file.write("caroid: %s\n"%(i))
        output_file.write("%s\n"%(list(soup.descendants)[i]))
        print (i)
        print (list(soup.descendants)[i])
    return node_abstract        
        
def F_P_static():
    book = open_workbook('/home/user/Desktop/R230D_2G/0_200pcs_1.xls')  
    sheet = book.sheet_by_index(3)  
    print sheet.name  
    print sheet.nrows  
    print sheet.ncols  
    for row_index in range(sheet.nrows):  
        #for col_index in range(sheet.ncols):  
        #    print cellname(row_index,col_index),'-',  
        #    print sheet.cell(row_index,col_index).value 
        if  sheet.cell(row_index -1 ,1).value <> sheet.cell(row_index ,1).value:
            if sheet.cell(row_index -1 ,1).value == "F":
                ws.write(i-1,j+7,"FAIL")
    return F_P_static    

def fail_comments():
    i = 1
    j = 1
    rb = open_workbook('/home/user/0_Daily_work/python/excel_styles.xls',formatting_info=True)  
    wb = copy(rb)
    ws = wb.get_sheet(2)
    sheet = rb.sheet_by_index(2)  
    print sheet.name  
    print sheet.nrows  
    print sheet.ncols 
    if True: 
        soup=BeautifulSoup(open('/home/user/0_Daily_work/python/081648011475_2016-06-07_09-51-15_F.html'),"lxml")
        for child in soup.descendants:
            if child.string == "FAIL":
                if i%2 == 1:
                    ws.write(j,1,child.previous.previous.string)
                    ws.write(j,2,child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string)
                    print(child.previous.previous.string)
                    print(child.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.next.string)
                    j = j + 1
                i += 1 
    wb.save('/home/user/0_Daily_work/python/excel_styles_1.xls')
    	

