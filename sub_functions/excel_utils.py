import os
import sys
from bs4 import BeautifulSoup
from xlrd import *
import xlwt
from xlwt import Workbook, easyxf
from xlutils import *
from xlutils.copy import copy

# o_filename = '/home/user/Desktop/R230D_2G/0_200pcs_2.xls'
# i_path = '/home/user/Desktop/r230d_0608/R230D_5G'
# i_filename =  filename   
def excel_rd_md_wr(i_path,i_filename,o_filename,sum_actual_SN): 
    rate_final_result_F = 0
    rate_final_result_P = 0
    err_7506_CONTROL_DONE_tx = 0
    err_7506_CONTROL_DONE_tx_OK = 0
    err_LP_VsaDataCapture_returned_error = 0
    err_LP_VsaDataCapture_returned_error_OK = 0
    err_RX_FAIL = 0
    err_RX_FAIL_OK = 0
    err_DUT_login_Error = 0
    err_DUT_login_Error_OK = 0
    err_Barcode_Query_value = 0
    err_Barcode_Query_value_OK = 0
    err_DUT_Start_Tx_Frame_Fail = 0
    err_DUT_Start_Tx_Frame_Fail_OK = 0
    err_EVM = 0
    err_EVM_OK = 0
    err_Channel = 0
    err_Channel_OK = 0   
    err_Ping_ONT_time_out = 0  
    err_Ping_ONT_time_out_OK = 0
    err_cd_uutap = 0
    err_cd_uutap_OK = 0
    err_Wifi_Calibration = 0
    err_Wifi_Calibration_OK = 0
           
    # The Colon number and Sheet index will be handle with. The values will be modified with the caller.
    COL_SN_Num = 1
    COL_Test_Result = 4
    COL_Final_Result = 8
    COL_Restart_Times = 9
    COL_Report_Hyperlink = 10
    COL_Fail_Comments = 11
    COL_remarks = 13
    SHEET_Log_Statistic = 3
    
    rowIndex = 1 # The value must be 1, because the "rowIndex - 1" operate
    
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')
    style1 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on',num_format_str='#0')
    style2 = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on',num_format_str='#,##0.00')
    style3 = xlwt.easyxf('font: name Times New Roman, color-index yellow, bold on',num_format_str='#,##0.00')
    style4 = xlwt.easyxf('pattern:pattern solid, fore_colour yellow;''align: vertical center, horizontal left;''font: bold true, colour black;',num_format_str='#0')
    
    i_filename_full = os.path.join(i_path, i_filename)  
    #i_filename_hyperlink = 'file://' + i_filename_full
    rb = open_workbook(i_filename_full,formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(SHEET_Log_Statistic)
    sheet = rb.sheet_by_index(SHEET_Log_Statistic)  
    print sheet.name  
    print sheet.nrows
    print sheet.ncols  
    for rowIndex in range(sheet.nrows):
        # for write hyperlink of log file  
        if  sheet.cell(rowIndex, COL_Restart_Times).value <> "" and rowIndex <> 0:
            ws.write(rowIndex,  COL_Report_Hyperlink, xlwt.Formula('Hyperlink("%s")'%sheet.cell(rowIndex, COL_remarks).value))
            ws.write(rowIndex, COL_remarks,"")
        elif sheet.cell(rowIndex, COL_Test_Result).value == "P":
            ws.write(rowIndex, COL_remarks,"")
        if  sheet.cell(rowIndex, COL_Test_Result).value == "F":
            ws.write(rowIndex,  COL_Report_Hyperlink, xlwt.Formula('Hyperlink("%s")'%sheet.cell(rowIndex, COL_remarks).value))
            ws.write(rowIndex, COL_remarks,"")
            if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:12] == "100 Frames @":
                ws.write(rowIndex ,COL_Fail_Comments,"100 Frames @")
            if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:20] == "7506 CONTROL DONE tx":
                ws.write(rowIndex ,COL_Fail_Comments,"7506 CONTROL DONE tx")
            if sheet.cell(rowIndex ,COL_Fail_Comments).value[0:32] == "LP_VsaDataCapture returned error":
                ws.write(rowIndex ,COL_Fail_Comments,"LP_VsaDataCapture returned error")
            if sheet.cell(rowIndex ,COL_Fail_Comments).value[0:15] == "DUT login Error":
                ws.write(rowIndex ,COL_Fail_Comments,"DUT login Error")   
            if sheet.cell(rowIndex ,COL_Fail_Comments).value[0:19] == "Barcode_Query_value":
                ws.write(rowIndex ,COL_Fail_Comments,"DUT login Error")
            if sheet.cell(rowIndex ,COL_Fail_Comments).value[0:23] == "DUT Start Tx Frame Fail":
                ws.write(rowIndex ,COL_Fail_Comments,"DUT login Error")
                                                                          
        # for count the final fail of test result
        if  sheet.cell(rowIndex -1 ,COL_SN_Num).value <> sheet.cell(rowIndex ,COL_SN_Num).value:
            if sheet.cell(rowIndex -1 ,COL_Test_Result).value == "F":
                ws.write(rowIndex -1,COL_Final_Result,"FAIL",style0)
                rate_final_result_F += 1
                print rowIndex -1
        # for count the final OK of test result
        if  sheet.cell(rowIndex ,COL_Final_Result).value == "OK":
            if  (sheet.cell(rowIndex -1,COL_Test_Result).value == "P") and (sheet.cell(rowIndex -1 ,COL_SN_Num).value == sheet.cell(rowIndex ,COL_SN_Num).value):
                ws.write(rowIndex ,COL_Final_Result,"")
            else :
                rate_final_result_P += 1
        # for count the error type of 7506_CONTROL_DONE
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:20] == "7506 CONTROL DONE tx":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_7506_CONTROL_DONE_tx_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_7506_CONTROL_DONE_tx += 1
        # for count the error type of LP_VsaDataCaptur_returned_error
        if sheet.cell(rowIndex ,COL_Fail_Comments).value[0:32] == "LP_VsaDataCapture returned error":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_LP_VsaDataCapture_returned_error_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_LP_VsaDataCapture_returned_error += 1
        # for count the error type of 7506_CONTROL_DONE
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:12] == "100 Frames @":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_RX_FAIL_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_RX_FAIL += 1            
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:15] == "DUT login Error":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_DUT_login_Error_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_DUT_login_Error += 1 
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:19] == "Barcode_Query_value":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_DUT_login_Error_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_DUT_login_Error += 1 
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:23] == "DUT Start Tx Frame Fail":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_DUT_Start_Tx_Frame_Fail_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_DUT_Start_Tx_Frame_Fail += 1 
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[5:8] == "EVM":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_EVM_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_EVM += 1 
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:7] == "Channel":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_Channel_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_Channel += 1                                 
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:17] == "Ping ONT time out":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_Ping_ONT_time_out_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_Ping_ONT_time_out += 1
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:8] == "cd uutap":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_cd_uutap_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_cd_uutap += 1
        if  sheet.cell(rowIndex ,COL_Fail_Comments).value[0:1] == "-":
            if (sheet.cell(rowIndex +1,COL_Final_Result).value == "OK") and (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value == sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_Wifi_Calibration_OK += 1
            if (sheet.cell(rowIndex ,COL_Test_Result).value == "F") and(sheet.cell(rowIndex ,COL_SN_Num).value <> sheet.cell(rowIndex +1,COL_SN_Num).value):
                err_Wifi_Calibration += 1                                                
    ws.write(sheet.nrows -1, COL_Restart_Times,"rate_final_result_F =")
    ws.write(sheet.nrows , COL_Restart_Times,rate_final_result_F)
    ws.write(sheet.nrows + 1, COL_Restart_Times,"%.4f%%"%(float(rate_final_result_F)/float(sum_actual_SN) * 100))
    
    ws.write(sheet.nrows -1, COL_Final_Result,"rate_final_result_P =")
    ws.write(sheet.nrows , COL_Final_Result,rate_final_result_P)
    ws.write(sheet.nrows + 1, COL_Final_Result,"%.4f%%"%(float(rate_final_result_P)/float(sum_actual_SN) * 100))
    
    ws.write(sheet.nrows -1, COL_Fail_Comments,"err_7506_CONTROL_DONE_tx =")
    ws.write(sheet.nrows , COL_Fail_Comments,err_7506_CONTROL_DONE_tx)
    ws.write(sheet.nrows + 1, COL_Fail_Comments,"%.4f%%"%(float(err_7506_CONTROL_DONE_tx)/float(sum_actual_SN) * 100))
    
    ws.write(sheet.nrows +3, COL_Fail_Comments,"err_LP_VsaDataCapture_returned_error =")
    ws.write(sheet.nrows +4, COL_Fail_Comments,err_LP_VsaDataCapture_returned_error)
    ws.write(sheet.nrows +5, COL_Fail_Comments,"%.4f%%"%(float(err_LP_VsaDataCapture_returned_error)/float(sum_actual_SN) * 100))
    
    ws.write(sheet.nrows +6, COL_Fail_Comments,"err_LP_VsaDataCapture_returned_error_OK =")
    ws.write(sheet.nrows +7, COL_Fail_Comments,err_LP_VsaDataCapture_returned_error_OK)
    ws.write(sheet.nrows +8, COL_Fail_Comments,"%.4f%%"%(float(err_LP_VsaDataCapture_returned_error_OK)/float(sum_actual_SN) * 100))
    
    ws.write(sheet.nrows +9, COL_Fail_Comments,"err_7506_CONTROL_DONE_tx_OK =")
    ws.write(sheet.nrows +10, COL_Fail_Comments,err_7506_CONTROL_DONE_tx_OK)
    ws.write(sheet.nrows +11, COL_Fail_Comments,"%.4f%%"%(float(err_7506_CONTROL_DONE_tx_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +12, COL_Fail_Comments,"err_RX_FAIL_OK =")
    ws.write(sheet.nrows +13, COL_Fail_Comments,err_RX_FAIL_OK)
    ws.write(sheet.nrows +14, COL_Fail_Comments,"%.4f%%"%(float(err_RX_FAIL_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +15, COL_Fail_Comments,"err_RX_FAIL =")
    ws.write(sheet.nrows +16, COL_Fail_Comments,err_RX_FAIL)
    ws.write(sheet.nrows +17, COL_Fail_Comments,"%.4f%%"%(float(err_RX_FAIL)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +18, COL_Fail_Comments,"err_DUT_login_Error_OK =")
    ws.write(sheet.nrows +19, COL_Fail_Comments,err_DUT_login_Error_OK)
    ws.write(sheet.nrows +20, COL_Fail_Comments,"%.4f%%"%(float(err_DUT_login_Error_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +21, COL_Fail_Comments,"err_DUT_login_Error =")
    ws.write(sheet.nrows +22, COL_Fail_Comments,err_DUT_login_Error)
    ws.write(sheet.nrows +23, COL_Fail_Comments,"%.4f%%"%(float(err_DUT_login_Error)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +24, COL_Fail_Comments,"err_Barcode_Query_value_OK =")
    ws.write(sheet.nrows +25, COL_Fail_Comments,err_Barcode_Query_value_OK)
    ws.write(sheet.nrows +26, COL_Fail_Comments,"%.4f%%"%(float(err_Barcode_Query_value_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +27, COL_Fail_Comments,"err_Barcode_Query_value =")
    ws.write(sheet.nrows +28, COL_Fail_Comments,err_Barcode_Query_value)
    ws.write(sheet.nrows +29, COL_Fail_Comments,"%.4f%%"%(float(err_Barcode_Query_value)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +30, COL_Fail_Comments,"err_DUT_Start_Tx_Frame_Fail_OK =")
    ws.write(sheet.nrows +31, COL_Fail_Comments,err_DUT_Start_Tx_Frame_Fail_OK)
    ws.write(sheet.nrows +32, COL_Fail_Comments,"%.4f%%"%(float(err_DUT_Start_Tx_Frame_Fail_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +33, COL_Fail_Comments,"err_DUT_Start_Tx_Frame_Fail =")
    ws.write(sheet.nrows +34, COL_Fail_Comments,err_DUT_Start_Tx_Frame_Fail)
    ws.write(sheet.nrows +35, COL_Fail_Comments,"%.4f%%"%(float(err_DUT_Start_Tx_Frame_Fail)/float(sum_actual_SN) * 100))
    
    ws.write(sheet.nrows +36, COL_Fail_Comments,"err_EVM_OK =")
    ws.write(sheet.nrows +37, COL_Fail_Comments,err_EVM_OK)
    ws.write(sheet.nrows +38, COL_Fail_Comments,"%.4f%%"%(float(err_EVM_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +39, COL_Fail_Comments,"err_EVM =")
    ws.write(sheet.nrows +40, COL_Fail_Comments,err_EVM)
    ws.write(sheet.nrows +41, COL_Fail_Comments,"%.4f%%"%(float(err_EVM)/float(sum_actual_SN) * 100))
    
    ws.write(sheet.nrows +42, COL_Fail_Comments,"err_Channel_OK =")
    ws.write(sheet.nrows +43, COL_Fail_Comments,err_Channel_OK)
    ws.write(sheet.nrows +44, COL_Fail_Comments,"%.4f%%"%(float(err_Channel_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +45, COL_Fail_Comments,"err_Channel =")
    ws.write(sheet.nrows +46, COL_Fail_Comments,err_Channel)
    ws.write(sheet.nrows +47, COL_Fail_Comments,"%.4f%%"%(float(err_Channel)/float(sum_actual_SN) * 100))        
    
    ws.write(sheet.nrows +48, COL_Fail_Comments,"err_Ping_ONT_time_out_OK =")
    ws.write(sheet.nrows +49, COL_Fail_Comments,err_Ping_ONT_time_out_OK)
    ws.write(sheet.nrows +50, COL_Fail_Comments,"%.4f%%"%(float(err_Ping_ONT_time_out_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +51, COL_Fail_Comments,"err_Ping_ONT_time_out =")
    ws.write(sheet.nrows +52, COL_Fail_Comments,err_Ping_ONT_time_out)
    ws.write(sheet.nrows +53, COL_Fail_Comments,"%.4f%%"%(float(err_Ping_ONT_time_out)/float(sum_actual_SN) * 100))        
    
    ws.write(sheet.nrows +54, COL_Fail_Comments,"err_cd_uutap_OK =")
    ws.write(sheet.nrows +55, COL_Fail_Comments,err_cd_uutap_OK)
    ws.write(sheet.nrows +56, COL_Fail_Comments,"%.4f%%"%(float(err_cd_uutap_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +57, COL_Fail_Comments,"err_cd_uutap =")
    ws.write(sheet.nrows +58, COL_Fail_Comments,err_cd_uutap)
    ws.write(sheet.nrows +59, COL_Fail_Comments,"%.4f%%"%(float(err_cd_uutap)/float(sum_actual_SN) * 100))        
    
    ws.write(sheet.nrows +60, COL_Fail_Comments,"err_Wifi_Calibration_OK =")
    ws.write(sheet.nrows +61, COL_Fail_Comments,err_Wifi_Calibration_OK)
    ws.write(sheet.nrows +62, COL_Fail_Comments,"%.4f%%"%(float(err_Wifi_Calibration_OK)/float(sum_actual_SN) * 100))

    ws.write(sheet.nrows +63, COL_Fail_Comments,"err_Wifi_Calibration =")
    ws.write(sheet.nrows +64, COL_Fail_Comments,err_Wifi_Calibration)
    ws.write(sheet.nrows +65, COL_Fail_Comments,"%.4f%%"%(float(err_Wifi_Calibration)/float(sum_actual_SN) * 100))        
                                                
    wb.save(o_filename)
    return excel_rd_md_wr
    


def show_color(sheet):
    colNum = 6
    width = 5000
    height = 500
    colors = ['aqua','black','blue','blue_gray','bright_green','brown','coral','cyan_ega','dark_blue','dark_blue_ega','dark_green','dark_green_ega','dark_purple','dark_red',
            'dark_red_ega','dark_teal','dark_yellow','gold','gray_ega','gray25','gray40','gray50','gray80','green','ice_blue','indigo','ivory','lavender',
            'light_blue','light_green','light_orange','light_turquoise','light_yellow','lime','magenta_ega','ocean_blue','olive_ega','olive_green','orange','pale_blue','periwinkle','pink',
            'plum','purple_ega','red','rose','sea_green','silver_ega','sky_blue','tan','teal','teal_ega','turquoise','violet','white','yellow']

    for colorIndex in range(len(colors)):
            rowIndex = colorIndex / colNum
            colIndex = colorIndex - rowIndex*colNum
            sheet.col(colIndex).width = width
            sheet.row(rowIndex).set_style(easyxf('font:height %s;'%height)) 
            color = colors[colorIndex]
            whiteStyle = easyxf('pattern:pattern solid, fore_colour %s;'
                                    'align: vertical center, horizontal center;'
                                    'font: bold true, colour white;' % color)
            blackStyle = easyxf('pattern:pattern solid, fore_colour %s;'
                                    'align: vertical center, horizontal center;'
                                    'font: bold true, colour black;' % color)


            if color == 'black':
                    sheet.write(rowIndex, colIndex, color, style = whiteStyle)
            else:
                    sheet.write(rowIndex, colIndex, color, style = blackStyle)

def show_size(sheet):
    widthStart = 100
    widthInterval = 100
    colNum = 255
    heightStart = 100
    heightInterval = 5
    rowNum = 255
    styles = (easyxf('pattern:pattern solid, fore_colour gray50;'
                    'align: vertical center, horizontal center;'
                    'font: bold true, colour white;'),
            easyxf('pattern:pattern solid, fore_colour gray80;'
                    'align: vertical center, horizontal center;'
                    'font: bold true, colour white;'))
    for rowIndex in range(rowNum):
            height = heightStart + heightInterval*rowIndex
            sheet.row(rowIndex).set_style(easyxf('font:height %s;'%height))
            styleIndex = rowIndex%2
            for colIndex in range(colNum):
                    width = widthStart + widthInterval*colIndex
                    sheet.col(colIndex).width = width
                    sheet.write(rowIndex, colIndex, '%sx%s'%(width,height), style = styles[styleIndex])
                    styleIndex = int(not styleIndex)


#if __name__ == '__main__':
def excel_styles():
    book = Workbook(encoding='utf-8')
    colorSheet = book.add_sheet('colors')
    sizeSheet = book.add_sheet('size')
    show_color(colorSheet)
    show_size(sizeSheet)
    styleFile = '/home/user/0_Daily_work/python/excel_styles.xls'
    book.save(styleFile)
    print 'saved to "%s"' % styleFile 
    return excel_styles   