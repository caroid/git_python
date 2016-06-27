'over time calculation
'ot() is the main function
'
Sub ot()
 Range("a150:a300").ClearContents
   Range("b150:b300").ClearContents
   
fill_fixedmsg

    'Just showing where the input data are
    Sheet1.Activate
    Range("A1").Select

weekofday
ot_1

Sheet2.Activate

report_out
'Worksheets.Add after:=Worksheets("节假日设定")
'Sheet4.Cells.Clear
'有空要将 考勤异常查询.xls中的人也统计到报告中。将报告中的人数补齐

Macro3


End Sub
Private Sub report_out()
Sheet2.Activate
Range("A1").Select

 
    Dim x As Long
    x = 1
For x = 1 To 200
    Sheet4.Cells(x, 1) = Sheet2.Cells(x, 12)
    Sheet4.Cells(x, 2) = Sheet2.Cells(x, 13)
    Sheet4.Cells(x, 3) = Sheet2.Cells(x, 14)
    Sheet4.Cells(x, 4) = Sheet2.Cells(x, 15)
    Sheet4.Cells(x, 5) = Sheet2.Cells(x, 16)
    Sheet4.Cells(x, 6) = Sheet2.Cells(x, 17)
    Sheet4.Cells(x, 7) = Sheet2.Cells(x, 18)

Next

    Range("A1:G1").Select
    Columns("E:E").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    
   Macro1
   Macro2
    
End Sub



Private Sub fill_fixedmsg()

Sheet2.Activate
Range("A1").Select

 Sheet2.Cells.Clear
 
    Dim x As Long
    x = 2
    y = 4
    m = 1
    
Do While Left(Sheet1.Cells(x, 1), 1) = "G"

    Sheet2.Cells(y, 1) = Sheet1.Cells(x, 1)
    Sheet2.Cells(y, 2) = Sheet1.Cells(x, 2)
    Sheet2.Cells(y, 3) = Sheet1.Cells(x, 3)
    Sheet2.Cells(y, 4) = Sheet1.Cells(x, 7)
    '忽略同一天多次刷卡记录，如果只刷一次，则扩展为同日同时的两次，如果三次以上，只保留第一次和最后一次。
    m = x
    Do While Left(Sheet1.Cells(m, 7), 10) = Left(Sheet1.Cells(m + 1, 7), 10)
        m = m + 1
    Loop
    '通宵加班判断条件:一次打卡，下一记录的工号与本记录相同，下一记录的日期比本记录大一（导出记录时要从1日到31日，不能垮月），下一记录打卡时间小于等于5点
    If (x = m) And (Sheet1.Cells(x, 1) = Sheet1.Cells(x + 1, 1)) And (Mid(Sheet1.Cells(x + 1, 7), 9, 2) = Mid(Sheet1.Cells(x, 7), 9, 2) + 1) And (Mid(Sheet1.Cells(x + 1, 7), 12, 2) <= 5) Then Sheet2.Cells(y, 20) = "通宵"
    
    x = m
    
    Sheet2.Cells(y + 1, 1) = Sheet1.Cells(x, 1)
    Sheet2.Cells(y + 1, 2) = Sheet1.Cells(x, 2)
    Sheet2.Cells(y + 1, 3) = Sheet1.Cells(x, 3)
    Sheet2.Cells(y + 1, 4) = Sheet1.Cells(x, 7)
    
       
    x = x + 1
    y = y + 2
Loop
   
    Sheet2.Cells(3, 1) = Sheet1.Cells(1, 1)
    Sheet2.Cells(3, 2) = Sheet1.Cells(1, 2)
    Sheet2.Cells(3, 3) = Sheet1.Cells(1, 3)
    Sheet2.Cells(3, 4) = Sheet1.Cells(1, 7)
    
    Sheet2.Cells(3, 16) = "工作日加班时长"
    Sheet2.Cells(3, 17) = "周末加班时长"
    Sheet2.Cells(3, 18) = "有效加班时长"
    Sheet2.Cells(3, 12) = "工号"
    Sheet2.Cells(3, 13) = "姓名"
    Sheet2.Cells(3, 14) = "部门"
    Sheet2.Cells(3, 15) = "刷卡时间"

    
   Macro1
   Macro2
    
End Sub

Private Sub weekofday()
Dim m, i As Integer
m = 4
i = 2
Do While Left(Sheet2.Cells(m, 4), 2) = "20"

    If Weekday(Left(Sheet2.Cells(m, 4), 10)) = 1 Then Sheet2.Cells(m, 5) = "星期日"
    If Weekday(Left(Sheet2.Cells(m, 4), 10)) = 2 Then Sheet2.Cells(m, 5) = "星期一"
    If Weekday(Left(Sheet2.Cells(m, 4), 10)) = 3 Then Sheet2.Cells(m, 5) = "星期二"
    If Weekday(Left(Sheet2.Cells(m, 4), 10)) = 4 Then Sheet2.Cells(m, 5) = "星期三"
    If Weekday(Left(Sheet2.Cells(m, 4), 10)) = 5 Then Sheet2.Cells(m, 5) = "星期四"
    If Weekday(Left(Sheet2.Cells(m, 4), 10)) = 6 Then Sheet2.Cells(m, 5) = "星期五"
    If Weekday(Left(Sheet2.Cells(m, 4), 10)) = 7 Then Sheet2.Cells(m, 5) = "星期六"
    '节假日及调休判断
    Do While Sheet3.Cells(i, 1) <> ""
    If Left(Sheet2.Cells(m, 4), 10) = Sheet3.Cells(i, 1) Then
        Sheet2.Cells(m, 5) = Sheet3.Cells(i, 3)
        End If
        
    i = i + 1
    Loop
    i = 2
    m = m + 1
Loop

End Sub

Private Sub ot_1()
Dim i, m As Integer
i = 4
m = 4
n = 4
Do While Sheet2.Cells(i, 1) <> ""
    m = i

    Do While Sheet2.Cells(i, 1) = Sheet2.Cells(i + 1, 1)
    '目前通宵加班是统计在平时加班时间中，以后根据实际来区分加载平时还是周末。
    If Sheet2.Cells(i, 20) = "通宵" And Sheet2.Cells(i, 5) <> "星期日" And Sheet2.Cells(i, 5) <> "星期六" And Sheet2.Cells(i, 5) <> "假日" And Sheet2.Cells(i, 5) <> "星期五" Then Sheet2.Cells(m, 6) = Sheet2.Cells(m, 6) + 5 + Mid(Sheet2.Cells(i, 4), 13, 1)
    If Sheet2.Cells(i, 20) = "通宵" And Sheet2.Cells(i, 5) = "星期五" Then
        Sheet2.Cells(m, 6) = Sheet2.Cells(m, 6) + 5
        Sheet2.Cells(m, 7) = Sheet2.Cells(m, 7) + Mid(Sheet2.Cells(i, 4), 13, 1)
    End If
    
    '平时加班时间统计
    If Mid(Sheet2.Cells(i, 4), 12, 2) > "18" And Sheet2.Cells(i, 5) <> "星期日" And Sheet2.Cells(i, 5) <> "星期六" And Sheet2.Cells(i, 5) <> "假日" Then
    Sheet2.Cells(m, 6) = Sheet2.Cells(m, 6) + (Mid(Sheet2.Cells(i, 4), 12, 2) - 19 + Mid(Sheet2.Cells(i, 4), 15, 2) / 60) * 2
    'If Mid(Sheet2.Cells(i, 4), 15, 2) / 60 >= 0.5 Then Sheet2.Cells(m, 6) = Sheet2.Cells(m, 6) + 0.5
    End If
    '周末和节假日加班时间统计
    If ((Sheet2.Cells(i, 5) = "星期日" And Sheet2.Cells(i + 1, 5) = "星期日") Or (Sheet2.Cells(i, 5) = "星期六" And Sheet2.Cells(i + 1, 5) = "星期六") Or (Sheet2.Cells(i, 5) = "假日" And Sheet2.Cells(i + 1, 5) = "假日")) And (i Mod 2 = 0) Then
    Sheet2.Cells(i, 10) = Mid(Sheet2.Cells(i + 1, 4), 15, 2)
    Sheet2.Cells(i, 11) = Mid(Sheet2.Cells(i, 4), 15, 2)
    '周末或假日加班8小时封顶，中午休息一小时还没有减去。
    If (Mid(Sheet2.Cells(i + 1, 4), 12, 2) - Mid(Sheet2.Cells(i, 4), 12, 2)) > 8 Then
        Sheet2.Cells(m, 7) = Sheet2.Cells(m, 7) + 8
    Else
        Sheet2.Cells(m, 7) = Sheet2.Cells(m, 7) + Mid(Sheet2.Cells(i + 1, 4), 12, 2) - Mid(Sheet2.Cells(i, 4), 12, 2)
    End If

    If Sheet2.Cells(i, 10) / 60 + 1 - Sheet2.Cells(i, 11) / 60 >= 0.5 Then Sheet2.Cells(m, 7) = Sheet2.Cells(m, 7) + 0.5
    End If
        Sheet2.Cells(n, 15) = Left(Sheet2.Cells(m, 4), 10) & " ~ " & Left(Sheet2.Cells(i + 1, 4), 10)
    i = i + 1
    Loop
    
        If Sheet2.Cells(m, 6) - Fix(Sheet2.Cells(m, 6)) >= 0.5 Then
            Sheet2.Cells(m, 6) = Fix(Sheet2.Cells(m, 6)) + 0.5
        Else
            Sheet2.Cells(m, 6) = Fix(Sheet2.Cells(m, 6))
        End If
           
        Sheet2.Cells(n, 12) = Sheet2.Cells(m, 1)
        Sheet2.Cells(n, 13) = Sheet2.Cells(m, 2)
        Sheet2.Cells(n, 14) = Sheet2.Cells(m, 3)
        Sheet2.Cells(n, 16) = Sheet2.Cells(m, 6)
        '周末加班记录为空的，添0. 20160120
        If Sheet2.Cells(m, 7) = "" Then Sheet2.Cells(m, 7) = 0
        Sheet2.Cells(n, 17) = Sheet2.Cells(m, 7)
        
        With Application
        Sheet2.Cells(n, 18) = .Min(Sheet2.Cells(n, 16).Value, Sheet2.Cells(n, 17).Value)
        End With
i = i + 1
n = n + 1
Loop
    

End Sub

Private Sub test_1()
Dim i As Long
i = 4
'Do While Sheet2.Cells(i, 1) <> ""
Sheet2.Cells(1, 20) = Mid(Sheet2.Cells(i + 1, 4), 15, 2) - 1
Sheet2.Cells(2, 20) = Mid(Sheet2.Cells(i + 1, 4), 12, 2) - 1

If Mid(Sheet2.Cells(i + 1, 4), 12, 2) > 5 Then Sheet2.Cells(3, 20) = 1


i = i + 1
'Loop

End Sub


Private Sub Macro1()
'
' Macro1 Macro
'

'
    Range("A1:G1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("E3:G3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1:G1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1:G1").Select
    ActiveCell.FormulaR1C1 = "加班时间统计"
    With ActiveCell.Characters(Start:=1, Length:=6).Font
        .Name = "宋体"
        .FontStyle = "加粗"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("A1:G1").Select
    Columns("E:E").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    


End Sub
Sub Macro2()
'
' Macro2 Macro
'

'
    Range("L1:R1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L3:R3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L1:R31").Select
    Selection.Columns.AutoFit
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("J16").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("L1:R1").Select
    ActiveCell.FormulaR1C1 = "加班时间统计"
    With ActiveCell.Characters(Start:=1, Length:=6).Font
        .Name = "宋体"
        .FontStyle = "加粗"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("S1").Select
End Sub

Sub Macro3()
'
' Macro3 Macro
'

'
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").EntireColumn.AutoFit
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
End Sub

