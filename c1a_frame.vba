Sub copyframe2(ws1 As Worksheet, ws2 As Worksheet, wbdict As Workbook)
    Dim lastcol1_t1 As Integer
    Dim lastcol_t2 As Integer
    Dim lastcol2 As Integer
    Dim lastcol3 As Integer
    Dim lastcol4 As Integer
    Dim lastrow1 As Integer
    Dim lastrow2 As Integer
    Dim lastrow3 As Integer
    Dim lastrow4 As Integer
    Dim ECUcol1 As Integer
    Dim ECUcol2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim keyFrame As String
    Dim keyy As Variant
    Dim foundCol As Integer
    Dim thisSheet As Worksheet
    Dim copyRange As Range
    
    Dim dictFrameList_base As Object
    Dim dictFrameList_comp As Object
'    Dim dictFrameList_ADAS As Object
    Dim dictFrameList_gen As Object
    
    ' Kh?i t?o các t? di?n
    Set dictFrameList_base = CreateObject("Scripting.Dictionary")
    Set dictFrameList_comp = CreateObject("Scripting.Dictionary")
'    Set dictFrameList_ADAS = CreateObject("Scripting.Dictionary")
    Set dictFrameList_gen = CreateObject("Scripting.Dictionary")
    Set thisSheet = wbdict.Sheets("Frame Synthesis")
    
    ' T?t các tính nang không c?n thi?t
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Kích ho?t trang tính ws1
    ws1.Activate
    
    lastcol1_t1 = ws1.Cells(7, ws1.Columns.Count).End(xlToLeft).Column
    lastrow1 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).row
    
    ' Ð?c d? li?u vào m?ng
    Dim data1 As Variant
    
    For i = 8 To lastrow1
    For j = 2 To lastcol1_t1
        If ws1.Cells(i, j).Value = "" Then
            ws1.Cells(i, lastcol1_t1 + 2).Value = ws1.Cells(i, lastcol1_t1 + 2).Value & "."
        Else
            ws1.Cells(i, lastcol1_t1 + 2).Value = ws1.Cells(i, lastcol1_t1 + 2).Value & ws1.Cells(i, j).Value
        End If
    Next j
    Next i

    data1 = ws1.Range(ws1.Cells(8, lastcol1_t1 + 2), ws1.Cells(lastrow1, lastcol1_t1 + 2)).Value
    
    ' Thêm các khung vào dictFrameList_base
    For i = 1 To UBound(data1, 1)
        If data1(i, 1) <> "" Then
            If dictFrameList_base.Exists(data1(i, 1)) = False Then dictFrameList_base.Add data1(i, 1), i
        End If
    Next i
    
    ' Kích ho?t trang tính ws2
    ws2.Activate
    
    lastcol1_t2 = ws2.Cells(7, ws2.Columns.Count).End(xlToLeft).Column
    If lastcol1_t2 <> lastcol1_t1 Then
        MsgBox "The number of ECU is not uniform!"
        Exit Sub
    End If
    
    lastrow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).row
    
    ' Ð?c d? li?u vào m?ng
    Dim data2 As Variant
    
    For i = 8 To lastrow2
    For j = 2 To lastcol1_t2
        If ws2.Cells(i, j).Value = "" Then
            ws2.Cells(i, lastcol1_t2 + 2).Value = ws2.Cells(i, lastcol1_t2 + 2).Value & "."
        Else
            ws2.Cells(i, lastcol1_t2 + 2).Value = ws2.Cells(i, lastcol1_t2 + 2).Value & ws2.Cells(i, j).Value
        End If
    Next j
    Next i

    data2 = ws2.Range(ws2.Cells(8, lastcol1_t2 + 2), ws2.Cells(lastrow2, lastcol1_t2 + 2)).Value
    
    ' Thêm các khung vào dictFrameList_comp
    For i = 1 To UBound(data2, 1)
        If data2(i, 1) <> "" Then
            If dictFrameList_comp.Exists(data2(i, 1)) = False Then dictFrameList_comp.Add data2(i, 1), i
        End If
    Next i
    
  
'--------------------------------------------
        
    
'
'    lastcol4 = ws3.Cells(1, ws3.Columns.Count).End(xlToLeft).Column
'    lastrow4 = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).row
'
'
'    Dim data3 As Variant
'
'    data3 = ws3.Range(ws3.Cells(2, 1), ws3.Cells(lastrow2, 1)).Value
'
'    For i = 1 To UBound(data3, 1)
'        If data3(i, 1) <> "" Then
'            If dictFrameList_ADAS.Exists(data3(i, 1)) = False Then dictFrameList_ADAS.Add data3(i, 1), i
'        End If
'    Next i
'
'
    
'-------------------------------------------------------

      ' Thêm các khung t? dictFrameList_base vào dictFrameList_gen
    For Each keyy In dictFrameList_base.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 8
        End If
    Next keyy
    
    ' Thêm các khung t? dictFrameList_comp vào dictFrameList_gen
    For Each keyy In dictFrameList_comp.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 8
        End If
    Next keyy
    
'    For Each keyy In dictFrameList_ADAS.Keys
'        If Not dictFrameList_ADAS.Exists(keyy) Then
'            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 8
'        End If
'    Next keyy
    
    ' Sao chép tiêu d? t? ws1 sang thisSheet
'-------------------------------------
    
    Set copyRange = ws1.Range(ws1.Cells(1, 1), ws1.Cells(7, lastcol1_t1))
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)

    ' Sao chép tiêu d? t? ws2 sang thisSheet
    Set copyRange = ws2.Range(ws2.Cells(1, 1), ws2.Cells(7, lastcol1_t2))
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + 2)
    copyRange.Copy Destination:=thisSheet.Cells(1, 2 * lastcol1_t1 + 3)
    
'    Set copyRange = ws3.Range(ws3.Cells(1, 1), ws3.Cells(1, lastcol4))
'    copyRange.Copy Destination:=thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 3)
    
    ' copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 3)

    ' Sao chép d? li?u t? ws1 và ws2 sang thisSheet d?a trên dictFrameList_gen
    For Each keyy In dictFrameList_gen.Keys
        If dictFrameList_base.Exists(keyy) Then
            Set copyRange = ws1.Range(ws1.Cells(dictFrameList_base(keyy) + 7, 1), ws1.Cells(dictFrameList_base(keyy) + 7, lastcol1_t1))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), 1)
        End If
        If dictFrameList_comp.Exists(keyy) Then
            Set copyRange = ws2.Range(ws2.Cells(dictFrameList_comp(keyy) + 7, 1), ws2.Cells(dictFrameList_comp(keyy) + 7, lastcol1_t2))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 + 2)
        End If
'         If dictFrameList_ADAS.Exists(keyy) Then
'            Set copyRange = ws3.Range(ws3.Cells(dictFrameList_ADAS(keyy) + 1, 1), ws3.Cells(dictFrameList_ADAS(keyy) + 1, lastcol1_t2))
'            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 + lastcol1_t2 + 3)
'        End If
    Next keyy
    
    thisSheet.Activate
    
    lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row > thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 3).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 3).End(xlUp).row)
  
   ' -------------
   ' boi xam base
    With thisSheet.Range(thisSheet.Cells(7, 1), thisSheet.Cells(lastrow3, lastcol1_t1))
        .AutoFilter
        .Rows(8).AutoFilter Field:=2, Criteria1:="="
    End With

    If thisSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
         For Each cell In thisSheet.AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible)
             If cell.Value = "" Then
               thisSheet.Range(thisSheet.Cells(cell.row, 1), thisSheet.Cells(cell.row, lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
             End If
         Next cell
    End If
     
    thisSheet.AutoFilterMode = False

    
'boi xam obj

    With thisSheet.Range(thisSheet.Cells(7, 1), thisSheet.Cells(lastrow3, 2 * lastcol1_t1 + 1))
        .AutoFilter
        .Rows(8).AutoFilter Field:=(lastcol1_t1 + 3), Criteria1:="="
    End With

    If thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + 2).SpecialCells(xlCellTypeVisible).Count > 1 Then
         For Each cell In thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + 3).SpecialCells(xlCellTypeVisible)
             If cell.Value = "" Then
                thisSheet.Range(thisSheet.Cells(cell.row, lastcol1_t1 + 2), thisSheet.Cells(cell.row, 2 * lastcol1_t1 + 1)).Interior.Color = RGB(191, 191, 191)
             End If
         Next cell
    End If
    
'boi xam ADASMsg
'
'    With thisSheet.Range(thisSheet.Cells(7, 1), thisSheet.Cells(lastrow3, lastcol1_t1 + lastcol1_t2 + 3 + 10))
'        .AutoFilter
'        .Rows(8).AutoFilter field:=(lastcol1_t1 + lastcol1_t2 + 3), Criteria1:="="
'    End With
'
'    If thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + lastcol1_t2 + 3).SpecialCells(xlCellTypeVisible).Count > 1 Then
'         For Each cell In thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + lastcol1_t2 + 3).SpecialCells(xlCellTypeVisible)
'             If cell.Value = "" Then
'                thisSheet.Range(thisSheet.Cells(cell.row, lastcol1_t1 + lastcol1_t2 + 3), thisSheet.Cells(cell.row, lastcol1_t1 + lastcol1_t2 + 3 + 10)).Interior.Color = RGB(176, 176, 176)
'             End If
'         Next cell
'    End If
    
    thisSheet.AutoFilterMode = False
    
    Call Compare.compare3(thisSheet, thisSheet.Range(thisSheet.Cells(8, 2 * lastcol1_t1 + 3), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 2)), 2 * lastcol1_t1 + 2, lastcol1_t1 + 1)
    Call Sumary.Summary(thisSheet, 7, lastcol1_t1 * 3 + 4, lastrow3, lastcol1_t1, lastcol1_t1 + 1)
    
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol1_t1 + 2), thisSheet.Cells(1, lastcol1_t1 + 2)))
    Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 2 * lastcol1_t1 + 3), thisSheet.Cells(1, 2 * lastcol1_t1 + 3)))
    Call DeleteRow(thisSheet, 2, 6)
   
'    Call compare(thisSheet, 8, lastrow3, lastcol1_t1, lastcol1_t1 + 1, 2 * (lastcol1_t1 + 1))
'    Call Summary(thisSheet, 7, lastcol1_t1 * 3 + 4, lastrow3, lastcol1_t1 * 3 + 4, lastcol1_t1)
    
    ' B?t l?i các tính nang
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

