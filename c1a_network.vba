
Sub copynet2(ws1 As Worksheet, ws2 As Worksheet, wbdict As Workbook)
    Dim lastcol1_t1 As Integer
    Dim lastcol_t2 As Integer
    Dim lastcol2 As Integer
    Dim lastcol3 As Integer
    Dim firstcol3 As Integer
    Dim lastrow1 As Integer
    Dim lastrow2 As Integer
    Dim lastrow3 As Integer
    Dim count_col_after_del As Integer
    Dim ECUcol1 As Integer
    Dim ECUcol2 As Integer
    Dim i As Integer
    Dim keyFrame As String
    Dim keyy As Variant
    Dim foundCol As Integer
    Dim thisSheet As Worksheet
    Dim copyRange As Range
    Dim dictFrameList_base As Object
    Dim dictFrameList_comp As Object
    Dim dictFrameList_gen As Object
    
    ' Kh?i t?o các t? di?n
    Set dictFrameList_base = CreateObject("Scripting.Dictionary")
    Set dictFrameList_comp = CreateObject("Scripting.Dictionary")
    Set dictFrameList_gen = CreateObject("Scripting.Dictionary")
    Set thisSheet = wbdict.Sheets("Network Path")
    
    ' T?t các tính nang không c?n thi?t
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
'base

    
    ' Kích ho?t trang tính ws1
    ws1.Activate
    
    lastcol1_t1 = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
    lastrow1 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).row
    
    If Sheet2.OptionButton2 Then
        GoTo copy_table_1
    End If
    
'--------------------------------------------
' C1A khong can dung doan code nay
    Dim coltoDel As Collection
    
    Set coltoDel = New Collection
    
    For i = 6 To lastcol1_t1 - 1
        If InStr(ws1.Cells(4, i).Value, "CH2-CAN") = 0 And InStr(ws1.Cells(4, i).Value, "ITS1-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS2-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS3-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS4-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS5-FD") = 0 Then
           coltoDel.Add i
        End If
    Next i
    
    count_col_after_del = lastcol1_t1 - coltoDel.Count
    
    For i = coltoDel.Count To 1 Step -1
        ws1.Columns(coltoDel(i)).Delete
        lastcol1_t1 = lastcol1_t1 - 1
    Next i

    

    Dim rowsToDelete As Collection
    
    Set rowsToDelete = New Collection
    
    
    For i = 5 To lastrow1
     If InStr(ws1.Cells(i, lastcol1_t1).Value, "ADAS") = 0 And InStr(ws1.Cells(i, lastcol1_t1).Value, "FrCamADAS") = 0 Then
        rowsToDelete.Add i
     End If
    Next i


    For i = rowsToDelete.Count To 1 Step -1
        ws1.Rows(rowsToDelete(i)).Delete
        lastrow1 = lastrow1 - 1
    Next i
'-------------------------------------------

copy_table_1:
      
    Dim data1 As Variant
    For i = 5 To lastrow1
      ws1.Cells(i, lastcol1_t1 + 2).Value = ws1.Cells(i, 2).Value & ws1.Cells(i, 3).Value & ws1.Cells(i, lastcol1_t1).Value
    Next i
    data1 = ws1.Range(ws1.Cells(5, lastcol1_t1 + 2), ws1.Cells(lastrow1, lastcol1_t1 + 2)).Value 'ganbien
    
    For i = 1 To UBound(data1, 1)
        If data1(i, 1) <> "" Then
            If dictFrameList_base.Exists(data1(i, 1)) = False Then dictFrameList_base.Add data1(i, 1), i + 2
        End If
    Next i
    
    

    ws2.Activate
    lastcol_t2 = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column
    lastrow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).row
    
    If Sheet2.OptionButton2 Then
        GoTo copy_table_2
    End If
    
'-----------------------------

    Dim coltoDel2 As Collection
    
    Set coltoDel2 = New Collection
    
    For i = 6 To lastcol_t2 - 1
        If InStr(ws2.Cells(4, i).Value, "CH2-CAN") = 0 And InStr(ws2.Cells(4, i).Value, "ITS1-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS2-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS3-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS4-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS5-FD") = 0 Then
           coltoDel2.Add i
        End If
    Next i
    
    For i = coltoDel.Count To 1 Step -1
        ws2.Columns(coltoDel(i)).Delete
        lastcol_t2 = lastcol1_t2 - 1
    Next i
    
    Dim rowsToDelete2 As Collection
    
    Set rowsToDelete2 = New Collection
    
    
    For i = 5 To lastrow2
     If InStr(ws2.Cells(i, lastcol1_t1).Value, "ADAS") = 0 And InStr(ws2.Cells(i, lastcol1_t1).Value, "FrCamADAS") = 0 Then
        rowsToDelete2.Add i
        lastrow2 = lastrow2 - 1
     End If
    Next i

    For i = rowsToDelete2.Count To 1 Step -1
        ws2.Rows(rowsToDelete2(i)).Delete
    Next i
'-----------------------------------------

copy_table_2:
    Dim data2 As Variant
    
    For i = 5 To lastrow2
      ws2.Cells(i, lastcol1_t1 + 2).Value = ws2.Cells(i, 2).Value & ws2.Cells(i, 3).Value & ws2.Cells(i, lastcol1_t1).Value
    Next i
    
    data2 = ws2.Range(ws2.Cells(5, lastcol1_t1 + 2), ws2.Cells(lastrow2, lastcol1_t1 + 2)).Value
 
    For i = 1 To UBound(data2, 1)
        If data2(i, 1) <> "" Then
            If dictFrameList_comp.Exists(data2(i, 1)) = False Then dictFrameList_comp.Add data2(i, 1), i + 2
        End If
    Next i
    
   
    For Each keyy In dictFrameList_base.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 5
        End If
    Next keyy
    
    For Each keyy In dictFrameList_comp.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 5
        End If
    Next keyy

'copy
    
    ' Sao chép tiêu d? t? ws1 sang thisSheet
    Set copyRange = ws1.Range(ws1.Cells(1, 1), ws1.Cells(4, lastcol1_t1))
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)

    ' Sao chép tiêu d? t? ws2 sang thisSheet
    Set copyRange = ws2.Range(ws2.Cells(1, 1), ws2.Cells(4, lastcol1_t1))
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + 2)
    
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 * 2 + 3)

    ' Sao chép d? li?u t? ws1 và ws2 sang thisSheet d?a trên dictFrameList_gen
    For Each keyy In dictFrameList_gen.Keys
        If dictFrameList_base.Exists(keyy) Then
            Set copyRange = ws1.Range(ws1.Cells(dictFrameList_base(keyy) + 2, 1), ws1.Cells(dictFrameList_base(keyy) + 2, lastcol1_t1))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), 1)
        End If
        If dictFrameList_comp.Exists(keyy) Then
            Set copyRange = ws2.Range(ws2.Cells(dictFrameList_comp(keyy) + 2, 1), ws2.Cells(dictFrameList_comp(keyy) + 2, lastcol1_t1))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 + 2)
        End If
    Next keyy
    
    
    thisSheet.Activate
    lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row > _
        thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 2).End(xlUp).row)
    
'--------------------
'boi xam base
    With thisSheet.Range(thisSheet.Cells(4, 1), thisSheet.Cells(lastrow3, lastcol1_t1))
        .AutoFilter
        .Rows(5).AutoFilter Field:=3, Criteria1:="="
    End With

    If thisSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
         For Each cell In thisSheet.AutoFilter.Range.Columns(3).SpecialCells(xlCellTypeVisible)
             If cell.Value = "" Then
                thisSheet.Range(thisSheet.Cells(cell.row, 1), thisSheet.Cells(cell.row, lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
             End If
         Next cell
    End If
     
    thisSheet.AutoFilterMode = False

    
'boi xam obj

    With thisSheet.Range(thisSheet.Cells(4, 1), thisSheet.Cells(lastrow3, 2 * lastcol1_t1 + 1))
        .AutoFilter
        .Rows(5).AutoFilter Field:=lastcol1_t1 + 4, Criteria1:="="
    End With

    If thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + 2).SpecialCells(xlCellTypeVisible).Count > 1 Then
         For Each cell In thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + 4).SpecialCells(xlCellTypeVisible)
             If cell.Value = "" Then
                thisSheet.Range(thisSheet.Cells(cell.row, lastcol1_t1 + 2), thisSheet.Cells(cell.row, 2 * lastcol1_t1 + 1)).Interior.Color = RGB(191, 191, 191)
             End If
         Next cell
    End If
    
    thisSheet.AutoFilterMode = False
'-----------------------------------------------------

    Call Compare.compare3(thisSheet, thisSheet.Range(thisSheet.Cells(5, 2 * lastcol1_t1 + 3), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 2)), 2 * lastcol1_t1 + 2, lastcol1_t1 + 1)
    Call Sumary.Summary(thisSheet, 4, lastcol1_t1 * 3 + 4, lastrow3, lastcol1_t1, lastcol1_t1 + 1)
    
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol1_t1 + 2), thisSheet.Cells(1, lastcol1_t1 + 2)))
    Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 2 * lastcol1_t1 + 3), thisSheet.Cells(1, 2 * lastcol1_t1 + 3)))
    Call DeleteRow(thisSheet, 2, 3)

'    Call compare(thisSheet, 5, lastrow3, lastcol1_t1, lastcol1_t1 + 1, 2 * (lastcol1_t1 + 1))
'    Call Summary(thisSheet, 4, lastcol1_t1 * 3 + 4, lastrow3, lastcol1_t1 * 3 + 4, lastcol1_t1)
    
    ' B?t l?i các tính nang
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub



