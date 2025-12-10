Sub copynet(ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, wbdict As Workbook)
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
    Dim dictFrameList_ADAS As Object
    Set dictFrameList_base = CreateObject("Scripting.Dictionary")
    Set dictFrameList_comp = CreateObject("Scripting.Dictionary")
    Set dictFrameList_gen = CreateObject("Scripting.Dictionary")
    Set dictFrameList_ADAS = CreateObject("Scripting.Dictionary")
    Set thisSheet = wbdict.Sheets("Network Path")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
''OBJ_Planning Sheet 1

    ws1.Activate
    lastcol1_t1 = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
    lastrow1 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).row
    
    If Sheet2.get_optBtn2_Click() Then
        GoTo copy_table_1
    End If
    
    Dim coltoDel As Collection
    Set coltoDel = New Collection
    
    For i = 6 To lastcol1_t1 - 1
        If InStr(ws1.Cells(4, i).Value, "CH3-CAN") = 0 And InStr(ws1.Cells(4, i).Value, "CH2-CAN") = 0 And InStr(ws1.Cells(4, i).Value, "ITS1-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS2-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS3-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS4-FD") = 0 And InStr(ws1.Cells(4, i).Value, "ITS5-FD") = 0 Then
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
    
copy_table_1:
    Dim data1 As Variant
    For i = 5 To lastrow1
      ws1.Cells(i, lastcol1_t1 + 2).Value = ws1.Cells(i, 2).Value & ws1.Cells(i, 3).Value & ws1.Cells(i, lastcol1_t1).Value
    Next i
    
    data1 = ws1.Range(ws1.Cells(5, lastcol1_t1 + 2), ws1.Cells(lastrow1, lastcol1_t1 + 2)).Value 'ganbien
    For i = 1 To UBound(data1, 1)
        If data1(i, 1) <> "" Then
            If dictFrameList_base.Exists(data1(i, 1)) = False Then dictFrameList_base.Add data1(i, 1), i
        End If
    Next i
    
'OBJ_Planning Sheet 2
    ws2.Activate
    lastcol_t2 = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column
    lastrow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).row
    
    If Sheet2.get_optBtn2_Click() Then
        GoTo copy_table_2
    End If
    
    Dim coltoDel2 As Collection
    Set coltoDel2 = New Collection
    For i = 6 To lastcol_t2 - 1
        If InStr(ws2.Cells(4, i).Value, "CH3-CAN") = 0 And InStr(ws2.Cells(4, i).Value, "CH2-CAN") = 0 And InStr(ws2.Cells(4, i).Value, "ITS1-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS2-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS3-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS4-FD") = 0 And InStr(ws2.Cells(4, i).Value, "ITS5-FD") = 0 Then
           coltoDel2.Add i
        End If
    Next i
    
    For i = coltoDel2.Count To 1 Step -1
        ws2.Columns(coltoDel2(i)).Delete
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
    Next
copy_table_2:

    Dim data2 As Variant
    
    For i = 5 To lastrow2
        ws2.Cells(i, lastcol1_t1 + 2).Value = ws2.Cells(i, 2).Value & ws2.Cells(i, 3).Value & ws2.Cells(i, lastcol1_t1).Value
    Next i
    
    data2 = ws2.Range(ws2.Cells(5, lastcol1_t1 + 2), ws2.Cells(lastrow2, lastcol1_t1 + 2)).Value
    
    For i = 1 To UBound(data2, 1)
        If data2(i, 1) <> "" Then
            If dictFrameList_comp.Exists(data2(i, 1)) = False Then dictFrameList_comp.Add data2(i, 1), i
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
'ADASMSg
    If Sheet2.ADASmsg.Value Then
        Dim data3 As Variant
        Dim lastrow4 As Integer
        Dim lastcol4 As Integer
        
        lastcol4 = ws3.Cells(1, ws3.Columns.Count).End(xlToLeft).Column
        lastrow4 = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).row
        
        Dim coltoDel3 As Collection
        Set coltoDel3 = New Collection
        For i = 5 To lastcol4 - 1
            If InStr(ws3.Cells(1, i).Value, "CH3-CAN") = 0 And InStr(ws3.Cells(1, i).Value, "CH2-CAN") = 0 And InStr(ws3.Cells(1, i).Value, "ITS1-FD") = 0 And InStr(ws3.Cells(1, i).Value, "ITS2-FD") = 0 And InStr(ws3.Cells(1, i).Value, "ITS3-FD") = 0 And InStr(ws3.Cells(1, i).Value, "ITS4-FD") = 0 And InStr(ws3.Cells(1, i).Value, "ITS5-FD") = 0 Then
               coltoDel3.Add i
            End If
        Next i
        
        For i = coltoDel3.Count To 1 Step -1
            ws3.Columns(coltoDel3(i)).Delete
            lastcol4 = lastcol4 - 1
        Next i
        
        For i = 2 To lastrow4
             ws3.Cells(i, lastcol4 + 2).Value = ws3.Cells(i, 1).Value & ws3.Cells(i, 2).Value & ws3.Cells(i, lastcol4).Value
        Next i
        
        data3 = ws3.Range(ws3.Cells(2, lastcol4 + 2), ws3.Cells(lastrow4, lastcol4 + 2)).Value
        
        For i = 1 To UBound(data3, 1)
            If data3(i, 1) <> "" Then
                If dictFrameList_ADAS.Exists(data3(i, 1)) = False Then dictFrameList_ADAS.Add data3(i, 1), i
            End If
        Next i
        
        For Each keyy In dictFrameList_ADAS.Keys
            If Not dictFrameList_gen.Exists(keyy) Then
                dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 5
            End If
        Next keyy
    End If
    
'---------------------------------------
    
'Copy title
 ' title planning sheet 1 copy
    Set copyRange = ws1.Range(ws1.Cells(1, 1), ws1.Cells(4, lastcol1_t1))
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)

 ' title planning sheet 2 copy
    Set copyRange = ws2.Range(ws2.Cells(1, 1), ws2.Cells(4, lastcol1_t1))
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + 2)
    
    If Sheet2.ADASmsg.Value Then
        'ADASMsg title
        Set copyRange = ws3.Range(ws3.Cells(1, 1), ws3.Cells(1, lastcol4))
        copyRange.Copy Destination:=thisSheet.Cells(4, lastcol1_t1 * 2 + 3)
        ' Planning sheet compare result title
        Set copyRange = thisSheet.Range(thisSheet.Cells(4, 14), thisSheet.Cells(4, lastcol1_t1 * 3 + 1))
        copyRange.Copy Destination:=thisSheet.Cells(4, 3 * lastcol1_t1 + 2)
    Else
        ' Planning sheet compare result title
        Set copyRange = thisSheet.Range(thisSheet.Cells(4, 14), thisSheet.Cells(4, lastcol1_t1 * 2 + 1))
        copyRange.Copy Destination:=thisSheet.Cells(4, 2 * lastcol1_t1 + 2)
    End If
'copy data
    For Each keyy In dictFrameList_gen.Keys
        If dictFrameList_base.Exists(keyy) Then
            Set copyRange = ws1.Range(ws1.Cells(dictFrameList_base(keyy) + 4, 1), ws1.Cells(dictFrameList_base(keyy) + 4, lastcol1_t1))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), 1)
        End If
        If dictFrameList_comp.Exists(keyy) Then
            Set copyRange = ws2.Range(ws2.Cells(dictFrameList_comp(keyy) + 4, 1), ws2.Cells(dictFrameList_comp(keyy) + 4, lastcol1_t1))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 + 2)
        End If
        If Sheet2.ADASmsg.Value Then
            If dictFrameList_ADAS.Exists(keyy) Then
                Set copyRange = ws3.Range(ws3.Cells(dictFrameList_ADAS(keyy) + 1, 1), ws3.Cells(dictFrameList_ADAS(keyy) + 1, lastcol4))
                copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 * 2 + 3)
            End If
        End If
    Next keyy
    
    thisSheet.Activate
    lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row > _
        thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 2).End(xlUp).row)
    If Sheet2.ADASmsg.Value Then
        lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2 * lastcol1_t1 + 3).End(xlUp).row > lastRow, thisSheet.Cells(thisSheet.Rows.Count, 2 * lastcol1_t1 + 3).End(xlUp).row, lastrow3)
    End If

'fill gray in base
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

    
'fill gray in obj
    If Sheet2.ADASmsg.Value Then
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
        
    End If
    thisSheet.AutoFilterMode = False
    
'fill gray in ADASmsg
    If Sheet2.ADASmsg.Value Then
        With thisSheet.Range(thisSheet.Cells(4, 1), thisSheet.Cells(lastrow3, 2 * lastcol1_t1 + 3))
            .AutoFilter
            .Rows(5).AutoFilter Field:=(2 * lastcol1_t1 + 3), Criteria1:="="
        End With
    
        If thisSheet.AutoFilter.Range.Columns(2 * lastcol1_t1 + 3).SpecialCells(xlCellTypeVisible).Count > 1 Then
             For Each cell In thisSheet.AutoFilter.Range.Columns(2 * lastcol1_t1 + 3).SpecialCells(xlCellTypeVisible)
                 If cell.Value = "" Then
                    thisSheet.Range(thisSheet.Cells(cell.row, 2 * lastcol1_t1 + 3), thisSheet.Cells(cell.row, 3 * lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
                 End If
             Next cell
        End If
    End If
    thisSheet.AutoFilterMode = False

'Compare
    If Sheet2.ADASmsg.Value Then
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(5, 3 * lastcol1_t1 + 3), thisSheet.Cells(lastrow3, 4 * lastcol1_t1 + 2)), 3 * lastcol1_t1 + 2, 2 * lastcol1_t1 + 1)
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(5, 4 * lastcol1_t1 + 4), thisSheet.Cells(lastrow3, 5 * lastcol1_t1 + 2)), 3 * lastcol1_t1 + 2, 2 * lastcol1_t1 + 1)
        Call Summary(thisSheet, 4, lastcol1_t1 * 5 + 6, lastrow3, lastcol1_t1 - 1, lastcol1_t1 + 2)
        Call Sumary2(thisSheet, thisSheet.Range(thisSheet.Cells(5, lastcol1_t1 * 3 + 3), thisSheet.Cells(lastrow3, lastcol1_t1 * 4 + 2)), thisSheet.Range(thisSheet.Cells(5, lastcol1_t1 * 5 + 4), thisSheet.Cells(lastrow3, lastcol1_t1 * 5 + 4)), lastrow3)
        ' sao chep tieu de sumary
        thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 8), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 13)).Merge
        thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 8).Value = "Œv‰æ‘‚ÆŒv‰æ‘”äŠr"
        With thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 8), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 13))
            .Interior.Color = RGB(0, 255, 0)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End With
        
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 6).Value = "Œv‰æ‘‚Ì·•ª"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 8).Value = "ˆê’v/•sˆê’v"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 9).Value = "”»’è"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 10).Value = "·•ª“à—e"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 11).Value = "Œ©‰ðE”õl"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 12).Value = "•â‘«î•ñ"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 13).Value = "Tag"
        With thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 6), thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 6))
            .Interior.Color = RGB(0, 255, 0)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
    
        With thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 8), thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 13))
            .Interior.Color = RGB(0, 255, 0)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
        
        thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 15), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 17)).Merge
        thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 15).Value = "‘O‰ñFB(ADAS5)"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 15).Value = "FB"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 16).Value = "FB“à—e"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 17).Value = "‘Î‰žó‹µ"
        
        With thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 15), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 17))
            .Interior.Color = RGB(0, 255, 0)
        End With
        With thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 15), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 2 * lastcol4 + 17)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 19), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 21)).Merge
        thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 19).Value = "‘O‰ñFB(ADCU)"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 19).Value = "FB"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 20).Value = "FB“à—e"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 21).Value = "‘Î‰žó‹µ"
        
        With thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 19), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 21))
            .Interior.Color = RGB(0, 255, 0)
        End With
        With thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 19), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 2 * lastcol4 + 21)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 23), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 24)).Merge
        thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 23).Value = "¡‰ñ‚Ì”»’f"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 23).Value = "‘ÎÛ•ñ"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 24).Value = "Œ‹˜_"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 2 * lastcol4 + 25).Value = "FB“à—e"
        
        With thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 23), thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 25))
            .Interior.Color = RGB(0, 255, 0)
        End With
        With thisSheet.Range(thisSheet.Cells(3, 3 * lastcol1_t1 + 2 * lastcol4 + 23), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 2 * lastcol4 + 25)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
    Else
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(5, 2 * lastcol1_t1 + 3), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 2)), 2 * lastcol1_t1 + 2, lastcol1_t1 + 1)
        ' tao bang tong ket keikakusho va messagelist
        thisSheet.Cells(4, 3 * lastcol1_t1 + 4).Value = "Œv‰æ‘‚Ì·•ª"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 5).Value = "”»’è"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 6).Value = "·•ª“à—e"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 7).Value = "Œ©‰ðE”õl"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 8).Value = "•â‘«î•ñ"
        thisSheet.Cells(4, 3 * lastcol1_t1 + 9).Value = "Tag"
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 5), thisSheet.Cells(1, 3 * lastcol1_t1 + 9)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 5), thisSheet.Cells(1, 3 * lastcol1_t1 + 9)).Value = "Œv‰æ‘‚ÆŒv‰æ‘”äŠr"
            
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 11), thisSheet.Cells(1, 3 * lastcol1_t1 + 13)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 11), thisSheet.Cells(1, 3 * lastcol1_t1 + 13)).Value = "‘O‰ñFB(ADAS5)"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 11).Value = "FB"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 12).Value = "FB“à—e"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 13).Value = "‘Î‰žó‹µ"
            
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 15), thisSheet.Cells(1, 3 * lastcol1_t1 + 17)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 15), thisSheet.Cells(1, 3 * lastcol1_t1 + 17)).Value = "‘O‰ñFB(ADCU)"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 15).Value = "FB"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 16).Value = "FB“à—e"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 17).Value = "‘Î‰žó‹µ"
            
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 19), thisSheet.Cells(1, 3 * lastcol1_t1 + 20)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 19), thisSheet.Cells(1, 3 * lastcol1_t1 + 20)).Value = "¡‰ñ‚Ì”»’f"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 19).Value = "‘ÎÛ•ñ"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 20).Value = "Œ‹˜_"
            thisSheet.Cells(4, 3 * lastcol1_t1 + 21).Value = "FB“à—e"
        
        With thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 4), thisSheet.Cells(4, 3 * lastcol1_t1 + 21))
            .Interior.Color = RGB(0, 255, 0)
        End With
        
        With thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 4), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 21))
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
        
        Call Summary(thisSheet, 4, lastcol1_t1 * 3 + 4, lastrow3, lastcol1_t1, lastcol1_t1 + 1)
    End If
    
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol1_t1 + 2), thisSheet.Cells(1, lastcol1_t1 + 2)))
    If Sheet2.ADASmsg.Value Then
        Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox3.Value), Range(thisSheet.Cells(1, 2 * lastcol1_t1 + 3), thisSheet.Cells(1, 2 * lastcol1_t1 + 3)))
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 2), thisSheet.Cells(1, 3 * lastcol1_t1 + 2)))
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆADASMsgList”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 4 * lastcol1_t1 + 3), thisSheet.Cells(1, 4 * lastcol1_t1 + 3)))
    Else
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 2 * lastcol1_t1 + 3), thisSheet.Cells(1, 2 * lastcol1_t1 + 3)))
    End If
    Call DeleteRow(thisSheet, 2, 3)
    
    ' B?t l?i các tính nang
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

