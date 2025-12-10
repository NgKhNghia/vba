Sub copyframe(ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, wbdict As Workbook)
    Dim lastcol1_t1 As Integer
    Dim lastcol1_t2 As Integer
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
    Dim keyFrame As String
    Dim keyy As Variant
    Dim foundCol As Integer
    Dim thisSheet As Worksheet
    Dim copyRange As Range
    
    Dim dictFrameList_base As Object
    Dim dictFrameList_comp As Object
    Dim dictFrameList_ADAS As Object
    Dim dictFrameList_gen As Object
    Set dictFrameList_base = CreateObject("Scripting.Dictionary")
    Set dictFrameList_comp = CreateObject("Scripting.Dictionary")
    Set dictFrameList_ADAS = CreateObject("Scripting.Dictionary")
    Set dictFrameList_gen = CreateObject("Scripting.Dictionary")
    Set thisSheet = wbdict.Sheets("Frame Synthesis")
    
   
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
'BASE
    ws1.Activate
    ws1.AutoFilterMode = False
    
    Call XuLyInput(ws1, ws2)
    
    lastcol1_t1 = ws1.Cells(7, ws1.Columns.Count).End(xlToLeft).Column
    lastrow1 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).row
    

    column_adas = ws1.Rows(7).Find("ADAS", LookIn:=xlValues, lookat:=xlWhole).Column
    column_adas_bridge = ws1.Rows(7).Find("ADAS_Bridge", LookIn:=xlValues, lookat:=xlWhole).Column
    ws1.Cells(8, lastcol1_t1 + 2).Formula = "=" & ws1.Cells(8, column_adas).Address(False, False) & " & " & ws1.Cells(8, column_adas_bridge).Address(False, False)
    ws1.Range(ws1.Cells(8, lastcol1_t1 + 2), ws1.Cells(lastrow1, lastcol1_t1 + 2)).FillDown
    
    
    ws1.Range(ws1.Cells(7, lastcol1_t1 + 2), ws1.Cells(lastrow1, lastcol1_t1 + 2)).AutoFilter Field:=1, Criteria1:="="
    ws1.Range(ws1.Cells(8, lastcol1_t1 + 2), ws1.Cells(lastrow1, lastcol1_t1 + 2)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ws1.AutoFilterMode = False
    
    
    Dim data1 As Variant
    data1 = ws1.Range(ws1.Cells(8, 2), ws1.Cells(lastrow1, 2)).Value

    For i = 1 To UBound(data1, 1)
        If data1(i, 1) <> "" Then
            If dictFrameList_base.Exists(data1(i, 1)) = False Then dictFrameList_base.Add data1(i, 1), i
        End If
    Next i
    
'OBJ
    
    ws2.Activate
    ws2.AutoFilterMode = False
    
    lastcol1_t2 = ws2.Cells(7, ws2.Columns.Count).End(xlToLeft).Column
    lastrow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).row
    
    column_adas = ws2.Rows(7).Find("ADAS", LookIn:=xlValues, lookat:=xlWhole).Column
    column_adas_bridge = ws2.Rows(7).Find("ADAS_Bridge", LookIn:=xlValues, lookat:=xlWhole).Column
    
    ws2.Cells(8, lastcol1_t2 + 2).Formula = "=" & ws2.Cells(8, column_adas).Address(False, False) & " & " & ws2.Cells(8, column_adas_bridge).Address(False, False)
    ws2.Range(ws2.Cells(8, lastcol1_t2 + 2), ws2.Cells(lastrow2, lastcol1_t2 + 2)).FillDown
    
    ws2.Range(ws2.Cells(7, lastcol1_t2 + 2), ws2.Cells(lastrow2, lastcol1_t2 + 2)).AutoFilter Field:=1, Criteria1:="="
    ws2.Range(ws2.Cells(8, lastcol1_t2 + 2), ws2.Cells(lastrow2, lastcol1_t2 + 2)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    
    ws2.AutoFilterMode = False
    
    Dim data2 As Variant
    data2 = ws2.Range(ws2.Cells(8, 2), ws2.Cells(lastrow2, 2)).Value
    For i = 1 To UBound(data2, 1)
        If data2(i, 1) <> "" Then
            If dictFrameList_comp.Exists(data2(i, 1)) = False Then dictFrameList_comp.Add data2(i, 1), i
        End If
    Next i
    
  
'ADASMsg

    If Sheet2.ADASmsg.Value Then
        lastcol4 = ws3.Cells(1, ws3.Columns.Count).End(xlToLeft).Column
        lastrow4 = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).row
    
        
        Dim data3 As Variant
        
        data3 = ws3.Range(ws3.Cells(2, 1), ws3.Cells(lastrow2, 1)).Value
         
        For i = 1 To UBound(data3, 1)
            If data3(i, 1) <> "" Then
                If dictFrameList_ADAS.Exists(data3(i, 1)) = False Then dictFrameList_ADAS.Add data3(i, 1), i
            End If
        Next i
    End If
    
'insert data to dic
    For Each keyy In dictFrameList_base.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 8
        End If
    Next keyy
    
    For Each keyy In dictFrameList_comp.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 8
        End If
    Next keyy
    
    If Sheet2.ADASmsg.Value Then
    For Each keyy In dictFrameList_ADAS.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 8
        End If
    Next keyy
    End If
    
'Copy tieu de
 ' title planning sheet 1 copy
    Set copyRange = ws1.Range(ws1.Cells(1, 1), ws1.Cells(7, lastcol1_t1))
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)
    
 ' title planning sheet 2 copy
    Set copyRange = ws2.Range(ws2.Cells(1, 1), ws2.Cells(7, lastcol1_t2))
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + 2)
    
    
    If Sheet2.ADASmsg.Value Then
        ' Planning sheet compare result title
        Set copyRange = thisSheet.Range(thisSheet.Cells(7, lastcol1_t1 + 2), thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 1))
        copyRange.Copy Destination:=thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 15)
        'ADASMsg title
        Set copyRange = ws3.Range(ws3.Cells(1, 1), ws3.Cells(1, lastcol4))
        copyRange.Copy Destination:=thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 3)
        
        'ADASMsg compare result title
        Set copyRange = thisSheet.Range(thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 3), thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 13))
        copyRange.Copy Destination:=thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 16 + lastcol1_t2)
    Else
         ' Planning sheet compare result title
        Set copyRange = thisSheet.Range(thisSheet.Cells(7, lastcol1_t1 + 2), thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 1))
        copyRange.Copy Destination:=thisSheet.Cells(7, lastcol1_t1 + lastcol1_t2 + 3)
    
    End If

    ' copy data
    For Each keyy In dictFrameList_gen.Keys
    'planning sheet 1
        If dictFrameList_base.Exists(keyy) Then
            Set copyRange = ws1.Range(ws1.Cells(dictFrameList_base(keyy) + 7, 1), ws1.Cells(dictFrameList_base(keyy) + 7, lastcol1_t1))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), 1)
        End If
    'planning sheet 2
        If dictFrameList_comp.Exists(keyy) Then
            Set copyRange = ws2.Range(ws2.Cells(dictFrameList_comp(keyy) + 7, 1), ws2.Cells(dictFrameList_comp(keyy) + 7, lastcol1_t2))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 + 2)
        End If
    'ADASMsg 2
        If Sheet2.ADASmsg.Value Then
            If dictFrameList_ADAS.Exists(keyy) Then
                Set copyRange = ws3.Range(ws3.Cells(dictFrameList_ADAS(keyy) + 1, 1), ws3.Cells(dictFrameList_ADAS(keyy) + 1, lastcol4))
                copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 + lastcol1_t2 + 3)
            End If
        End If
    Next keyy
    
    thisSheet.Activate

    lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row > thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 3).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 3).End(xlUp).row)
    If Sheet2.ADASmsg.Value Then
        lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2 * lastcol1_t1 + 3).End(xlUp).row > lastrow3, thisSheet.Cells(thisSheet.Rows.Count, 2 * lastcol1_t1 + 3).End(xlUp).row, lastrow3)
    End If

   ' Gray fill planning sheet 1
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
    
    ' Gray fill planning sheet 2

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
    
    ' Gray fill ADASMsg
    If Sheet2.ADASmsg.Value Then
        With thisSheet.Range(thisSheet.Cells(7, 1), thisSheet.Cells(lastrow3, lastcol1_t1 + lastcol1_t2 + 3 + 10))
            .AutoFilter
            .Rows(8).AutoFilter Field:=(lastcol1_t1 + lastcol1_t2 + 3), Criteria1:="="
        End With
    
        If thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + lastcol1_t2 + 3).SpecialCells(xlCellTypeVisible).Count > 1 Then
             For Each cell In thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + lastcol1_t2 + 3).SpecialCells(xlCellTypeVisible)
                 If cell.Value = "" Then
                    thisSheet.Range(thisSheet.Cells(cell.row, lastcol1_t1 + lastcol1_t2 + 3), thisSheet.Cells(cell.row, lastcol1_t1 + lastcol1_t2 + 3 + 10)).Interior.Color = RGB(191, 191, 191)
                 End If
             Next cell
        End If
    End If
    thisSheet.AutoFilterMode = False
    
    
    
'compare
    If Sheet2.ADASmsg.Value Then
    'planning sheet1 & planning sheet 2
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(8, 2 * lastcol1_t1 + 15), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 14)), 2 * lastcol1_t1 + 14, lastcol1_t1 + 13)
    'planning sheet 2 & ADASMsg
        Call CompareMsgListWithObject(thisSheet, 8, 3 * lastcol1_t1 + lastcol4 + 5, lastrow3, 3 * lastcol1_t1 + lastcol4 + 15, 2 * lastcol1_t1 + lastcol4 + 2, lastcol1_t1 + lastcol4 + 2)
    'sum result
        Call SummaryFrame(thisSheet, 8, 3 * lastcol1_t1 + 2 * lastcol4 + 6, lastrow3, 3 * lastcol1_t1 + 2 * lastcol4 + 13, lastcol1_t1 + lastcol4 + 2, lastcol4 + 1)
    ' tao bang tong ket keikakusho va messagelist
        thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 6).Value = "Œv‰æ‘‚Ì·•ª"
        thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 8).Value = "ˆê’v/•sˆê’v"
        thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 9).Value = "”»’è"
        thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 10).Value = "·•ª“à—e"
        thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 11).Value = "Œ©‰ðE”õl"
        thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 12).Value = "•â‘«î•ñ"
        thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 13).Value = "Tag"
    
        With thisSheet.Range(thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 6), thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 6))
            .Interior.Color = RGB(0, 255, 0)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
    
        With thisSheet.Range(thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 8), thisSheet.Cells(7, 3 * lastcol1_t1 + 2 * lastcol4 + 13))
            .Interior.Color = RGB(0, 255, 0)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
    Else
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(8, 2 * lastcol1_t1 + 3), thisSheet.Cells(lastrow3, 3 * lastcol1_t1 + 2)), 2 * lastcol1_t1 + 2, lastcol1_t1 + 1)
        ' tao bang tong ket keikakusho va messagelist
        
            thisSheet.Cells(7, 3 * lastcol1_t1 + 4).Value = "Œv‰æ‘‚Ì·•ª"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 5).Value = "”»’è"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 6).Value = "·•ª“à—e"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 7).Value = "Œ©‰ðE”õl"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 8).Value = "•â‘«î•ñ"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 9).Value = "Tag"
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 5), thisSheet.Cells(1, 3 * lastcol1_t1 + 9)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 5), thisSheet.Cells(1, 3 * lastcol1_t1 + 9)).Value = "Œv‰æ‘‚ÆŒv‰æ‘”äŠr"
            
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 10), thisSheet.Cells(1, 3 * lastcol1_t1 + 12)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 10), thisSheet.Cells(1, 3 * lastcol1_t1 + 12)).Value = "‘O‰ñFB(ADAS5)"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 10).Value = "FB"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 11).Value = "FB“à—e"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 12).Value = "‘Î‰žó‹µ"
            
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 14), thisSheet.Cells(1, 3 * lastcol1_t1 + 16)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 14), thisSheet.Cells(1, 3 * lastcol1_t1 + 16)).Value = "‘O‰ñFB(ADCU)"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 14).Value = "FB"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 15).Value = "FB“à—e"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 16).Value = "‘Î‰žó‹µ"
            
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 18), thisSheet.Cells(1, 3 * lastcol1_t1 + 19)).Merge
            thisSheet.Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 18), thisSheet.Cells(1, 3 * lastcol1_t1 + 19)).Value = "¡‰ñ‚Ì”»’f"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 18).Value = "‘ÎÛ•ñ"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 19).Value = "Œ‹˜_"
            thisSheet.Cells(7, 3 * lastcol1_t1 + 20).Value = "FB“à—e"
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
       Call SummaryFrame(thisSheet, 8, 3 * lastcol1_t1 + 4, lastrow3, 3 * lastcol1_t1 + 2, lastcol1_t1 + 1, 0)
    End If
'layout
    If Sheet2.ADASmsg.Value Then
        Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
        Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol1_t1 + 2), thisSheet.Cells(1, lastcol1_t1 + 2)))
        Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox3.Value), Range(thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 3), thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 3)))
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 15), thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 15)))
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆADASMsgList”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 16 + lastcol1_t2), thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 16 + lastcol1_t2)))
    Else
        Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
        Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol1_t1 + 2), thisSheet.Cells(1, lastcol1_t1 + 2)))
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 3), thisSheet.Cells(1, lastcol1_t1 + lastcol1_t2 + 3)))
    End If
    Call DeleteRow(thisSheet, 2, 6)
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


