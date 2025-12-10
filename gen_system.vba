Sub copysystem(ws1 As Worksheet, ws2 As Worksheet, wbdict As Workbook)
    Dim lastcol As Integer
    Dim lastcol2 As Integer
    Dim lastRow As Integer
    Dim lastrow2 As Integer
    Dim lastrow3 As Integer
    Dim i, k As Integer
    Dim rng1, rng2 As Range
    Dim foundRow As Integer
    Dim keyy As Variant
    Dim copyRange As Range
    Dim thisSheet As Worksheet
    Dim dictListECU_base As Object
    Dim dictListECU_comp As Object
    Dim dictListECU_gen As Object
    Dim cell As Range
    
    Set thisSheet = wbdict.Sheets("System")
    Set dictListECU_base = CreateObject("Scripting.Dictionary")
    Set dictListECU_comp = CreateObject("Scripting.Dictionary")
    Set dictListECU_gen = CreateObject("Scripting.Dictionary")
    
    If Sheet2.OptionButton2.Value Then

'       Set thisSheet = ThisWorkbook.Sheets("system")
    Set rng1 = ws1.Rows(23)
    Set rng2 = ws2.Rows(23)
   
    lastcol = ws1.Cells(8, ws1.Columns.Count).End(xlToLeft).Column
    lastRow = ws1.Cells(ws1.Rows.Count, 3).End(xlUp).row
    lastcol2 = ws2.Cells(8, ws2.Columns.Count).End(xlToLeft).Column
    If lastcol2 <> lastcol Then
        MsgBox "The number of variants is not uniform!"
        Exit Sub
    End If
    
    For i = lastRow To 1 Step -1
        If Left(ws1.Cells(i, 3).Value, 3) = "DLC" Then
            foundRow = i
            Exit For
        End If
    Next i
    
    For Each cell In rng1.Cells
        If cell.Value = "3chCGW" Or cell.Value = "NP1" Then
            k = cell.Column
            Exit For
        End If
    Next cell
    For i = 24 To foundRow
        ws1.Cells(i, 2).Value = ws1.Cells(i, 3).Value & ws1.Cells(i, k).Value & ws1.Cells(i, k + 1).Value
    Next i
    
    Dim data1 As Variant
    
    data1 = ws1.Range(ws1.Cells(24, 2), ws1.Cells(foundRow, 2)).Value 'ganbien
    
    For i = 1 To UBound(data1, 1)
        If data1(i, 1) <> "" Then
            If dictListECU_base.Exists(data1(i, 1)) = False Then dictListECU_base.Add data1(i, 1), i
        End If
    Next i
'    For i = 24 To foundRow
'        If ws1.Cells(i, 2).Value = "" Then GoTo nexti1
'        If dictListECU_base.Exists(ws1.Cells(i, 2).Value) = False Then
'            dictListECU_base.Add ws1.Cells(i, 2).Value, i
'        End If
'    Next i
    
    ws2.Activate
    lastrow2 = ws2.Cells(ws2.Rows.Count, 3).End(xlUp).row
    lastcol2 = ws2.Cells(8, ws2.Columns.Count).End(xlToLeft).Column
    
    For i = lastrow2 To 1 Step -1
        If Left(ws2.Cells(i, 3).Value, 3) = "DLC" Then
            foundRow = i
            Exit For
        End If
    Next i
    
    For Each cell In rng2.Cells
        If cell.Value = "3chCGW" Or cell.Value = "NP1" Then
            k = cell.Column
            Exit For
        End If
    Next cell
     For i = 24 To foundRow
        ws2.Cells(i, 2).Value = ws2.Cells(i, 3).Value & ws2.Cells(i, k).Value & ws2.Cells(i, k + 1).Value
    Next i
    
    Dim data2 As Variant
    
    data2 = ws2.Range(ws2.Cells(24, 2), ws2.Cells(foundRow, 2)).Value 'ganbien
    
    For i = 1 To UBound(data2, 1)
        If data2(i, 1) <> "" Then
            If dictListECU_comp.Exists(data2(i, 1)) = False Then dictListECU_comp.Add data2(i, 1), i
        End If
    Next i
'    For i = 24 To foundRow
'        If ws2.Cells(i, 2).Value = "" Then GoTo nexti2
'        If dictListECU_comp.Exists(ws2.Cells(i, 2).Value) = False Then
'            dictListECU_comp.Add ws2.Cells(i, 2).Value, i
'        End If
'    Next i
    
    For Each keyy In dictListECU_base.Keys
        If Not dictListECU_gen.Exists(keyy) Then
            dictListECU_gen.Add keyy, dictListECU_gen.Count
        End If
    Next keyy
    
    For Each keyy In dictListECU_comp.Keys
        If Not dictListECU_gen.Exists(keyy) Then
            dictListECU_gen.Add keyy, dictListECU_gen.Count
        End If
    Next keyy
    
'    ws1.Activate
    Set copyRange = ws1.Range(ws1.Cells(1, 3), ws1.Cells(23, lastcol))
'    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)
    
'    ws2.Activate
    Set copyRange = ws2.Range(ws2.Cells(1, 3), ws2.Cells(23, lastcol2))
'    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol)
    copyRange.Copy Destination:=thisSheet.Cells(1, 2 * lastcol - 1)                ' -3 don vi-----------------------------
    
    For Each keyy In dictListECU_gen.Keys
        If dictListECU_base.Exists(keyy) Then
            If IsNumeric(dictListECU_base(keyy)) Then ' Ki?m tra giá tr? là s?
'                ws1.Activate
                Set copyRange = ws1.Range(ws1.Cells(CLng(dictListECU_base(keyy)) + 23, 3), ws1.Cells(CLng(dictListECU_base(keyy)) + 23, lastcol))
'                thisSheet.Activate
                copyRange.Copy Destination:=thisSheet.Cells(dictListECU_gen(keyy) + 24, 1)
            End If
        End If
        If dictListECU_comp.Exists(keyy) Then
            If IsNumeric(dictListECU_comp(keyy)) Then ' Ki?m tra giá tr? là s?
'                ws2.Activate
                Set copyRange = ws2.Range(ws2.Cells(CLng(dictListECU_comp(keyy)) + 23, 3), ws2.Cells(CLng(dictListECU_comp(keyy)) + 23, lastcol2))
'                thisSheet.Activate
                copyRange.Copy Destination:=thisSheet.Cells(dictListECU_gen(keyy) + 24, lastcol)
            End If
        End If
    Next keyy
    
    thisSheet.Activate
    thisSheet.Cells(1, 1) = lastcol + lastcol2 + 3
    thisSheet.Cells(1, 2) = lastcol + lastcol2 + 2 + lastcol2
    thisSheet.Cells(1, 3) = WorksheetFunction.Max(thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row, _
                        thisSheet.Cells(thisSheet.Rows.Count, lastcol).End(xlUp).row)                                   ' -2------------------
                        
    lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row > thisSheet.Cells(thisSheet.Rows.Count, lastcol).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol).End(xlUp).row)
    Call compare2(thisSheet, thisSheet.Range(thisSheet.Cells(24, lastcol * 2 - 1), thisSheet.Cells(lastrow3, lastcol * 3 - 4)), lastcol * 2 - 2, lastcol - 1)
    Call Summary(thisSheet, 23, lastcol * 3 - 2, lastrow3, lastcol - 2, lastcol - 1)
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol), thisSheet.Cells(1, lastcol)))
    Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 2 * lastcol - 1), thisSheet.Cells(1, 2 * lastcol - 1)))
    Call DeleteRow(thisSheet, 2, 4)
    
'    Call compare(thisSheet, 24, lastrow3, lastcol - 2, lastcol - 1, 2 * (lastcol - 1))
'
'    Call Summary(thisSheet, 23, lastcol * 3 - 2, lastrow3, lastcol * 3 + 2, lastcol - 2)
    
    
    End If
'---------------------------------------------------------------------------------------------
    If Sheet2.OptionButton1.Value Then
    lastcol = ws1.Cells(6, ws1.Columns.Count).End(xlToLeft).Column
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).row
    lastcol2 = ws2.Cells(6, ws2.Columns.Count).End(xlToLeft).Column
    
    If lastcol2 <> lastcol Then
        MsgBox "The number of variants is not uniform!"
        Exit Sub
    End If
    
    For i = lastRow To 1 Step -1
        If Left(ws1.Cells(i, 1).Value, 3) = "DLC" Then
            foundRow = i
            Exit For
        End If
    Next i
    
    Dim data3 As Variant
    
    data3 = ws1.Range(ws1.Cells(18, 1), ws1.Cells(foundRow, 1)).Value 'ganbien
    
    For i = 1 To UBound(data3, 1)
        If data3(i, 1) <> "" Then
            If dictListECU_base.Exists(data3(i, 1)) = False Then dictListECU_base.Add data3(i, 1), i
        End If
    Next i
    
'    For i = 18 To foundRow
'        If ws1.Cells(i, 1).Value = "" Then GoTo nexti1
'        If dictListECU_base.Exists(ws1.Cells(i, 1).Value) = False Then dictListECU_base.Add ws1.Cells(i, 1).Value, i
'    Next i
    
    ws2.Activate
    lastrow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).row
    lastcol2 = ws2.Cells(6, ws2.Columns.Count).End(xlToLeft).Column
    
    For i = lastrow2 To 1 Step -1
        If Left(ws2.Cells(i, 1).Value, 3) = "DLC" Then
            foundRow = i
            Exit For
        End If
     Next i
     
    Dim data4 As Variant
    
    data4 = ws2.Range(ws2.Cells(18, 1), ws2.Cells(foundRow, 1)).Value 'ganbien
    
    For i = 1 To UBound(data3, 1)
        If data4(i, 1) <> "" Then
            If dictListECU_comp.Exists(data4(i, 1)) = False Then dictListECU_comp.Add data4(i, 1), i
        End If
    Next i
    
'    For i = 18 To foundRow
'        If ws2.Cells(i, 1).Value = "" Then GoTo nexti2
'        If dictListECU_comp.Exists(ws2.Cells(i, 1).Value) = False Then dictListECU_comp.Add ws2.Cells(i, 1).Value, i
'    Next i
    
    For Each keyy In dictListECU_base.Keys
        If Not dictListECU_gen.Exists(keyy) Then
            dictListECU_gen.Add keyy, dictListECU_gen.Count + 18
        End If
    Next keyy
    
    For Each keyy In dictListECU_comp.Keys
        If Not dictListECU_gen.Exists(keyy) Then
            dictListECU_gen.Add keyy, dictListECU_gen.Count + 18
        End If
    Next keyy
    
    ws1.Activate
    Set copyRange = ws1.Range(ws1.Cells(1, 1), ws1.Cells(17, lastcol))
    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)
    
    ws2.Activate
    Set copyRange = ws2.Range(ws2.Cells(1, 1), ws2.Cells(17, lastcol2))
    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol + 2)
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol + lastcol2 + 3)
    
    For Each keyy In dictListECU_gen.Keys
        If dictListECU_base.Exists(keyy) Then
           ' ws1.Activate
            Set copyRange = ws1.Range(ws1.Cells(dictListECU_base(keyy) + 17, 1), ws1.Cells(dictListECU_base(keyy) + 17, lastcol))
           ' thisSheet.Activate
            copyRange.Copy Destination:=thisSheet.Cells(dictListECU_gen(keyy), 1)
        End If
        If dictListECU_comp.Exists(keyy) Then
           ' ws2.Activate
            Set copyRange = ws2.Range(ws2.Cells(dictListECU_comp(keyy) + 17, 1), ws2.Cells(dictListECU_comp(keyy) + 17, lastcol2))
           ' thisSheet.Activate
            copyRange.Copy Destination:=thisSheet.Cells(dictListECU_gen(keyy), lastcol + 2)
        End If
    Next keyy
    
    thisSheet.Activate
    thisSheet.Cells(1, 1) = lastcol + lastcol2 + 3
    thisSheet.Cells(1, 2) = lastcol + lastcol2 + 2 + lastcol2
    thisSheet.Cells(1, 3) = WorksheetFunction.Max(thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row, _
                        thisSheet.Cells(thisSheet.Rows.Count, lastcol + 2).End(xlUp).row)
                        
    lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row > thisSheet.Cells(thisSheet.Rows.Count, lastcol + 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol + 2).End(xlUp).row)
    
    Call compare2(thisSheet, thisSheet.Range(thisSheet.Cells(18, lastcol * 2 + 3), thisSheet.Cells(lastrow3, lastcol * 3 + 2)), lastcol * 2 + 2, lastcol + 1)
    Call Summary(thisSheet, 17, lastcol * 3 + 4, lastrow3, lastcol, lastcol + 1)
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol + 2), thisSheet.Cells(1, lastcol + 2)))
    Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 2 * lastcol + 3), thisSheet.Cells(1, 2 * lastcol + 3)))
    Call DeleteRow(thisSheet, 2, 4)
    
'    Call compare(thisSheet, 18, lastrow3, lastcol, lastcol + 1, 2 * (lastcol + 1))
'    Call Summary(thisSheet, 17, lastcol * 3 + 4, lastrow3, lastcol * 3 + 4, lastcol)
    
    End If
                        
End Sub

