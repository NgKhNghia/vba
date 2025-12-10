
Sub copyframe1(ws3 As Worksheet, ws4 As Worksheet, wbdict As Workbook)
    Dim lastcol1_t As Integer
    Dim lastcol1 As Integer
    Dim lastcol2 As Integer
    Dim lastcol3 As Integer
    Dim firstcol3 As Integer
    Dim lastrow1 As Integer
    Dim lastrow3 As Integer
    Dim lastrow2 As Integer
    Dim ECUcol1 As Integer
    Dim ECUcol2 As Integer
    Dim i As Integer
    Dim keyFrame As String
    Dim keyy As Variant
    Dim foundCol As Integer
    Dim thisSheet As Worksheet
    Dim copyRange As Range
    Dim dictFrameList_base As Dictionary
    Dim dictFrameList_comp As Dictionary
    Dim dictFrameList_gen As Dictionary
    Dim lastcol1_check As Integer
    
    Set dictFrameList_base = New Dictionary
    Set dictFrameList_comp = New Dictionary
    Set dictFrameList_gen = New Dictionary
    Set thisSheet = wbdict.Sheets("Framelist")
    
    ws3.Activate
    lastcol1_t = ws3.Cells(6, ws3.Columns.Count).End(xlToLeft).Column
    lastrow1 = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).row
    For i = lastcol1_t To 1 Step -1
            If Left(ws3.Cells(6, i).Value, 3) = "The" Then
            lastcol1 = i
            Exit For
            End If
    Next i
    ECUcol1 = ws3.Cells.Find(What:="ECU Name", After:=ActiveCell, LookIn:=xlFormulas2, _
        lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Column

    For i = 7 To lastrow1
        ws3.Cells(i, lastcol1_t + 1).Value = ws3.Cells(i, 1).Value & ws3.Cells(i, 2).Value & ws3.Cells(i, 6).Value
    Next i

    
    For i = 7 To lastrow1
       For j = 11 To ECUcol1 - 1
            If ws3.Cells(i, j).Value = "" Then
             ws3.Cells(i, lastcol1_t + 1).Value = ws3.Cells(i, lastcol1_t + 1).Value & "."
            Else
            ws3.Cells(i, lastcol1_t + 1).Value = ws3.Cells(i, lastcol1_t + 1).Value & ws3.Cells(i, j).Value
            End If
       Next j
    Next i
    For i = 7 To lastrow1
        If ws3.Cells(i, lastcol1_t + 1).Value = "" Then GoTo nexti1
        If dictFrameList_base.Exists(ws3.Cells(i, lastcol1_t + 1).Value) = False Then dictFrameList_base.Add ws3.Cells(i, lastcol1_t + 1).Value, i
nexti1:
    Next i
    lastcol1_check = lastcol1
    
    ws4.Activate
    lastcol1_t = ws4.Cells(6, ws4.Columns.Count).End(xlToLeft).Column
    lastrow2 = ws4.Cells(ws4.Rows.Count, 1).End(xlUp).row
    For i = lastcol1_t To 1 Step -1
            If Left(ws4.Cells(6, i).Value, 3) = "The" Then
            lastcol2 = i
            Exit For
            End If
    Next i
    
    If lastcol1_check <> lastcoll2 Then
        MsgBox "The number of ECU/NP is not uniform!"
        Exit Sub
    End If
    
    ECUcol2 = ws4.Cells.Find(What:="ECU Name", After:=ActiveCell, LookIn:=xlFormulas2, _
        lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Column
        
    For i = 7 To lastrow2
      ws4.Cells(i, lastcol1_t + 1).Value = ws4.Cells(i, 1).Value & ws4.Cells(i, 2).Value & ws4.Cells(i, 6).Value
    Next i

     For i = 7 To lastrow2
       For j = 11 To ECUcol2 - 1
            If ws4.Cells(i, j).Value = "" Then
             ws4.Cells(i, lastcol1_t + 1).Value = ws4.Cells(i, lastcol1_t + 1).Value & "."
            Else
             ws4.Cells(i, lastcol1_t + 1).Value = ws4.Cells(i, lastcol1_t + 1).Value & ws4.Cells(i, j).Value
            End If
       Next j
    Next i
     For i = 7 To lastrow2
        If ws4.Cells(i, lastcol1_t + 1).Value = "" Then GoTo nexti2
        If dictFrameList_comp.Exists(ws4.Cells(i, lastcol1_t + 1).Value) = False Then dictFrameList_comp.Add ws4.Cells(i, lastcol1_t + 1).Value, i
nexti2:
    Next i
    
    For Each keyy In dictFrameList_base.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 7
        End If
    Next keyy
    
    For Each keyy In dictFrameList_comp.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 7
        End If
    Next keyy
    
    ws3.Activate
    Set copyRange = ws3.Range(ws3.Cells(1, 1), ws3.Cells(6, lastcol1))
    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)

    ws4.Activate
    Set copyRange = ws4.Range(ws4.Cells(1, 1), ws4.Cells(6, lastcol2))
    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1 + 2)
    copyRange.Copy Destination:=thisSheet.Cells(1, 2 * lastcol1 + 3)

    For Each keyy In dictFrameList_gen.Keys
        If dictFrameList_base.Exists(keyy) Then
            ws3.Activate
            Set copyRange = ws3.Range(ws3.Cells(dictFrameList_base(keyy), 1), ws3.Cells(dictFrameList_base(keyy), lastcol1))
            thisSheet.Activate
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), 1)
        End If
        If dictFrameList_comp.Exists(keyy) Then
            ws4.Activate
            Set copyRange = ws4.Range(ws4.Cells(dictFrameList_comp(keyy), 1), ws4.Cells(dictFrameList_comp(keyy), lastcol2))
            thisSheet.Activate
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1 + 2)
        End If
    Next keyy
    
    lastrow3 = IIf(thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row > thisSheet.Cells(thisSheet.Rows.Count, lastcol1 + 2).End(xlUp).row, _
        thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol1 + 2).End(xlUp).row)
    
    thisSheet.Activate
    
    Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(7, lastcol1 * 2 + 3), thisSheet.Cells(lastrow3, lastcol1 * 2 + lastcol1 + 2)), lastcol1 * 2 + 2, lastcol1 + 1)
    
    Call Sumary.Summary(thisSheet, 6, lastcol1 * 3 + 4, lastrow3, lastcol1, lastcol1 + 1)
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol1 + 2), thisSheet.Cells(1, lastcol1 + 2)))
    Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 2 * lastcol1 + 3), thisSheet.Cells(1, 2 * lastcol1 + 3)))
    Call DeleteRow(thisSheet, 2, 3)
'    Call SortDraftECU
    
End Sub
'Sub SortDraftECU()
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim firstCol_base, firstCol_draft, lastCol_base, lastCol_draft, lastCol_draftnew, firstCol_Comp, lastCol_Comp As Integer  ' column index in ECU area
'    Dim firstRow, lastRow_draft, lastRow_base As Integer
'    Dim thisSheet As Worksheet
'    Dim copyRange As Range
'    Dim keyy, temp_col As Variant
'    Dim dictECUList_base As Dictionary
'    Dim dictECUList_draft As Dictionary
'    Dim dictValue_draft As Dictionary
'    Dim dictECUList_gen As Dictionary
'
'    Set dictECUList_base = New Dictionary
'    Set dictECUList_draft = New Dictionary
'    Set dictValue_draft = New Dictionary
'    Set dictECUList_gen = New Dictionary
'    Set thisSheet = ThisWorkbook.Sheets("Framelist")
'
'    firstRow = 7
'    thisSheet.Activate
'
'    For i = 10 To Columns.Count
'       If Cells(5, i).Value = "ECU Name" And Cells(5, i).MergeCells = True Then
'            firstCol_base = i
'            lastCol_base = i + Cells(5, i).MergeArea.Count - 1
'           Exit For
'        End If
'    Next i
'    For i = lastCol_base + 9 To Columns.Count
'        If Cells(5, i).Value = "ECU Name" And Cells(5, i).MergeCells = True Then
'            firstCol_draft = i
'            lastCol_draft = i + Cells(5, i).MergeArea.Count - 1
'            Exit For
'        End If
'    Next i
'
'    lastRow_base = thisSheet.Cells(thisSheet.Rows.Count, 1).End(xlUp).row
'    lastRow_draft = thisSheet.Cells(thisSheet.Rows.Count, lastCol_base + 2).End(xlUp).row
'
'    For i = firstCol_base To lastCol_base
'        If thisSheet.Cells(6, i).Value = "" Then GoTo nexti1
'        If dictECUList_base.Exists(Cells(6, i)) = False Then dictECUList_base.Add Cells(6, i).Value, i
'nexti1:
'    Next i
'    For i = firstCol_draft To lastCol_draft
'        If thisSheet.Cells(6, i).Value = "" Then GoTo nexti2
'        If dictECUList_base.Exists(Cells(6, i)) = False Then
'            dictECUList_draft.Add Cells(6, i).Value, i
'            dictValue_draft.Add i, thisSheet.Range(Cells(6, i), Cells(lastRow_draft, i)).Value
'        End If
'nexti2:
'    Next i
'
'    For Each keyy In dictECUList_base.Keys
'        If Not dictECUList_gen.Exists(keyy) Then
'            dictECUList_gen.Add keyy, dictECUList_gen.Count + firstCol_draft
'        End If
'    Next keyy
'
'    For Each keyy In dictECUList_draft.Keys
'        If Not dictECUList_gen.Exists(keyy) Then
'            dictECUList_gen.Add keyy, dictECUList_gen.Count + firstCol_draft
'        End If
'    Next keyy
'
'    'action
'    thisSheet.Range(Cells(6, firstCol_draft), Cells(lastRow_draft, lastCol_draft)).ClearContents
'
'    For Each keyy In dictECUList_gen.Keys
'        If dictECUList_draft.Exists(keyy) Then
'            temp_col = dictECUList_gen(keyy)
'            thisSheet.Range(Cells(6, temp_col), Cells(lastRow_draft, temp_col)).Value = dictValue_draft(dictECUList_draft(keyy))
'        End If
'    Next keyy
'
'    'find the new last column then format the table
'    lastCol_draftnew = dictECUList_gen.Items(dictECUList_gen.Count - 1)
'    If lastCol_draftnew > lastCol_draft Then
'
'        If thisSheet.Cells(5, firstCol_draft).MergeCells = True Then
'            thisSheet.Range(Cells(5, firstCol_draft), Cells(5, lastCol_draft)).UnMerge
'        End If
'        thisSheet.Range(Cells(5, firstCol_draft), Cells(5, lastCol_draftnew)).Merge
'        With thisSheet.Range(Cells(5, firstCol_draft), Cells(5, lastCol_draftnew)).Borders
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = 0
'        End With
'
'        Range(Cells(6, lastCol_draft), Cells(lastRow_draft, lastCol_draft)).Copy
'        Range(Cells(6, lastCol_draft + 1), Cells(6, lastCol_draft + 1)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'        Range(Cells(6, lastCol_draftnew), Cells(6, lastCol_draftnew)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'        Application.CutCopyMode = False
'    End If
'
'End Sub
'
'Sub DeleteContents(sh As Worksheet)
'    sh.Activate
'    sh.Cells.Select
'    Selection.Delete Shift:=xlUp
'End Sub
'Sub tet()
'  ThisWorkbook.Sheets("Framelist").Range(Cells(5, 86), Cells(5, 86)).UnMerge
'End Sub
'

