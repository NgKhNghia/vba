Sub copycontainer(ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, wsframe1 As Worksheet, wsframe2 As Worksheet, wsnet1 As Worksheet, wsnet2 As Worksheet, wbdict As Workbook)
    Dim lastcol1_t1 As Integer
    Dim lastcol_t2 As Integer
    Dim lastcol2 As Integer
    Dim lastcol3 As Integer
    Dim firstcol3 As Integer
    Dim lastrow1 As Integer
    Dim lastrow2 As Integer
    Dim lastRow As Integer
    Dim lastrow3 As Integer
    Dim ECUcol1 As Integer
    Dim ECUcol2 As Integer
    Dim i, j As Integer
    Dim keyFrame, gopbien As String
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
    Set dictFrameList_gen = CreateObject("Scripting.Dictionary")
    Set dictFrameList_ADAS = CreateObject("Scripting.Dictionary")
    Set thisSheet = wbdict.Sheets("Construction of Container frame")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
'-----------------
    Call NetworkPathCheck(wsnet1, wsnet2, wbdict)
    Call FrameSynthCheck(wsframe1, wsframe2, wbdict)
'------------------
    
'BASE

    ws1.Activate
    
    lastcol1_t1 = ws1.Cells(5, ws1.Columns.Count).End(xlToLeft).Column
    lastrow1 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).row
     
'--------------------
    Dim StartCell As Excel.Range
    Dim StartCell_Synth As Excel.Range
    Dim StartCell_NP As Excel.Range
    Dim CheckFN As String
    Dim CheckPDU As String
    Dim CheckResult As Excel.Range

    
    Set StartCell = ws1.UsedRange.Find("Frame Name")
        With wbdict.Worksheets("Frame Synthesis (2)")
            Set StartCell_Synth = .UsedRange.Find("Frame Name")
        End With
        
        With wbdict.Worksheets("Network Path (2)")
            Set StartCell_NP = .UsedRange.Find("Frame Name")
        End With
        
        For i = lastrow1 To 6 Step -1
            UnusedFlag = 0
            CheckFN = ws1.Cells(i, StartCell.Column).Text
            If ws1.Cells(i, StartCell.Column).Font.Color = RGB(191, 191, 191) Then  '•¶??F‚ªƒOƒŒ[‚`ê‡‚Í?ù‚É–¢?g—p‚È‚`‚ÅA”»’è•s—v

            Else
                With wbdict.Worksheets("Frame Synthesis (2)")
                    'Set StartCell_Synth = .UsedRange.Find("Frame Name")
                    Set CheckResult = .Columns(StartCell_Synth.Column).Find(CheckFN)
                End With
                If CheckResult Is Nothing Then 'FrameSynthesis‚Å–¢?g—pFrame‚ÍƒOƒŒ[ƒAƒEƒg
                    ws1.Range(ws1.Cells(i, 1), ws1.Cells(i, lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
                    GoTo Continue
                End If
                
                With wbdict.Worksheets("Network Path (2)")
                    Set CheckResult = .Columns(StartCell_NP.Column).Find(CheckFN)
                End With
                If CheckResult Is Nothing Then 'NetworkPath‚Å–¢?g—pFrame‚ÍƒOƒŒ[ƒAƒEƒg
                    ws1.Range(ws1.Cells(i, 1), ws1.Cells(i, lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
                Else
                    '?g—p‚³‚ê‚Ä‚¢‚éFrame‚É‚Â‚¢‚ÄAPDUName’PˆÊ‚ÅANetworkPatha‚Å?g—p‚³‚ê‚Ä‚¢‚é‚©ƒ`ƒFƒbƒN
                    CheckPDU = ws1.Cells(i, StartCell.Column + 9).Text
                    
                    If (CheckFN = "VDC_A116C_FD" And CheckPDU = "VDC_A115") Or (CheckFN = "VDC_A117C_FD" And CheckPDU = "VDC_A12") Then
                        ws1.Range(ws1.Cells(i, 1), ws1.Cells(i, lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
                        
                    Else
                        With wbdict.Worksheets("Network Path (2)")
                            Set CheckResult = .Columns(StartCell_NP.Column - 1).Find(CheckPDU)
                        End With
                        If CheckResult Is Nothing Then 'NetworkPatha‚Å–¢?g—p‚`PDU‚dƒOƒŒ[ƒAƒEƒg
                            ws1.Range(ws1.Cells(i, 1), ws1.Cells(i, lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
                        End If
                    End If
                End If
            End If
Continue:
        Next i
'--------------------------------------------

    Dim data1 As Variant
    For i = 6 To lastrow1
     ws1.Cells(i, lastcol1_t1 + 2).Value = ws1.Cells(i, 2).Value & ws1.Cells(i, 3).Value & ws1.Cells(i, 10).Value & ws1.Cells(i, 11).Value
    Next i
    data1 = ws1.Range(ws1.Cells(6, lastcol1_t1 + 2), ws1.Cells(lastrow1, lastcol1_t1 + 2)).Value 'ganbien
    
    For i = 1 To UBound(data1, 1)
        If data1(i, 1) <> "" Then
            If dictFrameList_base.Exists(data1(i, 1)) = False Then dictFrameList_base.Add data1(i, 1), i
        End If
    Next i
    
    
    
'OBJ
    ws2.Activate
    lastcol_t2 = ws2.Cells(5, ws2.Columns.Count).End(xlToLeft).Column
    lastrow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).row
 '--------------------------------------------
        Set StartCell = ws2.UsedRange.Find("Frame Name")
        With wbdict.Worksheets("Frame Synthesis (3)")
            Set StartCell_Synth = .UsedRange.Find("Frame Name")
        End With
        
        With wbdict.Worksheets("Network Path (3)")
            Set StartCell_NP = .UsedRange.Find("Frame Name")
        End With
        
        For i = lastrow2 To 6 Step -1
            UnusedFlag = 0
            CheckFN = ws2.Cells(i, StartCell.Column).Text
            If ws2.Cells(i, StartCell.Column).Font.Color = RGB(191, 191, 191) Then  '•¶??F‚ªƒOƒŒ[‚`ê‡‚Í?ù‚É–¢?g—p‚È‚`‚ÅA”»’è•s—v
            
            Else
                With wbdict.Worksheets("Frame Synthesis (3)")
                    'Set StartCell_Synth = .UsedRange.Find("Frame Name")
                    Set CheckResult = .Columns(StartCell_Synth.Column).Find(CheckFN)
                End With
                If CheckResult Is Nothing Then 'FrameSynthesis‚Å–¢?g—pFrame‚ÍƒOƒŒ[ƒAƒEƒg
                     ws2.Range(ws2.Cells(i, 1), ws2.Cells(i, lastcol_t2)).Interior.Color = RGB(191, 191, 191)
                    GoTo Continue2
                End If
                
                With wbdict.Worksheets("Network Path (3)")
                    Set CheckResult = .Columns(StartCell_NP.Column).Find(CheckFN)
                End With
                If CheckResult Is Nothing Then 'NetworkPath‚Å–¢?g—pFrame‚ÍƒOƒŒ[ƒAƒEƒg
                     ws2.Range(ws2.Cells(i, 1), ws2.Cells(i, lastcol_t2)).Interior.Color = RGB(191, 191, 191)
                Else
                    '?g—p‚³‚ê‚Ä‚¢‚éFrame‚É‚Â‚¢‚ÄAPDUName’PˆÊ‚ÅANetworkPatha‚Å?g—p‚³‚ê‚Ä‚¢‚é‚©ƒ`ƒFƒbƒN
                    CheckPDU = ws2.Cells(i, StartCell.Column + 9).Text
                    
                    If (CheckFN = "VDC_A116C_FD" And CheckPDU = "VDC_A115") Or (CheckFN = "VDC_A117C_FD" And CheckPDU = "VDC_A12") Then
                        ws2.Range(ws2.Cells(i, 1), ws2.Cells(i, lastcol_t2)).Interior.Color = RGB(191, 191, 191)
                        
                    Else
                        With wbdict.Worksheets("Network Path (3)")
                            Set CheckResult = .Columns(StartCell_NP.Column - 1).Find(CheckPDU)
                        End With
                        If CheckResult Is Nothing Then 'NetworkPatha‚Å–¢?g—p‚`PDU‚dƒOƒŒ[ƒAƒEƒg
                            ws2.Range(ws2.Cells(i, 1), ws2.Cells(i, lastcol_t2)).Interior.Color = RGB(191, 191, 191)
                        End If
                    End If
                End If
            End If
Continue2:
        Next i
    
    Dim data2 As Variant
    
    For i = 6 To lastrow2
      ws2.Cells(i, lastcol_t2 + 2).Value = ws2.Cells(i, 2).Value & ws2.Cells(i, 3).Value & ws2.Cells(i, 10).Value & ws2.Cells(i, 11).Value
    Next i
    
    data2 = ws2.Range(ws2.Cells(6, lastcol_t2 + 2), ws2.Cells(lastrow2, lastcol_t2 + 2)).Value

    ' Thêm các khung vào dictFrameList_comp
    For i = 1 To UBound(data2, 1)
        If data2(i, 1) <> "" Then
            If dictFrameList_comp.Exists(data2(i, 1)) = False Then dictFrameList_comp.Add data2(i, 1), i
        End If
    Next i
    
'-----------------------------
'Insert to dic
    For Each keyy In dictFrameList_base.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 6
        End If
    Next keyy
    
    For Each keyy In dictFrameList_comp.Keys
        If Not dictFrameList_gen.Exists(keyy) Then
            dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 6
        End If
    Next keyy

'---------------------------------------
'ADASMsg
    If Sheet2.ADASmsg.Value Then
        ws3.Activate
        
' ep sang kieu string de so sanh
'        thisSheet.Cells.NumberFormat = "@"
        
        Dim lastrow4 As Integer
        Dim lastcol4 As Integer
        
        lastcol4 = ws3.Cells(4, ws3.Columns.Count).End(xlToLeft).Column
        lastrow4 = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).row
        
        Dim data3 As Variant
        For i = 5 To lastrow4
          ws3.Cells(i, lastcol4 + 2).Value = ws3.Cells(i, 1).Value & ws3.Cells(i, 2).Value & ws3.Cells(i, 9).Value & ws3.Cells(i, 10).Value
        Next i
        
        data3 = ws3.Range(ws3.Cells(5, lastcol4 + 2), ws3.Cells(lastrow4, lastcol4 + 2)).Value
        For i = 1 To UBound(data3, 1)
            If data3(i, 1) <> "" Then
                If dictFrameList_ADAS.Exists(data3(i, 1)) = False Then dictFrameList_ADAS.Add data3(i, 1), i
            End If
        Next i
    
         For Each keyy In dictFrameList_ADAS.Keys
            If Not dictFrameList_gen.Exists(keyy) Then
                dictFrameList_gen.Add keyy, dictFrameList_gen.Count + 6
            End If
        Next keyy
        
    End If

'-----------------------------
'Copy title
'planning sheet 1
    ws1.Activate
    Set copyRange = ws1.Range(ws1.Cells(1, 1), ws1.Cells(5, lastcol1_t1))
    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, 1)
'planning sheet 2
    ws2.Activate
    Set copyRange = ws2.Range(ws2.Cells(1, 1), ws2.Cells(5, lastcol_t2))
    thisSheet.Activate
    copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + 2)

   If Sheet2.ADASmsg.Value Then
   'ADASMsg
        ws2.Activate
        Set copyRange = ws2.Range(ws2.Cells(1, 2), ws2.Cells(5, lastcol_t2))
        thisSheet.Activate
        copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + lastcol_t2 + 3)
        
    'planning sheet 1 & planning sheet 2
        Set copyRange = thisSheet.Range(thisSheet.Cells(4, 1), thisSheet.Cells(5, lastcol1_t1))
        copyRange.Copy Destination:=thisSheet.Cells(4, 3 * lastcol1_t1 + 3)
        
    'planning sheet 2 & ADASmsg
        Set copyRange = thisSheet.Range(thisSheet.Cells(4, 2), thisSheet.Cells(5, lastcol1_t1))
        copyRange.Copy Destination:=thisSheet.Cells(4, 4 * lastcol1_t1 + 4)
    Else
        'planning sheet 1 & planning sheet 2
        Set copyRange = ws2.Range(ws2.Cells(1, 1), ws2.Cells(5, lastcol_t2))
        thisSheet.Activate
        copyRange.Copy Destination:=thisSheet.Cells(1, lastcol1_t1 + lastcol_t2 + 3)
    End If

'Copy data
    For Each keyy In dictFrameList_gen.Keys
        If dictFrameList_base.Exists(keyy) Then
            Set copyRange = ws1.Range(ws1.Cells(dictFrameList_base(keyy) + 5, 1), ws1.Cells(dictFrameList_base(keyy) + 5, lastcol1_t1))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), 1)
        End If
        
        If dictFrameList_comp.Exists(keyy) Then
            Set copyRange = ws2.Range(ws2.Cells(dictFrameList_comp(keyy) + 5, 1), ws2.Cells(dictFrameList_comp(keyy) + 5, lastcol_t2))
            copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), lastcol1_t1 + 2)
        End If
        
        If Sheet2.ADASmsg.Value Then
            If dictFrameList_ADAS.Exists(keyy) Then
                Set copyRange = ws3.Range(ws3.Cells(dictFrameList_ADAS(keyy) + 4, 1), ws3.Cells(dictFrameList_ADAS(keyy) + 4, lastcol_t2))
                copyRange.Copy Destination:=thisSheet.Cells(dictFrameList_gen(keyy), 2 * lastcol1_t1 + 3)
            End If
         End If
    Next keyy
    
    thisSheet.Activate

    'update lastRow
    lastRow = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row > thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 3).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, 2).End(xlUp).row, thisSheet.Cells(thisSheet.Rows.Count, lastcol1_t1 + 3).End(xlUp).row)
    
    If Sheet2.ADASmsg.Value Then
        lastRow = IIf(thisSheet.Cells(thisSheet.Rows.Count, 2 * lastcol1_t1 + 3).End(xlUp).row > lastRow, thisSheet.Cells(thisSheet.Rows.Count, 2 * lastcol1_t1 + 3).End(xlUp).row, lastRow)
    End If
'--------------------
'Fill gray base
    With thisSheet.Range(thisSheet.Cells(5, 1), thisSheet.Cells(lastRow, lastcol1_t1))
        .AutoFilter
        .Rows(6).AutoFilter Field:=2, Criteria1:="="
    End With

    If thisSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
         For Each cell In thisSheet.AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible)
             If cell.Value = "" Then
                thisSheet.Range(thisSheet.Cells(cell.row, 1), thisSheet.Cells(cell.row, lastcol1_t1)).Interior.Color = RGB(191, 191, 191)
             End If
         Next cell
    End If
     
    thisSheet.AutoFilterMode = False

    
'Fill gray obj

    With thisSheet.Range(thisSheet.Cells(5, 1), thisSheet.Cells(lastRow, 2 * lastcol1_t1 + 1))
        .AutoFilter
        .Rows(6).AutoFilter Field:=lastcol1_t1 + 3, Criteria1:="="
    End With

    If thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + 2).SpecialCells(xlCellTypeVisible).Count > 1 Then
         For Each cell In thisSheet.AutoFilter.Range.Columns(lastcol1_t1 + 3).SpecialCells(xlCellTypeVisible)
             If cell.Value = "" Then
                thisSheet.Range(thisSheet.Cells(cell.row, lastcol1_t1 + 2), thisSheet.Cells(cell.row, 2 * lastcol1_t1 + 1)).Interior.Color = RGB(191, 191, 191)
             End If
         Next cell
    End If
    
    thisSheet.AutoFilterMode = False
    
    
'Fill gray messageList
    If Sheet2.ADASmsg.Value Then
        With thisSheet.Range(thisSheet.Cells(5, 2 * lastcol1_t1 + 3), thisSheet.Cells(lastRow, 3 * lastcol1_t1 + 1))
            .AutoFilter
            .Rows(6).AutoFilter Field:=2, Criteria1:="="
        End With
    
        If thisSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
             For Each cell In thisSheet.AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible)
                 If cell.Value = "" Then
                    thisSheet.Range(thisSheet.Cells(cell.row, 2 * lastcol1_t1 + 3), thisSheet.Cells(cell.row, 3 * lastcol1_t1 + 1)).Interior.Color = RGB(191, 191, 191)
                 End If
             Next cell
        End If
    End If
     
    thisSheet.AutoFilterMode = False






'-----------------------------------------------------Compare-----------------------
    If Sheet2.ADASmsg.Value Then
        ' tao bang tong ket keikakusho va messagelist
        
        For Each cell In thisSheet.UsedRange
        If Not IsEmpty(cell.Value) Then
            cell.Value = CStr(cell.Value)
        End If
        Next cell

        thisSheet.Cells(5, 5 * lastcol1_t1 + 4).Value = "Œv‰æ‘‚Ì·•ª"
        thisSheet.Cells(5, 5 * lastcol1_t1 + 6).Value = "ˆê’v/•sˆê’v"
        thisSheet.Cells(5, 5 * lastcol1_t1 + 7).Value = "”»’è"
        thisSheet.Cells(5, 5 * lastcol1_t1 + 8).Value = "·•ª“à—e"
        thisSheet.Cells(5, 5 * lastcol1_t1 + 9).Value = "Œ©‰ðE”õl"
        thisSheet.Cells(5, 5 * lastcol1_t1 + 10).Value = "•â‘«î•ñ"
        thisSheet.Cells(5, 5 * lastcol1_t1 + 11).Value = "Tag"
        With thisSheet.Range(thisSheet.Cells(5, 5 * lastcol1_t1 + 4), thisSheet.Cells(5, 5 * lastcol1_t1 + 4))
            .Interior.Color = RGB(0, 255, 0)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
    
        With thisSheet.Range(thisSheet.Cells(5, 5 * lastcol1_t1 + 6), thisSheet.Cells(5, 5 * lastcol1_t1 + 11))
            .Interior.Color = RGB(0, 255, 0)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(6, 3 * lastcol1_t1 + 3), thisSheet.Cells(lastRow, 4 * lastcol1_t1 + 2)), 3 * lastcol1_t1 + 2, 2 * lastcol1_t1 + 1)
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(6, 4 * lastcol1_t1 + 4), thisSheet.Cells(lastRow, 5 * lastcol1_t1 + 2)), 3 * lastcol1_t1 + 1, 2 * lastcol1_t1 + 1)
        Call Summary(thisSheet, 5, lastcol1_t1 * 5 + 6, lastRow, lastcol1_t1 - 1, lastcol1_t1 + 2)
        Call Sumary2(thisSheet, thisSheet.Range(thisSheet.Cells(6, lastcol1_t1 * 3 + 3), thisSheet.Cells(lastRow, lastcol1_t1 * 4 + 2)), thisSheet.Range(thisSheet.Cells(6, lastcol1_t1 * 5 + 4), thisSheet.Cells(lastRow, lastcol1_t1 * 5 + 4)), lastRow)
    Else
        Call compare3(thisSheet, thisSheet.Range(thisSheet.Cells(6, 2 * lastcol1_t1 + 3), thisSheet.Cells(lastRow, 3 * lastcol1_t1 + 2)), 2 * lastcol1_t1 + 2, lastcol1_t1 + 1)
        Call Summary(thisSheet, 5, lastcol1_t1 * 3 + 4, lastRow, lastcol1_t1, lastcol1_t1 + 1)
            thisSheet.Cells(5, 3 * lastcol1_t1 + 4).Value = "Œv‰æ‘‚Ì·•ª"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 5).Value = "”»’è"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 6).Value = "·•ª“à—e"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 7).Value = "Œ©‰ðE”õl"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 8).Value = "•â‘«î•ñ"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 9).Value = "Tag"
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 5), thisSheet.Cells(4, 3 * lastcol1_t1 + 9)).Merge
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 5), thisSheet.Cells(4, 3 * lastcol1_t1 + 9)).Value = "Œv‰æ‘‚ÆŒv‰æ‘”äŠr"
            
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 11), thisSheet.Cells(4, 3 * lastcol1_t1 + 13)).Merge
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 11), thisSheet.Cells(4, 3 * lastcol1_t1 + 13)).Value = "‘O‰ñFB(ADAS5)"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 11).Value = "FB"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 12).Value = "FB“à—e"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 13).Value = "‘Î‰žó‹µ"
            
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 15), thisSheet.Cells(4, 3 * lastcol1_t1 + 17)).Merge
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 15), thisSheet.Cells(4, 3 * lastcol1_t1 + 17)).Value = "‘O‰ñFB(ADCU)"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 15).Value = "FB"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 16).Value = "FB“à—e"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 17).Value = "‘Î‰žó‹µ"
            
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 19), thisSheet.Cells(4, 3 * lastcol1_t1 + 20)).Merge
            thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 19), thisSheet.Cells(4, 3 * lastcol1_t1 + 20)).Value = "¡‰ñ‚Ì”»’f"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 19).Value = "‘ÎÛ•ñ"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 20).Value = "Œ‹˜_"
            thisSheet.Cells(5, 3 * lastcol1_t1 + 21).Value = "FB“à—e"
            
        
        With thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 4), thisSheet.Cells(4, 3 * lastcol1_t1 + 21))
            .Interior.Color = RGB(0, 255, 0)
        End With
        
        With thisSheet.Range(thisSheet.Cells(4, 3 * lastcol1_t1 + 4), thisSheet.Cells(lastRow, 3 * lastcol1_t1 + 21))
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
        
    End If

    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox1.Value), Range(thisSheet.Cells(1, 1), thisSheet.Cells(1, 1)))
    Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox2.Value), Range(thisSheet.Cells(1, lastcol1_t1 + 2), thisSheet.Cells(1, lastcol1_t1 + 2)))
    If Sheet2.ADASmsg.Value Then
        Call SetNameInFirstCell(thisSheet, GetFileName(Sheet2.TextBox3.Value), Range(thisSheet.Cells(1, 2 * lastcol1_t1 + 3), thisSheet.Cells(1, 2 * lastcol1_t1 + 3)))
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 3 * lastcol1_t1 + 4), thisSheet.Cells(1, 3 * lastcol1_t1 + 4)))
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆADASMsgList”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 4 * lastcol1_t1 + 5), thisSheet.Cells(1, 4 * lastcol1_t1 + 5)))
    Else
        Call SetNameInFirstCell(thisSheet, "Œv‰æ‘‚ÆŒv‰æ‘”äŠrŒ‹‰Ê", Range(thisSheet.Cells(1, 2 * lastcol1_t1 + 3), thisSheet.Cells(1, 2 * lastcol1_t1 + 3)))
    End If
    Call DeleteRow(thisSheet, 2, 3)
    
    ' B?t l?i các tính nang
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub



