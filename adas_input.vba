Public title_Base As Scripting.Dictionary
Public title_draft As Scripting.Dictionary
Public title_Msg As Scripting.Dictionary

Dim keywordColumn_Base As Integer
Public lastRow_Base As Integer

Dim keywordColumn_Draft As Integer
Public lastRow_Draft As Integer

Dim keywordColumn_Msg As Integer
Public lastRow_Msg As Integer

Public firstDataRow_Out As Integer

Dim tmp1 As Range
Dim tmp2 As Range
Dim cell As Range

Dim ws As Worksheet

Function SheetExists(wb As Workbook, sheetName As String) As Boolean
' kiem tra su ton tai cua cac sheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing And ws.name = sheetName
    On Error GoTo 0
End Function

Sub Refresh()
' bo loc, bo an
    wsBase.AutoFilterMode = False
    wsDraft.AutoFilterMode = False
    If flag_Msg Then wsMsg.AutoFilterMode = False

    wsBase.Columns.Hidden = False
    wsDraft.Columns.Hidden = False
    If flag_Msg Then wsMsg.Columns.Hidden = False
    
    wsBase.Rows.Hidden = False
    wsDraft.Rows.Hidden = False
    If flag_Msg Then wsMsg.Rows.Hidden = False
End Sub

Sub DeleteEmptyRow()
' dam bao du lieu bat dau tu dong 4 -> dong nhat thi de lam
    If InStr(wsBase.name, "Frame Synthesis") Then
        wsBase.Rows("1:4").Delete
        wsDraft.Rows("1:4").Delete
        If flag_Msg Then wsMsg.Rows("1:2").Insert                   ' cai nay thi phai chen them dong
        
    ElseIf InStr(wsBase.name, "Construction") Then
        wsBase.Rows("1:2").Delete
        wsDraft.Rows("1:2").Delete
        If flag_Msg Then wsMsg.Rows("1").Delete
        
    ElseIf InStr(wsBase.name, "Network") Then
        wsBase.Rows("1").Delete
        wsDraft.Rows("1").Delete
        If flag_Msg Then wsMsg.Rows("1:2").Insert                   ' cai nay thi phai chen them dong
    End If
End Sub

Sub ConfirmTitle()
'' xoa bo nhung title thua trong moi sheet
'' do moi sheet co title can xoa khac nhau nen khong the gop code
'    Dim rngDelete As Range
'
'    If wsBase.name = "Frame Synthesis" Then
'        Set rngDelete = wsBase.Columns(10000)
'        For Each cell In Range(wsBase.Cells(3, 11), wsBase.Cells(3, 11).End(xlToRight))
'            If cell.Value <> "ADAS" And cell.Value <> "ADAS_Bridge" Then Set rngDelete = Union(rngDelete, cell.EntireColumn)
'        Next cell
'        rngDelete.Delete
'
'        Set rngDelete = wsDraft.Columns(10000)
'        For Each cell In Range(wsDraft.Cells(3, 11), wsDraft.Cells(3, 11).End(xlToRight))
'            If cell.Value <> "ADAS" And cell.Value <> "ADAS_Bridge" Then Set rngDelete = Union(rngDelete, cell.EntireColumn)
'        Next cell
'        rngDelete.Delete
'
'    ElseIf wsBase.name = "Network Path" Then
'        Set rngDelete = wsBase.Columns(10000)
'        For Each cell In Range(wsBase.Cells(3, 6), wsBase.Cells(3, 6).End(xlToRight).Offset(0, -1))
'            If cell.Value <> "CH2-CAN" And cell.Value <> "CH3-CAN" And cell.Value <> "ITS1-FD" And cell.Value <> "ITS2-FD" And _
'                cell.Value <> "ITS3-FD" And cell.Value <> "ITS4-FD" And cell.Value <> "ITS5-FD" Then
'                Set rngDelete = Union(rngDelete, cell.EntireColumn)
'            End If
'        Next cell
'        rngDelete.Delete
'        If wsBase.Cells(3, 7).Value <> "CH3-CAN" Then
'            wsBase.Columns(7).Insert shift:=xlToRight
'            wsBase.Cells(3, 7).Value = "CH3-CAN"
'        End If
'
'        Set rngDelete = wsDraft.Columns(10000)
'        For Each cell In Range(wsDraft.Cells(3, 6), wsDraft.Cells(3, 6).End(xlToRight).Offset(0, -1))
'            If cell.Value <> "CH2-CAN" And cell.Value <> "CH3-CAN" And cell.Value <> "ITS1-FD" And cell.Value <> "ITS2-FD" And _
'                cell.Value <> "ITS3-FD" And cell.Value <> "ITS4-FD" And cell.Value <> "ITS5-FD" Then
'                Set rngDelete = Union(rngDelete, cell.EntireColumn)
'            End If
'        Next cell
'        rngDelete.Delete
'        If wsDraft.Cells(3, 7).Value <> "CH3-CAN" Then
'            wsDraft.Columns(7).Insert shift:=xlToRight
'            wsDraft.Cells(3, 7).Value = "CH3-CAN"
'        End If
'
'        If flag_Msg Then
'            If wsMsg.Cells(3, 6).Value <> "CH3-CAN" Then
'                wsMsg.Columns(6).Insert shift:=xlToRight
'                wsMsg.Cells(3, 6).Value = "CH3-CAN"
'            End If
'        End If
'    End If
' =========================================================================

' xac nhan so luong ecu giua base voi draft, thieu thi chen them cot xam
' xac nhan so luong kenh truyen giua base voi draft, thieu thi chem them cot xam
'    Dim base_dict As Scripting.Dictionary           ' chi can key;  value khong quan trong
'    Dim draft_dict As Scripting.Dictionary          ' chi can key;  value khong quan trong
    
    If InStr(wsBase.name, "Frame Synthesis") Then
'        ' lay thong tin toan bo ecu
'        For Each cell In Range(wsBase.Cells(3, 11), wsBase.Cells(3, 11).End(xlToRight))
'            base_dict.Add cell.Value, cell.Column
'        Next cell
'
'        For Each cell In Range(wsDraft.Cells(3, 11), wsDraft.Cells(3, 11).End(xlToRight))
'            draft_dict.Add cell.Value, cell.Column
'        Next cell
'
'        ' kiem tra so luong ecu
'        ' kiem tra ecu cua base co trong draft khong?
'        For Each cell In Range(wsBase.Cells(3, 11), wsBase.Cells(3, 11).End(xlToRight))
'            if draft_dict.Exists(cell.Value)
'        Next cell


        wsBase.Cells(2, 11).Formula = "=MATCH(R[1]C,'[" & wbDraft.name & "]Frame Synthesis'!R3C1:R3C" & _
            wsBase.Cells(3, 11).End(xlToRight).Column & ",0)"
        Range(wsBase.Cells(2, 11), wsBase.Cells(3, 11).End(xlToRight)).FillRight
        wsBase.Columns("11:" & wsBase.Cells(3, 11).End(xlToRight).Column).Sort _
            Key1:=wsBase.Cells(2, 11), _
            Order1:=xlAscending, _
            Header:=xlNo, _
            Orientation:=xlLeftToRight
        For i = 11 To wsBase.Cells(3, 11).End(xlToRight).Column
            If wsBase.Cells(2, 0 + i).Value > i Then
                
            ElseIf wsBase.Cells(2, 0 + i).Value > i Then
                
            End If
        Next i
        
        
        
    ElseIf InStr(wsBase.name, "Network") Then
        
    End If
End Sub

' ham nay ko can thiet -> co the bo
Sub LoadTitle()
' do sheet Construction co 2 dong tieu de nen can xu ly dac biet
    Set title_Base = New Scripting.Dictionary
    Set title_draft = New Scripting.Dictionary
    Set title_Msg = New Scripting.Dictionary
    
    If InStr(wsBase.name, "Construction ") > 0 Then
        wsBase.Cells(2, 1).UnMerge
        wsBase.Cells(3, 1).Value = wsBase.Cells(2, 1).Value
        wsDraft.Cells(2, 1).UnMerge
        wsDraft.Cells(3, 1).Value = wsDraft.Cells(2, 1).Value
    End If
          
    For Each cell In Range(wsBase.Cells(3, 1), wsBase.Cells(3, 1).End(xlToRight))
        title_Base.Add cell.Value, cell.Column
    Next cell
    
    For Each cell In Range(wsDraft.Cells(3, 1), wsDraft.Cells(3, 1).End(xlToRight))
        title_draft.Add cell.Value, cell.Column
    Next cell
    
    If flag_Msg Then
        For Each cell In Range(wsMsg.Cells(3, 1), wsMsg.Cells(3, 1).End(xlToRight))
            title_Msg.Add cell.Value, cell.Column
        Next cell
    End If
        
    If InStr(wsBase.name, "Construction ") > 0 Then
        wsBase.Cells(3, 1).Value = ""
        wsDraft.Cells(3, 1).Value = ""
    End If
End Sub

Sub CreatKeyword()
    lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    If flag_Msg Then lastRow_Msg = wsMsg.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    keywordColumn_Base = wsBase.Cells(3, 2).End(xlToRight).Column + 2
    keywordColumn_Draft = wsDraft.Cells(3, 2).End(xlToRight).Column + 2
    If flag_Msg Then keywordColumn_Msg = wsMsg.Cells(3, 2).End(xlToRight).Column + 2

    If InStr(wsBase.name, "Frame Synthesis") Then
        wsBase.Cells(4, keywordColumn_Base).Formula2 = "=RC[-12]"
        wsDraft.Cells(4, keywordColumn_Base).Formula2 = "=RC[-12]"
        If flag_Msg Then wsMsg.Cells(4, keywordColumn_Msg).Formula2 = "=RC[-12]"
    ElseIf InStr(wsBase.name, "Construction") Then
        wsBase.Cells(4, keywordColumn_Base).Formula2 = "=RC[-16]&RC[-7]"
        wsDraft.Cells(4, keywordColumn_Base).Formula2 = "=RC[-16]&RC[-7]"
        If flag_Msg Then wsMsg.Cells(4, keywordColumn_Msg).Formula2 = "=RC[-16]&RC[-7]"
    ElseIf InStr(wsBase.name, "Network") Then
        wsBase.Cells(4, keywordColumn_Base).Formula2 = "=RC[-13]&RC[-12]&RC[-2]"
        wsDraft.Cells(4, keywordColumn_Base).Formula2 = "=RC[-13]&RC[-12]&RC[-2]"
        If flag_Msg Then wsMsg.Cells(4, keywordColumn_Msg).Formula2 = "=RC[-13]&RC[-12]&RC[-2]"
    End If
    
    Range(wsBase.Cells(4, keywordColumn_Base), wsBase.Cells(lastRow_Base, keywordColumn_Base)).FillDown
    Range(wsBase.Cells(4, keywordColumn_Base), wsBase.Cells(lastRow_Base, keywordColumn_Base)).Copy
    Range(wsBase.Cells(4, keywordColumn_Base), wsBase.Cells(lastRow_Base, keywordColumn_Base)).PasteSpecial xlPasteValues
    
    Range(wsDraft.Cells(4, keywordColumn_Draft), wsDraft.Cells(lastRow_Draft, keywordColumn_Draft)).FillDown
    Range(wsDraft.Cells(4, keywordColumn_Draft), wsDraft.Cells(lastRow_Draft, keywordColumn_Draft)).Copy
    Range(wsDraft.Cells(4, keywordColumn_Draft), wsDraft.Cells(lastRow_Draft, keywordColumn_Draft)).PasteSpecial xlPasteValues
    
    If flag_Msg Then
        Range(wsMsg.Cells(4, keywordColumn_Msg), wsMsg.Cells(lastRow_Msg, keywordColumn_Msg)).FillDown
        Range(wsMsg.Cells(4, keywordColumn_Msg), wsMsg.Cells(lastRow_Msg, keywordColumn_Msg)).Copy
        Range(wsMsg.Cells(4, keywordColumn_Msg), wsMsg.Cells(lastRow_Msg, keywordColumn_Msg)).PasteSpecial xlPasteValues
    End If
End Sub

Sub DeleteData()
    ' xoa data hang 2 frame, network
    If InStr(wsBase.name, "Construction") <= 0 Then
        wsBase.Rows("1:2").ClearContents
        wsDraft.Rows("1:2").ClearContents
    Else
        wsBase.Rows("1").ClearContents
        wsDraft.Rows("1").ClearContents
    End If

    ' backup data
    wsBase.Cells(1, 21).Value = "backup"
    lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    Range(wsBase.Cells(2, 1), wsBase.Cells(lastRow_Base, title_Base.Count)).Copy wsBase.Cells(2, 21)
    wsDraft.Cells(1, 21).Value = "backup"
    lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    Range(wsDraft.Cells(2, 1), wsDraft.Cells(lastRow_Draft, title_draft.Count)).Copy wsDraft.Cells(2, 21)
    
    '
    If InStr(wsBase.name, "Frame Synthesis") Then
        ' loc cot adad va adas_bridge khong co du lieu
        wsBase.AutoFilterMode = False
        lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        With Range(wsBase.Cells(3, 1), wsBase.Cells(lastRow_Base, title_Base.Count))
            .AutoFilter Field:=title_Base.Item("ADAS"), Criteria1:="="
            .AutoFilter Field:=title_Base.Item("ADAS_Bridge"), Criteria1:="="
        End With
        Range(wsBase.Cells(4, 1), wsBase.Cells(lastRow_Base, title_Base.Count)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        wsBase.AutoFilterMode = False
        lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        
        wsDraft.AutoFilterMode = False
        lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        With Range(wsDraft.Cells(3, 1), wsDraft.Cells(lastRow_Draft, title_draft.Count))
            .AutoFilter Field:=title_draft.Item("ADAS"), Criteria1:="="
            .AutoFilter Field:=title_draft.Item("ADAS_Bridge"), Criteria1:="="
        End With
        Range(wsDraft.Cells(4, 1), wsDraft.Cells(lastRow_Draft, title_draft.Count)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        wsDraft.AutoFilterMode = False
        lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        
        ' xoa cot adas va adas_bridge bi gach
        For Each cell In Range(wsBase.Cells(4, 11), wsBase.Cells(lastRow_Base, 12))
            For i = 1 To Len(cell.Value)
                If cell.Value = "" Then Exit For
                If cell.Characters(i, 1).Font.Strikethrough Then
                    cell.Characters(i, 1).Text = ""
                    i = i - 1
                End If
            Next i
        Next cell
        
        For Each cell In Range(wsDraft.Cells(4, 11), wsDraft.Cells(lastRow_Draft, 12))
            For i = 1 To Len(cell.Value)
                If cell.Value = "" Then Exit For
                If cell.Characters(i, 1).Font.Strikethrough Then
                    cell.Characters(i, 1).Text = ""
                    i = i - 1
                End If
            Next i
        Next cell
        
    ElseIf InStr(wsBase.name, "Construction") Then
        ' loc nhung frame ben sheet frame synthesis va PDU ma adas khong dung, boi xam
        lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        wsBase.AutoFilterMode = False
        wsBase.Cells(4, 17).Formula2 = "=XLOOKUP(RC[-15],'Frame Synthesis'!C2,'Frame Synthesis'!C2,0)"
        wsBase.Cells(4, 19).Formula2 = "=XLOOKUP(RC[-8],'Network Path'!C2,'Network Path'!C2,0)"
        Range(wsBase.Cells(4, 17), wsBase.Cells(lastRow_Base, 17)).FillDown
        Range(wsBase.Cells(4, 19), wsBase.Cells(lastRow_Base, 19)).FillDown
        Range(wsBase.Cells(3, 17), wsBase.Cells(lastRow_Base, 17)).AutoFilter Field:=1, Criteria1:=0
        Range(wsBase.Cells(4, 1), wsBase.Cells(lastRow_Base, 16)).SpecialCells(xlCellTypeVisible).EntireRow.Interior.Color = RGB(128, 128, 128)
        wsBase.AutoFilterMode = False
        Range(wsBase.Cells(3, 19), wsBase.Cells(lastRow_Base, 19)).AutoFilter Field:=1, Criteria1:=0
        Range(wsBase.Cells(4, 1), wsBase.Cells(lastRow_Base, 16)).SpecialCells(xlCellTypeVisible).EntireRow.Interior.Color = RGB(128, 128, 128)
        wsBase.AutoFilterMode = False
        wsBase.Columns(17).ClearContents
        wsBase.Columns(19).ClearContents
        ' xoa cac chu bi boi xam trong cot 5 va cot 16
        For i = 4 To lastRow_Base
            If wsBase.Cells(i, 1).Value <> "N/A" Then
                If wsBase.Cells(i, 5).Value <> "" Then
                    For j = 1 To Len(wsBase.Cells(i, 5).Value)
                        If wsBase.Cells(i, 5).Value = "" Then
                            Exit For
                        End If
                        If wsBase.Cells(i, 5).Characters(j, 1).Font.Strikethrough Then
                            wsBase.Cells(i, 5).Characters(j, 1).Text = ""
                            j = j - 1
                        End If
                    Next j
                End If
                
                If wsBase.Cells(i, 16).Value <> "" Then
                    For j = 1 To Len(wsBase.Cells(i, 16).Value)
                        If wsBase.Cells(i, 16).Value = "" Then
                            Exit For
                        End If
                        If wsBase.Cells(i, 16).Characters(j, 1).Font.Strikethrough Then
                            wsBase.Cells(i, 16).Characters(j, 1).Text = ""
                            j = j - 1
                        End If
                    Next j
                End If
            End If
        Next i
        
        ' loc nhung frame ben sheet frame synthesis va PDU ma adas khong dung
        lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        wsDraft.AutoFilterMode = False
        wsDraft.Cells(4, 17).Formula2 = "=XLOOKUP(RC[-15],'Frame Synthesis'!C2,'Frame Synthesis'!C2,0)"
        wsDraft.Cells(4, 19).Formula2 = "=XLOOKUP(RC[-8],'Network Path'!C2,'Network Path'!C2,0)"
        Range(wsDraft.Cells(4, 17), wsDraft.Cells(lastRow_Draft, 17)).FillDown
        Range(wsDraft.Cells(4, 19), wsDraft.Cells(lastRow_Draft, 19)).FillDown
        Range(wsDraft.Cells(3, 17), wsDraft.Cells(lastRow_Draft, 17)).AutoFilter Field:=1, Criteria1:=0
        Range(wsDraft.Cells(4, 1), wsDraft.Cells(lastRow_Draft, 16)).SpecialCells(xlCellTypeVisible).EntireRow.Interior.Color = RGB(128, 128, 128)
        wsDraft.AutoFilterMode = False
        Range(wsDraft.Cells(3, 19), wsDraft.Cells(lastRow_Draft, 19)).AutoFilter Field:=1, Criteria1:=0
        Range(wsDraft.Cells(4, 1), wsDraft.Cells(lastRow_Draft, 16)).SpecialCells(xlCellTypeVisible).EntireRow.Interior.Color = RGB(128, 128, 128)
        wsDraft.AutoFilterMode = False
        wsDraft.Columns(17).ClearContents
        wsDraft.Columns(19).ClearContents
        ' xoa cac chu bi boi xam trong cot 5 va cot 16
        For i = 4 To lastRow_Draft
            If wsDraft.Cells(i, 1).Value <> "N/A" Then
                For j = 1 To Len(wsDraft.Cells(i, 5).Value)
                    If wsDraft.Cells(i, 5).Value = "" Then
                        Exit For
                    End If
                    If wsDraft.Cells(i, 5).Characters(j, 1).Font.Strikethrough Then
                        wsDraft.Cells(i, 5).Characters(j, 1).Text = ""
                        j = j - 1
                    End If
                Next j
                For j = 1 To Len(wsDraft.Cells(i, 16).Value)
                    If wsDraft.Cells(i, 16).Value = "" Then
                        Exit For
                    End If
                    If wsDraft.Cells(i, 16).Characters(j, 1).Font.Strikethrough Then
                        wsDraft.Cells(i, 16).Characters(j, 1).Text = ""
                        j = j - 1
                    End If
                Next j
            End If
        Next i
        
    ElseIf InStr(wsBase.name, "Network Path") Then
        ' xoa cac bus khong co adas
        wsBase.AutoFilterMode = False
        lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        Range(wsBase.Cells(3, 13), wsBase.Cells(lastRow_Base, 13)).AutoFilter Field:=1, Criteria1:="<>*adas*"
        On Error Resume Next
        Range(wsBase.Cells(4, 13), wsBase.Cells(lastRow_Base, 13)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        wsBase.AutoFilterMode = False
        
        wsDraft.AutoFilterMode = False
        lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        Range(wsDraft.Cells(3, 13), wsDraft.Cells(lastRow_Draft, 13)).AutoFilter Field:=1, Criteria1:="<>*adas*"
        On Error Resume Next
        Range(wsDraft.Cells(4, 13), wsDraft.Cells(lastRow_Draft, 13)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        wsDraft.AutoFilterMode = False
        
        ' xoa bus co FrCamADAS
        wsBase.AutoFilterMode = False
        lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        Range(wsBase.Cells(3, 13), wsBase.Cells(lastRow_Base, 13)).AutoFilter Field:=1, Criteria1:="*frcamadas*"
        On Error Resume Next
        Range(wsBase.Cells(4, 13), wsBase.Cells(lastRow_Base, 13)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        wsBase.AutoFilterMode = False
        
        wsDraft.AutoFilterMode = False
        lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        Range(wsDraft.Cells(3, 13), wsDraft.Cells(lastRow_Draft, 13)).AutoFilter Field:=1, Criteria1:="*frcamadas*"
        On Error Resume Next
        Range(wsDraft.Cells(4, 13), wsDraft.Cells(lastRow_Draft, 13)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        wsDraft.AutoFilterMode = False
    End If
    
    On Error Resume Next
    
    ' xoa N/A
    wsBase.AutoFilterMode = False
    Range(wsBase.Cells(3, 1), wsBase.Cells(lastRow_Base, title_Base.Count)).AutoFilter Field:=1, Criteria1:="N/A"
    Range(wsBase.Cells(4, 1), wsBase.Cells(lastRow_Base, title_Base.Count)).SpecialCells(xlCellTypeVisible).ClearContents
    wsBase.AutoFilterMode = False

    wsDraft.AutoFilterMode = False
    Range(wsDraft.Cells(3, 1), wsDraft.Cells(lastRow_Draft, title_draft.Count)).AutoFilter Field:=1, Criteria1:="N/A"
    Range(wsDraft.Cells(4, 1), wsDraft.Cells(lastRow_Draft, title_draft.Count)).SpecialCells(xlCellTypeVisible).ClearContents
    wsDraft.AutoFilterMode = False
    
    ' xoa mau xam
    wsBase.AutoFilterMode = False
    Range(wsBase.Cells(3, 1), wsBase.Cells(lastRow_Base, title_Base.Count)).AutoFilter Field:=1, Criteria1:=RGB(128, 128, 128), Operator:=xlFilterCellColor
    Range(wsBase.Cells(4, 1), wsBase.Cells(lastRow_Base, title_Base.Count)).SpecialCells(xlCellTypeVisible).ClearContents
    wsBase.AutoFilterMode = False
    
    wsDraft.AutoFilterMode = False
    Range(wsDraft.Cells(3, 1), wsDraft.Cells(lastRow_Draft, title_draft.Count)).AutoFilter Field:=1, Criteria1:=RGB(128, 128, 128), Operator:=xlFilterCellColor
    Range(wsDraft.Cells(4, 1), wsDraft.Cells(lastRow_Draft, title_draft.Count)).SpecialCells(xlCellTypeVisible).ClearContents
    wsDraft.AutoFilterMode = False
    
    On Error GoTo 0
End Sub

Sub SortData2()
    lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    If flag_Msg Then lastRow_Msg = wsMsg.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    firstDataRow_Out = 4

    wsDraft.Cells(4, keywordColumn_Base + 1).Formula2 = "=IFERROR(MATCH(RC[-1],'[" & wbBase.name & "]" & wsBase.name & "'!C[-1],0), 9999)"
    Range(wsDraft.Cells(4, keywordColumn_Draft + 1), wsDraft.Cells(lastRow_Draft, keywordColumn_Draft + 1)).FillDown
    wsDraft.Rows("4:" & lastRow_Draft).Sort Key1:=wsDraft.Cells(4, keywordColumn_Draft + 1), Order1:=xlAscending, Header:=xlNo
    
    For i = 4 To lastRow_Base                   ' dich du lieu den hang cuoi cua base thoi nhe  --__--
        If i <> wsDraft.Cells(i, keywordColumn_Draft + 1).Value Then
            wsDraft.Rows(i).Insert
            wsDraft.Rows(i).Interior.Color = RGB(128, 128, 128)
            wsDraft.Cells(i, keywordColumn_Draft).Value = wsBase.Cells(i, keywordColumn_Base).Value
            wsDraft.Cells(i, keywordColumn_Draft + 1).Value = i
        End If
    Next i
    lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    If flag_Msg Then
        wsMsg.Cells(4, keywordColumn_Msg + 1).Formula2 = "=IFERROR(MATCH(RC[-1],'[" & wbDraft.name & "]" & wsBase.name & "'!C,0), 9999)"
        Range(wsMsg.Cells(4, keywordColumn_Msg + 1), wsMsg.Cells(lastRow_Msg, keywordColumn_Msg + 1)).FillDown
        wsMsg.Rows("4:" & lastRow_Msg).Sort Key1:=wsMsg.Cells(4, keywordColumn_Msg + 1), Order1:=xlAscending, Header:=xlNo
        
        For i = 4 To lastRow_Draft                   ' dich du lieu den hang cuoi cua draft thoi nhe  --__--
            If i <> wsMsg.Cells(i, keywordColumn_Msg + 1).Value Then
                wsMsg.Rows(i).Insert
                wsMsg.Rows(i).Interior.Color = RGB(128, 128, 128)
                wsMsg.Cells(i, keywordColumn_Msg).Value = wsDraft.Cells(i, keywordColumn_Draft).Value
                wsMsg.Cells(i, keywordColumn_Msg + 1).Value = i
            End If
        Next i
    End If
End Sub

Sub SortData()
' sap xep du lieu, lay base lam goc
    lastRow_Base = wsBase.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastRow_Draft = wsDraft.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    If flag_Msg Then lastRow_Msg = wsMsg.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    firstDataRow_Out = 4

    For i = firstDataRow_Out To lastRow_Base
        Set tmp1 = wsDraft.Columns(keywordColumn_Draft).Find(What:=wsBase.Cells(i, keywordColumn_Base).Value, LookIn:=xlValues, lookat:=xlWhole)
        If tmp1 Is Nothing Then
            wsDraft.Rows(i).Insert
            wsDraft.Rows(i).Interior.Color = RGB(128, 128, 128)
            wsDraft.Cells(i, keywordColumn_Draft).Value = wsBase.Cells(i, keywordColumn_Base).Value
            lastRow_Draft = lastRow_Draft + 1
        Else
            If tmp1.row <> i Then
                wsDraft.Rows(tmp1.row).Cut
                wsDraft.Rows(i).Insert
            End If
        End If
        
        If flag_Msg Then
            Set tmp2 = wsMsg.Columns(keywordColumn_Msg).Find(What:=wsDraft.Cells(i, keywordColumn_Draft).Value, LookIn:=xlValues, lookat:=xlWhole)
            If tmp2 Is Nothing Then
                wsMsg.Rows(i).Insert
                wsMsg.Rows(i).Interior.Color = RGB(128, 128, 128)
                wsMsg.Cells(i, keywordColumn_Msg).Value = wsDraft.Cells(i, keywordColumn_Draft).Value
                lastRow_Msg = lastRow_Msg + 1
            Else
                If tmp2.row <> i Then
                    wsMsg.Rows(tmp2.row).Cut
                    wsMsg.Rows(i).Insert
                End If
            End If
        End If
    Next i
    
' xu ly phan du cua draft va phan du cua msg
    If flag_Msg Then
        For i = lastRow_Base + 1 To lastRow_Draft
            Set tmp2 = wsMsg.Columns(keywordColumn_Msg).Find(What:=wsDraft.Cells(i, keywordColumn_Draft).Value, LookIn:=xlValues, lookat:=xlWhole)
            If tmp2 Is Nothing Then
                wsMsg.Rows(i).Insert
                wsMsg.Rows(i).Interior.Color = RGB(128, 128, 128)
                wsMsg.Cells(i, keywordColumn_Msg).Value = wsDraft.Cells(i, keywordColumn_Draft).Value
                lastRow_Msg = lastRow_Msg + 1
            Else
                If tmp2.row <> i Then
                    wsMsg.Rows(tmp2.row).Cut
                    wsMsg.Rows(i).Insert
                End If
            End If
        Next i
    End If
End Sub

Sub Copy()
' copy du lieu tu input sang output
    pivotBase = 1
    wsOut.Cells(1, pivotBase).Value = "‘O‰ñ: " & wbBase.name
    Range(wsBase.Cells(2, 1), wsBase.Cells(lastRow_Base, title_Base.Items(title_Base.Count - 1))).Copy Destination:=wsOut.Cells(2, pivotBase)
    pivotDraft = title_Base.Count + 2
    wsOut.Cells(1, pivotDraft).Value = "¡‰ñ: " & wbDraft.name
    Range(wsDraft.Cells(2, 1), wsDraft.Cells(lastRow_Draft, title_draft.Items(title_draft.Count - 1))).Copy Destination:=wsOut.Cells(2, pivotDraft)
    If flag_Msg Then
        pivotMsg = title_Base.Count + title_draft.Count + 3
        wsOut.Cells(1, pivotMsg).Value = "ADASMsgList: " & wbMsg.name
        Range(wsMsg.Cells(2, 1), wsMsg.Cells(lastRow_Msg, title_Msg.Items(title_Msg.Count - 1))).Copy Destination:=wsOut.Cells(2, pivotMsg)
    End If
End Sub



