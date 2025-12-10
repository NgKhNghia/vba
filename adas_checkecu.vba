Sub XuLyInput(ws_base As Worksheet, ws_draft As Worksheet)
'    Dim ws_base As Worksheet
'    Dim ws_draft As Worksheet
'
'    Set ws_base = Workbooks.Open(fileName:="C:\Users\KNT23265\Desktop\tool_CAN\tool so sanh kei vs kei vs msgl\testcase\base_P42S26MY(US)2025FEB11_PT2toSOP-1.xlsx", ReadOnly:=False, Password:="canxes").Worksheets("Frame Synthesis")
'    Set ws_draft = Workbooks.Open(fileName:="C:\Users\KNT23265\Desktop\tool_CAN\tool so sanh kei vs kei vs msgl\testcase\draft_P42S26MY(US)2025JUL10_PT2toSOP-1.xlsx", ReadOnly:=False, Password:="canxes").Worksheets("Frame Synthesis")
    
    Dim dict_base As Scripting.Dictionary
    Dim dict_draft As Scripting.Dictionary
    Dim dict_merge As Scripting.Dictionary
    
    Set dict_base = New Scripting.Dictionary
    Set dict_draft = New Scripting.Dictionary
    Set dict_merge = New Scripting.Dictionary
    
    Dim last_column_base As Integer
    Dim last_column_draft As Integer
    
    last_column_base = ws_base.Cells(7, 2).End(xlToRight).Column
    last_column_draft = ws_draft.Cells(7, 2).End(xlToRight).Column
    
    Dim last_row_base As Integer
    Dim last_row_draft As Integer
    
    last_row_base = ws_base.Cells(7, 2).End(xlDown).row
    last_row_draft = ws_draft.Cells(7, 2).End(xlDown).row
        
    Dim cell As Range
    Dim key As String
    Dim col_base As Integer
    Dim col_draft As Integer
    
    ' gop cac ecu lai de sap xep, giu nguyen thu tu theo base
    For Each cell In ws_base.Range(ws_base.Cells(7, 11), ws_base.Cells(7, last_column_base))
        dict_base.Add cell.Value, cell.Column
    Next cell
    
    For Each cell In ws_draft.Range(ws_draft.Cells(7, 11), ws_draft.Cells(7, last_column_draft))
        dict_draft.Add cell.Value, cell.Column
    Next cell
    
    For i = 0 To dict_base.Count - 1
        key = dict_base.Keys(i)
        If dict_draft.Exists(key) Then
            dict_merge.Add key, Array(dict_base.Item(key), dict_draft.Item(key))
        Else
            dict_merge.Add key, Array(dict_base.Item(key), 0)
        End If
    Next i
    
    For i = 0 To dict_draft.Count - 1
        key = dict_draft.Keys(i)
        If Not dict_merge.Exists(key) Then
            dict_merge.Add key, Array(0, dict_draft.Item(key))
        End If
    Next i
    
    ' bat dau them cac ecu khong co vao trong tung tai lieu tuong ung
    col_base = 11
    col_draft = 11
    For i = 0 To dict_merge.Count - 1
        If dict_merge.Items(i)(0) = 0 Then
            ws_base.Columns(col_base).Insert shift:=xlToRight
            ws_base.Columns(col_base).Interior.Color = RGB(191, 191, 191)
            ws_base.Cells(7, col_base).Value = dict_merge.Keys(i)
            last_column_base = last_column_base + 1
        ElseIf dict_merge.Items(i)(1) = 0 Then
            ws_draft.Columns(col_draft).Insert shift:=xlToRight
            ws_draft.Columns(col_draft).Interior.Color = RGB(191, 191, 191)
            ws_draft.Cells(7, col_draft).Value = dict_merge.Keys(i)
            last_column_draft = last_column_draft + 1
        End If
        col_base = col_base + 1
        col_draft = col_draft + 1
    Next i
    
    ' ke khung
    With ws_base.Range(ws_base.Cells(7, 1), ws_base.Cells(last_row_base, last_column_base)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With ws_draft.Range(ws_draft.Cells(7, 1), ws_draft.Cells(last_row_draft, last_column_draft)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

