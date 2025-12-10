Public lastDataRow_Out As Integer

Dim adas_Base As Integer
Dim adas_bridge_Base As Integer

Dim adas_draft As Integer
Dim adas_bridge_draft As Integer

Dim adas_Msg As Integer
Dim adas_bridge_Msg As Integer

Dim adas_CompareAdas1 As Integer
Dim adas_bridge_CompareAdas1 As Integer

Public pivotBase As Integer
Public pivotDraft As Integer
Public pivotMsg As Integer
Public pivotCompareAdas1 As Integer
Public pivotCompareAdas2 As Integer
Public pivotSummaryAdas As Integer

Dim tmp As Range

Sub CompareAdas(content As String, pivot1 As Integer, pivot2 As Integer, pivot3 As Integer, firstTime As Boolean)
' content la noi dung cua bang
' pivot1 la cot dau tien cua bang so sanh thu nhat
' pivot2 la cot dau tien cua bang so sanh thu hai
' pivot3 la cot dau tien cua bang ket qua
' firstTime la bang so sanh dau tien
    If firstTime Then
        lastDataRow_Out = wsOut.Cells.SpecialCells(xlCellTypeLastCell).row
'        LastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
        wsOut.Cells(1, pivot3).Value = content
        Range(wsOut.Cells(2, pivot2), wsOut.Cells(firstDataRow_Out - 1, pivot2 + title_draft.Count)).Copy _
            Destination:=wsOut.Cells(2, pivot3)             ' copy tieu de
                
        ' so sanh
        wsOut.Cells(firstDataRow_Out, pivot3).Formula2 = "="""" & " & wsOut.Cells(firstDataRow_Out, pivot1).Address(False, False) & " = " & _
            """"" & " & wsOut.Cells(firstDataRow_Out, pivot2).Address(False, False)
        With Range(wsOut.Cells(firstDataRow_Out, pivot3), wsOut.Cells(lastDataRow_Out, pivot3 + title_draft.Count - 1))
            .FillRight
            .FillDown
            .Copy
            .PasteSpecial xlPasteValues
            .Borders.LineStyle = xlContinuous
            .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE").Interior.Color = RGB(255, 255, 0)
        End With
        
        ' so sanh cac cot dac biet
        If Not flag_Msg Then
            If InStr(wsBase.name, "Construction") Then
                Call SubCompareAdas(5, 22, 38)
                Call SubCompareAdas(16, 33, 49)
            End If
        Else
            If InStr(wsBase.name, "Construction") Then
                Call SubCompareAdas(5, 22, 55)
                Call SubCompareAdas(16, 33, 66)
            End If
        End If
        
        ' tra lai du lieu goc
        Range(wsBase.Cells(firstDataRow_Out, 21), wsBase.Cells(lastRow_Base, 21 + title_Base.Count - 1)).Copy wsOut.Cells(4, pivot1)
        Range(wsDraft.Cells(firstDataRow_Out, 21), wsDraft.Cells(lastRow_Draft, 21 + title_draft.Count - 1)).Copy wsOut.Cells(4, pivot2)
        
    Else
        lastDataRow_Out = wsOut.Cells.SpecialCells(xlCellTypeLastCell).row
        wsOut.Cells(1, pivot3).Value = content
        Range(wsOut.Cells(2, pivot2), wsOut.Cells(firstDataRow_Out - 1, pivot2 + title_Msg.Count)).Copy _
            Destination:=wsOut.Cells(2, pivot3)             ' copy tieu de
        
        ' xoa hang N/A
        wsOut.AutoFilterMode = False
        Range(wsOut.Cells(firstDataRow_Out - 1, pivot1), wsOut.Cells(lastDataRow_Out, pivot1)).AutoFilter Field:=1, Criteria1:="N/A"
        Range(wsOut.Cells(firstDataRow_Out, pivot1), wsOut.Cells(lastDataRow_Out, pivot1 + title_draft.Count - 1)).SpecialCells(xlCellTypeVisible).ClearContents
        wsOut.AutoFilterMode = False
'
'        ' xoa mau xam
'        If InStr(wsBase.Name, "Construction") Then
'            wsOut.AutoFilterMode = False
'            Range(wsOut.Cells(firstDataRow_Out - 1, pivot1), wsOut.Cells(lastDataRow_Out, pivot1)).AutoFilter Field:=1, Criteria1:=RGB(128, 128, 128), Operator:=xlFilterCellColor
'            Range(wsOut.Cells(firstDataRow_Out, pivot1), wsOut.Cells(lastDataRow_Out, pivot1 + title_Base.Count - 1)).SpecialCells(xlCellTypeVisible).ClearContents
'            wsOut.AutoFilterMode = False
'
''            wsOut.AutoFilterMode = False
''            Range(wsOut.Cells(firstDataRow_Out - 1, pivot2), wsOut.Cells(lastDataRow_Out, pivot2)).AutoFilter Field:=1, Criteria1:=RGB(128, 128, 128), Operator:=xlFilterCellColor
''            Range(wsOut.Cells(firstDataRow_Out, pivot2), wsOut.Cells(lastDataRow_Out, pivot2 + title_Draft.Count - 1)).SpecialCells(xlCellTypeVisible).ClearContents
''            wsOut.AutoFilterMode = False
'        End If

        ' so sanh
        wsOut.Cells(firstDataRow_Out, pivot3).Formula2 = "="""" & " & wsOut.Cells(firstDataRow_Out, pivot1 + 1).Address(False, False) & " = " & _
            """"" & " & wsOut.Cells(firstDataRow_Out, pivot2).Address(False, False)
        With Range(wsOut.Cells(firstDataRow_Out, pivot3), wsOut.Cells(lastDataRow_Out, pivot3 + title_Msg.Count - 1))
            .FillRight
            .FillDown
            .Copy
            .PasteSpecial xlPasteValues
            .Borders.LineStyle = xlContinuous
            .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE").Interior.Color = RGB(255, 255, 0)
        End With
        
        ' so sanh cac cot dac biet
        If InStr(wsBase.name, "Construction") Then
            Call SubCompareAdas(22, 38, 71)
            Call SubCompareAdas(33, 49, 82)
'        ElseIf InStr(wsBase.Name, "Network") Then
'            Call SubCompareAdas(17, 30, 57)
        End If
        
        ' tra lai du lieu goc
        Range(wsDraft.Cells(firstDataRow_Out, 21), wsDraft.Cells(lastRow_Draft, 21 + title_draft.Count - 1)).Copy wsOut.Cells(4, pivot1)
'        Range(wsMsg.Cells(firstDataRow_Out, 1), wsMsg.Cells(lastRow_Msg, title_Msg.Count)).Copy wsOut.Cells(4, pivot2)
    End If
End Sub

Sub SubCompareAdas(columnFirst As Integer, columnSecond As Integer, columnThird As Integer)
' so sanh Tx unit va original Tx unit
' tach chuoi roi so sanh tung phan tu cua mang
    Dim arrFirst() As String, arrSecond() As String
    Dim flagElement As Boolean, flagArray As Boolean
    For i = firstDataRow_Out To lastDataRow_Out
        arrFirst = Split(wsOut.Cells(i, columnFirst), "/")
        arrSecond = Split(wsOut.Cells(i, columnSecond), "/")
        
        If UBound(arrFirst) <> UBound(arrSecond) Then
            wsOut.Cells(i, columnThird).Value = False
        Else
            flagArray = True
            For j = LBound(arrFirst) To UBound(arrFirst)
                flagElement = False
                For k = LBound(arrSecond) To UBound(arrSecond)
                    If arrFirst(j) = arrSecond(k) Then
                        flagElement = True
                        Exit For
                    End If
                Next k
                If Not flagElement Then
                    wsOut.Cells(i, columnThird).Value = False
                    flagArray = False
                    Exit For
                End If
            Next j
            wsOut.Cells(i, columnThird).Value = True
        End If
    Next i
End Sub

Sub SummaryAdas(pivot3 As Integer, pivot1 As Integer, Optional pivot2 As Integer = 0)
' pivot1 la cot dau tien cua bang so sanh thu nhat
' pivot2 la cot dau tien cua bang so sanh thu hai, neu co
' pivot3 la cot dau tien cua bang ket qua
    If Not flag_Msg Then
        With Range(wsOut.Cells(3, pivot3), wsOut.Cells(3, pivot3 + 5))
            .Value = Array("ˆê’v/•sˆê’v", "”»’è", "·•ª“à—e", "Œ©‰ðE”õl", "·•ªî•ñ", "Tag")
            .Font.Bold = True
            .Interior.Color = RGB(0, 255, 0)
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        wsOut.Cells(4, pivot3).Formula = "=IF(AND(" & wsOut.Cells(4, pivot1).Address(False, False) & ":" & _
            wsOut.Cells(4, pivot1 + title_draft.Count - 1).Address(False, False) & "), ""ˆê’v"", ""•sˆê’v"")"
        wsOut.Cells(4, pivot3 + 1).Formula2 = "=IF(" & wsOut.Cells(4, pivot3).Address(False, False) & "=""ˆê’v"", ""OK"", """")"
        Range(wsOut.Cells(4, pivot3), wsOut.Cells(lastDataRow_Out, pivot3 + 1)).FillDown
        Range(wsOut.Cells(4, pivot3), wsOut.Cells(lastDataRow_Out, pivot3)) _
            .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""•sˆê’v""").Interior.Color = RGB(255, 255, 0)
        
        With Range(wsOut.Cells(3, pivot3), wsOut.Cells(lastDataRow_Out, pivot3 + 5))
            .Copy
            .PasteSpecial xlPasteValues
            .Borders.LineStyle = xlContinuous
        End With
        
    Else
        With Range(wsOut.Cells(3, pivot3), wsOut.Cells(3, pivot3 + 6))
            .Value = Array("Œv‰æ‘·•ª", "ˆê’v/•sˆê’v", "”»’è", "·•ª“à—e", "Œ©‰ðE”õl", "·•ªî•ñ", "Tag")
            .Font.Bold = True
            .Interior.Color = RGB(0, 255, 0)
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        wsOut.Cells(4, pivot3).Formula = "=IF(AND(" & wsOut.Cells(4, pivot1).Address(False, False) & ":" & _
            wsOut.Cells(4, pivot1 + title_draft.Count - 1).Address(False, False) & "), ""ˆê’v"", ""•sˆê’v"")"
        wsOut.Cells(4, pivot3 + 1).Formula = "=IF(AND(" & wsOut.Cells(4, pivot2).Address(False, False) & ":" & _
            wsOut.Cells(4, pivot2 + title_Msg.Count - 1).Address(False, False) & "), ""ˆê’v"", ""•sˆê’v"")"
        wsOut.Cells(4, pivot3 + 2).Formula = "=IF(" & wsOut.Cells(4, pivot3 + 1).Address(False, False) & "=""ˆê’v"", ""OK"", """")"
    
        With Range(wsOut.Cells(4, pivot3), wsOut.Cells(lastDataRow_Out, pivot3 + 2))
            .FillDown
            .Copy
            .PasteSpecial xlPasteValues
        End With
        
        Range(wsOut.Cells(4, pivot3), wsOut.Cells(lastDataRow_Out, pivot3 + 1)) _
            .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""•sˆê’v""").Interior.Color = RGB(255, 255, 0)
    
        Range(wsOut.Cells(3, pivot3), wsOut.Cells(lastDataRow_Out, pivot3 + 6)).Borders.LineStyle = xlContinuous
    End If
    
'    Rows("3:" & lastDataRow_Out).AutoFilter
End Sub


