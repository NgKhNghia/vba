
'''' tao bang tong ket
Sub Summary(thisWorksheet As Worksheet, startRow As Integer, startColumn As Integer, endRow As Integer, lengthTable As Integer, distance As Integer)
' start la vi tri bat dau cua bang ket qua
' end la vi tri ket thuc cua bang ket qua
' lengthTable la so luong cot cua bang true/false
' distance la khoang cach giua bang true/false va bang ket qua
    
    Dim check As Boolean

'    'tao tieu de
    thisWorksheet.Activate
    thisWorksheet.Range(thisWorksheet.Cells(startRow, startColumn), _
        thisWorksheet.Cells(startRow, startColumn + 5)).Interior.Color = RGB(0, 255, 0)

    ' check ˆê’v/•sˆê’v
    For row = startRow + 1 To endRow
        check = True
        For col = startColumn - distance To startColumn - distance + lengthTable - 1
            check = check And thisWorksheet.Cells(row, col)
            If check = False Then
                thisWorksheet.Cells(row, startColumn).Interior.Color = RGB(255, 165, 0)
                thisWorksheet.Cells(row, startColumn) = "•sˆê’v"
                Exit For
            End If
        Next col
        If check = True Then
            thisWorksheet.Cells(row, startColumn) = "ˆê’v"
        End If
    Next row

    With Range(Cells(startRow, startColumn), Cells(endRow, startColumn + 5)).Borders
        .LineStyle = xlContinuous '
        .Weight = xlThin '
        .Color = RGB(0, 0, 0)
    End With
    
'    Range(Cells(startRow, startColumn), Cells(endRow, startColumn + 5)).AutoFilter
End Sub

Sub Sumary2(ws As Worksheet, selectedRange1 As Range, selectedRange2 As Range, lastRow As Integer)
' selectedRange1 la bang true/false
' selectedRange2 la bang ˆê’v/•sˆê’v
    
    check = True
    startColumn1 = selectedRange1.Columns(1).Column
    endColumn1 = selectedRange1.Columns(selectedRange1.Columns.Count).Column
    startrow1 = selectedRange1.Rows(1).row
    For Each cell In selectedRange1
        check = check And cell.Value
        If cell.Column = endColumn1 Then
            If check Then
                ws.Cells(cell.row, selectedRange2.Columns(1).Column).Value = "ˆê’v"
            Else
                With ws.Cells(cell.row, selectedRange2.Columns(1).Column)
                    .Value = "•sˆê’v"
                    .Interior.Color = RGB(255, 165, 0)
                End With
            End If
            check = True
        End If
    Next cell
     With Range(Cells(startrow1, selectedRange2.Columns(1).Column), Cells(lastRow, selectedRange2.Columns(1).Column)).Borders
        .LineStyle = xlContinuous '
        .Weight = xlThin '
        .Color = RGB(0, 0, 0)
    End With
    
    Range(Cells(startrow1 - 1, selectedRange2.Columns(1).Column), Cells(lastRow, selectedRange2.Columns(1).Column)).AutoFilter
End Sub

Sub SummaryFrame(ws As Worksheet, startRow As Integer, startColumn As Integer, endRow As Integer, endColumn As Integer, distance13 As Integer, distance23 As Integer)
' start la o tinh dau tien chua ket qua
' end la o tinh cuoi cung chua ket qua
' distance13 la khoang cach tu bang ket qua so sanh 2 keikakusho voi bang tong ket
' distance23 la khoang cach tu bang ket qua so sanh keikakusho voi messagelist voi bang tong ket
    Dim check As Boolean
    ' xac dinh cot chua adas
    indexAdas = ws.Range(ws.Cells(7, startColumn - distance13), ws.Cells(7, startColumn - distance23 - 3)).Find(What:="ADAS", LookIn:=xlValues, lookat:=xlWhole).Column

    ' cot Œv‰æ‘‚Ì·•ª
    check = True
    lastRow = startRow
    For Each cell In ws.Range(ws.Cells(startRow, startColumn - distance13), ws.Cells(endRow, startColumn - distance13 + 9))
        If lastRow <> cell.row Then
            check = check And ws.Cells(lastRow, indexAdas) And ws.Cells(lastRow, indexAdas + 1)
            If check Then
                ws.Cells(lastRow, startColumn) = "ˆê’v"
            Else
                With ws.Cells(lastRow, startColumn)
                    .Value = "•sˆê’v"
                    .Interior.Color = RGB(255, 165, 0)
                End With
            End If
            check = True
        End If
        check = check And cell.Value
        lastRow = cell.row
    Next cell
    ' xu ly o cuoi cung
    If check Then
        ws.Cells(lastRow, startColumn) = "ˆê’v"
    Else
        With ws.Cells(lastRow, startColumn)
            .Value = "•sˆê’v"
            .Interior.Color = RGB(255, 165, 0)
        End With
    End If

    If Sheet2.ADASmsg.Value Then
    ' check ˆê’v/•sˆê’v
        check = True
        lastRow = startRow
        For Each cell In ws.Range(ws.Cells(startRow, startColumn - distance23), ws.Cells(endRow, startColumn - distance23 + 10))
            If lastRow <> cell.row Then
    '            check = check And ws.Cells(lastRow, indexAdas) And ws.Cells(lastRow, indexAdas + 1)
                If check Then
                    ws.Cells(lastRow, startColumn + 2) = "ˆê’v"
                Else
                    With ws.Cells(lastRow, startColumn + 2)
                        .Value = "•sˆê’v"
                        .Interior.Color = RGB(255, 165, 0)
                    End With
                End If
                check = True
            End If
            check = check And cell.Value
            lastRow = cell.row
        Next cell
            If check Then
                ws.Cells(lastRow, startColumn) = "ˆê’v"
            Else
                With ws.Cells(lastRow, startColumn + 2)
                    .Value = "•sˆê’v"
                    .Interior.Color = RGB(255, 165, 0)
                End With
            End If

            With Range(Cells(startRow, startColumn), Cells(endRow, startColumn)).Borders
                .LineStyle = xlContinuous '
                .Weight = xlThin '
                .Color = RGB(0, 0, 0)
            End With

            With Range(Cells(startRow, startColumn + 2), Cells(endRow, startColumn + 7)).Borders
                .LineStyle = xlContinuous '
                .Weight = xlThin '
                .Color = RGB(0, 0, 0)
            End With
    End If

    Range(Cells(startRow - 1, startColumn), Cells(endRow, startColumn + 7)).AutoFilter
End Sub




