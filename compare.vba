
Sub Compare(ws As Worksheet, Start As Integer, last_row_table As Integer, last_col_table As Integer, distance_12 As Integer, distance_13 As Integer)
' start la hang bat dau vong for
' distance_12 la khoang cach bang 1 voi bang 2
' distance_13 la khoang cach bang 1 voi bang 3
    
    Dim row As Integer
    Dim col As Integer
        
    For row = Start To last_row_table Step 1                          ' row
        For col = 1 To last_col_table Step 1                              ' col
            If ws.Cells(row, col) = "N/A" And ws.Cells(row, col + distance_12) = "" Then
                ws.Cells(row, col + distance_13).Value = True
            ElseIf ws.Cells(row, col) = "" And ws.Cells(row, col + distance_12) = "N/A" Then
                ws.Cells(row, col + distance_13).Value = True
            ElseIf ws.Cells(row, col) = "" And ws.Cells(row, col + distance_12).Font.Strikethrough = True Then
                ws.Cells(row, col + distance_13).Value = True
            ElseIf ws.Cells(row, col).Font.Strikethrough = True And ws.Cells(row, col + distance_12) = "" Then
                ws.Cells(row, col + distance_13).Value = True
            ElseIf ws.Cells(row, col) <> ws.Cells(row, col + distance_12) Then
                ws.Cells(row, col + distance_13).Value = False
                ws.Cells(row, col + distance_13).Interior.Color = RGB(255, 165, 0)
            Else  ' kiem tra tung cap ky tu tuong ung
                Dim checked As Boolean: checked = True
                For i = 1 To Len(ws.Cells(row, col)) Step 1
                    If ws.Cells(row, col).Characters(i, 1).Font.Strikethrough <> ws.Cells(row, col + distance_12).Characters(i, 1).Font.Strikethrough Then
                        ws.Cells(row, col + distance_13).Value = False
                        ws.Cells(row, col + distance_13).Interior.Color = RGB(255, 165, 0)
                        checked = False
                        Exit For
                    End If
                Next i
                
                If checked = True Then
                    ws.Cells(row, col + distance_13).Value = True
                End If
            End If
            
Continue:
        Next col
    Next row
    
    With ws.Range(ws.Cells(Start, 1 + distance_13), ws.Cells(last_row_table, last_col_table + distance_13)).Borders
        .LineStyle = xlContinuous '
        .Weight = xlThin '
        .Color = RGB(0, 0, 0)
    End With
    
End Sub

Sub compare2(ws As Worksheet, selectedRange As Range, distance13 As Integer, distance23 As Integer)
' start la cell dau tien cua bang ket qua
' end la cell cuoi cung cua bang ket qua
' distance12 la khoang cach giua bang thu nhat voi bang ket qua
' distance23 la khoang cach giua bang thu hai voi bang ket qua

' duyet theo hang, het hang thi thuc hien logic and() o o chua ket qua
    firstColumn = selectedRange.Columns(1).Column
    lastColumn = selectedRange.Columns(selectedRange.Columns.Count).Column
    reachLastColumn = False
    For Each cell In selectedRange
        If (IsNumeric(cell.Offset(0, -distance13).Value) And Not IsNumeric(cell.Offset(0, -distance23).Value)) Or _
           (Not IsNumeric(cell.Offset(0, -distance13).Value) And IsNumeric(cell.Offset(0, -distance23).Value)) Then
               With cell
                    .Value = False
                    .Interior.Color = RGB(255, 165, 0)
                End With
        ElseIf IsNumeric(cell.Offset(0, -distance13).Value) And IsNumeric(cell.Offset(0, -distance23).Value) Or _
           IsNumeric(cell.Offset(0, -distance13).Value) And cell.Offset(0, -distance23).Value = "" Or _
           cell.Offset(0, -distance13).Value = "" And IsNumeric(cell.Offset(0, -distance23).Value) Then
        ' du lieu la so hoac rong
            On Error Resume Next
            ' reset gia tri
            num1 = 0
            num2 = 0

            If cell.Offset(0, -distance13).Font.Strikethrough Or cell.Offset(0, -distance13).Interior.Color = RGB(191, 191, 191) Then
            ' neu font ctr+5 hoac mau xam thi doi thanh gia tri 0 de so sanh
                num1 = 0
            Else
                num1 = cell.Offset(0, -distance13).Value
            End If
            
            If cell.Offset(0, -distance23).Font.Strikethrough Or cell.Offset(0, -distance13).Interior.Color = RGB(191, 191, 191) Then
            ' neu font ctr+5 thi doi thanh gia tri 0 de so sanh
                num2 = 0
            Else
                num2 = cell.Offset(0, -distance23).Value
            End If
            
            If num1 = num2 Then
                cell.Value = True
            Else
                With cell
                    .Value = False
                    .Interior.Color = RGB(255, 165, 0)
                End With
            End If

            On Error GoTo 0
        Else
        ' du lieu la chu
            ' reset du lieu
            str1 = ""
            str2 = ""
            
            If cell.Offset(0, -distance13).Interior.Color <> RGB(191, 191, 191) And cell.Offset(0, -distance13).Value <> "N/A" Then
                If cell.Offset(0, -distance13).HasFormula Then
                    cell.Offset(0, -distance13).Value = cell.Offset(0, -distance13).Value
                End If
                For i = 1 To Len(cell.Offset(0, -distance13).Value) Step 1
                
                    char1 = cell.Offset(0, -distance13).Characters(i, 1).Text
                    If Not cell.Offset(0, -distance13).Characters(i, 1).Font.Strikethrough Then
                        str1 = str1 & char1
                    End If
                Next i
            End If
            
            If cell.Offset(0, -distance23).Interior.Color <> RGB(191, 191, 191) And cell.Offset(0, -distance23).Value <> "N/A" Then
                If cell.Offset(0, -distance23).HasFormula Then
                    cell.Offset(0, -distance23).Value = cell.Offset(0, -distance23).Value
                End If
                For i = 1 To Len(cell.Offset(0, -distance23).Value) Step 1
                    char2 = cell.Offset(0, -distance23).Characters(i, 1).Text
                    If Not cell.Offset(0, -distance23).Characters(i, 1).Font.Strikethrough Then
                        str2 = str2 & char2
                    End If
                Next i
            End If
            
            If str1 = str2 Then
                cell.Value = True
            Else
                With cell
                    .Value = False
                    .Interior.Color = RGB(255, 165, 0)
                End With
            End If
        End If
    Next cell
End Sub

Sub CompareMsgListWithObject(ws As Worksheet, startRow As Integer, startColumn As Integer, endRow As Integer, endColumn As Integer, distance12 As Integer, distance13 As Integer)
' start la cell dau tien cua bang ket qua
' end la cell cuoi cung cua bang ket qua
' distance12 la khoang cach giua bang keikakusho voi bang ket qua, khong tinh cot Appli-cation
' distance13 la khoang cach giua bang messagelist voi bang ket qua

    ' compare ecu
    For row = startRow To endRow Step 1
        For col = startColumn To endColumn - 2 Step 1
            If IsNumeric(ws.Cells(row, col - distance12).Value) Then
            ' neu du lieu la so thi xu ly rieng
                On Error Resume Next                ' du lieu chua xu ly duoc
                
                Dim numObj As Integer: numObj = 0
                Dim numMsgList As Integer: numMsgList = 0
                
                numObj = ws.Cells(row, col - distance12).Value
                numMsgList = ws.Cells(row, col - distance13).Value

                If ws.Cells(row, col - distance12).Font.Strikethrough Then
                ' neu font ctr+5 thi doi thanh gia tri 0 de so sanh
                    numObj = 0
                End If
                
                If numObj = numMsgList Then
                    ws.Cells(row, col) = True
                Else
                    With ws.Cells(row, col)
                        .Value = False
                        .Interior.Color = RGB(255, 165, 0)
                    End With
                End If
                
                On Error GoTo 0
            Else
            ' neu cell.value la text thi loai bo cac ky tu co font ctr+5
                strObj = ""        ' chi chua cac ky tu khong bi ctrl+5
                strMsgList = ""    ' chi chua cac ky tu khong bi ctrl+5
                
                For i = 1 To Len(ws.Cells(row, col - distance12)) Step 1
                    charObj = ws.Cells(row, col - distance12).Characters(i, 1).Text
                    If Not ws.Cells(row, col - distance12).Characters(i, 1).Font.Strikethrough Then
                        strObj = strObj & charObj
                    End If
                Next i
                
                For i = 1 To Len(ws.Cells(row, col - distance13)) Step 1
                    charMsgList = ws.Cells(row, col - distance13).Characters(i, 1).Text
                    If Not ws.Cells(row, col - distance13).Characters(i, 1).Font.Strikethrough Then
                        strMsgList = strMsgList & charMsgList
                    End If
                Next i
                
                If strObj = strMsgList Then
                    ws.Cells(row, col) = True
                Else
                    With ws.Cells(row, col)
                        .Value = False
                        .Interior.Color = RGB(255, 165, 0)
                    End With
                End If
            End If
        Next col
    Next row

    ' compaare adas va adas_bridge
    index_adas = ws.Range(ws.Cells(7, startColumn - distance12), ws.Cells(7, endColumn - distance13 - 2)).Find(What:="ADAS", LookIn:=xlValues, lookat:=xlWhole).Column
    For row = startRow To endRow Step 1
        For col = endColumn - 1 To endColumn Step 1
            strObj = ""        ' chi chua cac ky tu khong bi ctrl+5
            strMsgList = ""    ' chi chua cac ky tu khong bi ctrl+5
            
            For i = 1 To Len(ws.Cells(row, col - endColumn + index_adas + 1)) Step 1
                charObj = ws.Cells(row, col - endColumn + index_adas + 1).Characters(i, 1).Text
                If Not ws.Cells(row, col - endColumn + index_adas + 1).Characters(i, 1).Font.Strikethrough Then
                    strObj = strObj & charObj
                End If
            Next i
            
            For i = 1 To Len(ws.Cells(row, col - distance13)) Step 1
                charMsgList = ws.Cells(row, col - distance13).Characters(i, 1).Text
                If Not ws.Cells(row, col - distance13).Characters(i, 1).Font.Strikethrough Then
                    strMsgList = strMsgList & charMsgList
                End If
            Next i
            
            If strObj = strMsgList Then
                ws.Cells(row, col) = True
            Else
                With ws.Cells(row, col)
                    .Value = False
                    .Interior.Color = RGB(255, 165, 0)
                End With
            End If
            
        Next col
    Next row

    With ws.Range(ws.Cells(startRow, startColumn), ws.Cells(endRow, endColumn)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
End Sub

Sub compare3(ws As Worksheet, selectedRange As Range, distance13 As Integer, distance23 As Integer)
    Dim firstColumn As Integer
    Dim lastColumn As Integer
    Dim num1 As Variant
    Dim num2 As Variant
    Dim str1 As String
    Dim str2 As String
    Dim char1 As String
    Dim char2 As String
    Dim i As Integer
    Dim cell As Range

    firstColumn = selectedRange.Columns(1).Column
    lastColumn = selectedRange.Columns(selectedRange.Columns.Count).Column

    For Each cell In selectedRange
        ' Ki?m tra ði?u ki?n cho s?
        If cell.Offset(0, -distance13).HasFormula Then
            cell.Offset(0, -distance13).Value = cell.Offset(0, -distance13).Value
        End If
        If cell.Offset(0, -distance23).HasFormula Then
            cell.Offset(0, -distance23).Value = cell.Offset(0, -distance23).Value
        End If
        If IsNumeric(cell.Offset(0, -distance13).Value) And IsNumeric(cell.Offset(0, -distance23).Value) Then
            ' N?u c? hai ô ð?u có d? li?u
            num1 = IIf(cell.Offset(0, -distance13).Font.Strikethrough Or cell.Offset(0, -distance13).Interior.Color = RGB(191, 191, 191), 0, cell.Offset(0, -distance13).Value)
            num2 = IIf(cell.Offset(0, -distance23).Font.Strikethrough Or cell.Offset(0, -distance23).Interior.Color = RGB(191, 191, 191), 0, cell.Offset(0, -distance23).Value)

            If num1 = num2 Then
                cell.Value = True
            Else
                cell.Value = False
                cell.Interior.Color = RGB(255, 165, 0) ' Màu cam
            End If

        ' Ki?m tra ði?u ki?n cho chu?i
        ElseIf Not IsEmpty(cell.Offset(0, -distance13).Value) Or Not IsEmpty(cell.Offset(0, -distance23).Value) Then
            ' Reset d? li?u
            str1 = ""
            str2 = ""

            If cell.Offset(0, -distance13).Interior.Color <> RGB(191, 191, 191) And cell.Offset(0, -distance13).Value <> "N/A" Then
                For i = 1 To Len(cell.Offset(0, -distance13).Value)
                    char1 = cell.Offset(0, -distance13).Characters(i, 1).Text
                    If Not cell.Offset(0, -distance13).Characters(i, 1).Font.Strikethrough Then
                        str1 = str1 & char1
                    End If
                Next i
            End If

            If cell.Offset(0, -distance23).Interior.Color <> RGB(191, 191, 191) And cell.Offset(0, -distance23).Value <> "N/A" Then
                For i = 1 To Len(cell.Offset(0, -distance23).Value)
                    char2 = cell.Offset(0, -distance23).Characters(i, 1).Text
                    If Not cell.Offset(0, -distance23).Characters(i, 1).Font.Strikethrough Then
                        str2 = str2 & char2
                    End If
                Next i
            End If

            ' N?u m?t ô có d? li?u không b? g?ch ngang và ô ðó b? bôi xám
            If (cell.Offset(0, -distance13).Interior.Color = RGB(191, 191, 191) And str1 <> "") Then
                cell.Value = True
            ElseIf str1 = str2 Then
                cell.Value = True
            Else
                cell.Value = False
                cell.Interior.Color = RGB(255, 165, 0) ' Màu cam
            End If
        Else
            cell.Value = False
            cell.Interior.Color = RGB(255, 165, 0) ' Màu cam
        End If
    Next cell
End Sub

Sub compare4(ws As Worksheet, selectedRange As Range, distance13 As Integer, distance23 As Integer)
    Dim firstColumn As Integer
    Dim lastColumn As Integer
    Dim num1 As Variant
    Dim num2 As Variant
    Dim str1 As String
    Dim str2 As String
    Dim char1 As String
    Dim char2 As String
    Dim i As Integer
    Dim cell As Range

    firstColumn = selectedRange.Columns(1).Column
    lastColumn = selectedRange.Columns(selectedRange.Columns.Count).Column

    For Each cell In selectedRange
        ' Ki?m tra ði?u ki?n cho s?
        If cell.Offset(0, -distance13).HasFormula Then
            cell.Offset(0, -distance13).Value = cell.Offset(0, -distance13).Value
        End If
        If cell.Offset(0, -distance23).HasFormula Then
            cell.Offset(0, -distance23).Value = cell.Offset(0, -distance23).Value
        End If
        If IsNumeric(cell.Offset(0, -distance13).Value) And IsNumeric(cell.Offset(0, -distance23).Value) Then
            ' N?u c? hai ô ð?u có d? li?u
            num1 = IIf(cell.Offset(0, -distance13).Font.Strikethrough Or cell.Offset(0, -distance13).Interior.Color = RGB(191, 191, 191), 0, cell.Offset(0, -distance13).Value)
            num2 = IIf(cell.Offset(0, -distance23).Font.Strikethrough Or cell.Offset(0, -distance23).Interior.Color = RGB(191, 191, 191), 0, cell.Offset(0, -distance23).Value)

            If num1 = num2 Then
                cell.Value = True
            Else
                cell.Value = False
                cell.Interior.Color = RGB(255, 165, 0) ' Màu cam
            End If

        ' Ki?m tra ði?u ki?n cho chu?i
        ElseIf Not IsEmpty(cell.Offset(0, -distance13).Value) Or Not IsEmpty(cell.Offset(0, -distance23).Value) Then
            ' Reset d? li?u
            str1 = ""
            str2 = ""

            If cell.Offset(0, -distance13).Interior.Color <> RGB(191, 191, 191) And cell.Offset(0, -distance13).Value <> "N/A" Then
                For i = 1 To Len(cell.Offset(0, -distance13).Value)
                    char1 = cell.Offset(0, -distance13).Characters(i, 1).Text
                    If Not cell.Offset(0, -distance13).Characters(i, 1).Font.Strikethrough Then
                        str1 = str1 & char1
                    End If
                Next i
            End If

            If cell.Offset(0, -distance23).Interior.Color <> RGB(191, 191, 191) And cell.Offset(0, -distance23).Value <> "N/A" Then
                For i = 1 To Len(cell.Offset(0, -distance23).Value)
                    char2 = cell.Offset(0, -distance23).Characters(i, 1).Text
                    If Not cell.Offset(0, -distance23).Characters(i, 1).Font.Strikethrough Then
                        str2 = str2 & char2
                    End If
                Next i
            End If

            ' N?u m?t ô có d? li?u không b? g?ch ngang và ô ðó b? bôi xám
            If (cell.Offset(0, -distance13).Interior.Color = RGB(191, 191, 191) And str1 <> "") Then
                cell.Value = True
            ElseIf str1 = str2 Then
                cell.Value = True
            Else
                cell.Value = False
                cell.Interior.Color = RGB(255, 165, 0) ' Màu cam
            End If
        Else
            cell.Value = False
            cell.Interior.Color = RGB(255, 165, 0) ' Màu cam
        End If
    Next cell
End Sub
