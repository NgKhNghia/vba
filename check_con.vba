
Sub NetworkPathCheck(wsnet1 As Worksheet, wsnet2 As Worksheet, wbdict As Workbook)
' check pdu co N/A hay khong
' neu khong -> to mau
    '"Network Path"ƒV[ƒg‚ðƒRƒs[‚µ‚ÄAN/As‚ðíœ
    
    Call wsnet1.Copy(After:=wbdict.Sheets(1))
    With wbdict.Worksheets("Network Path (2)")
    
        .Rows(4).AutoFilter 2, RGB(191, 191, 191), xlFilterFontColor
        If .AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
            .Range("A5:A" & .UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete shift:=xlUp
        End If
        .ShowAllData
    End With
    
    Call wsnet2.Copy(After:=wbdict.Sheets(1))
    With wbdict.Worksheets("Network Path (3)")
    
        .Rows(4).AutoFilter 2, RGB(191, 191, 191), xlFilterFontColor
        If .AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
            .Range("A5:A" & .UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete shift:=xlUp
        End If
        .ShowAllData
    End With
    
End Sub

Sub FrameSynthCheck(wsframe1 As Worksheet, wsframe2 As Worksheet, wbdict As Workbook)
' check frame co ton tai va lien quan toi adas hay khong
' neu khong -> to mau
    
    Dim TargetCell As Excel.Range
    Dim TargetCol1 As Integer
    Dim TargetCol2 As Integer
    Dim FrameNameCol As Integer
    Dim DataRow As Integer
    Dim Bottom As Integer
    
    '"Frame Synthesis all"ƒV[ƒg‚ðƒRƒs[‚µ‚ÄAADAS–¢Žg—ps‚ðíœ
    Call wsframe1.Copy(Before:=wbdict.Worksheets(1))
    
    With wbdict.Worksheets(1)
        Set TargetCell = .Rows(7).Find("ADAS")
        TargetCol1 = TargetCell.Column
        Set TargetCell = .Rows(7).Find("ADAS_Bridge")
        TargetCol2 = TargetCell.Column
        Set TargetCell = .Rows(7).Find("Frame Name")
        FrameNameCol = TargetCell.Column
        
        DataRow = TargetCell.row + 1
        Bottom = .Cells(Rows.Count, FrameNameCol).End(xlUp).row
        
        .Rows(7).AutoFilter TargetCol1, RGB(191, 191, 191), xlFilterFontColor
        If .AutoFilter.Range.Columns(FrameNameCol).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
            .Range(.Cells(DataRow, TargetCol1), .Cells(Bottom, TargetCol1)).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .ShowAllData
        
        .Rows(7).AutoFilter TargetCol2, RGB(191, 191, 191), xlFilterFontColor
        If .AutoFilter.Range.Columns(FrameNameCol).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
            .Range(.Cells(DataRow, TargetCol2), .Cells(Bottom, TargetCol2)).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .ShowAllData
        
        .Rows(7).AutoFilter TargetCol1, ""
        .Rows(7).AutoFilter TargetCol2, ""
        If .AutoFilter.Range.Columns(FrameNameCol).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
             .Range("A8:A" & .UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete shift:=xlUp
        End If
        .ShowAllData
        
    End With
    
    Call wsframe2.Copy(Before:=wbdict.Worksheets(1))
    
    With wbdict.Worksheets(1)
        Set TargetCell = .Rows(7).Find("ADAS")
        TargetCol1 = TargetCell.Column
        Set TargetCell = .Rows(7).Find("ADAS_Bridge")
        TargetCol2 = TargetCell.Column
        Set TargetCell = .Rows(7).Find("Frame Name")
        FrameNameCol = TargetCell.Column
        
        DataRow = TargetCell.row + 1
        Bottom = .Cells(Rows.Count, FrameNameCol).End(xlUp).row
        
        .Rows(7).AutoFilter TargetCol1, RGB(191, 191, 191), xlFilterFontColor
        If .AutoFilter.Range.Columns(FrameNameCol).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
            .Range(.Cells(DataRow, TargetCol1), .Cells(Bottom, TargetCol1)).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .ShowAllData
        
        .Rows(7).AutoFilter TargetCol2, RGB(191, 191, 191), xlFilterFontColor
        If .AutoFilter.Range.Columns(FrameNameCol).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
            .Range(.Cells(DataRow, TargetCol2), .Cells(Bottom, TargetCol2)).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .ShowAllData
        
        .Rows(7).AutoFilter TargetCol1, ""
        .Rows(7).AutoFilter TargetCol2, ""
        If .AutoFilter.Range.Columns(FrameNameCol).SpecialCells(xlCellTypeVisible).Count = 1 Then
        Else
            .Range("A8:A" & .UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete shift:=xlUp
        End If
        .ShowAllData
        
    End With
End Sub
