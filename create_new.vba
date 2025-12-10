
Sub Creat_New_Workbook(wb As Workbook, path1 As String, path2 As String)
    Set wb = Workbooks.Add
    wb.SaveAs fileName:=ThisWorkbook.path & "\compare_" & GetFileNamesWithoutExtension(path1, path2) & ".xlsx"
    wb.Activate
'    Application.Calculation = xlCalculationAutomatic
'    If Sheet2.get_optBtn3_Click() Then
'        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "Frame Synthesis"
'        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "Construction of Container frame"
'        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "Network Path"
'        If Sheet2.CheckBox2.Value Then
'            wb.Sheets(1).name = "System"
'        Else
'            Application.DisplayAlerts = False
'            wb.Sheets(1).Delete
'            Application.DisplayAlerts = True
'        End If
'
'    End If
    
    If Sheet2.OptionButton2.Value Then
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "Frame Synthesis"
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "Network Path"
        If Sheet2.CheckBox2.Value Then
            wb.Sheets(1).name = "System"
        Else
            Application.DisplayAlerts = False
            wb.Sheets(1).Delete
            Application.DisplayAlerts = True
        End If
    End If
    
    If Sheet2.OptionButton1.Value Then
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "FrameList"
        If Sheet2.CheckBox2.Value Then
            wb.Sheets(1).name = "System"
        Else
            Application.DisplayAlerts = False
            wb.Sheets(1).Delete
            Application.DisplayAlerts = True
        End If
    End If

End Sub

Function GetFileNamesWithoutExtension(path1 As String, path2 As String) As String
    Dim fileName1 As String
    Dim fileName2 As String
    Dim fileNameWithoutExt1 As String
    Dim fileNameWithoutExt2 As String

    ' get file name from path
    fileName1 = Mid(path1, InStrRev(path1, "\") + 1)
    fileName2 = Mid(path2, InStrRev(path2, "\") + 1)

    ' delete extension
    fileNameWithoutExt1 = Left(fileName1, InStrRev(fileName1, ".") - 1)
    fileNameWithoutExt2 = Left(fileName2, InStrRev(fileName2, ".") - 1)

    GetFileNamesWithoutExtension = fileNameWithoutExt1 & "_with_" & fileNameWithoutExt2
End Function
Function GetFileName(path As String) As String
    Dim fileName As String
    Dim fileNameWithoutExt As String
    
    fileName = Mid(path, InStrRev(path, "\") + 1)
    fileNameWithoutExt = Left(fileName, InStrRev(fileName, ".") - 1)
    
    GetFileName = fileNameWithoutExt
End Function

