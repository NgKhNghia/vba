Public wbBase As Workbook
Public wbDraft As Workbook
Public wbMsg As Workbook
Public wbOut As Workbook

Public wsBase As Worksheet
Public wsDraft As Worksheet
Public wsMsg As Worksheet
Public wsOut As Worksheet

Public flag_Msg As Boolean

Sub main()
'    ThisWorkbook.Save
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Dim timePoint As Double
    'timePoint = Timer
    
    ' check xem input day du chua
    If ThisWorkbook.Sheets("Main").TextBox1.Text = "" Then
        MsgBox "Check Input file!"
        Exit Sub
    End If
    
    If ThisWorkbook.Sheets("Main").TextBox2.Text = "" Then
        MsgBox "Check Input file!"
        Exit Sub
    End If
    
    If ThisWorkbook.Sheets("Main").CheckBox1.Value Then
        If ThisWorkbook.Sheets("Main").TextBox3.Text = "" Then
            MsgBox "Check ADASŽd—l file!"
        End If
    End If
    
    
    
    Set wbBase = Workbooks.Open(fileName:=ThisWorkbook.Sheets("Main").TextBox1.Text, UpdateLinks:=False, Password:="canxes", ReadOnly:=True)
    Set wbDraft = Workbooks.Open(fileName:=ThisWorkbook.Sheets("Main").TextBox2.Text, UpdateLinks:=False, Password:="canxes", ReadOnly:=True)
    ' check xem dung tai lieu chua
    If Not adas_new_input.SheetExists(wbBase, "Frame Synthesis") Then
        MsgBox "The input file has incorrect content!"
        Exit Sub
    ElseIf Not adas_new_input.SheetExists(wbBase, "Construction of Container frame") Then
        MsgBox "The input file has incorrect content."
        Exit Sub
    ElseIf Not adas_new_input.SheetExists(wbBase, "Network Path") Then
        MsgBox "The input file has incorrect content."
        Exit Sub
    ElseIf Not adas_new_input.SheetExists(wbDraft, "Frame Synthesis") Then
        MsgBox "The input file has incorrect content."
        Exit Sub
    ElseIf Not adas_new_input.SheetExists(wbDraft, "Construction of Container frame") Then
        MsgBox "The input file has incorrect content."
        Exit Sub
    ElseIf Not adas_new_input.SheetExists(wbDraft, "Network Path") Then
        MsgBox "The input file has incorrect content."
        Exit Sub
    End If
    
    If ThisWorkbook.Worksheets("Main").CheckBox1.Value Then
        Set wbMsg = Workbooks.Open(fileName:=ThisWorkbook.Sheets("Main").TextBox3.Text, UpdateLinks:=False, Password:="canxes", ReadOnly:=True)
        flag_Msg = True
        ' check xem dung tai lieu chua
        If Not adas_new_input.SheetExists(wbMsg, "Frame Synthesis (FD+HS) all CAN") Then
            MsgBox "ADASŽd—l file has incorrect content."
            Exit Sub
        ElseIf Not adas_new_input.SheetExists(wbMsg, "Construction of Container frame") Then
            MsgBox "ADASŽd—l file has incorrect content."
            Exit Sub
        ElseIf Not adas_new_input.SheetExists(wbMsg, "Network Path") Then
            MsgBox "ADASŽd—l file has incorrect content."
            Exit Sub
        End If
    End If
    
    ' tao output
    Set wbOut = Workbooks.Add
'   System
    
    If Sheet2.CheckBox2.Value Then
    Set wsBase = wbBase.Worksheets("System")
    Set wsDraft = wbDraft.Worksheets("System")
    Set wsOut = wbOut.Worksheets.Add
    wsOut.name = "System"
    Call Newcopysystem(wsBase, wsDraft, wsOut)
    End If
    
'    frame synthesis
    Set wsBase = wbBase.Worksheets("Frame Synthesis")
    Set wsDraft = wbDraft.Worksheets("Frame Synthesis")
    If flag_Msg Then Set wsMsg = wbMsg.Worksheets("Frame Synthesis (FD+HS) all CAN")
    Set wsOut = wbOut.Worksheets.Add
    wsOut.name = "Frame Synthesis (FD+HS) all CAN"

    Call Refresh
    Call DeleteEmptyRow
    Call ConfirmTitle
    Call LoadTitle
    Call CreatKeyword
    Call DeleteData
    Call SortData2
    Call Copy
    If Not flag_Msg Then
        Call CompareAdas("”äŠr: ‘O‰ñ‚Æ¡‰ñ", pivotBase, pivotDraft, title_Base.Count + title_draft.Count + 3, True)
        Call SummaryAdas(title_Base.Count + title_draft.Count + title_draft.Count + 4, title_Base.Count + title_draft.Count + 3)
    Else
        Call CompareAdas("”äŠr: ‘O‰ñ‚Æ¡‰ñ", pivotBase, pivotDraft, title_Base.Count + title_draft.Count + title_Msg.Count + 4, True)
        Call CompareAdas("”äŠr: ¡‰ñ‚ÆMsg", pivotDraft, pivotMsg, title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + 5, False)
        Call SummaryAdas(title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + title_Msg.Count + 6, _
            title_Base.Count + title_draft.Count + title_Msg.Count + 4, title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + 5)
    End If


'   network
    Set wsBase = wbBase.Worksheets("Network Path")
    Set wsDraft = wbDraft.Worksheets("Network Path")
    If flag_Msg Then Set wsMsg = wbMsg.Worksheets("Network Path")
    Set wsOut = wbOut.Worksheets.Add
    wsOut.name = "Network Path"

    Call Refresh
    Call DeleteEmptyRow
    Call ConfirmTitle
    Call LoadTitle
    Call CreatKeyword
    Call DeleteData
    Call SortData2
    Call Copy

    If Not flag_Msg Then
        Call CompareAdas("”äŠr: ‘O‰ñ‚Æ¡‰ñ", pivotBase, pivotDraft, title_Base.Count + title_draft.Count + 3, True)
        Call SummaryAdas(title_Base.Count + title_draft.Count + title_draft.Count + 4, title_Base.Count + title_draft.Count + 3)
    Else
        Call CompareAdas("”äŠr: ‘O‰ñ‚Æ¡‰ñ", pivotBase, pivotDraft, title_Base.Count + title_draft.Count + title_Msg.Count + 4, True)
        Call CompareAdas("”äŠr: ¡‰ñ‚ÆMsg", pivotDraft, pivotMsg, title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + 5, False)
        Call SummaryAdas(title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + title_Msg.Count + 6, _
            title_Base.Count + title_draft.Count + title_Msg.Count + 4, title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + 5)
    End If


'    construction
    Set wsBase = wbBase.Worksheets("Construction of Container frame")
    Set wsDraft = wbDraft.Worksheets("Construction of Container frame")
    If flag_Msg Then Set wsMsg = wbMsg.Worksheets("Construction of Container frame")
    Set wsOut = wbOut.Worksheets.Add
    wsOut.name = "Construction of Container frame"

    Call Refresh
    Call DeleteEmptyRow
    Call ConfirmTitle
    Call LoadTitle
    Call CreatKeyword
    Call DeleteData
    Call SortData2
    Call Copy

    If Not flag_Msg Then
        Call CompareAdas("”äŠr: ‘O‰ñ‚Æ¡‰ñ", pivotBase, pivotDraft, title_Base.Count + title_draft.Count + 3, True)
        Call SummaryAdas(title_Base.Count + title_draft.Count + title_draft.Count + 4, title_Base.Count + title_draft.Count + 3)
    Else
        Call CompareAdas("”äŠr: ‘O‰ñ‚Æ¡‰ñ", pivotBase, pivotDraft, title_Base.Count + title_draft.Count + title_Msg.Count + 4, True)
        Call CompareAdas("”äŠr: ¡‰ñ‚ÆMsg", pivotDraft, pivotMsg, title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + 5, False)
        Call SummaryAdas(title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + title_Msg.Count + 6, _
            title_Base.Count + title_draft.Count + title_Msg.Count + 4, title_Base.Count + title_draft.Count + title_Msg.Count + title_draft.Count + 5)
    End If
    
    


    wbOut.SaveAs ThisWorkbook.path & "\CompareAdas " & Left(wbBase.name, Len(wbBase.name) - 5) & " vs " & Left(wbDraft.name, Len(wbDraft.name) - 5) & ".xlsx"
    wbBase.Close SaveChanges:=False
    wbDraft.Close SaveChanges:=False
    If flag_Msg Then wbMsg.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    'MsgBox Timer() - timePoint
    MsgBox "Done."
End Sub




