Attribute VB_Name = "Module1"
Option Compare Text

Public tbl As ListObject
Public idx As Long
Public word As String
Public read As String
Public chinese As String
Public meaning As String
Public answers(1 To 4) As String

Sub Tao_Cau_Hoi()
'    chon data base
    Set tbl = ThisWorkbook.Worksheets("kanji").ListObjects("table_kanji")

'    random
    Randomize
    idx = Int(tbl.ListRows.Count * Rnd) + 1
    word = tbl.ListColumns("Word").DataBodyRange.Cells(idx).Value
    read = tbl.ListColumns("Read").DataBodyRange.Cells(idx).Value
    chinese = tbl.ListColumns("Chinese").DataBodyRange.Cells(idx).Value
    meaning = tbl.ListColumns("Meaning").DataBodyRange.Cells(idx).Value
    
'    reset user form
    usrf_cau_hoi.lbl_word.Caption = word
    usrf_cau_hoi.txt_read = ""
    usrf_cau_hoi.txt_chinese = ""
    
    usrf_cau_hoi.lbl_dap_an_read.Caption = ""
    usrf_cau_hoi.lbl_dap_an_chinese.Caption = ""
    
    usrf_cau_hoi.txt_read.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.txt_chinese.BackColor = RGB(240, 240, 240)
    
    usrf_cau_hoi.cmd_meaning_1.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_2.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_3.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_4.BackColor = RGB(240, 240, 240)
    
'    tao 4 dap an dap an
    Dim i As Integer
    Dim j As Integer
    Dim r As Integer
    Dim used As Object
    Dim tmp As String
    
    Set used = CreateObject("Scripting.Dictionary")
    answers(1) = meaning
    used.Add answers(1), True
    i = 2
    Do While i <= 4
        r = Int(tbl.ListRows.Count * Rnd) + 1
        tmp = tbl.ListColumns("Meaning").DataBodyRange.Cells(r).Value
        If tmp = "" Then
            Debug.Print r; Rnd; tmp
        End If
        
'        kiem tra xem tmp da ton tai trong answers chua
        For j = 1 To i - 1
            If tmp = answers(j) Then
                Exit For
            End If
        Next j
            
'        neu chua ton tai thi them tmp lam answer
        If j >= i Then
            answers(i) = tmp
            i = i + 1
        End If
        
    Loop
    
'    tron 4 dap an
    For i = 4 To 2 Step -1
        j = Int(i * Rnd) + 1
        
        tmp = answers(i)
        answers(i) = answers(j)
        answers(j) = tmp
    Next i
    
'    chuyen dap an vao form
    usrf_cau_hoi.cmd_meaning_1.Caption = answers(1)
    usrf_cau_hoi.cmd_meaning_2.Caption = answers(2)
    usrf_cau_hoi.cmd_meaning_3.Caption = answers(3)
    usrf_cau_hoi.cmd_meaning_4.Caption = answers(4)
End Sub

Sub main()
    usrf_man_hinh_chinh.Show
End Sub
