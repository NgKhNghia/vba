VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrf_cau_hoi 
   Caption         =   "Question"
   ClientHeight    =   8860.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   13970
   OleObjectBlob   =   "usrf_cau_hoi.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "usrf_cau_hoi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub check_chinese_Click()
    If LCase(Trim(chinese)) = LCase(Trim(usrf_cau_hoi.txt_chinese)) Then
        usrf_cau_hoi.txt_chinese.BackColor = RGB(0, 255, 0)
    Else
        usrf_cau_hoi.txt_chinese.BackColor = RGB(255, 0, 0)
        usrf_cau_hoi.lbl_dap_an_chinese.Caption = chinese
    End If
End Sub

Private Sub check_read_Click()
    If LCase(Trim(read)) = LCase(Trim(usrf_cau_hoi.txt_read)) Then
        usrf_cau_hoi.txt_read.BackColor = RGB(0, 255, 0)
    Else
        usrf_cau_hoi.txt_read.BackColor = RGB(255, 0, 0)
        usrf_cau_hoi.lbl_dap_an_read.Caption = read
    End If
End Sub

Private Sub cmd_close_Click()
    Unload Me
    Unload usrf_man_hinh_chinh
End Sub

Private Sub cmd_home_Click()
    Me.Hide
    usrf_man_hinh_chinh.Show
End Sub

Private Sub cmd_meaning_1_Click()
    If usrf_cau_hoi.cmd_meaning_1.Caption = meaning Then
        usrf_cau_hoi.cmd_meaning_1.BackColor = RGB(0, 255, 0)
    Else
        usrf_cau_hoi.cmd_meaning_1.BackColor = RGB(255, 0, 0)
    End If
    
    usrf_cau_hoi.cmd_meaning_2.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_3.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_4.BackColor = RGB(240, 240, 240)
End Sub

Private Sub cmd_meaning_2_Click()
    If usrf_cau_hoi.cmd_meaning_2.Caption = meaning Then
        usrf_cau_hoi.cmd_meaning_2.BackColor = RGB(0, 255, 0)
    Else
        usrf_cau_hoi.cmd_meaning_2.BackColor = RGB(255, 0, 0)
    End If
    
    usrf_cau_hoi.cmd_meaning_1.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_3.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_4.BackColor = RGB(240, 240, 240)
End Sub

Private Sub cmd_meaning_3_Click()
    If usrf_cau_hoi.cmd_meaning_3.Caption = meaning Then
        usrf_cau_hoi.cmd_meaning_3.BackColor = RGB(0, 255, 0)
    Else
        usrf_cau_hoi.cmd_meaning_3.BackColor = RGB(255, 0, 0)
    End If
    
    usrf_cau_hoi.cmd_meaning_1.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_2.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_4.BackColor = RGB(240, 240, 240)
End Sub

Private Sub cmd_meaning_4_Click()
    If usrf_cau_hoi.cmd_meaning_4.Caption = meaning Then
        usrf_cau_hoi.cmd_meaning_4.BackColor = RGB(0, 255, 0)
    Else
        usrf_cau_hoi.cmd_meaning_4.BackColor = RGB(255, 0, 0)
    End If
    
    usrf_cau_hoi.cmd_meaning_1.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_2.BackColor = RGB(240, 240, 240)
    usrf_cau_hoi.cmd_meaning_3.BackColor = RGB(240, 240, 240)
End Sub

Private Sub cmd_next_Click()
    Call Tao_Cau_Hoi
End Sub

