VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrf_man_hinh_chinh 
   Caption         =   " Home"
   ClientHeight    =   5620
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9820.001
   OleObjectBlob   =   "usrf_man_hinh_chinh.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "usrf_man_hinh_chinh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_start_Click()
    Me.Hide
    Call Tao_Cau_Hoi
    usrf_cau_hoi.Show
End Sub

