VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1323
   ClientLeft      =   21
   ClientTop       =   336
   ClientWidth     =   1988
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright @2025-2035 Zhuo Li, All Rights Reserved.
' Email:9031003831@qq.com
' Date 2025.4.14
' 用户窗体代码（在frmCountdown中）
Private Sub UserForm_Initialize()
    Me.Caption = "倒计时"
    Me.lblTime.Font.Size = 30
    Me.lblTime.Caption = "05:00"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 防止用户手动关闭窗体
    If CloseMode <> vbFormCode Then
        Cancel = True
    End If
End Sub
