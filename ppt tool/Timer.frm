VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Timer 
   Caption         =   "Timer"
   ClientHeight    =   1320
   ClientLeft      =   15
   ClientTop       =   330
   ClientWidth     =   1995
   OleObjectBlob   =   "Timer.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Copyright @2025-2035 shuilin, All Rights Reserved.
' Email:www.github.com/shuilin123/
' Date 2025.4.14
' 用户窗体代码（在frmCountdown中）
Private Sub UserForm_Initialize()
    Me.Caption = "倒计时"
    Me.lblTime.Font.Size = 30
    Me.lblTime.Caption = "05:00"
    '设置显示位置
    Me.StartUpPosition = 0
    Me.Top = Application.Top
    Me.Left = Application.Left
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 防止用户手动关闭窗体
    If CloseMode <> vbFormCode Then
        Cancel = True
    End If
End Sub
