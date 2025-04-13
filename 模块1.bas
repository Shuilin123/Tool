Attribute VB_Name = "模块1"
' Copyright @2025-2035 Zhuo Li, All Rights Reserved.
' Email:9031003831@qq.com
' Date 2025.4.14

' 标准模块中的代码（如 Module1）
Option Explicit
' 64位兼容的API声明

' 32位声明（兼容旧版Office）
Private Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerProc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" _
    Alias "sndPlaySoundA" ( _
    ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long


Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000
' 全局变量
Private mTimerID As Long
Public RemainingTime As Long
Public TimeForm As UserForm1 ' 保存窗体实例的引用

' 回调函数
Public Sub TimerProc(ByVal hWnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal idEvent As Long, _
                     ByVal dwTime As Long)
    RemainingTime = RemainingTime - 1
    ' 更新窗体显示
    On Error Resume Next
    If Not TimeForm Is Nothing Then
        TimeForm.lblTime.Caption = Format(RemainingTime \ 60, "00") & ":" & _
                                  Format(RemainingTime Mod 60, "00")
        DoEvents  ' 允许界面更新
    End If
        '此时ppt已经关闭了 关闭定时器
    If ActivePresentation.SlideShowWindow Is Nothing Then
         #If VBA7 Then
            KillTimer 0, mTimerID
        #Else
            KillTimer 0, mTimerID
        #End If
         ' 关闭窗体并释放引用
        Unload TimeForm
        Set TimeForm = Nothing
    End If
    ' 最后60秒提示
    If RemainingTime = 60 Then
             sndPlaySound "C:\Windows\Media\Alarm01.wav", SND_ASYNC Or SND_FILENAME
    End If
     ' 倒计时结束处理
    If RemainingTime <= 0 Then
          ' 退出幻灯片放映
         #If VBA7 Then
            KillTimer 0, mTimerID
        #Else
            KillTimer 0, mTimerID
        #End If
         ' 关闭窗体并释放引用
        If Not TimeForm Is Nothing Then
             Unload TimeForm
             Set TimeForm = Nothing
        End If
        If Not ActivePresentation.SlideShowWindow Is Nothing Then
            MsgBox "时间到！", vbInformation
            ActivePresentation.SlideShowWindow.View.Exit
        End If
    End If
    Exit Sub
ErrorHandler:
    ' 处理错误（如窗体意外关闭）
    If Err.Number = 91 Then  ' 对象变量未设置
        Set TimeForm = Nothing
    End If
End Sub

' 启动倒计时
Public Sub StartCountdown()
    RemainingTime = 90
    ' 显示窗体
    If TimeForm Is Nothing Then
        Set TimeForm = New UserForm1
        TimeForm.Show vbModeless  ' 非模态显示
    End If
    ' 启动定时器
    #If VBA7 Then
        mTimerID = SetTimer(0, 0, 1000, AddressOf TimerProc)
    #Else
        mTimerID = SetTimer(0, 0, 1000, AddressOf TimerProc)
    #End If
End Sub
