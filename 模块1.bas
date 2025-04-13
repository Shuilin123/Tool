Attribute VB_Name = "ģ��1"
' Copyright @2025-2035 Zhuo Li, All Rights Reserved.
' Email:9031003831@qq.com
' Date 2025.4.14

' ��׼ģ���еĴ��루�� Module1��
Option Explicit
' 64λ���ݵ�API����

' 32λ���������ݾɰ�Office��
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
' ȫ�ֱ���
Private mTimerID As Long
Public RemainingTime As Long
Public TimeForm As UserForm1 ' ���洰��ʵ��������

' �ص�����
Public Sub TimerProc(ByVal hWnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal idEvent As Long, _
                     ByVal dwTime As Long)
    RemainingTime = RemainingTime - 1
    ' ���´�����ʾ
    On Error Resume Next
    If Not TimeForm Is Nothing Then
        TimeForm.lblTime.Caption = Format(RemainingTime \ 60, "00") & ":" & _
                                  Format(RemainingTime Mod 60, "00")
        DoEvents  ' ����������
    End If
        '��ʱppt�Ѿ��ر��� �رն�ʱ��
    If ActivePresentation.SlideShowWindow Is Nothing Then
         #If VBA7 Then
            KillTimer 0, mTimerID
        #Else
            KillTimer 0, mTimerID
        #End If
         ' �رմ��岢�ͷ�����
        Unload TimeForm
        Set TimeForm = Nothing
    End If
    ' ���60����ʾ
    If RemainingTime = 60 Then
             sndPlaySound "C:\Windows\Media\Alarm01.wav", SND_ASYNC Or SND_FILENAME
    End If
     ' ����ʱ��������
    If RemainingTime <= 0 Then
          ' �˳��õ�Ƭ��ӳ
         #If VBA7 Then
            KillTimer 0, mTimerID
        #Else
            KillTimer 0, mTimerID
        #End If
         ' �رմ��岢�ͷ�����
        If Not TimeForm Is Nothing Then
             Unload TimeForm
             Set TimeForm = Nothing
        End If
        If Not ActivePresentation.SlideShowWindow Is Nothing Then
            MsgBox "ʱ�䵽��", vbInformation
            ActivePresentation.SlideShowWindow.View.Exit
        End If
    End If
    Exit Sub
ErrorHandler:
    ' ��������細������رգ�
    If Err.Number = 91 Then  ' �������δ����
        Set TimeForm = Nothing
    End If
End Sub

' ��������ʱ
Public Sub StartCountdown()
    RemainingTime = 90
    ' ��ʾ����
    If TimeForm Is Nothing Then
        Set TimeForm = New UserForm1
        TimeForm.Show vbModeless  ' ��ģ̬��ʾ
    End If
    ' ������ʱ��
    #If VBA7 Then
        mTimerID = SetTimer(0, 0, 1000, AddressOf TimerProc)
    #Else
        mTimerID = SetTimer(0, 0, 1000, AddressOf TimerProc)
    #End If
End Sub
