Attribute VB_Name = "Kernel"
Option Explicit
'===========COLORS===========
'INFO &H00FFFFFF& (SENDPAR=WHITE)
'WARNING &H0000FFFF& (SENDPAR=BLUE)
'ERROR &H00FFFF00& (SENDPAR=RED)

Public Const ShowCntPerMsg As Long = 50
Public Const TwipsPerCM As Long = 567
Public Const Development As Long = 1
Public Const TitleHi As Long = 495
Public Const DefCnt As Long = 1

Public PageWidth As Long, PageHeight As Long, TopMargin As Long, BotMargin As Long, LeftMargin As Long, RightMargin As Long, AutoCls As Long

Function ReverseColor(Color As Long) As Long
    If Color = vbRed Then ReverseColor = RGB(15, 255, 255)
    If Color = vbBlue Then ReverseColor = RGB(255, 255, 15)
    If Color = vbBlack Then ReverseColor = vbWhite
End Function

Sub RaiseSysErr(Detail As String, Module As String)
    SystemError.ErrDetail.Caption = "An error occured, some operations won't be excuted. Please Report the following contents to us :)" & vbCrLf & "Module: " & Module & vbCrLf & "Details:" & vbCrLf & Detail & vbCrLf & "Time: " & Now & vbCrLf & vbCrLf & "Double click to close the window after 10 seconds. Press PrtSc to take a capture (THIS OPERATION WILL COVER YOUR CLIPBOARD). "
    SystemError.CurrentTime.Caption = Now
    SystemError.Show 1
End Sub

Function InputText(StartText As String, Optional MultiLine As Boolean = True) As String
    InputWin.Text2.Text = StartText
    If MultiLine = False Then
        InputWin.Text1.Visible = True
        InputWin.Text2.Visible = False
        InputWin.Text1.Text = StartText
    End If
    InputWin.Show 1
    If MultiLine = False Then InputWin.Text2.Text = InputWin.Text1.Text
    If InputWin.Caption = "UserCancel" Then InputText = StartText Else InputText = InputWin.Text2.Text
    Unload InputWin
End Function

Sub InitPreview()
    If AutoCls = 1 Then Preview.Picture2.Cls
    Preview.Picture2.Height = PageHeight
    Preview.Picture2.Width = PageWidth
'    Preview.HScroll1.Max = PageWidth
 '   Preview.VScroll1.Max = PageHeight
    Preview.Show
    MainFrm.WIP.Left = 0
End Sub

Sub Main()
    AutoCls = GetSetting("FreeExam", "Create", "AutoCls", 1)
    DevWin.Show
End Sub

Function Max(a, b)
    If a > b Then Max = a Else Max = b
End Function
