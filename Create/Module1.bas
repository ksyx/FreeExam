Attribute VB_Name = "Kernel"
Option Explicit
'===========COLORS===========
'INFO &H00FFFFFF& (SENDPAR=WHITE)
'WARNING &H0000FFFF& (SENDPAR=BLUE)
'ERROR &H00FFFF00& (SENDPAR=RED)

Public Const ShowCntPerMsg As Integer = 50
Public Const TwipsPerCM As Integer = 567
Public Const Development As Integer = 1

Public PageWidth As Long, PageHeight As Long, TopMargin As Long, BotMargin As Long, LeftMargin As Long, RightMargin As Long

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

Sub InitPreview()
    Preview.Picture2.Height = PageHeight
    Preview.Picture2.Width = PageWidth
    Preview.HScroll1.Max = PageWidth
    Preview.VScroll1.Max = PageHeight
    Preview.Show
End Sub
