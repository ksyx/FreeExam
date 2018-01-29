VERSION 5.00
Begin VB.Form DevWin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Developer's Window"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CheckBox Check1 
      Caption         =   "AutoCls"
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2145
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   -15
      ScaleHeight     =   225
      ScaleWidth      =   5670
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2685
      Width           =   5730
      Begin VB.Shape ProgressBar 
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   0
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Message 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.ListBox MsgContentList 
      Height          =   450
      ItemData        =   "DevWin.frx":0000
      Left            =   1635
      List            =   "DevWin.frx":0002
      TabIndex        =   3
      Top             =   2070
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgColorList 
      Height          =   450
      ItemData        =   "DevWin.frx":0004
      Left            =   1635
      List            =   "DevWin.frx":0006
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgTypeList 
      Height          =   450
      ItemData        =   "DevWin.frx":0008
      Left            =   2310
      List            =   "DevWin.frx":000A
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1095
      Top             =   1800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Open Window"
      Height          =   1275
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4350
      Begin VB.CommandButton Command2 
         Caption         =   "Main"
         Height          =   360
         Left            =   1290
         TabIndex        =   7
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PgSetting&s"
         Height          =   360
         Left            =   195
         TabIndex        =   6
         Top             =   255
         Width           =   990
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9000
      Top             =   990
      Width           =   210
   End
End
Attribute VB_Name = "DevWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim showcnt As Integer, current As Integer
Sub NewMessage(Content As String, Color As Long, Optional ClearList As Boolean = False, Optional ClearOnly = False)
    current = -1
    If (ClearOnly And Not ClearList) Then
        RaiseSysErr "Clear message list only and do not clear message list were both turned on.", "Create/PageSettings/NewEvent"
        Exit Sub
    End If
    If ClearList Then
        MsgContentList.Clear
        MsgColorList.Clear
        MsgTypeList.Clear
        If Message.Caption <> "" Then Message.Caption = Message.Caption & "(Expired)"
        If ClearOnly Then Exit Sub
    End If
    MsgContentList.AddItem Content
    MsgColorList.AddItem Color
    Select Case Color
        Case vbBlack: MsgTypeList.AddItem "[Info]"
        Case vbBlue: MsgTypeList.AddItem "[Warning]"
        Case vbRed: MsgTypeList.AddItem "[Error]"
    End Select
    showcnt = 49
    Timer1_Timer
End Sub

Private Sub Check1_Click()
    AutoCls = Check1.Value
    SaveSetting "FreeExam", "Create", "AutoCls", AutoCls
End Sub

Private Sub Command1_Click()
    PageSettings.Show
End Sub

Private Sub Command2_Click()
    MainFrm.Show
End Sub

Private Sub Form_Load()
    current = -1
    If Development <> 1 Then
       Shape1.Left = 0
       Shape1.Top = 0
       Shape1.Height = Me.Height
       Shape1.Width = Me.Width
       Me.Caption = "Contents can't be shown"
       RaiseSysErr "Access Denied - You don't have enough privilege to access here. By the way, there is nothing interesting.", "DevWin/PrivCheck"
    End If
    NewMessage "Authentication Passed.", vbBlack
    Check1.Value = AutoCls
End Sub

Private Sub Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
    Dim first As Integer
    If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 100
    showcnt = showcnt + 1
    If MsgContentList.ListCount = 0 Then
        Message.Caption = "No new messages."
        Message.ForeColor = vbWhite
        showcnt = ShowCntPerMsg - 1
        GoTo rrr
    End If
'    If MsgContentList.ListCount <= 1 Then
'        first = showcnt
'        showcnt = ShowCntPerMsg
'        Message.Caption = ""
'        If MsgContentList.ListCount = 1 Then
'            current = 0
'            MsgContentList.ListIndex = current
'            MsgColorList.ListIndex = current
'            MsgTypeList.ListIndex = current
'            Message.Caption = MsgTypeList.Text & MsgContentList.Text
'            Message.ForeColor = ReverseColor(MsgColorList.Text)
'        End If
'        If showcnt <> first Then ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
'        Exit Sub
'    End If
    If current >= MsgContentList.ListCount Then
        Message.Caption = "No new messages."
        Message.ForeColor = vbWhite
        showcnt = ShowCntPerMsg - 1
        GoTo rrr
    End If
    If showcnt = ShowCntPerMsg Then
        current = current + 1
        showcnt = 0
        If MsgContentList.ListCount = 0 Then
            ProgressBar.Width = 15
            Message.Caption = ""
            Exit Sub
        End If
        If current >= MsgContentList.ListCount Then
            Message.Caption = "No new messages."
            Message.ForeColor = vbWhite
            showcnt = ShowCntPerMsg - 1
            GoTo rrr
        End If
        MsgContentList.ListIndex = current
        MsgColorList.ListIndex = current
        MsgTypeList.ListIndex = current
        Message.Caption = MsgTypeList.Text & MsgContentList.Text
        Message.ForeColor = ReverseColor(MsgColorList.Text)
rrr:
    End If
    ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
End Sub

Private Sub Label3_Click()
End Sub
