VERSION 5.00
Begin VB.Form Preview 
   Caption         =   "Preview Window"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7290
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
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7290
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Export 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2325
      Left            =   675
      ScaleHeight     =   2265
      ScaleWidth      =   6360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   855
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.ListBox MsgColorList 
      Height          =   450
      ItemData        =   "Preview.frx":0000
      Left            =   2160
      List            =   "Preview.frx":0002
      TabIndex        =   8
      Top             =   870
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgTypeList 
      Height          =   450
      ItemData        =   "Preview.frx":0004
      Left            =   1785
      List            =   "Preview.frx":0006
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   -15
      ScaleHeight     =   225
      ScaleWidth      =   7185
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   7245
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
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.ListBox MsgContentList 
      Height          =   450
      ItemData        =   "Preview.frx":0008
      Left            =   3570
      List            =   "Preview.frx":000A
      TabIndex        =   4
      Top             =   315
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   -15
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2820
      Left            =   210
      ScaleHeight     =   2820
      ScaleWidth      =   6885
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   495
      Width           =   6885
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1965
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   4920
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   -15
         Width           =   4980
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2940
      Left            =   0
      TabIndex        =   1
      Top             =   495
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   6915
   End
End
Attribute VB_Name = "Preview"
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

Private Sub Form_Load()
    current = -1
    NewMessage "The size of the preview is NEAR the actual size.", vbBlack
End Sub

Private Sub Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub
Private Sub Timer1_Timer()
    If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 100
    showcnt = showcnt + 1
    If MsgContentList.ListCount <= 1 Then
        showcnt = ShowCntPerMsg
        If MsgContentList.ListCount = 1 Then
            current = 0
            MsgContentList.ListIndex = current
            MsgColorList.ListIndex = current
            MsgTypeList.ListIndex = current
            Message.Caption = MsgTypeList.Text & MsgContentList.Text
            Message.ForeColor = ReverseColor(MsgColorList.Text)
        End If
        ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
        Exit Sub
    End If
    If showcnt = ShowCntPerMsg Then
        current = current + 1
        showcnt = 0
        If MsgContentList.ListCount = 0 Then
            ProgressBar.Width = 15
            Message.Caption = ""
            Exit Sub
        End If
        If current >= MsgContentList.ListCount Then current = 0
        MsgContentList.ListIndex = current
        MsgColorList.ListIndex = current
        MsgTypeList.ListIndex = current
        Message.Caption = MsgTypeList.Text & MsgContentList.Text
        Message.ForeColor = ReverseColor(MsgColorList.Text)
rrr:
    End If
    ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    HScroll1.Width = Me.Width - 567 * 5
    VScroll1.Height = Me.Height - 567 * 5
    Picture1.Width = Me.Width
    Picture1.Height = Me.Height
    Picture3.Width = Me.Width
End Sub

Private Sub HScroll1_Change()
    Picture2.Left = -HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    Picture2.Top = -VScroll1.Value
End Sub
