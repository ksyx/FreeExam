VERSION 5.00
Begin VB.Form Preview 
   BackColor       =   &H00A0ACBA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preview Window"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13605
   ControlBox      =   0   'False
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
   ScaleHeight     =   7065
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   7590
      Top             =   900
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15
   End
   Begin VB.PictureBox Exporter 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   5400
      ScaleHeight     =   2325
      ScaleWidth      =   6420
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1635
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A0ACBA&
      Caption         =   "Advanced View"
      Height          =   825
      Left            =   9765
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   1785
      Begin VB.TextBox Text1 
         Height          =   240
         Left            =   0
         TabIndex        =   13
         Text            =   "0"
         Top             =   0
         Width           =   480
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Text            =   "0"
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label1 
         Height          =   225
         Left            =   960
         TabIndex        =   16
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label2 
         Height          =   225
         Left            =   1005
         TabIndex        =   15
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label3 
         Caption         =   "Up"
         Height          =   210
         Left            =   0
         TabIndex        =   14
         Top             =   195
         Width           =   315
      End
   End
   Begin VB.PictureBox Exports 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   1080
      ScaleHeight     =   2325
      ScaleWidth      =   6420
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3900
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.PictureBox Export 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   7350
      ScaleHeight     =   2325
      ScaleWidth      =   6420
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1740
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
      BackColor       =   &H00656D76&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   600
      ScaleHeight     =   225
      ScaleWidth      =   6450
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   15
      Width           =   6450
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
      Left            =   3705
      List            =   "Preview.frx":000A
      TabIndex        =   4
      Top             =   435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7170
      Top             =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Height          =   2820
      Left            =   -15
      ScaleHeight     =   2820
      ScaleWidth      =   6885
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   255
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
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "You are not in a normal view, click me to get back to the normal view."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   480
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   5010
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2940
      LargeChange     =   100
      Left            =   30
      Max             =   10000
      SmallChange     =   10
      TabIndex        =   1
      Top             =   2220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   4335
      Max             =   100
      TabIndex        =   0
      Top             =   6975
      Visible         =   0   'False
      Width           =   6915
   End
   Begin VB.Label Label63 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00B4BFCC&
      Caption         =   " Close "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim showcnt As Long, current As Long, lasttop As Long
Public SystemCall As Long, WinStat As Long
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
        If Message.Caption <> "" Then Message.Caption = Message.Caption & translate("(Expired)")
        If ClearOnly Then Exit Sub
    End If
    MsgContentList.AddItem Content
    MsgColorList.AddItem Color
    Select Case Color
        Case vbBlack: MsgTypeList.AddItem translate("[Info]")
        Case vbBlue: MsgTypeList.AddItem translate("[Warning]")
        Case vbRed: MsgTypeList.AddItem translate("[Error]")
    End Select
    showcnt = 49
    Timer1_Timer
End Sub

Private Sub Form_Load()
    current = -1
    translatecontrol Me.Name
    NewMessage translate("The size of the preview is NEAR the actual size."), vbBlack
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If SystemCall <> SystemCallFlag And Me.Caption = translate("Preview Window - Rendering, please wait, you can't close this window while rendering") Then
        Cancel = 1
        NewMessage translate("Rending work in progress, you can't close it now!"), vbRed
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainFrm.WIP.Left = 99999
End Sub

Private Sub Label11_Click()
    Picture2.Top = 0
End Sub

Private Sub Label63_Click()
    Unload Me
End Sub

Private Sub Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub

Private Sub Text1_Change()
    If IsNumeric(Text1.Text) Then
        Picture2.Left = -Val(Text1.Text)
    End If
End Sub

Private Sub Text2_Change()
    If IsNumeric(Text2.Text) Then
    Picture2.Top = -Val(Text2.Text)
    End If
End Sub

Private Sub Text3_Change()
    Text3.Text = ""
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyD Then
        Picture2.Top = Picture2.Top - (Me.Height - 1500)
    End If
    If KeyCode = vbKeyU Then
        Picture2.Top = Picture2.Top + (Me.Height - 1500)
    End If
End Sub

Private Sub Timer1_Timer()
    Dim first As Long
    If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 100
    showcnt = showcnt + 1
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
        Message.Caption = translate("No new messages.")
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
            Message.Caption = translate("No new messages.")
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

Private Sub Form_Resize()
    On Error Resume Next
    HScroll1.Width = Me.Width - 567 * 5
    VScroll1.Height = Me.Height - 567 * 5
    Picture1.Width = Me.Width
    Picture1.Height = Me.Height
    Picture3.Width = Me.Width
End Sub

Private Sub HScroll1_Change()
    Picture2.Left = -(Picture2.Width / HScroll1.Value)
    Text1.Text = HScroll1.Value
End Sub

Private Sub Timer2_Timer()
    If Picture2.Top <> lasttop Then
        lasttop = Picture2.Top
        Me.Caption = translate("Preview Window - Rendering, please wait, you can't close this window while rendering")
        Label63.Visible = False
    Else
        Me.Caption = translate("Preview Window")
        Label63.Visible = True
    End If
End Sub

Private Sub VScroll1_Change()
    Dim r As Long
    r = PageHeight * PresetPageNumber / 100 * VScroll1.Value
    Picture2.Top = -r - PageHeight
    Debug.Print Picture2.Top
    Text2.Text = VScroll1.Value
End Sub
