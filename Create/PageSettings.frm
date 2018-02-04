VERSION 5.00
Begin VB.Form PageSettings 
   BackColor       =   &H00A0ACBA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PageSettings"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
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
   ScaleHeight     =   5745
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   4830
   End
   Begin VB.ListBox MsgTypeList 
      Height          =   450
      Left            =   2310
      TabIndex        =   9
      Top             =   4875
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgColorList 
      Height          =   450
      Left            =   1635
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgContentList 
      Height          =   450
      Left            =   1650
      TabIndex        =   7
      Top             =   4815
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00656D76&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   7035
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5430
      Width           =   7095
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
         TabIndex        =   6
         Top             =   15
         Width           =   45
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00A0ACBA&
      Caption         =   "Margins"
      ForeColor       =   &H00656D76&
      Height          =   3225
      Left            =   135
      TabIndex        =   3
      Top             =   1545
      Width           =   6705
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   375
         Index           =   3
         Left            =   1095
         TabIndex        =   18
         Text            =   "3.17"
         Top             =   1710
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   375
         Index           =   2
         Left            =   1095
         TabIndex        =   17
         Text            =   "3.17"
         Top             =   1365
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   375
         Index           =   1
         Left            =   1095
         TabIndex        =   16
         Text            =   "3.17"
         Top             =   1005
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   375
         Index           =   0
         Left            =   1095
         TabIndex        =   15
         Text            =   "3.17"
         Top             =   645
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   3315
         TabIndex        =   22
         Top             =   1725
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   3300
         TabIndex        =   21
         Top             =   1380
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   3300
         TabIndex        =   20
         Top             =   1005
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   3315
         TabIndex        =   19
         Top             =   645
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   525
         TabIndex        =   14
         Top             =   615
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   375
         TabIndex        =   13
         Top             =   1725
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   540
         TabIndex        =   12
         Top             =   1365
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
         Height          =   360
         Left            =   60
         TabIndex        =   11
         Top             =   990
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A0ACBA&
      Caption         =   "PageType"
      ForeColor       =   &H00656D76&
      Height          =   1455
      Left            =   135
      TabIndex        =   0
      Top             =   15
      Width           =   6735
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00B4BFCC&
         ForeColor       =   &H00656D76&
         Height          =   315
         ItemData        =   "PageSettings.frx":0000
         Left            =   120
         List            =   "PageSettings.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   255
         Width           =   6435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"PageSettings.frx":004D
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
         Height          =   615
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   6345
      End
   End
   Begin VB.Label PreviewButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00B4BFCC&
      Caption         =   " Preview(&P) "
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
      Left            =   4905
      TabIndex        =   10
      Top             =   4905
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00B4BFCC&
      Caption         =   " OK(&S) "
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
      Left            =   6135
      TabIndex        =   4
      Top             =   4905
      Width           =   660
   End
End
Attribute VB_Name = "PageSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim showcnt As Long, current As Long

Private Sub Label2_Click()
    NewMessage "", vbRed, True, True
    Select Case Left(Combo1.Text, 5)
        Case "A4(21"
            PageWidth = 21 * TwipsPerCM
            PageHeight = 29.7 * TwipsPerCM
        Case "8K[1]"
            PageWidth = 18.5 * TwipsPerCM
            PageHeight = 26 * TwipsPerCM
        Case "8K[2]"
            PageWidth = 21 * TwipsPerCM
            PageHeight = 28.5 * TwipsPerCM
        Case Else
            NewMessage "Unable to get the page size that you've chosen.", vbRed
            Exit Sub
    End Select
    If Not IsNumeric(Text1(0).Text) Or Not IsNumeric(Text1(1).Text) Or Not IsNumeric(Text1(2).Text) Or Not IsNumeric(Text1(3).Text) Then
        NewMessage "The margin that you've inputed is invaild", vbRed
        Exit Sub
    End If
    TopMargin = Val(Text1(0).Text) * TwipsPerCM
    BotMargin = Val(Text1(1).Text) * TwipsPerCM
    LeftMargin = Val(Text1(2).Text) * TwipsPerCM
    RightMargin = Val(Text1(3).Text) * TwipsPerCM
'    InitPreview
'    Preview.Picture2.Line (LeftMargin, TopMargin)-(LeftMargin, Preview.Picture2.Height - BotMargin)
'    Preview.Picture2.Line (Preview.Picture2.Width - RightMargin, TopMargin)-(Preview.Picture2.Width - RightMargin, Preview.Picture2.Height - BotMargin)
'    Preview.Picture2.Line (LeftMargin, TopMargin)-(Preview.Picture2.Width - RightMargin, TopMargin)
'    Preview.Picture2.Line (LeftMargin, Preview.Picture2.Height - BotMargin)-(Preview.Picture2.Width - RightMargin, Preview.Picture2.Height - BotMargin)
'    Preview.NewMessage "Available area to edit is the area in the rectangle.", vbBlack
    TopMargin = Val(Text1(0).Text) * TwipsPerCM
    BotMargin = PageHeight - Val(Text1(1).Text) * TwipsPerCM
    LeftMargin = Val(Text1(2).Text) * TwipsPerCM
    RightMargin = PageWidth - Val(Text1(3).Text) * TwipsPerCM
    Unload Me
End Sub

Private Sub Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub
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
End Sub

Private Sub PreviewButton_Click()
    NewMessage "", vbRed, True, True
    Select Case Left(Combo1.Text, 5)
        Case "A4(21"
            PageWidth = 21 * TwipsPerCM
            PageHeight = 29.7 * TwipsPerCM
        Case "8K[1]"
            PageWidth = 18.5 * TwipsPerCM
            PageHeight = 26 * TwipsPerCM
        Case "8K[2]"
            PageWidth = 21 * TwipsPerCM
            PageHeight = 28.5 * TwipsPerCM
        Case Else
            NewMessage "Unable to get the page size that you've chosen.", vbRed
            Exit Sub
    End Select
    If Not IsNumeric(Text1(0).Text) Or Not IsNumeric(Text1(1).Text) Or Not IsNumeric(Text1(2).Text) Or Not IsNumeric(Text1(3).Text) Then
        NewMessage "The margin that you've inputed is invaild", vbRed
        Exit Sub
    End If
    TopMargin = Val(Text1(0).Text) * TwipsPerCM
    BotMargin = Val(Text1(1).Text) * TwipsPerCM
    LeftMargin = Val(Text1(2).Text) * TwipsPerCM
    RightMargin = Val(Text1(3).Text) * TwipsPerCM
    InitPreview
    Preview.Picture2.Line (LeftMargin, TopMargin)-(LeftMargin, Preview.Picture2.Height - BotMargin)
    Preview.Picture2.Line (Preview.Picture2.Width - RightMargin, TopMargin)-(Preview.Picture2.Width - RightMargin, Preview.Picture2.Height - BotMargin)
    Preview.Picture2.Line (LeftMargin, TopMargin)-(Preview.Picture2.Width - RightMargin, TopMargin)
    Preview.Picture2.Line (LeftMargin, Preview.Picture2.Height - BotMargin)-(Preview.Picture2.Width - RightMargin, Preview.Picture2.Height - BotMargin)
    Preview.NewMessage "Available area to edit is the area in the rectangle.", vbBlack
    TopMargin = Val(Text1(0).Text) * TwipsPerCM
    BotMargin = PageHeight - Val(Text1(1).Text) * TwipsPerCM
    LeftMargin = Val(Text1(2).Text) * TwipsPerCM
    RightMargin = PageWidth - Val(Text1(3).Text) * TwipsPerCM
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
    If MsgContentList.ListCount = 0 Then
        Message.Caption = "No new messages."
        Message.ForeColor = vbWhite
        showcnt = ShowCntPerMsg - 1
        GoTo rrr
    End If
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
