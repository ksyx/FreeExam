VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ExamPaper Editor"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
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
   ScaleHeight     =   5595
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   45
      TabIndex        =   29
      Top             =   1935
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.PictureBox Manage 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   60
      ScaleHeight     =   1380
      ScaleWidth      =   7770
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   495
      Visible         =   0   'False
      Width           =   7770
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MergePreview "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   32
         Top             =   135
         Width           =   1410
      End
   End
   Begin VB.ListBox MsgContentList 
      Height          =   450
      ItemData        =   "Main.frx":0000
      Left            =   4755
      List            =   "Main.frx":0002
      TabIndex        =   21
      Top             =   4410
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1110
      Top             =   1800
   End
   Begin VB.ListBox MsgTypeList 
      Height          =   450
      ItemData        =   "Main.frx":0004
      Left            =   3330
      List            =   "Main.frx":0006
      TabIndex        =   19
      Top             =   -15
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgColorList 
      Height          =   450
      ItemData        =   "Main.frx":0008
      Left            =   4785
      List            =   "Main.frx":000A
      TabIndex        =   18
      Top             =   105
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   45
      ScaleHeight     =   225
      ScaleWidth      =   7725
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5310
      Width           =   7785
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
         TabIndex        =   17
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format"
      Height          =   1545
      Left            =   135
      TabIndex        =   6
      Top             =   1995
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   825
         MaxLength       =   3
         TabIndex        =   15
         Top             =   570
         Width           =   1890
      End
      Begin VB.ComboBox FontCombo 
         Height          =   315
         ItemData        =   "Main.frx":000C
         Left            =   825
         List            =   "Main.frx":0019
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   225
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   270
      End
      Begin VB.CheckBox Check2 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   855
         Width           =   270
      End
      Begin VB.ComboBox AlignCombo 
         Height          =   315
         ItemData        =   "Main.frx":004E
         Left            =   810
         List            =   "Main.frx":005B
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   195
         Left            =   465
         TabIndex        =   14
         Top             =   585
         Width           =   285
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         Height          =   195
         Left            =   435
         TabIndex        =   12
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shape"
         Height          =   195
         Left            =   345
         TabIndex        =   9
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   1155
         Width           =   705
      End
   End
   Begin VB.PictureBox General 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   45
      ScaleHeight     =   1380
      ScaleWidth      =   7770
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   465
      Width           =   7770
      Begin VB.Label PreviewButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Text(&T) "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1485
      Left            =   3150
      TabIndex        =   20
      Top             =   2055
      Width           =   4620
      Begin VB.TextBox Text2 
         Height          =   1050
         Left            =   555
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   315
         Width           =   375
      End
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Preview(&T) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2265
      TabIndex        =   30
      Top             =   1965
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Remove(&D) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2265
      TabIndex        =   31
      Top             =   2340
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6195
      TabIndex        =   27
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Save(&S) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6315
      TabIndex        =   26
      Top             =   3780
      Width           =   870
   End
   Begin VB.Label Temp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      Height          =   195
      Left            =   735
      TabIndex        =   25
      Top             =   3765
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Preview(&T) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5175
      TabIndex        =   24
      Top             =   3780
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7095
      TabIndex        =   4
      Top             =   135
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubjectN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2130
      TabIndex        =   2
      Top             =   135
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   135
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   345
      Left            =   60
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str() As String, wholestr As String
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
    Dim i As Integer
    FontCombo.Clear
    Init.Show
    For i = 1 To Screen.FontCount
        Init.Message.Caption = "Loading Fonts(" & i & "/" & Screen.FontCount & ")"
        DoEvents
        FontCombo.AddItem Screen.Fonts(i)
    Next
    Unload Init
    Shape1.Left = Label1.Left
    Shape1.Width = Label1.Width
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


Private Sub Label1_Click()
    Shape1.Left = Label1.Left
    Shape1.Width = Label1.Width
    Frame1.Visible = Not False
    Frame2.Visible = Not False
    Label10.Visible = Not False
    Label11.Visible = Not False
    List1.Visible = Not True
    Label14.Visible = Not True
    Label13.Visible = Not True
    Manage.Visible = Not True
End Sub

Private Sub Label10_Click()
    On Error Resume Next
    Dim i As Integer
    Dim delta As Integer, bound As Integer, start As Integer
    wholestr = Text2.Text
    NewMessage "", vbGreen, True, True
    str = Split(Text2.Text, vbCrLf)
    bound = UBound(str)
    If FontCombo.Text = "" Or Not IsNumeric(Text1.Text) Or AlignCombo.Text = "" Then
        NewMessage "Invaild format.", vbRed, True
        Exit Sub
    End If
    For i = bound To 0 Step -1
        If str(i) <> "" Then Exit For Else bound = bound - 1
    Next
    For i = 0 To bound
        If str(i) <> "" Then Exit For Else start = i + 1
    Next
    If bound - start + 1 < 1 Or Text2.Text = "" Then
        NewMessage "Nothing can be previewed.", vbRed, True
        Exit Sub
    End If
    For i = start To bound Step 1
        Text2.Text = str(i)
        Temp.FontName = FontCombo.Text
        Temp.FontSize = Val(Text1.Text)
        Temp.Alignment = Val(Left(AlignCombo.Text, 1))
        Temp.Caption = Text2.Text
        InitPreview
        If Temp.Width > PageWidth - LeftMargin - (PageWidth - RightMargin) Or Temp.Height + delta > PageHeight - TopMargin - (PageHeight - BotMargin) Then Preview.NewMessage "The final image is out of page, something may be invisible.", vbBlue
        With Preview.Picture2
            .BorderStyle = 0
            .FontName = FontCombo.Text
            .FontSize = Val(Text1.Text)
            If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = RightMargin - Temp.Width Else .CurrentX = (LeftMargin + RightMargin) / 2 - Temp.Width / 2
            .CurrentY = TopMargin + delta
            Preview.Picture2.Print Text2.Text
        End With
        delta = delta + Temp.Height
    Next
'    wholestr = str(0)
'    If UBound(str) > 0 Then wholestr = wholestr & vbCrLf
'    For i = 1 To UBound(str)
'        wholestr = wholestr & str(i) & vbCrLf
'    Next
    Text2.Text = wholestr
End Sub

Private Sub Label11_Click()
    On Error Resume Next
    NewMessage "", vbGreen, True, True
    wholestr = Text2.Text
    Dim usage As Integer
    Dim i As Integer
    Dim delta As Integer, bound As Integer, start As Integer
    str = Split(Text2.Text, vbCrLf)
    bound = UBound(str)
    If FontCombo.Text = "" Or Not IsNumeric(Text1.Text) Or AlignCombo.Text = "" Then
        NewMessage "Invaild format.", vbRed, True
        Exit Sub
    End If
    For i = bound To 0 Step -1
        If str(i) <> "" Then Exit For Else bound = bound - 1
    Next
    For i = 0 To bound
        If str(i) <> "" Then Exit For Else start = i + 1
    Next
    If bound - start + 1 < 1 Or Text2.Text = "" Then
        NewMessage "Nothing can be saved.", vbRed, True
        Exit Sub
    End If
    For i = start To bound Step 1
        Text2.Text = str(i)
        Temp.FontName = FontCombo.Text
        Temp.FontSize = Val(Text1.Text)
        Temp.Alignment = Val(Left(AlignCombo.Text, 1))
        Temp.Caption = Text2.Text
        InitPreview
        If Temp.Width > PageWidth - LeftMargin - RightMargin Or Temp.Height > PageHeight - TopMargin - (PageHeight - BotMargin) Then Preview.NewMessage "The final image is out of page, something may be invisible.", vbBlue
        With Preview.Picture2
            .BorderStyle = 0
            DoEvents
            .FontName = FontCombo.Text
            .FontSize = Val(Text1.Text)
            If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = RightMargin - Temp.Width Else .CurrentX = (LeftMargin + RightMargin) / 2 - Temp.Width / 2
            .CurrentY = TopMargin + delta
            Preview.Picture2.Print Text2.Text
            delta = delta + Temp.Height
        End With
        Preview.Export.Visible = True '
    Next
    With Preview.Export
        .Width = RightMargin - LeftMargin
        .Height = delta
        DoEvents
        .BorderStyle = 0
        .PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, TopMargin, RightMargin - LeftMargin, delta
        usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
        If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
        SavePicture .Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
        SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
    End With
    Unload Preview
    List1.AddItem usage + 1
'    wholestr = str(0)
'    For i = 1 To UBound(str)
'        wholestr = wholestr & vbCrLf & str(i)
'    Next
    Text2.Text = wholestr
End Sub

Private Sub Label12_Click()
    Shape1.Left = Label12.Left
    Shape1.Width = Label12.Width
    Frame1.Visible = False
    Frame2.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    List1.Visible = True
    Label13.Visible = True
    Label14.Visible = True
    Manage.Visible = True
End Sub

Private Sub Label13_Click()
    On Error GoTo err
    InitPreview
    Preview.Picture2.AutoSize = True
    Preview.Picture2.Picture = LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg")
    Preview.Picture2.AutoSize = False
err:
    Preview.NewMessage "Image which tracknumber=" & List1.Text & " not found", vbBlue
End Sub

Private Sub Label14_Click()
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub Label15_Click()
    Dim i As Integer
    On Error Resume Next
    Dim X As Integer
    X = TopMargin
    InitPreview
    For i = 0 To List1.ListCount - 1
        List1.ListIndex = i
        Debug.Print X
        Preview.Picture2.PaintPicture LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg"), LeftMargin, X
        Preview.Export.Width = 1
        Preview.Export.Height = 1
        Preview.Export.Visible = True
        Preview.Export.Picture = LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg")
        X = X + Preview.Export.Height
        Debug.Print Preview.Export.Height
    Next
    Preview.Export.Visible = False
End Sub

