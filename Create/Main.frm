VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ExamPaper Editor"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7920
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
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame InsText 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3420
      Left            =   45
      TabIndex        =   17
      Top             =   1860
      Width           =   7785
      Begin VB.Frame Frame3 
         Caption         =   "Parts"
         Height          =   1755
         Left            =   0
         TabIndex        =   33
         Top             =   1590
         Visible         =   0   'False
         Width           =   3405
         Begin VB.ListBox List2 
            Height          =   840
            Left            =   105
            TabIndex        =   34
            Top             =   180
            Width           =   3195
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Pay attention: After selecting, it will copy to Text box, click the last one to restore your orginal text."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   90
            TabIndex        =   35
            Top             =   1065
            Width           =   3165
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Format"
         Height          =   1545
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   2895
         Begin VB.ComboBox AlignCombo 
            Height          =   315
            ItemData        =   "Main.frx":0000
            Left            =   810
            List            =   "Main.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1080
            Width           =   1905
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
            TabIndex        =   27
            Top             =   855
            Width           =   270
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
            TabIndex        =   26
            Top             =   840
            Width           =   270
         End
         Begin VB.ComboBox FontCombo 
            Height          =   315
            ItemData        =   "Main.frx":0042
            Left            =   825
            List            =   "Main.frx":004F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   225
            Width           =   1905
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   825
            MaxLength       =   3
            TabIndex        =   24
            Top             =   570
            Width           =   1890
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alignment"
            Height          =   195
            Left            =   90
            TabIndex        =   32
            Top             =   1155
            Width           =   705
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shape"
            Height          =   195
            Left            =   345
            TabIndex        =   31
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font"
            Height          =   195
            Left            =   435
            TabIndex        =   30
            Top             =   285
            Width           =   330
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            Height          =   195
            Left            =   465
            TabIndex        =   29
            Top             =   585
            Width           =   285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Options"
         Height          =   1560
         Left            =   3060
         TabIndex        =   18
         Top             =   0
         Width           =   4620
         Begin VB.TextBox Text2 
            Height          =   1050
            Left            =   555
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   330
            Width           =   3900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text"
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   315
            Width           =   375
         End
      End
      Begin VB.Label Temp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         Height          =   195
         Left            =   4155
         TabIndex        =   36
         Top             =   2130
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
         Left            =   4365
         TabIndex        =   22
         Top             =   2700
         Width           =   1125
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
         Left            =   5505
         TabIndex        =   21
         Top             =   2700
         Width           =   870
      End
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   90
      TabIndex        =   13
      Top             =   1965
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
      TabIndex        =   12
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
         TabIndex        =   16
         Top             =   135
         Width           =   1410
      End
   End
   Begin VB.ListBox MsgContentList 
      Height          =   450
      ItemData        =   "Main.frx":0084
      Left            =   4755
      List            =   "Main.frx":0086
      TabIndex        =   10
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
      ItemData        =   "Main.frx":0088
      Left            =   3330
      List            =   "Main.frx":008A
      TabIndex        =   9
      Top             =   -15
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgColorList 
      Height          =   450
      ItemData        =   "Main.frx":008C
      Left            =   4785
      List            =   "Main.frx":008E
      TabIndex        =   8
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
      TabIndex        =   6
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
         TabIndex        =   7
         Top             =   0
         Width           =   45
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
      TabIndex        =   14
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
      TabIndex        =   15
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
      TabIndex        =   11
      Top             =   120
      Width           =   810
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
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str() As String, wholestr As String, outputs As String
Dim showcnt As Integer, current As Integer, strs() As String
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
    Integrated.Show
    Integrated.WinMode = 1
    Integrated.InitWindow
    For i = 1 To Screen.FontCount
        Integrated.Message.Caption = "Loading Fonts(" & i & "/" & Screen.FontCount & ")"
        DoEvents
        FontCombo.AddItem Screen.Fonts(i)
    Next
    Unload Integrated
    Shape1.Left = Label1.Left
    Shape1.Width = Label1.Width
End Sub

Private Sub List2_Click()
    Text2.Text = strs(List2.ListIndex)
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
    If MsgContentList.ListCount <= 1 Then
        first = showcnt
        showcnt = ShowCntPerMsg
        Message.Caption = ""
        If MsgContentList.ListCount = 1 Then
            current = 0
            MsgContentList.ListIndex = current
            MsgColorList.ListIndex = current
            MsgTypeList.ListIndex = current
            Message.Caption = MsgTypeList.Text & MsgContentList.Text
            Message.ForeColor = ReverseColor(MsgColorList.Text)
        End If
        If showcnt <> first Then ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
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
    Dim delta As Integer, partid As Integer, bound As Integer, start As Integer, length As Integer, j As Integer, xdelta As Integer, tmpstr As String, issel As Boolean
    wholestr = Text2.Text
    outputs = ""
    partid = 1
    If List2.ListIndex = List2.ListCount - 1 Or List2.Text <> Text2.Text Then
        List2.Clear
        Frame3.Visible = False
        ReDim strs(0)
    End If
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
    Temp.FontName = FontCombo.Text
        Temp.FontSize = Val(Text1.Text)
        If Check2.Value = 1 Then Temp.FontBold = True Else Temp.FontBold = False
        If Check1.Value = 1 Then Temp.FontItalic = True Else Temp.FontItalic = False
        If Check2.Value = 1 Then Preview.Picture2.FontBold = True Else Preview.Picture2.FontBold = False
        If Check1.Value = 1 Then Preview.Picture2.FontItalic = True Else Preview.Picture2.FontItalic = False
    InitPreview
    For i = start To bound Step 1
        length = Len(str(i))
        Text2.Text = str(i)
        
        Temp.Alignment = Val(Left(AlignCombo.Text, 1))
        Temp.Caption = Text2.Text
        With Preview.Picture2
            .BorderStyle = 0
            .FontName = FontCombo.Text
            .FontSize = Val(Text1.Text)
            If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
            .CurrentY = TopMargin + delta
            If .CurrentY + Temp.Height >= BotMargin Then
                If partid = 1 Then
                    NewMessage "The input will be split into multi parts", vbBlue
                    NewMessage "select parts that you want to preview in the list.", vbBlue
                    'Exit Sub
                End If
                .Cls
                delta = 0
                partid = partid + 1
                ReDim Preserve strs(partid - 2)
                strs(partid - 2) = outputs
                .CurrentY = TopMargin
                Preview.Picture2.Print "test";
                If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                List2.AddItem outputs
                outputs = ""
                If length = 0 Then GoTo cont
            End If
        End With
        For j = 1 To length
            Text2.Text = Mid(str(i), j, 1)
            'Debug.Print Text2.Text
            Temp.FontName = FontCombo.Text
            Temp.FontSize = Val(Text1.Text)
            Temp.Alignment = Val(Left(AlignCombo.Text, 1))
            Temp.Caption = Text2.Text
            Temp.Visible = True
            With Preview.Picture2
                tmpstr = Text2.Text
                If Temp.Width + .CurrentX > RightMargin Then
                    If Temp.Alignment <> 0 Then
                        NewMessage "Auto split line is unsupportted for alignment mode 1 or 2.", vbRed
                        Text2.Text = wholestr
                        On Error Resume Next
                        Unload Preview
                        Exit Sub
                    End If
                    delta = delta + Temp.Height
                    If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                    .CurrentY = TopMargin + delta
                    Print BotMargin
                    'Preview.Picture2.Line (0, Preview.Picture2.CurrentY + Temp.Height)-(Preview.Picture2.Width, Preview.Picture2.CurrentY + Temp.Height), vbRed
                    If .CurrentY + Temp.Height >= BotMargin Then
                        If partid = 1 Then
                            NewMessage "The input will be split into multi parts", vbBlue
                            NewMessage "select parts that you want to preview in the list.", vbBlue
                            'Exit Sub
                        End If
                        .Cls
                        delta = 0
                        partid = partid + 1
                        ReDim Preserve strs(partid - 2)
                        strs(partid - 2) = outputs
                        .CurrentY = TopMargin
                        Preview.Picture2.Print "test";
                        If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                        List2.AddItem outputs
                        outputs = ""
                        
                    End If
                End If
                If Temp.Width > PageWidth - LeftMargin - (PageWidth - RightMargin) Or Temp.Height > PageHeight - TopMargin - (PageHeight - BotMargin) Then
                    NewMessage "The target size is too large, we are unable to process it.", vbRed
                    Text2.Text = wholestr
                    On Error Resume Next
                    Unload Preview
                    Exit Sub
                End If
            Preview.Picture2.Print tmpstr;
            outputs = outputs & tmpstr
            'Debug.Print .CurrentX
            End With
            DoEvents
        Next
        delta = delta + Temp.Height
        
        If i <> bound Then outputs = outputs & vbCrLf
cont:
    Next
'    wholestr = str(0)
'    If UBound(str) > 0 Then wholestr = wholestr & vbCrLf
'    For i = 1 To UBound(str)
'        wholestr = wholestr & str(i) & vbCrLf
'    Next
    If partid > 1 Then
        List2.AddItem outputs
        List2.AddItem wholestr
        ReDim Preserve strs(partid)
        strs(partid - 1) = wholestr
        strs(partid - 0) = wholestr
        Frame3.Visible = True
        On Error Resume Next
        Unload Preview
    End If
    Text2.Text = wholestr
End Sub

Private Sub Label11_Click()
    On Error Resume Next
    Dim i As Integer, usage As Long
    Dim delta As Integer, lastcapt As Integer, partid As Integer, bound As Integer, start As Integer, length As Integer, j As Integer, xdelta As Integer, tmpstr As String, issel As Boolean
    wholestr = Text2.Text
    outputs = ""
    partid = 1
    lastcapt = TopMargin
    If List2.ListIndex = List2.ListCount - 1 Or List2.Text <> Text2.Text Then
        List2.Clear
        Frame3.Visible = False
        ReDim strs(0)
    End If
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
    Temp.FontName = FontCombo.Text
        Temp.FontSize = Val(Text1.Text)
        If Check2.Value = 1 Then Temp.FontBold = True Else Temp.FontBold = False
        If Check1.Value = 1 Then Temp.FontItalic = True Else Temp.FontItalic = False
        If Check2.Value = 1 Then Preview.Picture2.FontBold = True Else Preview.Picture2.FontBold = False
        If Check1.Value = 1 Then Preview.Picture2.FontItalic = True Else Preview.Picture2.FontItalic = False
    InitPreview
    For i = start To bound Step 1
        length = Len(str(i))
        Text2.Text = str(i)
        
        Temp.Alignment = Val(Left(AlignCombo.Text, 1))
        Temp.Caption = Text2.Text
        With Preview.Picture2
            .BorderStyle = 0
            .FontName = FontCombo.Text
            .FontSize = Val(Text1.Text)
            If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
            .CurrentY = TopMargin + delta
            If .CurrentY + Temp.Height >= BotMargin Then
                If partid = 1 Then
                    NewMessage "The input will be split into multi parts", vbBlue
                    NewMessage "select parts that you want to preview in the list.", vbBlue
                    'Exit Sub
                End If
                .Cls
                delta = 0
                partid = partid + 1
                ReDim Preserve strs(partid - 2)
                strs(partid - 2) = outputs
                .CurrentY = TopMargin
                Preview.Picture2.Print "test";
                If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                List2.AddItem outputs
                outputs = ""
                If length = 0 Then GoTo cont
            End If
        End With
        For j = 1 To length
            Text2.Text = Mid(str(i), j, 1)
            'Debug.Print Text2.Text
            Temp.FontName = FontCombo.Text
            Temp.FontSize = Val(Text1.Text)
            Temp.Alignment = Val(Left(AlignCombo.Text, 1))
            Temp.Caption = Text2.Text
            Temp.Visible = True
            With Preview.Picture2
                tmpstr = Text2.Text
                If Temp.Width + .CurrentX > RightMargin Then
                    If Temp.Alignment <> 0 Then
                        NewMessage "Auto split line is unsupportted for alignment mode 1 or 2.", vbRed
                        Text2.Text = wholestr
                        On Error Resume Next
                        Unload Preview
                        Exit Sub
                    End If
                    delta = delta + Temp.Height
                    If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                    .CurrentY = TopMargin + delta
                    Print BotMargin
                    'Preview.Picture2.Line (0, Preview.Picture2.CurrentY + Temp.Height)-(Preview.Picture2.Width, Preview.Picture2.CurrentY + Temp.Height), vbRed
                    With Preview.Export
                        .Width = RightMargin - LeftMargin
                        .Height = Temp.Height
                        DoEvents
                        .BorderStyle = 0
                        .PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, lastcapt, RightMargin - LeftMargin, Temp.Height
                        lastcapt = lastcapt + Temp.Height
                        usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
                        If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
                        SavePicture .Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
                        SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
                        List1.AddItem usage + 1
                    End With
                    If .CurrentY + Temp.Height >= BotMargin Then
                        If partid = 1 Then
                            NewMessage "The input will be split into multi parts", vbBlue
                            NewMessage "select parts that you want to preview in the list.", vbBlue
                            'Exit Sub
                        End If
                        .Cls
                        delta = 0
                        partid = partid + 1
                        ReDim Preserve strs(partid - 2)
                        strs(partid - 2) = outputs
                        .CurrentY = TopMargin
                        Preview.Picture2.Print "test";
                        If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                        List2.AddItem outputs
                        outputs = ""
                        
                    End If
                    
                End If
                If Temp.Width > PageWidth - LeftMargin - (PageWidth - RightMargin) Or Temp.Height > PageHeight - TopMargin - (PageHeight - BotMargin) Then
                    NewMessage "The target size is too large, we are unable to process it.", vbRed
                    Text2.Text = wholestr
                    On Error Resume Next
                    Unload Preview
                    Exit Sub
                End If
            Preview.Picture2.Print tmpstr;
            outputs = outputs & tmpstr
            'Debug.Print .CurrentX
            End With
            DoEvents
        Next
        With Preview.Export
            .Width = RightMargin - LeftMargin
            .Height = Temp.Height
            DoEvents
            .BorderStyle = 0
            .PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, lastcapt, RightMargin - LeftMargin, Temp.Height
            lastcapt = lastcapt + Temp.Height
            usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
            If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
            SavePicture .Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
            SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
            List1.AddItem usage + 1
        End With
        delta = delta + Temp.Height
        If i <> bound Then outputs = outputs & vbCrLf
cont:
    Next
'    wholestr = str(0)
'    If UBound(str) > 0 Then wholestr = wholestr & vbCrLf
'    For i = 1 To UBound(str)
'        wholestr = wholestr & str(i) & vbCrLf
'    Next
    If partid > 1 Then
        List2.AddItem outputs
        List2.AddItem wholestr
        ReDim Preserve strs(partid)
        strs(partid - 1) = wholestr
        strs(partid - 0) = wholestr
        Frame3.Visible = True
        On Error Resume Next
        Unload Preview
    End If
    Text2.Text = wholestr
'    With Preview.Export
'        .Width = RightMargin - LeftMargin
'        .Height = delta
'        DoEvents
'        .BorderStyle = 0
'        .PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, TopMargin, RightMargin - LeftMargin, delta
'        usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
'        If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
'        SavePicture .Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
'        SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
'    End With
    Unload Preview
   
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

