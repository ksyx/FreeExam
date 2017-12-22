VERSION 5.00
Begin VB.Form Integrated 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Initiating"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1800
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1500
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   3570
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   435
         TabIndex        =   13
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   390
         TabIndex        =   12
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1530
         TabIndex        =   11
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press Enter to insert, Esc to exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   10
         Top             =   1530
         Width           =   2685
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1800
      Left            =   15
      TabIndex        =   2
      Top             =   30
      Width           =   5655
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1500
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   3570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press Enter to insert, Esc to exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   7
         Top             =   1530
         Width           =   2685
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1530
         TabIndex        =   6
         Top             =   720
         Width           =   3570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   390
         TabIndex        =   5
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   435
         TabIndex        =   4
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   525
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      Begin VB.Label Message 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Fonts...(233/233)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5025
      End
   End
End
Attribute VB_Name = "Integrated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WinMode As Integer

Sub InitWindow()
    If WinMode = 1 Then 'ProgressMode
        Frame1.Visible = True
        Frame2.Visible = False
        Frame3.Visible = False
        Me.Width = Frame1.Width
        Me.Height = Frame1.Height + TitleHi
        Me.Caption = "Initiating"
    End If
    If WinMode = 2 Then 'SpecialInput
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = False
        Me.Width = Frame2.Width
        Me.Height = Frame2.Height + TitleHi
        Me.Caption = "Special Input"
    End If
    If WinMode = 3 Then 'SpecialInput
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = True
        Me.Width = Frame3.Width
        Me.Height = Frame3.Height + TitleHi
        Me.Caption = "Format"
    End If
End Sub

Private Sub Text1_Change()
Debug.Print "CHANGE"
    'Me.Height = Me.Height - Label3.Height
    Label3.FontSize = 23
    'Frame1.Height = Frame1.Height - Label3.Height
    Select Case LCase(Text1.Text)
        Case "delta": Label3.Caption = "Δ"
        Case "pi": Label3.Caption = "π"
        Case "rou": Label3.Caption = "ρ"
        Case "density": Label3.Caption = "ρ"
        Case "p": Label3.Caption = "ρ"
        Case "s": Label3.Caption = "∽"
        Case "s=": Label3.Caption = "≌"
        Case "s": Label3.Caption = "∽"
        Case "<": Label3.Caption = "＜"
        Case ">": Label3.Caption = "＞"
        Case "<=": Label3.Caption = "≤"
        Case ">=": Label3.Caption = "≥"
        Case "oo": Label3.Caption = "∞"
        Case "inf": Label3.Caption = "∞"
        Case "o.": Label3.Caption = "⊙"
        Case "circle": Label3.Caption = "⊙"
        Case "because": Label3.Caption = "∵"
        Case "so": Label3.Caption = "∴"
        Case "alpha": Label3.Caption = "α"
        Case "gamma": Label3.Caption = "γ"
        Case "eta": Label3.Caption = "η"
        Case "micro": Label3.Caption = "μ"
        Case "a": Label3.Caption = "α"
        Case "y": Label3.Caption = "γ"
        Case "n": Label3.Caption = "η"
        Case "u": Label3.Caption = "μ"
        Case "x": Label3.Caption = "χ"
        Case "w": Label3.Caption = "ω"
        Case "%": Label3.Caption = "％"
        Case "%.": Label3.Caption = "‰"
        Case "%。": Label3.Caption = "‰"
        Case "duc": Label3.Caption = "℃"
        Case "degreec": Label3.Caption = "℃"
        Case "f": Label3.Caption = "℃"
        Case "duf": Label3.Caption = "℃"
        Case "degreef": Label3.Caption = "℃"
        Case "f": Label3.Caption = "℃"
        Case "'": Label3.Caption = "′"
        Case "''": Label3.Caption = "″"
        Case "+": Label3.Caption = "＋"
        Case "-": Label3.Caption = "－"
        Case "*": Label3.Caption = "×"
        Case "/": Label3.Caption = "÷"
        Case "+-": Label3.Caption = "±"
        Case "=": Label3.Caption = "＝"
        Case "~=": Label3.Caption = "≈"
        Case "-=": Label3.Caption = "≡"
        Case "/=": Label3.Caption = "≠"
        Case "o": Label3.Caption = "°"
        Case Else: Label3.Caption = ""
    End Select
    Label3.FontSize = 12
    If Left(Text1.Text, 3) = "tot" Then
        Label3.Caption = "共"
    End If
    If Left(Text1.Text, 3) = "no" Then
        Label3.Caption = "第"
    End If
    If Left(Text1.Text, 6) = "listen" Then
        Label3.Caption = "听第" & (Right(Text1.Text, Len(Text1.Text) - 6)) & "段对话，回答第"
    End If
    If Left(Text1.Text, 3) = "pts" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "分"
    End If
    If Left(Text1.Text, 3) = "sub" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "小题"
    End If
    If Left(Text1.Text, 3) = "big" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "小题"
    End If
    If Left(Text1.Text, 3) = "spc" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "空"
    End If
    If Left(Text1.Text, 4) = "prob" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 4)) & "题"
    End If
    If Left(Text1.Text, 3) = "sel" Then
        Label3.Caption = "每个小题有四个备选答案，从其中选出最符合题意的一个。"
    End If
    
   ' Me.Height = Label3.Height + Me.Height
    'Frame1.Height = Label3.Height + Frame1.Height
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        MainFrm.Text2.SelText = Label3.Caption
        KeyCode = 0
        Unload Me
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        MainFrm.Text2.SelText = Label6.Caption
        MainFrm.Text2.SelStart = MainFrm.Text2.SelStart - Len(Label6.Caption) / 2
        KeyCode = 0
        Unload Me
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Text2_Change()
    Select Case LCase(Text2.Text)
        Case "b": Label6.Caption = "^b^^b^"
        Case "i": Label6.Caption = "^i^^i^"
        Case "u": Label6.Caption = "^u^^u^"
        Case "e": Label6.Caption = "^ee^^ed^"
        Case "s": Label6.Caption = "^se^^sd^"
    End Select
End Sub
