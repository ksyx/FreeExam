VERSION 5.00
Begin VB.Form EdgeSettings 
   Caption         =   "EdgeSettings"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7320
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
   ScaleHeight     =   4830
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   4695
      TabIndex        =   10
      Text            =   "2.54"
      Top             =   2625
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Text            =   "2.54"
      Top             =   2295
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   4680
      TabIndex        =   8
      Text            =   "3.18"
      Top             =   1950
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Text            =   "3.18"
      Top             =   1575
      Width           =   1770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4725
      TabIndex        =   2
      Top             =   4095
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   360
      Left            =   5775
      TabIndex        =   1
      Top             =   4095
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4205
      Left            =   450
      ScaleHeight     =   4200
      ScaleWidth      =   2970
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   270
      Width           =   2975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cm"
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
      Left            =   6525
      TabIndex        =   15
      Top             =   2625
      Width           =   315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cm"
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
      Left            =   6510
      TabIndex        =   14
      Top             =   2325
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cm"
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
      Left            =   6495
      TabIndex        =   13
      Top             =   1965
      Width           =   315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   210
      TabIndex        =   12
      Top             =   360
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cm"
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
      Left            =   6495
      TabIndex        =   11
      Top             =   1575
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bottom"
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
      Left            =   3870
      TabIndex        =   6
      Top             =   2595
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
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
      Left            =   4215
      TabIndex        =   5
      Top             =   2265
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   525
   End
End
Attribute VB_Name = "EdgeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Text1_Change (0)
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo err
    Picture1.Cls
    Picture1.Line (567 * Text1(0).Text / 5, 0)-(567 * Text1(0).Text / 5, Picture1.Height)
    Picture1.Line (Picture1.Width - 567 * Text1(1).Text / 5, 0)-(Picture1.Width - 567 * Text1(1).Text / 5, Picture1.Height)
    Picture1.Line (0, 567 * Text1(2).Text / 5)-(Picture1.Width, 567 * Text1(2).Text / 5)
    Picture1.Line (0, Picture1.Height - 567 * Text1(3).Text / 5)-(Picture1.Width, Picture1.Height - 567 * Text1(3).Text / 5)
    Exit Sub
err:
    Picture1.ForeColor = vbRed
    Picture1.FontSize = 32
    Picture1.Print "Error"
    Picture1.FontSize = 18
    Picture1.Print "Invaild Values"
    Picture1.ForeColor = vbBlack
End Sub
