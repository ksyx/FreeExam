VERSION 5.00
Begin VB.Form PageType 
   Caption         =   "Form2"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8670
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
   ScaleHeight     =   5040
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Caption         =   "Presets"
      Height          =   1605
      Left            =   3645
      TabIndex        =   10
      Top             =   2040
      Width           =   3480
      Begin VB.CommandButton Command4 
         Caption         =   "8K"
         Height          =   360
         Left            =   1155
         TabIndex        =   12
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton Command3 
         Caption         =   "A4"
         Height          =   360
         Left            =   105
         TabIndex        =   11
         Top             =   165
         Width           =   990
      End
      Begin VB.Frame Frame2 
         Caption         =   "Choose a type of 8K"
         Height          =   915
         Left            =   1110
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   2220
         Begin VB.CommandButton Command6 
            Caption         =   "Confirm"
            Height          =   240
            Left            =   60
            TabIndex        =   16
            Top             =   585
            Width           =   990
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cancel"
            Height          =   240
            Left            =   1065
            TabIndex        =   15
            Top             =   585
            Width           =   990
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "PageType.frx":0000
            Left            =   150
            List            =   "PageType.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   255
            Width           =   1905
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4205
      Left            =   345
      ScaleHeight     =   4200
      ScaleWidth      =   2970
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   330
      Width           =   2975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   360
      Left            =   5565
      TabIndex        =   3
      Top             =   3825
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4515
      TabIndex        =   2
      Top             =   3825
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   4470
      TabIndex        =   1
      Text            =   "21"
      Top             =   1305
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   4470
      TabIndex        =   0
      Text            =   "29.7"
      Top             =   1680
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
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
      Left            =   3780
      TabIndex        =   9
      Top             =   1290
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   1650
      Width           =   690
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
      Left            =   6285
      TabIndex        =   7
      Top             =   1305
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
      Left            =   0
      TabIndex        =   6
      Top             =   90
      Width           =   555
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
      Left            =   6285
      TabIndex        =   5
      Top             =   1695
      Width           =   315
   End
End
Attribute VB_Name = "PageType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub SetVal(Width As Double, Height As Double)
    Text1(0).Text = Width
    Text1(1).Text = Height
    Picture1.Width = 567 / 5 * Width
    Picture1.Height = 567 / 5 * Height
End Sub

Private Sub Command3_Click()
    SetVal 21, 29.7
End Sub

Private Sub Command4_Click()
    Frame2.Visible = True
End Sub

Private Sub Command5_Click()
    Frame2.Visible = False
End Sub

Private Sub Command6_Click()
    If Combo1.Text = "260mm¡Á370mm" Then SetVal 37, 26
    If Combo1.Text = "420mm¡Á285mm" Then SetVal 42, 28.5
    Frame2.Visible = False
End Sub

