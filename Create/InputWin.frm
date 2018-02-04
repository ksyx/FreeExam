VERSION 5.00
Begin VB.Form InputWin 
   BackColor       =   &H00A0ACBA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input"
   ClientHeight    =   6765
   ClientLeft      =   -75
   ClientTop       =   390
   ClientWidth     =   9675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00B4BFCC&
      BorderStyle     =   0  'None
      ForeColor       =   &H00656D76&
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   9705
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00B4BFCC&
      BorderStyle     =   0  'None
      ForeColor       =   &H00656D76&
      Height          =   6435
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9705
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00B4BFCC&
      Caption         =   " Cancel "
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
      Left            =   8415
      TabIndex        =   2
      Top             =   6480
      Width           =   720
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00B4BFCC&
      Caption         =   " OK "
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
      Left            =   9210
      TabIndex        =   1
      Top             =   6480
      Width           =   390
   End
End
Attribute VB_Name = "InputWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    Me.Caption = "UserCancel"
    Me.Hide
End Sub

Private Sub Label28_Click()
    Me.Hide
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Integrated.WinMode = 2
        Integrated.InitWindow
        Integrated.Show 1
    End If
    If KeyCode = vbKeyF3 Then
        Integrated.WinMode = 3
        Integrated.InitWindow
        Integrated.Show 1
    End If
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Integrated.WinMode = 2
        Integrated.InitWindow
        Integrated.Show 1
    End If
    If KeyCode = vbKeyF3 Then
        Integrated.WinMode = 3
        Integrated.InitWindow
        Integrated.Show 1
    End If
End Sub

