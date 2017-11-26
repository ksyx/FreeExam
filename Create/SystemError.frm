VERSION 5.00
Begin VB.Form SystemError 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
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
   ScaleHeight     =   4905
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4455
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3540
      Top             =   270
   End
   Begin VB.Label CurrentTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Left            =   3900
      TabIndex        =   2
      Top             =   30
      Width           =   465
   End
   Begin VB.Label ErrDetail 
      BackStyle       =   0  'Transparent
      Caption         =   "Error!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   165
      TabIndex        =   1
      Top             =   780
      Width           =   7080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SystemError"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   2730
   End
End
Attribute VB_Name = "SystemError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CurrentTime_DblClick()
    Form_DblClick
End Sub

Private Sub ErrDetail_DblClick()
    Form_DblClick
End Sub

Private Sub Form_DblClick()
    If Timer1.Enabled = False Then Unload Me
End Sub

Private Sub Label1_DblClick()
    Form_DblClick
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    CurrentTime.Caption = Now
End Sub
