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
   Begin VB.TextBox Text9 
      BackColor       =   &H00B4BFCC&
      BorderStyle     =   0  'None
      ForeColor       =   &H00656D76&
      Height          =   285
      Left            =   1155
      TabIndex        =   4
      Text            =   "99999"
      Top             =   6450
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00B4BFCC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   9705
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00B4BFCC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   6435
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9705
   End
   Begin VB.Label Label68 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00A0ACBA&
      Caption         =   "AutoSpace"
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
      Left            =   45
      TabIndex        =   5
      Top             =   6480
      Width           =   990
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
      Left            =   9225
      TabIndex        =   1
      Top             =   6480
      Width           =   390
   End
   Begin VB.Shape Shape0 
      BackColor       =   &H00B4BFCC&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00B4BFCC&
      Height          =   300
      Index           =   22
      Left            =   8385
      Top             =   6465
      Width           =   795
   End
   Begin VB.Shape Shape0 
      BackColor       =   &H00B4BFCC&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00B4BFCC&
      Height          =   330
      Index           =   0
      Left            =   9195
      Top             =   6465
      Width           =   450
   End
End
Attribute VB_Name = "InputWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Text9.Text = GetSetting("FreeExam", "Create", "AutoSpace", 99999)
    translatecontrol Me.Name
End Sub

Private Sub Label1_Click()
    Me.Caption = translate("UserCancel")
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

Private Sub Text9_Change()
    Dim oldstr As String, v As Long, newstr As String
    If Not IsNumeric(Text9.Text) Or Int(Val(Text9.Text)) <> Val(Text9.Text) Or Val(Text9.Text) < 1 Then
        Text9.Text = GetSetting("FreeExam", "Create", "AutoSpace", 99999)
        Exit Sub
    End If
    v = GetSetting("FreeExam", "Create", "AutoSpace", 99999)
    Dim i As Long
    If v = 99999 Then
        v = Val(Text9.Text)
        GoTo save
    End If
    oldstr = "^u^"
    newstr = "^u^"
    For i = 1 To v
        oldstr = oldstr & " "
    Next
    v = Val(Text9.Text)
    For i = 1 To v
        newstr = newstr & " "
    Next
    newstr = newstr & "^u^"
    oldstr = oldstr & "^u^"
    Text1.Text = Replace(Text1.Text, oldstr, newstr)
    Text2.Text = Replace(Text2.Text, oldstr, newstr)
save:
    SaveSetting "FreeExam", "Create", "AutoSpace", v
End Sub
