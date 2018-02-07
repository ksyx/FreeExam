VERSION 5.00
Begin VB.Form Lang 
   BackColor       =   &H00A0ACBA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose Language / 选择语言"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0ACBA&
      Caption         =   "简体中文"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   495
      Width           =   3480
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0ACBA&
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Value           =   -1  'True
      Width           =   3480
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00B4BFCC&
      Caption         =   " Next "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   210
      Left            =   2820
      TabIndex        =   2
      Top             =   930
      Width           =   630
   End
End
Attribute VB_Name = "Lang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label28_Click()
    If Option2.Value = True Then EnableTranslation = 1
    PageSettings.Show
    Unload Me
End Sub

Private Sub Option1_Click()
    Label28.Caption = "Next"
    Label28.FontName = "Tahoma"
End Sub

Private Sub Option2_Click()
    Label28.Caption = "下一步"
    Label28.FontName = "黑体"
End Sub
