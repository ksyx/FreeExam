VERSION 5.00
Begin VB.Form Integrated 
   BackColor       =   &H00A0ACBA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Initiating"
   ClientHeight    =   540
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
   ScaleHeight     =   540
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00656D76&
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   15
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
         ForeColor       =   &H00656D76&
         Height          =   525
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5025
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1725
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   2190
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Click the shape to insert, Press Esc to exit"
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
         Height          =   390
         Left            =   75
         TabIndex        =   44
         Top             =   1305
         Width           =   2100
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   27
         Left            =   990
         TabIndex        =   43
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   26
         Left            =   990
         TabIndex        =   42
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   25
         Left            =   990
         TabIndex        =   41
         Top             =   555
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   24
         Left            =   975
         TabIndex        =   40
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   23
         Left            =   975
         TabIndex        =   39
         Top             =   285
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   22
         Left            =   750
         TabIndex        =   38
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   21
         Left            =   735
         TabIndex        =   37
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   20
         Left            =   735
         TabIndex        =   36
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   19
         Left            =   495
         TabIndex        =   35
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   18
         Left            =   495
         TabIndex        =   34
         Top             =   825
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   17
         Left            =   495
         TabIndex        =   33
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   16
         Left            =   480
         TabIndex        =   32
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   15
         Left            =   480
         TabIndex        =   31
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   14
         Left            =   240
         TabIndex        =   30
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Index           =   13
         Left            =   225
         TabIndex        =   29
         Top             =   825
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   12
         Left            =   240
         TabIndex        =   28
         Top             =   555
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Index           =   11
         Left            =   225
         TabIndex        =   27
         Top             =   285
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   10
         Left            =   240
         TabIndex        =   26
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   20
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   825
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   555
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   285
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H00656D76&
         Height          =   870
         Left            =   1530
         TabIndex        =   15
         Top             =   720
         Width           =   225
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1800
      Left            =   15
      TabIndex        =   2
      Top             =   -45
      Width           =   5655
      Begin VB.TextBox Text1 
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
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
         ForeColor       =   &H00656D76&
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
         ForeColor       =   &H00656D76&
         Height          =   870
         Left            =   1500
         TabIndex        =   6
         Top             =   810
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
         ForeColor       =   &H00656D76&
         Height          =   435
         Left            =   390
         TabIndex        =   5
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00656D76&
         Height          =   435
         Left            =   435
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1800
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text2 
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00656D76&
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
         ForeColor       =   &H00656D76&
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
         ForeColor       =   &H00656D76&
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
         ForeColor       =   &H00656D76&
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
         ForeColor       =   &H00656D76&
         Height          =   195
         Left            =   2940
         TabIndex        =   10
         Top             =   1515
         Width           =   2685
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
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
      Height          =   240
      Index           =   9
      Left            =   0
      TabIndex        =   25
      Top             =   1035
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
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
      Height          =   240
      Index           =   8
      Left            =   0
      TabIndex        =   24
      Top             =   810
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
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
      Height          =   240
      Index           =   7
      Left            =   0
      TabIndex        =   23
      Top             =   540
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
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
      Height          =   240
      Index           =   6
      Left            =   0
      TabIndex        =   22
      Top             =   270
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
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
      Height          =   240
      Index           =   5
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "Integrated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WinMode As Long

Sub InitWindow()
    If WinMode = 1 Then 'ProgressMode
        Frame1.Visible = True
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Me.Width = Frame1.Width
        Me.Height = Frame1.Height + TitleHi
        Me.Caption = translate("Initiating")
    End If
    If WinMode = 2 Then 'SpecialInput
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = False
        Frame4.Visible = False
        Me.Width = Frame2.Width
        Me.Height = Frame2.Height + TitleHi
        Me.Caption = translate("Special Input")
        Text1.Text = ""
    End If
    If WinMode = 3 Then 'SpecialInput
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = True
        Frame4.Visible = False
        Me.Width = Frame3.Width
        Me.Height = Frame3.Height + TitleHi
        Me.Caption = translate("Format")
        Text2.Text = ""
    End If
    If WinMode = 4 Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = True
        Me.Width = Frame4.Width
        Me.Height = Frame4.Height + TitleHi
        Me.Caption = translate("Table Maker")
    End If
    If EnableTranslation = 1 Then
        Message.Font = "����"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    translatecontrol Me.Name
End Sub

Private Sub Label9_Click(Index As Integer)
    InputWin.Text2.SelText = Label9(Index).Caption
End Sub

Private Sub Text1_Change()
Debug.Print "CHANGE"
    'Me.Height = Me.Height - Label3.Height
    Label3.FontSize = 23
    'Frame1.Height = Frame1.Height - Label3.Height
    Select Case LCase(Text1.Text)
        Case "delta": Label3.Caption = "��"
        Case "pi": Label3.Caption = "��"
        Case "rou": Label3.Caption = "��"
        Case "density": Label3.Caption = "��"
        Case "p": Label3.Caption = "��"
        Case "s": Label3.Caption = "��"
        Case "s=": Label3.Caption = "��"
        Case "s": Label3.Caption = "��"
        Case "<": Label3.Caption = "��"
        Case ">": Label3.Caption = "��"
        Case "<=": Label3.Caption = "��"
        Case ">=": Label3.Caption = "��"
        Case "oo": Label3.Caption = "��"
        Case "inf": Label3.Caption = "��"
        Case "o.": Label3.Caption = "��"
        Case "circle": Label3.Caption = "��"
        Case "yuan": Label3.Caption = "��"
        Case "because": Label3.Caption = "��"
        Case "so": Label3.Caption = "��"
        Case "alpha": Label3.Caption = "��"
        Case "gamma": Label3.Caption = "��"
        Case "eta": Label3.Caption = "��"
        Case "micro": Label3.Caption = "��"
        Case "a": Label3.Caption = "��"
        Case "y": Label3.Caption = "��"
        Case "n": Label3.Caption = "��"
        Case "u": Label3.Caption = "��"
        Case "x": Label3.Caption = "��"
        Case "w": Label3.Caption = "��"
        Case "%": Label3.Caption = "��"
        Case "%.": Label3.Caption = "��"
        Case "%��": Label3.Caption = "��"
        Case "duc": Label3.Caption = "��"
        Case "degreec": Label3.Caption = "��"
        Case "f": Label3.Caption = "�H"
        Case "duf": Label3.Caption = "�H"
        Case "degreef": Label3.Caption = "�H"
        Case "'": Label3.Caption = "��"
        Case "''": Label3.Caption = "��"
        Case "+": Label3.Caption = "��"
        Case "-": Label3.Caption = "��"
        Case "*": Label3.Caption = "��"
        Case "/": Label3.Caption = "��"
        Case "+-": Label3.Caption = "��"
        Case "=": Label3.Caption = "��"
        Case "~=": Label3.Caption = "��"
        Case "-=": Label3.Caption = "��"
        Case "/=": Label3.Caption = "��"
        Case "o": Label3.Caption = "��"
        Case "du": Label3.Caption = "��"
        Case "a1": Label3.Caption = "��"
        Case "a2": Label3.Caption = "��"
        Case "a3": Label3.Caption = "��"
        Case "a4": Label3.Caption = "��"
        Case "u1": Label3.Caption = "��"
        Case "u2": Label3.Caption = "��"
        Case "u3": Label3.Caption = "��"
        Case "u4": Label3.Caption = "��"
        Case "o1": Label3.Caption = "��"
        Case "o2": Label3.Caption = "��"
        Case "o3": Label3.Caption = "��"
        Case "o4": Label3.Caption = "��"
        Case "i1": Label3.Caption = "��"
        Case "i2": Label3.Caption = "��"
        Case "i3": Label3.Caption = "��"
        Case "i4": Label3.Caption = "��"
        Case "e1": Label3.Caption = "��"
        Case "e2": Label3.Caption = "��"
        Case "e3": Label3.Caption = "��"
        Case "e4": Label3.Caption = "��"
        Case "v1": Label3.Caption = "��"
        Case "v2": Label3.Caption = "��"
        Case "v3": Label3.Caption = "��"
        Case "v4": Label3.Caption = "��"
        Case "v0": Label3.Caption = "��"
        Case "1a": Label3.Caption = "��"
        Case "2a": Label3.Caption = "��"
        Case "3a": Label3.Caption = "��"
        Case "4a": Label3.Caption = "��"
        Case "5a": Label3.Caption = "��"
        Case "6a": Label3.Caption = "��"
        Case "7a": Label3.Caption = "��"
        Case "8a": Label3.Caption = "��"
        Case "9a": Label3.Caption = "��"
        Case "10a": Label3.Caption = "��"
        Case "11a": Label3.Caption = "��"
        Case "12a": Label3.Caption = "��"
        Case "13a": Label3.Caption = "��"
        Case "14a": Label3.Caption = "��"
        Case "15a": Label3.Caption = "��"
        Case "16a": Label3.Caption = "��"
        Case "17a": Label3.Caption = "��"
        Case "18a": Label3.Caption = "��"
        Case "19a": Label3.Caption = "��"
        Case "20a": Label3.Caption = "��"
        
        Case "1b": Label3.Caption = "��"
        Case "2b": Label3.Caption = "��"
        Case "3b": Label3.Caption = "��"
        Case "4b": Label3.Caption = "��"
        Case "5b": Label3.Caption = "��"
        Case "6b": Label3.Caption = "��"
        Case "7b": Label3.Caption = "��"
        Case "8b": Label3.Caption = "��"
        Case "9b": Label3.Caption = "��"
        Case "10b": Label3.Caption = "��"
        
        Case "1c": Label3.Caption = "��"
        Case "2c": Label3.Caption = "��"
        Case "3c": Label3.Caption = "��"
        Case "4c": Label3.Caption = "��"
        Case "5c": Label3.Caption = "��"
        Case "6c": Label3.Caption = "��"
        Case "7c": Label3.Caption = "��"
        Case "8c": Label3.Caption = "��"
        Case "9c": Label3.Caption = "��"
        Case "10c": Label3.Caption = "��"
        Case "11c": Label3.Caption = "��"
        Case "12c": Label3.Caption = "��"
        Case "13c": Label3.Caption = "��"
        Case "14c": Label3.Caption = "��"
        Case "15c": Label3.Caption = "��"
        Case "16c": Label3.Caption = "��"
        Case "17c": Label3.Caption = "��"
        Case "18c": Label3.Caption = "��"
        Case "19c": Label3.Caption = "��"
        Case "20c": Label3.Caption = "��"
        
        Case "1d": Label3.Caption = "��"
        Case "2d": Label3.Caption = "��"
        Case "3d": Label3.Caption = "��"
        Case "4d": Label3.Caption = "��"
        Case "5d": Label3.Caption = "��"
        Case "6d": Label3.Caption = "��"
        Case "7d": Label3.Caption = "��"
        Case "8d": Label3.Caption = "��"
        Case "9d": Label3.Caption = "��"
        Case "10d": Label3.Caption = "��"
        Case "11d": Label3.Caption = "��"
        Case "12d": Label3.Caption = "��"
        Case "13d": Label3.Caption = "��"
        Case "14d": Label3.Caption = "��"
        Case "15d": Label3.Caption = "��"
        Case "16d": Label3.Caption = "��"
        Case "17d": Label3.Caption = "��"
        Case "18d": Label3.Caption = "��"
        Case "19d": Label3.Caption = "��"
        Case "20d": Label3.Caption = "��"
        
        Case "1e": Label3.Caption = "�� "
        Case "2e": Label3.Caption = "�� "
        Case "3e": Label3.Caption = "�� "
        Case "4e": Label3.Caption = "�� "
        Case "5e": Label3.Caption = "�� "
        Case "6e": Label3.Caption = "�� "
        Case "7e": Label3.Caption = "�� "
        Case "8e": Label3.Caption = "�� "
        Case "9e": Label3.Caption = "�� "
        Case "10e": Label3.Caption = "�� "
        Case "11e": Label3.Caption = "�� "
        Case "12e": Label3.Caption = "�� "
        Case "13e": Label3.Caption = "�� "
        Case "14e": Label3.Caption = "�� "
        Case "15e": Label3.Caption = "�� "
        Case "16e": Label3.Caption = "�� "
        Case "17e": Label3.Caption = "�� "
        Case "18e": Label3.Caption = "�� "
        Case "19e": Label3.Caption = "�� "
        Case "20e": Label3.Caption = "�� "

        Case Else: Label3.Caption = ""
    End Select
    If Label3.Caption <> "" Then Exit Sub
    Label3.FontSize = 12
    If Left(Text1.Text, 3) = "tot" Then
        Label3.Caption = "��"
    End If
    If Left(Text1.Text, 3) = "no" Then
        Label3.Caption = "��"
    End If
    If Left(Text1.Text, 6) = "listen" Then
        Label3.Caption = "����" & (Right(Text1.Text, Len(Text1.Text) - 6)) & "�ζԻ����ش��"
    End If
    If Left(Text1.Text, 3) = "pts" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "��"
    End If
    If Left(Text1.Text, 3) = "sub" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "С��"
    End If
    If Left(Text1.Text, 3) = "big" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "����"
    End If
    If Left(Text1.Text, 3) = "spc" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 3)) & "��"
    End If
    If Left(Text1.Text, 4) = "prob" Then
        Label3.Caption = (Right(Text1.Text, Len(Text1.Text) - 4)) & "��"
    End If
    If Left(Text1.Text, 3) = "sel" Then
        Label3.Caption = "ÿ��С�����ĸ���ѡ�𰸣�������ѡ������������һ����"
    End If
    
   ' Me.Height = Label3.Height + Me.Height
    'Frame1.Height = Label3.Height + Frame1.Height
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        InputWin.Text2.SelText = Label3.Caption
        InputWin.Text1.SelText = Label3.Caption
        KeyCode = 0
        Me.Hide
    End If
    If KeyCode = vbKeyEscape Then Me.Hide
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        InputWin.Text2.SelText = Label6.Caption
        InputWin.Text1.SelText = Label6.Caption
        If Len(Text2.Text) = 1 Then InputWin.Text2.SelStart = InputWin.Text2.SelStart - Len(Label6.Caption) / 2
        KeyCode = 0
        Me.Hide
    End If
    If KeyCode = vbKeyEscape Then Me.Hide
    
End Sub

Private Sub Text2_Change()
    Select Case LCase(Text2.Text)
        Case "b": Label6.Caption = "^b^^b^"
        Case "i": Label6.Caption = "^i^^i^"
        Case "u": Label6.Caption = "^u^^u^"
        Case "d": Label6.Caption = "^d^^d^"
        Case "e": Label6.Caption = "^ee^^ed^"
        Case "s": Label6.Caption = "^se^^sd^"
    End Select
    If LCase(Text2.Text) = "as" Then
        Dim v As Long
        v = GetSetting("FreeExam", "Create", "AutoSpace", 99999)
        If v = 99999 Then Exit Sub
        Label6.Caption = "^u^"
        Dim i As Long
        For i = 1 To v
            Label6.Caption = Label6.Caption & " "
        Next
        Label6.Caption = Label6.Caption & "^u^ "
    End If
End Sub
