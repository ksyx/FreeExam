VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ExamPaper Editor"
   ClientHeight    =   8700
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
   ScaleHeight     =   8700
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox WIP 
      Height          =   8250
      Left            =   99999
      ScaleHeight     =   8190
      ScaleWidth      =   7815
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   0
      Width           =   7875
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "The preview window is opening. You should close it before using this window."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   15
         TabIndex        =   45
         Top             =   1125
         Width           =   7755
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work in progress"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1890
         TabIndex        =   44
         Top             =   465
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   510
      Top             =   3570
   End
   Begin VB.ListBox MsgTypeList 
      Height          =   450
      ItemData        =   "Main.frx":0000
      Left            =   3330
      List            =   "Main.frx":0002
      TabIndex        =   8
      Top             =   -15
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgColorList 
      Height          =   450
      ItemData        =   "Main.frx":0004
      Left            =   4785
      List            =   "Main.frx":0006
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   7725
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8310
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
         TabIndex        =   6
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox General 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   75
      ScaleHeight     =   1380
      ScaleWidth      =   7770
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   525
      Width           =   7770
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Answerline(&T) "
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
         Left            =   930
         TabIndex        =   47
         Top             =   45
         Width           =   1365
      End
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
         Left            =   60
         TabIndex        =   46
         Top             =   45
         Width           =   840
      End
   End
   Begin VB.PictureBox Manage 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   75
      ScaleHeight     =   1380
      ScaleWidth      =   7770
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   7770
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Logs "
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
         Left            =   1590
         TabIndex        =   63
         Top             =   120
         Width           =   585
      End
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
         Left            =   105
         TabIndex        =   11
         Top             =   120
         Width           =   1410
      End
   End
   Begin VB.Frame InsText 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6315
      Left            =   75
      TabIndex        =   12
      Top             =   1950
      Width           =   7785
      Begin VB.Frame Frame10 
         Caption         =   "Logger"
         Height          =   825
         Left            =   3630
         TabIndex        =   60
         Top             =   4545
         Width           =   3675
         Begin VB.CheckBox Check17 
            Caption         =   "Auto"
            Height          =   210
            Left            =   1920
            TabIndex        =   64
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ThisPage "
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
            Left            =   975
            TabIndex        =   62
            Top             =   225
            Width           =   1755
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Format "
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
            Left            =   105
            TabIndex        =   61
            Top             =   225
            Width           =   810
         End
      End
      Begin VB.CheckBox Check16 
         Caption         =   "English Mode"
         Height          =   285
         Left            =   3600
         TabIndex        =   59
         Top             =   5415
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Caption         =   "Options"
         Height          =   1560
         Left            =   3045
         TabIndex        =   13
         Top             =   0
         Width           =   4620
         Begin VB.TextBox Text2 
            Height          =   1050
            Left            =   555
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   330
            Width           =   3900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   315
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Text with Image"
         Height          =   2850
         Left            =   15
         TabIndex        =   32
         Top             =   1575
         Width           =   7695
         Begin VB.Frame Frame5 
            Height          =   2745
            Left            =   30
            TabIndex        =   33
            Top             =   15
            Width           =   7650
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " With Image(&I) "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   555
               Left            =   2220
               TabIndex        =   34
               Top             =   1170
               Width           =   3030
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Options"
            Height          =   2490
            Left            =   5490
            TabIndex        =   38
            Top             =   210
            Width           =   2070
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "Main.frx":0008
               Left            =   75
               List            =   "Main.frx":0012
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   540
               Width           =   1905
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Without Image(&X) "
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
               Left            =   105
               TabIndex        =   41
               Top             =   2025
               Width           =   1800
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               Height          =   195
               Left            =   120
               TabIndex        =   40
               Top             =   300
               Width           =   555
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Select a image"
            Height          =   2520
            Left            =   90
            TabIndex        =   35
            Top             =   210
            Width           =   5310
            Begin VB.FileListBox File1 
               Appearance      =   0  'Flat
               Height          =   1980
               Left            =   2655
               Pattern         =   "*.JPG;*.PNG"
               TabIndex        =   42
               Top             =   255
               Width           =   2370
            End
            Begin VB.DirListBox Dir1 
               Appearance      =   0  'Flat
               Height          =   1665
               Left            =   105
               TabIndex        =   37
               Top             =   570
               Width           =   2565
            End
            Begin VB.DriveListBox Drive1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   2535
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Parts"
         Height          =   1785
         Left            =   30
         TabIndex        =   28
         Top             =   4425
         Visible         =   0   'False
         Width           =   3405
         Begin VB.ListBox List2 
            Height          =   840
            Left            =   105
            TabIndex        =   29
            Top             =   180
            Width           =   3195
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Pay attention: After selecting, it will be copied to Text box, click the last one to restore your orginal text."
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
            TabIndex        =   30
            Top             =   1065
            Width           =   3165
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Format"
         Height          =   1545
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2895
         Begin VB.ComboBox AlignCombo 
            Height          =   315
            ItemData        =   "Main.frx":0040
            Left            =   810
            List            =   "Main.frx":004D
            Style           =   2  'Dropdown List
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   840
            Width           =   270
         End
         Begin VB.ComboBox FontCombo 
            Height          =   315
            ItemData        =   "Main.frx":0082
            Left            =   825
            List            =   "Main.frx":008F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   225
            Width           =   1905
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   825
            MaxLength       =   3
            TabIndex        =   19
            Top             =   570
            Width           =   1890
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alignment"
            Height          =   195
            Left            =   90
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
            Top             =   585
            Width           =   285
         End
      End
      Begin VB.Label Temp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Preview]"
         Height          =   195
         Left            =   5460
         TabIndex        =   31
         Top             =   5445
         Visible         =   0   'False
         Width           =   690
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
         Left            =   5220
         TabIndex        =   17
         Top             =   5790
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
         Left            =   6450
         TabIndex        =   16
         Top             =   5790
         Width           =   870
      End
   End
   Begin VB.Frame LogMgr 
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   6210
      Left            =   45
      TabIndex        =   81
      Top             =   1920
      Width           =   7725
      Begin VB.Frame Frame14 
         Caption         =   "Details"
         Height          =   3060
         Left            =   15
         TabIndex        =   85
         Top             =   2940
         Width           =   7530
         Begin VB.TextBox Text5 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   2820
            Left            =   75
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   86
            Top             =   165
            Width           =   7395
         End
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   285
         Left            =   5850
         TabIndex        =   82
         Top             =   2730
         Width           =   1815
         Begin VB.OptionButton Option2 
            Caption         =   "Formats"
            Height          =   195
            Left            =   855
            TabIndex        =   84
            Top             =   45
            Width           =   930
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pages"
            Height          =   195
            Left            =   30
            TabIndex        =   83
            Top             =   45
            Value           =   -1  'True
            Width           =   810
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Log List"
         Height          =   2865
         Left            =   15
         TabIndex        =   87
         Top             =   0
         Width           =   7560
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Caption         =   "Use"
            Height          =   330
            Left            =   6840
            TabIndex        =   88
            Top             =   2025
            Visible         =   0   'False
            Width           =   540
            Begin VB.Label Label31 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Use "
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
               Left            =   30
               TabIndex        =   89
               Top             =   15
               Width           =   450
            End
         End
         Begin VB.Frame Frame16 
            BorderStyle     =   0  'None
            Caption         =   "Use"
            Height          =   330
            Left            =   6840
            TabIndex        =   90
            Top             =   1680
            Visible         =   0   'False
            Width           =   540
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Del "
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
               Left            =   15
               TabIndex        =   91
               Top             =   0
               Width           =   480
            End
         End
         Begin VB.ListBox ListPage 
            Appearance      =   0  'Flat
            Height          =   2370
            Left            =   75
            TabIndex        =   92
            Top             =   210
            Width           =   7170
         End
         Begin VB.ListBox ListFormat 
            Appearance      =   0  'Flat
            Height          =   2370
            Left            =   75
            TabIndex        =   93
            Top             =   210
            Visible         =   0   'False
            Width           =   7170
         End
      End
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   6255
      Left            =   45
      TabIndex        =   48
      Top             =   1875
      Width           =   7800
      Begin VB.ListBox MsgContentList 
         Height          =   450
         ItemData        =   "Main.frx":00C4
         Left            =   2775
         List            =   "Main.frx":00C6
         TabIndex        =   50
         Top             =   1830
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Left            =   525
         TabIndex        =   49
         Top             =   1020
         Visible         =   0   'False
         Width           =   2100
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
         Left            =   2700
         TabIndex        =   52
         Top             =   1410
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   2685
         TabIndex        =   51
         Top             =   1035
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.Frame AnswerLine 
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      Height          =   6195
      Left            =   0
      TabIndex        =   53
      Top             =   1890
      Width           =   7920
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2490
         MaxLength       =   3
         TabIndex        =   67
         Top             =   390
         Width           =   2445
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   450
         MaxLength       =   3
         TabIndex        =   66
         Top             =   135
         Width           =   1890
      End
      Begin VB.CheckBox Check15 
         Caption         =   "cm"
         Height          =   255
         Left            =   2415
         TabIndex        =   65
         Top             =   120
         Width           =   1755
      End
      Begin VB.Frame Frame9 
         Caption         =   "Options"
         Height          =   1830
         Left            =   90
         TabIndex        =   68
         Top             =   780
         Width           =   2490
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            Height          =   195
            Left            =   1230
            TabIndex        =   80
            Top             =   420
            Width           =   195
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Check4"
            Height          =   240
            Left            =   420
            TabIndex        =   79
            Top             =   420
            Width           =   225
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Check4"
            Height          =   240
            Left            =   810
            TabIndex        =   78
            Top             =   420
            Width           =   225
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Check3"
            Height          =   195
            Left            =   1245
            TabIndex        =   77
            Top             =   1260
            Width           =   195
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Check4"
            Height          =   240
            Left            =   375
            TabIndex        =   76
            Top             =   1260
            Width           =   225
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Check4"
            Height          =   240
            Left            =   795
            TabIndex        =   75
            Top             =   1245
            Width           =   225
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Check4"
            Height          =   240
            Left            =   780
            TabIndex        =   74
            Top             =   165
            Width           =   225
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Check4"
            Height          =   240
            Left            =   795
            TabIndex        =   73
            Top             =   1545
            Width           =   225
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Check4"
            Height          =   240
            Left            =   1485
            TabIndex        =   72
            Top             =   840
            Width           =   225
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Check4"
            Height          =   240
            Left            =   135
            TabIndex        =   71
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Check4"
            Height          =   240
            Left            =   420
            TabIndex        =   70
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Check4"
            Height          =   240
            Left            =   1215
            TabIndex        =   69
            Top             =   840
            Width           =   225
         End
         Begin VB.Line Line1 
            X1              =   225
            X2              =   225
            Y1              =   270
            Y2              =   1650
         End
         Begin VB.Line Line3 
            X1              =   1590
            X2              =   1590
            Y1              =   285
            Y2              =   1665
         End
         Begin VB.Line Line2 
            X1              =   225
            X2              =   1605
            Y1              =   270
            Y2              =   270
         End
         Begin VB.Line Line4 
            X1              =   225
            X2              =   1605
            Y1              =   1650
            Y2              =   1650
         End
         Begin VB.Line Line5 
            X1              =   240
            X2              =   1590
            Y1              =   285
            Y2              =   1650
         End
         Begin VB.Line Line7 
            X1              =   240
            X2              =   1605
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line8 
            X1              =   900
            X2              =   900
            Y1              =   270
            Y2              =   1665
         End
         Begin VB.Line Line6 
            X1              =   1575
            X2              =   225
            Y1              =   285
            Y2              =   1635
         End
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "----"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1110
         TabIndex        =   95
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label33 
         Caption         =   "TrackNumber"
         Height          =   285
         Left            =   105
         TabIndex        =   94
         Top             =   4095
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Count(-1 for as much as possible)"
         Height          =   225
         Left            =   45
         TabIndex        =   58
         Top             =   435
         Width           =   2475
      End
      Begin VB.Label Label26 
         Caption         =   "Label26"
         Height          =   885
         Left            =   5370
         TabIndex        =   57
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Save "
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
         Left            =   960
         TabIndex        =   56
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Preview "
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
         Left            =   75
         TabIndex        =   55
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Size"
         Height          =   225
         Left            =   60
         TabIndex        =   54
         Top             =   150
         Width           =   1710
      End
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
      TabIndex        =   9
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
Dim showcnt As Integer, deltachange As Integer, current As Integer, strs() As String, heightdata As Integer, special() As Integer, specialinfo() As Integer, stats() As Boolean
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

Function GetCharID(X As String) As Integer
    If (Asc(X) >= Asc("a") And Asc(X) <= Asc("z")) Or (Asc(X) >= Asc("A") And Asc(X) <= Asc("Z")) Then
        GetCharID = 200
        Exit Function
    End If
    If X = " " Then
        GetCharID = 333
        Exit Function
    End If
    If Asc(X) > 32 And Asc(X) < 127 Then GetCharID = 999 Else GetCharID = 1233
End Function

Private Sub Check13_Click()
    Check14.Value = Check13.Value
End Sub

Private Sub Check14_Click()
    Check13.Value = Check14.Value
End Sub

Private Sub Check17_Click()
    'If Check17.Value = 1 Then Label29.Enabled = False Else Label29.Enabled = True
End Sub

Private Sub Check3_Click()
    Check7.Value = Check3.Value
End Sub

Private Sub Check4_Click()
    Check5.Value = Check4.Value
End Sub

Private Sub Check5_Click()
    Check4.Value = Check5.Value
End Sub

Private Sub Check6_Click()
    Check8.Value = Check6.Value
End Sub

Private Sub Check7_Click()
    Check3.Value = Check7.Value
End Sub

Private Sub Check8_Click()
    Check6.Value = Check8.Value
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    current = -1
    Dim i As Long
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

Private Sub Frame11_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label17_Click()
    Frame5.Visible = False
End Sub

Private Sub Label19_Click()
    Frame5.Visible = True
End Sub

Private Sub Label22_Click()
    AnswerLine.Visible = Not False
    InsText.Visible = Not True
End Sub

Private Sub Label24_Click()
    If Not IsNumeric(Text3.Text) Or Not IsNumeric(Text4.Text) Then
        NewMessage "Invaild Format.", vbRed, True
        Exit Sub
    End If
    'Label26.FontSize = Val(Text3.Text)
    Dim lim As Long, succ As Boolean
    succ = False
    lim = Val(Text4.Text)
    If lim = -1 Then lim = 32767
    InitPreview
    Dim p As Long
    p = LeftMargin
    Label26.Width = Val(Text3.Text) * (1 + (TwipsPerCM - 1) * Check15.Value)
    Label26.Height = Val(Text3.Text) * (1 + (TwipsPerCM - 1) * Check15.Value)
    While p + Label26.Width <= RightMargin And lim > 0
        succ = True
        lim = lim - 1
        If Check9.Value = 1 Then
            Preview.Picture2.Line (p, TopMargin)-(p + Label26.Width, TopMargin)
        End If
        If Check10.Value = 1 Then
            Preview.Picture2.Line (p, Label26.Height + TopMargin)-(p + Label26.Width, Label26.Height + TopMargin)
        End If
        If Check12.Value = 1 Then
            Preview.Picture2.Line (p, TopMargin)-(p, Label26.Height + TopMargin)
        End If
        If Check11.Value = 1 Then
            Preview.Picture2.Line (p + Label26.Width, TopMargin)-(p + Label26.Width, Label26.Height + TopMargin)
        End If
        If Check13.Value = 1 Then
            Preview.Picture2.Line (p, Label26.Height / 2 + TopMargin)-(p + Label26.Width, Label26.Height / 2 + TopMargin)
        End If
        If Check3.Value = 1 Then
            Preview.Picture2.Line (p, TopMargin)-(p + Label26.Width, Label26.Height + TopMargin)
        End If
        If Check5.Value = 1 Then
            Preview.Picture2.Line (p + Label26.Width, TopMargin)-(p, Label26.Height + TopMargin)
        End If
        If Check6.Value = 1 Then
            Preview.Picture2.Line (p + Label26.Width / 2, TopMargin)-(p + Label26.Width / 2, Label26.Height + TopMargin)
        End If
        p = p + Label26.Width
    Wend
    If Not succ Then
        NewMessage "The size you've input is too large", vbRed
        On Error Resume Next
        Unload Preview
    End If
End Sub

Private Sub Label25_Click()
    If Not IsNumeric(Text3.Text) Or Not IsNumeric(Text4.Text) Then
        NewMessage "Invaild Format.", vbRed, True
        Exit Sub
    End If
    'Label26.FontSize = Val(Text3.Text)
    Dim lim As Long, succ As Boolean
    succ = False
    lim = Val(Text4.Text)
    If lim = -1 Then lim = 32767
    InitPreview
    Dim p As Long
    p = LeftMargin
    Label26.Width = Val(Text3.Text) * (1 + (TwipsPerCM - 1) * Check15.Value)
    Label26.Height = Val(Text3.Text) * (1 + (TwipsPerCM - 1) * Check15.Value)
    While p + Label26.Width <= RightMargin And lim > 0
        succ = True
        lim = lim - 1
        If Check9.Value = 1 Then
            Preview.Picture2.Line (p, TopMargin)-(p + Label26.Width, TopMargin)
        End If
        If Check10.Value = 1 Then
            Preview.Picture2.Line (p, Label26.Height + TopMargin)-(p + Label26.Width, Label26.Height + TopMargin)
        End If
        If Check12.Value = 1 Then
            Preview.Picture2.Line (p, TopMargin)-(p, Label26.Height + TopMargin)
        End If
        If Check11.Value = 1 Then
            Preview.Picture2.Line (p + Label26.Width, TopMargin)-(p + Label26.Width, Label26.Height + TopMargin)
        End If
        If Check13.Value = 1 Then
            Preview.Picture2.Line (p, Label26.Height / 2 + TopMargin)-(p + Label26.Width, Label26.Height / 2 + TopMargin)
        End If
        If Check3.Value = 1 Then
            Preview.Picture2.Line (p, TopMargin)-(p + Label26.Width, Label26.Height + TopMargin)
        End If
        If Check5.Value = 1 Then
            Preview.Picture2.Line (p + Label26.Width, TopMargin)-(p, Label26.Height + TopMargin)
        End If
        If Check6.Value = 1 Then
            Preview.Picture2.Line (p + Label26.Width / 2, TopMargin)-(p + Label26.Width / 2, Label26.Height + TopMargin)
        End If
        p = p + Label26.Width
    Wend
    Preview.Exports.Height = Label26.Height + 50
    Preview.Exports.Width = p
    'Preview.Exports.BorderStyle = 0
    Preview.Exports.PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, TopMargin, p, Label26.Height + 50
    Preview.Exports.Visible = True
    If Not succ Then
        NewMessage "The size you've input is too large", vbRed
        On Error Resume Next
        Unload Preview
    End If
    Dim usage As Long
    usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
    If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
    SavePicture Preview.Exports.Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
    SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
    Label34.Caption = usage + 1
    List1.AddItem usage + 1
    On Error Resume Next
    Unload Preview
End Sub

Private Sub Label28_Click()
    Dim outstr As String
    outstr = FontCombo.Text & "," & Text1.Text & "," & Check2.Value & "," & Check1.Value & "," & AlignCombo.Text
    ListFormat.AddItem outstr
End Sub

Private Sub Label29_Click()
    Dim outstr As String
    outstr = FontCombo.Text & "," & Text1.Text & "," & Check2.Value & "," & Check1.Value & "," & AlignCombo.Text & "," & Text2.Text & "," & Frame5.Visible & "," & Drive1.Drive & "," & Dir1.Path & "," & File1.FileName & "," & Combo1.Text
    ListPage.AddItem outstr
End Sub

Private Sub Label30_Click()
    InsText.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = True
    Frame8.Visible = False
End Sub

Private Sub List2_Click()
    Text2.Text = strs(List2.ListIndex)
End Sub

Private Sub ListFormat_Click()
    Dim qqq() As String, i As Long
    Text5.Text = ""
    qqq = Split(ListFormat.Text, ",")
    Text5.Text = Text5.Text & "FontName: " & qqq(0) & vbCrLf
    Text5.Text = Text5.Text & "FontSize: " & qqq(1) & vbCrLf
    Text5.Text = Text5.Text & "Bold: " & qqq(2) & vbCrLf
    Text5.Text = Text5.Text & "Italtic: " & qqq(3) & vbCrLf
    Text5.Text = Text5.Text & "Alignment: " & qqq(4) & vbCrLf
    Text5.Text = Text5.Text & "* For Italtic/Bold: 1 is true, 0 is false"
End Sub

Private Sub ListPage_Click()
    Dim qqq() As String, i As Long
    Text5.Text = ""
    qqq = Split(ListPage.Text, ",")
    Text5.Text = Text5.Text & "FontName: " & qqq(0) & vbCrLf
    Text5.Text = Text5.Text & "FontSize: " & qqq(1) & vbCrLf
    Text5.Text = Text5.Text & "Bold: " & qqq(2) & vbCrLf
    Text5.Text = Text5.Text & "Italtic: " & qqq(3) & vbCrLf
    Text5.Text = Text5.Text & "Alignment: " & qqq(4) & vbCrLf
    Text5.Text = Text5.Text & "Text: " & qqq(5) & vbCrLf
    Text5.Text = Text5.Text & "DisabledImage: " & qqq(6) & vbCrLf
    If qqq(6) = "False" Then
        Text5.Text = Text5.Text & "Drive: " & qqq(7) & vbCrLf
        Text5.Text = Text5.Text & "Path: " & qqq(8) & vbCrLf
        Text5.Text = Text5.Text & "File: " & qqq(9) & vbCrLf
        Text5.Text = Text5.Text & "Position: " & qqq(10) & vbCrLf
    Else
        Text5.Text = Text5.Text & "[Image Information Unavailable]" & vbCrLf
    End If
    Text5.Text = Text5.Text & "* For Italtic/Bold: 1 is true, 0 is false"
End Sub

Private Sub Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub

Private Sub Option1_Click()
    ListPage.Visible = True
    ListFormat.Visible = False
End Sub

Private Sub Option2_Click()
    ListFormat.Visible = True
    ListPage.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 1000
End Sub


Private Sub PreviewButton_Click()
    InsText.Visible = True
    Frame8.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = False
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

Private Sub Timer1_Timer()
    Dim first As Long
    If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 100
    showcnt = showcnt + 1
'    If MsgContentList.ListCount <= 1 Then
'        first = showcnt
'        showcnt = ShowCntPerMsg
'        Message.Caption = ""
'        If MsgContentList.ListCount = 1 Then
'            current = 0
'            MsgContentList.ListIndex = current
'            MsgColorList.ListIndex = current
'            MsgTypeList.ListIndex = current
'            Message.Caption = MsgTypeList.Text & MsgContentList.Text
'            Message.ForeColor = ReverseColor(MsgColorList.Text)
'        End If
'        If showcnt <> first Then ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
'        Exit Sub
'    End If
    If MsgContentList.ListCount = 0 Then
        Message.Caption = "No new messages."
        Message.ForeColor = vbWhite
        showcnt = ShowCntPerMsg - 1
        GoTo rrr
    End If
    If current >= MsgContentList.ListCount Then
        Message.Caption = "No new messages."
        Message.ForeColor = vbWhite
        showcnt = ShowCntPerMsg - 1
        GoTo rrr
    End If
    If showcnt = ShowCntPerMsg Then
        current = current + 1
        showcnt = 0
        If MsgContentList.ListCount = 0 Then
            ProgressBar.Width = 15
            Message.Caption = ""
            Exit Sub
        End If
        If current >= MsgContentList.ListCount Then
            Message.Caption = "No new messages."
            Message.ForeColor = vbWhite
            showcnt = ShowCntPerMsg - 1
            GoTo rrr
        End If
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
    InsText.Visible = True
    AnswerLine.Visible = False
    LogMgr.Visible = False
    General.Visible = True
End Sub

Sub RegisterStat(StatName As String)
    If StatName = "b" Then
        Temp.FontBold = Not Temp.FontBold
        Preview.Picture2.FontBold = Not Preview.Picture2.FontBold
    End If
    If StatName = "i" Then
        Temp.FontItalic = Not Temp.FontItalic
        Preview.Picture2.FontItalic = Not Preview.Picture2.FontItalic
    End If
    If StatName = "u" Then
        Temp.FontUnderline = Not Temp.FontUnderline
        Preview.Picture2.FontUnderline = Not Preview.Picture2.FontUnderline
    End If
    If StatName = "ee" Then
        Temp.FontSize = Temp.FontSize * 0.7
        Preview.Picture2.FontSize = Preview.Picture2.FontSize * 0.7
    End If
    If StatName = "ed" Then
        Temp.FontSize = Temp.FontSize / 0.7
        Preview.Picture2.FontSize = Preview.Picture2.FontSize / 0.7
    End If
    If StatName = "se" Then
        heightdata = Temp.Height
        Temp.FontSize = Temp.FontSize * 0.7
        deltachange = deltachange + heightdata - Temp.Height
        'delta = delta + heightdata - Temp.Height
        Preview.Picture2.FontSize = Preview.Picture2.FontSize * 0.7
        stats(1) = True
    End If
    If StatName = "sd" Then
        Temp.AutoSize = True
        deltachange = deltachange - (heightdata - Temp.Height)
        'delta = delta - (heightdata - Temp.Height)
        Temp.FontSize = Temp.FontSize / 0.7
        Preview.Picture2.FontSize = Preview.Picture2.FontSize / 0.7
        stats(1) = False
    End If
End Sub

Private Sub Label10_Click()
    On Error Resume Next
    Dim i As Long
    ReDim stats(DefCnt)
    Dim statstr As String, recording As Boolean, delta As Long, reced As Boolean, partid As Long, bound As Long, start As Long, length As Long, j As Long, xdelta As Long, tmpstr As String, issel As Boolean
    delta = 0
    deltachange = 0
    wholestr = Text2.Text
    outputs = ""
    recording = False
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
    If Frame5.Visible = False And Combo1.Text = "" Then
        NewMessage "You have not selected the position of the image.", vbRed
        Exit Sub
    End If
    If Check17.Value = 1 Then Label29_Click
    InitPreview
    If Frame5.Visible = False Then
        On Error GoTo err
        reced = False
        Preview.Exports.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
        If Preview.Exports.Height > BotMargin - TopMargin Or Preview.Exports.Width > RightMargin - LeftMargin Then
            NewMessage "Your input is so large that we can't process it.", vbRed
            On Error Resume Next
            Unload Preview
            Exit Sub
        End If
        Debug.Print Preview.Exports.Height, BotMargin - TopMargin, Preview.Exports.Width, RightMargin - LeftMargin
        If Left(Combo1.Text, 1) = "0" Then
            Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
            LeftMargin = LeftMargin + Preview.Exports.Width
        Else
            Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
            RightMargin = RightMargin - Preview.Exports.Width
            Debug.Print RightMargin
        End If
        GoTo ooi
err:
        NewMessage "[TYPE=RUNTIME_ERROR][ERRORID=" & err.Number & "][ERRDESC.=" & err.Description & "]", vbRed
        On Error Resume Next
        Unload Preview
        Exit Sub
    Else
        reced = True
    End If
ooi:
'Debug.Print Text2.Text
    Temp.FontName = FontCombo.Text
    Temp.FontSize = Val(Text1.Text)
    Temp.Alignment = Val(Left(AlignCombo.Text, 1))
    
    Temp.Visible = True
    If Check2.Value = 1 Then Temp.FontBold = True Else Temp.FontBold = False
    If Check1.Value = 1 Then Temp.FontItalic = True Else Temp.FontItalic = False
    If Check2.Value = 1 Then Preview.Picture2.FontBold = True Else Preview.Picture2.FontBold = False
    If Check1.Value = 1 Then Preview.Picture2.FontItalic = True Else Preview.Picture2.FontItalic = False
    
    For i = start To bound Step 1
    str(i) = str(i) & " "
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
'                If partid = 1 Then
'                    NewMessage "The input will be split into multi parts", vbBlue
'                    NewMessage "select parts that you want to preview in the list.", vbBlue
'                    'Exit Sub
'                End If
                .Height = .Height + PageHeight
                '.Cls
                'delta = 0
                partid = partid + 1
                ReDim Preserve strs(partid - 2)
                strs(partid - 2) = outputs
                '.CurrentY = TopMargin
                'Preview.Picture2.Print "test";
                If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                List2.AddItem outputs
                outputs = ""
                If length = 0 Then GoTo cont
            End If
        End With
        For j = 1 To length
            Preview.Picture2.Top = -delta
            If Mid(str(i), j, 1) = "^" Then
                If statstr = "" Then
                    recording = True
                    GoTo nextj
                End If
                If statstr <> "" Then
                    RegisterStat statstr
                    statstr = ""
                    recording = False
                    GoTo nextj
                End If
            End If
            If recording Then
                statstr = statstr & Mid(str(i), j, 1)
                GoTo nextj
            End If
            Text2.Text = Mid(str(i), j, 1)
            If GetCharID(Text2.Text) = 200 And Check16.Value = 1 Then
                Dim ptr As Long
                ptr = j + 1
                While ptr <= length And (GetCharID(Mid(str(i), ptr, 1)) = 1233 Or GetCharID(Mid(str(i), ptr, 1)) = 200)
                    Text2.Text = Text2.Text & Mid(str(i), ptr, 1)
                    ptr = ptr + 1
                Wend
                j = ptr - 1
            End If
            Temp.Caption = Text2.Text
            With Preview.Picture2
                tmpstr = Text2.Text
                If Temp.Width + .CurrentX > RightMargin Then
                    If Temp.Alignment <> 0 Then
                        NewMessage "Auto split line is unsupportted for alignment mode 1 or 2.", vbRed
                        Text2.Text = wholestr
                        On Error Resume Next
                        If Not reced Then
                            reced = True
                            If Left(Combo1.Text, 1) = "0" Then
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                                LeftMargin = LeftMargin - Preview.Exports.Width
                            Else
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                                RightMargin = RightMargin + Preview.Exports.Width
                            End If
                        End If
                        Unload Preview
                        Exit Sub
                    End If
                    delta = delta + Temp.Height
                    If Not reced And delta > Preview.Exports.Height Then
                        reced = True
                        If Left(Combo1.Text, 1) = "0" Then
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                            LeftMargin = LeftMargin - Preview.Exports.Width
                        Else
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                            RightMargin = RightMargin + Preview.Exports.Width
                        End If
                    End If
                    If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                    .CurrentY = TopMargin + delta
                    'Print BotMargin
                    'Preview.Picture2.Line (0, Preview.Picture2.CurrentY + Temp.Height)-(Preview.Picture2.Width, Preview.Picture2.CurrentY + Temp.Height), vbRed
                    If .CurrentY + Temp.Height >= BotMargin Then
                        If partid = 1 Then
                            NewMessage "The input will be split into multi parts", vbBlue
                            NewMessage "select parts that you want to preview in the list.", vbBlue
                            'Exit Sub
                        End If
                        '.Cls
                        .Height = .Height + PageHeight
                        partid = partid + 1
                        ReDim Preserve strs(partid - 2)
                        strs(partid - 2) = outputs
                        '.CurrentY = TopMargin
                        'Preview.Picture2.Print "test";
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
                    If Not reced Then
                        reced = True
                        If Left(Combo1.Text, 1) = "0" Then
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                            LeftMargin = LeftMargin - Preview.Exports.Width
                        Else
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                            RightMargin = RightMargin + Preview.Exports.Width
                        End If
                    End If
                    Exit Sub
                End If
            Preview.Picture2.CurrentY = Preview.Picture2.CurrentY + deltachange
            Preview.Picture2.Print tmpstr;
            Preview.Picture2.CurrentY = Preview.Picture2.CurrentY - deltachange
            outputs = outputs & tmpstr
            'Debug.Print .CurrentX
            End With
            DoEvents
nextj:
        Next
        delta = delta + Temp.Height
        If Not reced And delta > Preview.Exports.Height Then
            reced = True
            If Left(Combo1.Text, 1) = "0" Then
                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                LeftMargin = LeftMargin - Preview.Exports.Width
            Else
                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                RightMargin = RightMargin + Preview.Exports.Width
            End If
        End If
        If i <> bound Then outputs = outputs & vbCrLf
cont:
    Next
'    wholestr = str(0)
'    If UBound(str) > 0 Then wholestr = wholestr & vbCrLf
'    For i = 1 To UBound(str)
'        wholestr = wholestr & str(i) & vbCrLf
'    Next
    If partid > 1 Then
'        List2.AddItem outputs
'        List2.AddItem wholestr
'        ReDim Preserve strs(partid)
'        strs(partid - 1) = wholestr
'        strs(partid - 0) = wholestr
'        Frame3.Visible = True
'        On Error Resume Next
'        Unload Preview
    End If
    If Not reced Then
        reced = True
        If Left(Combo1.Text, 1) = "0" Then
            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
            LeftMargin = LeftMargin - Preview.Exports.Width
        Else
            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
            RightMargin = RightMargin + Preview.Exports.Width
        End If
    End If
    Preview.VScroll1.Max = Preview.Picture2.Height
    Text2.Text = wholestr
End Sub

Private Sub Label11_Click()
    On Error Resume Next
    Dim i As Long, usage As Long
    Dim delta As Long, recordid As Long, recording As Boolean, statstr As String, lastcapt As Long, orglen As Long, reced As Boolean, partid As Long, bound As Long, start As Long, length As Long, j As Long, xdelta As Long, tmpstr As String, issel As Boolean
    orglen = RightMargin - LeftMargin
    deltachange = 0
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
    If Frame5.Visible = False And Combo1.Text = "" Then
        NewMessage "You have not selected the position of the image.", vbRed
        Exit Sub
    End If
    If Check17.Value = 1 Then Label29_Click
    InitPreview
    If Frame5.Visible = False Then
        On Error GoTo err
        reced = False
        Preview.Exports.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
        If Preview.Exports.Height > BotMargin - TopMargin Or Preview.Exports.Width > RightMargin - LeftMargin Then
            NewMessage "Your input is so large that we can't process it.", vbRed
            On Error Resume Next
            Unload Preview
            Exit Sub
        End If
        If Left(Combo1.Text, 1) = "0" Then
            'Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
            LeftMargin = LeftMargin + Preview.Exports.Width
        Else
            'Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
            RightMargin = RightMargin - Preview.Exports.Width
            Debug.Print RightMargin
        End If
        usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
        If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
        SavePicture Preview.Exports.Picture, App.Path & "\Cache\" & usage + 1 & ".jpg"
        SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
        List1.AddItem "P" & Left(Combo1.Text, 1) & usage + 1
        recordid = List1.ListCount - 1
        GoTo ooi
err:
        NewMessage "[TYPE=RUNTIME_ERROR][ERRORID=" & err.Number & "][ERRDESC.=" & err.Description & "]", vbRed
        On Error Resume Next
        Unload Preview
        Exit Sub
    Else
        reced = True
    End If
ooi:
    'Debug.Print Text2.Text
    Temp.FontName = FontCombo.Text
    Temp.FontSize = Val(Text1.Text)
    Temp.Alignment = Val(Left(AlignCombo.Text, 1))
    
    If Check2.Value = 1 Then Temp.FontBold = True Else Temp.FontBold = False
    If Check1.Value = 1 Then Temp.FontItalic = True Else Temp.FontItalic = False
    If Check2.Value = 1 Then Preview.Picture2.FontBold = True Else Preview.Picture2.FontBold = False
    If Check1.Value = 1 Then Preview.Picture2.FontItalic = True Else Preview.Picture2.FontItalic = False
    For i = start To bound Step 1
        str(i) = str(i) & " "
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
                '.Cls
                'delta = 0
                partid = partid + 1
                .Height = .Height + PageWidth
                ReDim Preserve strs(partid - 2)
                strs(partid - 2) = outputs
                '.CurrentY = TopMargin
                'Preview.Picture2.Print "test";
                If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                List2.AddItem outputs
                outputs = ""
                If length = 0 Then GoTo cont
            End If
        End With
        For j = 1 To length
            Preview.Picture2.Top = -delta
            If Mid(str(i), j, 1) = "^" Then
                If statstr = "" Then
                    recording = True
                    GoTo nextj
                End If
                If statstr <> "" Then
                    RegisterStat statstr
                    statstr = ""
                    recording = False
                    GoTo nextj
                End If
            End If
            If recording Then
                statstr = statstr & Mid(str(i), j, 1)
                GoTo nextj
            End If
            Text2.Text = Mid(str(i), j, 1)
            If GetCharID(Text2.Text) = 200 And Check16.Value = 1 Then
                Dim ptr As Long
                ptr = j + 1
                While ptr <= length And (GetCharID(Mid(str(i), ptr, 1)) = 1233 Or GetCharID(Mid(str(i), ptr, 1)) = 200)
                    Text2.Text = Text2.Text & Mid(str(i), ptr, 1)
                    ptr = ptr + 1
                Wend
                j = ptr - 1
            End If
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
                        If Not reced Then
                            reced = True
                            If Left(Combo1.Text, 1) = "0" Then
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                                LeftMargin = LeftMargin - Preview.Exports.Width
                            Else
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                                RightMargin = RightMargin + Preview.Exports.Width
                            End If
                        End If
                        Exit Sub
                    End If
                    delta = delta + Temp.Height
                    If Not reced And delta > Preview.Exports.Height Then
                        reced = True
                        If Left(Combo1.Text, 1) = "0" Then
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                            LeftMargin = LeftMargin - Preview.Exports.Width
                        Else
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                            RightMargin = RightMargin + Preview.Exports.Width
                        End If
                    End If
                    If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                    .CurrentY = TopMargin + delta
                    Print BotMargin
                    'Preview.Picture2.Line (0, Preview.Picture2.CurrentY + Temp.Height)-(Preview.Picture2.Width, Preview.Picture2.CurrentY + Temp.Height), vbRed
                    With Preview.Export
                        'Temp.Caption = outputs
                        .Width = orglen
                        .Height = Temp.Height
                        DoEvents
                        .BorderStyle = 0
                        If Not reced Then
                            If Left(Combo1.Text, 1) = "0" Then
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
                                LeftMargin = LeftMargin - Preview.Exports.Width
                                Debug.Print "from " & LeftMargin
                            Else
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
                                RightMargin = RightMargin + Preview.Exports.Width
                                Debug.Print "from " & RightMargin
                            End If
                        End If
                        .PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, lastcapt, orglen, Temp.Height
                        If Not reced Then
                            If Left(Combo1.Text, 1) = "0" Then
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
                                LeftMargin = LeftMargin + Preview.Exports.Width
                                Debug.Print " to " & LeftMargin
                            Else
                                'Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
                                RightMargin = RightMargin - Preview.Exports.Width
                                Debug.Print " to " & RightMargin
                            End If
                        End If
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
                        '.Cls
                        'delta = 0
                        .Height = .Height + PageHeight
                        partid = partid + 1
                        ReDim Preserve strs(partid - 2)
                        strs(partid - 2) = outputs
                        '.CurrentY = TopMargin
                        'Preview.Picture2.Print "test";
                        If Temp.Alignment = 0 Then .CurrentX = LeftMargin Else If Temp.Alignment = 1 Then .CurrentX = Max(RightMargin - Temp.Width, LeftMargin) Else .CurrentX = Max((LeftMargin + RightMargin) / 2 - Temp.Width / 2, LeftMargin)
                        List2.AddItem outputs
                        outputs = ""
                        
                    End If
                    
                End If
                If Temp.Width > PageWidth - LeftMargin - (PageWidth - RightMargin) Or Temp.Height > PageHeight - TopMargin - (PageHeight - BotMargin) Then
                    NewMessage "The target size is too large, we are unable to process it.", vbRed
                    Text2.Text = wholestr
                    If Not reced Then
                        reced = True
                        If Left(Combo1.Text, 1) = "0" Then
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                            LeftMargin = LeftMargin - Preview.Exports.Width
                        Else
                            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                            RightMargin = RightMargin + Preview.Exports.Width
                        End If
                    End If
                    On Error Resume Next
                    Unload Preview
                    Exit Sub
                End If
            Preview.Picture2.Print tmpstr;
            outputs = outputs & tmpstr
            'Debug.Print .CurrentX
            End With
            DoEvents
nextj:
        Next
        With Preview.Export
            .Width = orglen
            .Height = Temp.Height
            DoEvents
            .BorderStyle = 0
            If Not reced Then
                If Left(Combo1.Text, 1) = "0" Then
                    'Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
                    LeftMargin = LeftMargin - Preview.Exports.Width
                    Debug.Print "from " & LeftMargin
                Else
                    'Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
                    RightMargin = RightMargin + Preview.Exports.Width
                    Debug.Print "from " & RightMargin
                End If
            End If
            .PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, lastcapt, orglen, Temp.Height
            If Not reced Then
                If Left(Combo1.Text, 1) = "0" Then
                    'Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
                    LeftMargin = LeftMargin + Preview.Exports.Width
                    Debug.Print " to " & LeftMargin
                Else
                    'Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
                    RightMargin = RightMargin - Preview.Exports.Width
                    Debug.Print " to " & RightMargin
                End If
            End If
            lastcapt = lastcapt + Temp.Height
            usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
            If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
            SavePicture .Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
            SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
            List1.AddItem usage + 1
        End With
        delta = delta + Temp.Height
        If Not reced And delta > Preview.Exports.Height Then
            reced = True
            If Left(Combo1.Text, 1) = "0" Then
                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
                LeftMargin = LeftMargin - Preview.Exports.Width
            Else
                'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
                RightMargin = RightMargin + Preview.Exports.Width
            End If
        End If
        If i <> bound Then outputs = outputs & vbCrLf
cont:
    Next
'    wholestr = str(0)
'    If UBound(str) > 0 Then wholestr = wholestr & vbCrLf
'    For i = 1 To UBound(str)
'        wholestr = wholestr & str(i) & vbCrLf
'    Next
'    If partid > 1 Then
'        List2.AddItem outputs
'        List2.AddItem wholestr
'        ReDim Preserve strs(partid)
'        strs(partid - 1) = wholestr
'        strs(partid - 0) = wholestr
'        Frame3.Visible = True
'        On Error Resume Next
'        Unload Preview
'    End If
    If Not reced Then
        reced = True
        List1.AddItem "BT" & Preview.Exports.Height
        If Left(Combo1.Text, 1) = "0" Then
            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, LeftMargin
            LeftMargin = LeftMargin - Preview.Exports.Width
        Else
            'Preview.Picture2.PaintPicture Preview.Exports.Picture, TopMargin, RightMargin - Preview.Exports.Width
            RightMargin = RightMargin + Preview.Exports.Width
        End If
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
    InsText.Visible = False
    AnswerLine.Visible = False
    General.Visible = False
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
    Dim i As Long
    On Error Resume Next
    Dim X As Long, cnt As Long, pos As Long
    X = TopMargin
    InitPreview
    ReDim special(0)
    ReDim specialinfo(0)
    For i = 0 To List1.ListCount - 1
        List1.ListIndex = i
        Debug.Print X
        If Left(List1.Text, 1) = "P" Then
            cnt = cnt + 1
            ReDim Preserve special(cnt)
            ReDim Preserve specialinfo(cnt)
            special(cnt) = i
            specialinfo(cnt) = X
            Debug.Print "CNT=" & List1.ListIndex
            GoTo nfor
        End If
        If Left(List1.Text, 1) = "B" Then
            If Mid(List1.Text, 2, 1) = "T" Then X = X + Val(Right(List1.Text, Len(List1.Text) - 2))
        End If
        If X > Preview.Picture2.Height Then
            Preview.Picture2.Height = Preview.Picture2.Height + PageWidth
        End If
        Preview.Picture2.Top = -X
        DoEvents
        Debug.Print "CurrentI" & i
        Preview.Picture2.PaintPicture LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg"), LeftMargin, X
        Preview.Export.Width = 1
        Preview.Export.Height = 1
        Preview.Export.Visible = True
        Preview.Export.Picture = LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg")
        X = X + Preview.Export.Height
        Debug.Print Preview.Export.Height
nfor:
    Next
    For i = 1 To cnt Step 1
        List1.ListIndex = special(i)
        Debug.Print "Index=" & List1.ListIndex
        If Left(List1.Text, 1) = "P" Then
            X = specialinfo(i)
            Preview.Exports.Picture = LoadPicture(App.Path & "\Cache\" & Right(List1.Text, Len(List1.Text) - 2) & ".jpg")
            If Mid(List1.Text, 2, 1) = "0" Then pos = LeftMargin Else pos = RightMargin - Preview.Exports.Width
            Debug.Print "TrackID=" & Right(List1.Text, Len(List1.Text) - 2)
            Preview.Picture2.PaintPicture Preview.Exports.Picture, pos, X
        End If
    Next
    Preview.Export.Visible = False
End Sub

