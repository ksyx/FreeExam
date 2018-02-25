VERSION 5.00
Begin VB.Form MainFrm 
   BackColor       =   &H00A0ACBA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ExamPaper Editor"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7920
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   7920
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox WIP 
      BackColor       =   &H00A0ACBA&
      ForeColor       =   &H00B4BFCC&
      Height          =   8250
      Left            =   99999
      ScaleHeight     =   8190
      ScaleWidth      =   7815
      TabIndex        =   42
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
         ForeColor       =   &H00656D76&
         Height          =   765
         Left            =   15
         TabIndex        =   44
         Top             =   1125
         Width           =   7755
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00656D76&
         Height          =   630
         Left            =   60
         TabIndex        =   43
         Top             =   465
         Width           =   7710
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6870
      Top             =   5175
   End
   Begin VB.ListBox MsgTypeList 
      Height          =   450
      ItemData        =   "Main.frx":0000
      Left            =   3330
      List            =   "Main.frx":0002
      TabIndex        =   7
      Top             =   -15
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox MsgColorList 
      Height          =   450
      ItemData        =   "Main.frx":0004
      Left            =   2760
      List            =   "Main.frx":0006
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00656D76&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   7785
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8325
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
         TabIndex        =   5
         Top             =   10
         Width           =   45
      End
   End
   Begin VB.Frame InsText 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6315
      Left            =   75
      TabIndex        =   11
      Top             =   1950
      Width           =   7785
      Begin VB.Frame Frame10 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Logger"
         ForeColor       =   &H00656D76&
         Height          =   825
         Left            =   3645
         TabIndex        =   58
         Top             =   4545
         Width           =   3675
         Begin VB.CheckBox Check17 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            Caption         =   "Auto"
            ForeColor       =   &H00656D76&
            Height          =   225
            Left            =   1875
            TabIndex        =   62
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00B4BFCC&
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
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   960
            TabIndex        =   60
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00B4BFCC&
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
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   105
            TabIndex        =   59
            Top             =   225
            Width           =   780
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   330
            Index           =   12
            Left            =   90
            Top             =   195
            Width           =   825
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   330
            Index           =   13
            Left            =   945
            Top             =   210
            Width           =   1830
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Options"
         ForeColor       =   &H00656D76&
         Height          =   1560
         Left            =   3030
         TabIndex        =   12
         Top             =   30
         Width           =   4620
         Begin VB.CheckBox Check16 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "English Mode"
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   165
            TabIndex        =   132
            Top             =   660
            Value           =   1  'Checked
            Width           =   1320
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H00656D76&
            Height          =   225
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   1005
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label Label51 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            Caption         =   " Click to edit "
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
            Left            =   780
            TabIndex        =   131
            Top             =   330
            Width           =   3555
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text &^"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   315
            Width           =   495
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   330
            Index           =   9
            Left            =   765
            Top             =   300
            Width           =   3585
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Text with Image"
         ForeColor       =   &H00656D76&
         Height          =   2850
         Left            =   15
         TabIndex        =   31
         Top             =   1605
         Width           =   7695
         Begin VB.Frame Frame5 
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   2850
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   7695
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00B4BFCC&
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
               ForeColor       =   &H00656D76&
               Height          =   525
               Left            =   2340
               TabIndex        =   33
               Top             =   1155
               Width           =   3000
            End
            Begin VB.Shape Shape0 
               BackColor       =   &H00B4BFCC&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00B4BFCC&
               Height          =   525
               Index           =   10
               Left            =   2325
               Top             =   1155
               Width           =   3030
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00A0ACBA&
            Caption         =   "Options"
            ForeColor       =   &H00656D76&
            Height          =   2490
            Left            =   5490
            TabIndex        =   37
            Top             =   210
            Width           =   2070
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00B4BFCC&
               ForeColor       =   &H00656D76&
               Height          =   315
               ItemData        =   "Main.frx":0008
               Left            =   75
               List            =   "Main.frx":0012
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   540
               Width           =   1905
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00B4BFCC&
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
               ForeColor       =   &H00656D76&
               Height          =   255
               Left            =   105
               TabIndex        =   40
               Top             =   2025
               Width           =   1770
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               ForeColor       =   &H00656D76&
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   300
               Width           =   555
            End
            Begin VB.Shape Shape0 
               BackColor       =   &H00B4BFCC&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00B4BFCC&
               Height          =   330
               Index           =   11
               Left            =   90
               Top             =   1980
               Width           =   1845
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00A0ACBA&
            Caption         =   "Select a image"
            ForeColor       =   &H00656D76&
            Height          =   2520
            Left            =   90
            TabIndex        =   34
            Top             =   210
            Width           =   5310
            Begin VB.FileListBox File1 
               Appearance      =   0  'Flat
               BackColor       =   &H00B4BFCC&
               ForeColor       =   &H00656D76&
               Height          =   1980
               Left            =   2655
               Pattern         =   "*.JPG;*.PNG"
               TabIndex        =   41
               Top             =   255
               Width           =   2370
            End
            Begin VB.DirListBox Dir1 
               Appearance      =   0  'Flat
               BackColor       =   &H00B4BFCC&
               ForeColor       =   &H00656D76&
               Height          =   1665
               Left            =   105
               TabIndex        =   36
               Top             =   570
               Width           =   2565
            End
            Begin VB.DriveListBox Drive1 
               Appearance      =   0  'Flat
               BackColor       =   &H00B4BFCC&
               ForeColor       =   &H00656D76&
               Height          =   315
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   2535
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Parts"
         ForeColor       =   &H00656D76&
         Height          =   1785
         Left            =   1350
         TabIndex        =   27
         Top             =   6180
         Visible         =   0   'False
         Width           =   3405
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   810
            Left            =   105
            TabIndex        =   28
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
            ForeColor       =   &H00656D76&
            Height          =   645
            Left            =   90
            TabIndex        =   29
            Top             =   1065
            Width           =   3165
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Format"
         ForeColor       =   &H00656D76&
         Height          =   1545
         Left            =   30
         TabIndex        =   17
         Top             =   0
         Width           =   2895
         Begin VB.ComboBox AlignCombo 
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   315
            ItemData        =   "Main.frx":0040
            Left            =   810
            List            =   "Main.frx":004D
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1080
            Width           =   1905
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
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
            ForeColor       =   &H00656D76&
            Height          =   225
            Left            =   825
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   855
            Width           =   270
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00B4BFCC&
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
            ForeColor       =   &H00656D76&
            Height          =   225
            Left            =   1110
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   855
            Width           =   270
         End
         Begin VB.ComboBox FontCombo 
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   315
            ItemData        =   "Main.frx":0082
            Left            =   825
            List            =   "Main.frx":008F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   225
            Width           =   1905
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   825
            MaxLength       =   3
            TabIndex        =   18
            Top             =   570
            Width           =   1890
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alignment"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   90
            TabIndex        =   26
            Top             =   1155
            Width           =   705
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shape"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   345
            TabIndex        =   25
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   405
            TabIndex        =   24
            Top             =   285
            Width           =   360
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   435
            TabIndex        =   23
            Top             =   585
            Width           =   315
         End
      End
      Begin VB.Label Temp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Preview]"
         Height          =   195
         Left            =   1095
         TabIndex        =   30
         Top             =   5145
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   5310
         TabIndex        =   16
         Top             =   5790
         Width           =   1095
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   6630
         TabIndex        =   15
         Top             =   5790
         Width           =   840
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   14
         Left            =   5310
         Top             =   5775
         Width           =   1110
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   15
         Left            =   6630
         Top             =   5775
         Width           =   915
      End
   End
   Begin VB.Frame SaveLoad 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   6465
      Left            =   15
      TabIndex        =   162
      Top             =   1800
      Width           =   7905
      Begin VB.Frame Frame23 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Load"
         ForeColor       =   &H00656D76&
         Height          =   1095
         Left            =   105
         TabIndex        =   166
         Top             =   855
         Width           =   7560
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   825
            TabIndex        =   169
            Top             =   315
            Width           =   1440
         End
         Begin VB.Label Label70 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "LoadID"
            ForeColor       =   &H00656D76&
            Height          =   180
            Left            =   270
            TabIndex        =   168
            Top             =   345
            Width           =   570
         End
         Begin VB.Label Label69 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00B4BFCC&
            Caption         =   " Load "
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
            Left            =   270
            TabIndex        =   167
            Top             =   720
            Width           =   570
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   330
            Index           =   32
            Left            =   240
            Top             =   690
            Width           =   630
         End
      End
      Begin VB.Frame Frame22 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Save"
         ForeColor       =   &H00656D76&
         Height          =   660
         Left            =   105
         TabIndex        =   163
         Top             =   150
         Width           =   7560
         Begin VB.Label Label66 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00B4BFCC&
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
            ForeColor       =   &H00656D76&
            Height          =   255
            Left            =   255
            TabIndex        =   165
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label67 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   180
            Left            =   855
            TabIndex        =   164
            Top             =   705
            Width           =   3750
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   330
            Index           =   31
            Left            =   240
            Top             =   240
            Width           =   630
         End
      End
   End
   Begin VB.Frame Blk 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   6465
      Left            =   -30
      TabIndex        =   105
      Top             =   1965
      Width           =   7905
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00656D76&
         Height          =   285
         Left            =   2730
         TabIndex        =   110
         Text            =   "1"
         Top             =   210
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CheckBox Check18 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0ACBA&
         Caption         =   "cm"
         ForeColor       =   &H00656D76&
         Height          =   195
         Left            =   1980
         TabIndex        =   108
         Top             =   255
         Width           =   510
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00656D76&
         Height          =   285
         Left            =   495
         TabIndex        =   107
         Top             =   195
         Width           =   1440
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   150
         TabIndex        =   112
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0ACBA&
         Caption         =   "Lines"
         ForeColor       =   &H00656D76&
         Height          =   180
         Left            =   4200
         TabIndex        =   111
         Top             =   270
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label40 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0ACBA&
         Caption         =   "X"
         ForeColor       =   &H00656D76&
         Height          =   180
         Left            =   2520
         TabIndex        =   109
         Top             =   255
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0ACBA&
         Caption         =   "Size"
         ForeColor       =   &H00656D76&
         Height          =   180
         Left            =   90
         TabIndex        =   106
         Top             =   225
         Width           =   390
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   18
         Left            =   120
         Top             =   540
         Width           =   870
      End
   End
   Begin VB.PictureBox Copyright 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Height          =   7590
      Left            =   60
      ScaleHeight     =   7590
      ScaleWidth      =   7770
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   7770
      Begin VB.Label Label71 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0ACBA&
         Caption         =   "Version: Demo 2.8.1.20180226"
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
         Left            =   180
         TabIndex        =   170
         Top             =   435
         Width           =   2910
      End
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0ACBA&
         Caption         =   " Copyright (c) ksyx 2018, All Rights Reserved. "
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
         Left            =   120
         TabIndex        =   160
         Top             =   105
         Width           =   4290
      End
   End
   Begin VB.PictureBox Tools 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   60
      ScaleHeight     =   1380
      ScaleWidth      =   7770
      TabIndex        =   171
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   7770
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: All of the tools there is provided by system."
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
         Left            =   135
         TabIndex        =   175
         Top             =   480
         Width           =   4515
      End
      Begin VB.Label Label73 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Paint "
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
         Left            =   2430
         TabIndex        =   174
         Top             =   150
         Width           =   570
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   35
         Left            =   2370
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label76 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Calcutator "
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
         Left            =   105
         TabIndex        =   173
         Top             =   120
         Width           =   1065
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   33
         Left            =   90
         Top             =   105
         Width           =   1155
      End
      Begin VB.Label Label75 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Notepad "
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
         Left            =   1335
         TabIndex        =   172
         Top             =   150
         Width           =   900
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   315
         Index           =   34
         Left            =   1320
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Manage 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   75
      ScaleHeight     =   1380
      ScaleWidth      =   7770
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   7770
      Begin VB.Label Label64 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Save/Load "
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
         Left            =   3480
         TabIndex        =   161
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label63 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " List "
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
         Left            =   3000
         TabIndex        =   158
         Top             =   165
         Width           =   420
      End
      Begin VB.Label Label60 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Merge "
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
         Left            =   2235
         TabIndex        =   144
         Top             =   135
         Width           =   675
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   1590
         TabIndex        =   61
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   120
         Width           =   1380
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   19
         Left            =   90
         Top             =   105
         Width           =   1410
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   315
         Index           =   20
         Left            =   1605
         Top             =   105
         Width           =   540
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   21
         Left            =   2250
         Top             =   105
         Width           =   660
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   22
         Left            =   2970
         Top             =   120
         Width           =   450
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   30
         Left            =   3465
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.PictureBox General 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   75
      ScaleHeight     =   1380
      ScaleWidth      =   7770
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   525
      Width           =   7770
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " ABCD(&O) "
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
         Left            =   4785
         TabIndex        =   113
         Top             =   60
         Width           =   945
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Blank Area(&B) "
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
         Left            =   3390
         TabIndex        =   96
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Picture(&P) "
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
         Left            =   2310
         TabIndex        =   94
         Top             =   45
         Width           =   1020
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   930
         TabIndex        =   46
         Top             =   45
         Width           =   1335
      End
      Begin VB.Label PreviewButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   75
         TabIndex        =   45
         Top             =   45
         Width           =   810
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   285
         Index           =   25
         Left            =   30
         Top             =   30
         Width           =   870
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   300
         Index           =   26
         Left            =   930
         Top             =   30
         Width           =   1365
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   300
         Index           =   27
         Left            =   2310
         Top             =   30
         Width           =   1035
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   285
         Index           =   28
         Left            =   3360
         Top             =   60
         Width           =   1395
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   285
         Index           =   29
         Left            =   4770
         Top             =   45
         Width           =   990
      End
   End
   Begin VB.Frame Merge 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   6165
      Left            =   0
      TabIndex        =   145
      Top             =   1965
      Visible         =   0   'False
      Width           =   7920
      Begin VB.Frame Frame19 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Header"
         ForeColor       =   &H00656D76&
         Height          =   660
         Left            =   210
         TabIndex        =   154
         Top             =   345
         Width           =   7560
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   120
            TabIndex        =   156
            Top             =   255
            Width           =   4485
         End
         Begin VB.CheckBox Check23 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Split Line"
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   4710
            TabIndex        =   155
            Top             =   240
            Value           =   1  'Checked
            Width           =   1320
         End
         Begin VB.Label Label62 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   180
            Left            =   855
            TabIndex        =   157
            Top             =   705
            Width           =   3750
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Footer"
         ForeColor       =   &H00656D76&
         Height          =   1020
         Left            =   180
         TabIndex        =   147
         Top             =   1230
         Width           =   7560
         Begin VB.CheckBox Check22 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Split Line"
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   6390
            TabIndex        =   153
            Top             =   270
            Value           =   1  'Checked
            Width           =   1110
         End
         Begin VB.ComboBox Combo5 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   315
            ItemData        =   "Main.frx":00C4
            Left            =   4020
            List            =   "Main.frx":00D4
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   255
            Width           =   2295
         End
         Begin VB.ComboBox Combo3 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   315
            ItemData        =   "Main.frx":0115
            Left            =   3345
            List            =   "Main.frx":0125
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   240
            Width           =   660
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   120
            TabIndex        =   148
            Top             =   255
            Width           =   3120
         End
         Begin VB.Label Label65 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00A0ACBA&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   855
            TabIndex        =   152
            Top             =   705
            Width           =   105
         End
         Begin VB.Label Label61 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Preview"
            ForeColor       =   &H00656D76&
            Height          =   180
            Left            =   165
            TabIndex        =   150
            Top             =   675
            Width           =   585
         End
      End
      Begin VB.Label Label59 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
         Caption         =   " Merge "
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
         Left            =   120
         TabIndex        =   146
         Top             =   2430
         Width           =   675
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   8
         Left            =   75
         Top             =   2415
         Width           =   810
      End
   End
   Begin VB.Frame AnswerLine 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      Height          =   6195
      Left            =   -30
      TabIndex        =   52
      Top             =   1860
      Width           =   7920
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00656D76&
         Height          =   285
         Left            =   2490
         MaxLength       =   3
         TabIndex        =   65
         Top             =   390
         Width           =   2445
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00656D76&
         Height          =   285
         Left            =   495
         MaxLength       =   3
         TabIndex        =   64
         Top             =   135
         Width           =   1890
      End
      Begin VB.CheckBox Check15 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0ACBA&
         Caption         =   "cm"
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   2415
         TabIndex        =   63
         Top             =   120
         Width           =   1755
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Options"
         ForeColor       =   &H00656D76&
         Height          =   1830
         Left            =   90
         TabIndex        =   66
         Top             =   780
         Width           =   2490
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check3"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1230
            TabIndex        =   78
            Top             =   420
            Width           =   195
         End
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   420
            TabIndex        =   77
            Top             =   420
            Width           =   225
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   810
            TabIndex        =   76
            Top             =   420
            Width           =   225
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check3"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1245
            TabIndex        =   75
            Top             =   1260
            Width           =   195
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   375
            TabIndex        =   74
            Top             =   1260
            Width           =   225
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   795
            TabIndex        =   73
            Top             =   1245
            Width           =   225
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   780
            TabIndex        =   72
            Top             =   165
            Width           =   225
         End
         Begin VB.CheckBox Check10 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   795
            TabIndex        =   71
            Top             =   1545
            Width           =   225
         End
         Begin VB.CheckBox Check11 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1485
            TabIndex        =   70
            Top             =   840
            Width           =   225
         End
         Begin VB.CheckBox Check12 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   135
            TabIndex        =   69
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox Check13 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   420
            TabIndex        =   68
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox Check14 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Check4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   67
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
         BackColor       =   &H00A0ACBA&
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
         ForeColor       =   &H00656D76&
         Height          =   675
         Left            =   1110
         TabIndex        =   93
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label33 
         BackColor       =   &H00A0ACBA&
         Caption         =   "TrackNumber"
         ForeColor       =   &H00656D76&
         Height          =   285
         Left            =   105
         TabIndex        =   92
         Top             =   4095
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Count(-1 for as much as possible)"
         ForeColor       =   &H00656D76&
         Height          =   225
         Left            =   45
         TabIndex        =   57
         Top             =   435
         Width           =   2475
      End
      Begin VB.Label Label26 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Label26"
         ForeColor       =   &H00656D76&
         Height          =   885
         Left            =   5370
         TabIndex        =   56
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   1245
         TabIndex        =   55
         Top             =   2850
         Width           =   585
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   210
         TabIndex        =   54
         Top             =   2835
         Width           =   825
      End
      Begin VB.Label Label23 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Size"
         ForeColor       =   &H00656D76&
         Height          =   225
         Left            =   60
         TabIndex        =   53
         Top             =   150
         Width           =   435
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   16
         Left            =   210
         Top             =   2835
         Width           =   855
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   17
         Left            =   1110
         Top             =   2835
         Width           =   885
      End
   End
   Begin VB.Frame LogMgr 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   6210
      Left            =   45
      TabIndex        =   79
      Top             =   1920
      Width           =   7725
      Begin VB.Frame Frame12 
         BackColor       =   &H00A0ACBA&
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   285
         Left            =   5700
         TabIndex        =   80
         Top             =   2895
         Width           =   1815
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Pages"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   30
            TabIndex        =   81
            Top             =   45
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "Formats"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   885
            TabIndex        =   82
            Top             =   45
            Width           =   915
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Log List"
         ForeColor       =   &H00656D76&
         Height          =   2865
         Left            =   30
         TabIndex        =   85
         Top             =   255
         Width           =   7560
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Caption         =   "Use"
            Height          =   330
            Left            =   6840
            TabIndex        =   86
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
               TabIndex        =   87
               Top             =   15
               Width           =   450
            End
         End
         Begin VB.Frame Frame16 
            BorderStyle     =   0  'None
            Caption         =   "Use"
            Height          =   330
            Left            =   6840
            TabIndex        =   88
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
               TabIndex        =   89
               Top             =   0
               Width           =   480
            End
         End
         Begin VB.ListBox ListPage 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   2370
            Left            =   75
            TabIndex        =   90
            Top             =   210
            Width           =   7170
         End
         Begin VB.ListBox ListFormat 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   2370
            Left            =   75
            TabIndex        =   91
            Top             =   210
            Visible         =   0   'False
            Width           =   7170
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Details"
         ForeColor       =   &H00656D76&
         Height          =   3060
         Left            =   45
         TabIndex        =   83
         Top             =   3090
         Width           =   7530
         Begin VB.TextBox Text5 
            BackColor       =   &H00A0ACBA&
            BorderStyle     =   0  'None
            ForeColor       =   &H00656D76&
            Height          =   2820
            Left            =   75
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   84
            Top             =   195
            Width           =   7365
         End
      End
   End
   Begin VB.Frame InsPic 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Insert Picture"
      Height          =   6240
      Left            =   -30
      TabIndex        =   97
      Top             =   1995
      Width           =   7800
      Begin VB.Frame Frame20 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Select a image"
         ForeColor       =   &H00656D76&
         Height          =   2325
         Left            =   15
         TabIndex        =   99
         Top             =   0
         Width           =   7755
         Begin VB.DriveListBox Drive2 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   315
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   3945
         End
         Begin VB.DirListBox Dir2 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   1665
            Left            =   105
            TabIndex        =   101
            Top             =   570
            Width           =   3960
         End
         Begin VB.FileListBox File2 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            ForeColor       =   &H00656D76&
            Height          =   1980
            Left            =   4080
            Pattern         =   "*.JPG;*.PNG"
            TabIndex        =   100
            Top             =   255
            Width           =   3630
         End
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0ACBA&
         ForeColor       =   &H00656D76&
         Height          =   315
         ItemData        =   "Main.frx":013B
         Left            =   30
         List            =   "Main.frx":0148
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   2325
         Width           =   7740
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   270
         Left            =   5715
         TabIndex        =   104
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   6765
         TabIndex        =   103
         Top             =   2775
         Width           =   840
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   315
         Index           =   0
         Left            =   5700
         Top             =   2745
         Width           =   1005
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   1
         Left            =   6735
         Top             =   2745
         Width           =   900
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      ForeColor       =   &H00A0ACBA&
      Height          =   6255
      Left            =   45
      TabIndex        =   47
      Top             =   1875
      Width           =   7800
      Begin VB.ListBox MsgContentList 
         Height          =   450
         ItemData        =   "Main.frx":017D
         Left            =   2775
         List            =   "Main.frx":017F
         TabIndex        =   49
         Top             =   1830
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4BFCC&
         ForeColor       =   &H00656D76&
         Height          =   3150
         Left            =   525
         TabIndex        =   48
         Top             =   1020
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   2700
         TabIndex        =   51
         Top             =   1365
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   2700
         TabIndex        =   50
         Top             =   1020
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   2
         Left            =   2685
         Top             =   975
         Width           =   1215
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   3
         Left            =   2685
         Top             =   1335
         Width           =   1200
      End
   End
   Begin VB.Frame ABCD 
      BackColor       =   &H00A0ACBA&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   6300
      Left            =   15
      TabIndex        =   114
      Top             =   2010
      Visible         =   0   'False
      Width           =   7905
      Begin VB.Frame Frame18 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Texts &^"
         ForeColor       =   &H00656D76&
         Height          =   1545
         Left            =   0
         TabIndex        =   122
         Top             =   1185
         Width           =   7635
         Begin VB.CheckBox Check21 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "D"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   135
            TabIndex        =   141
            Top             =   1200
            Value           =   1  'Checked
            Width           =   390
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   -765
            TabIndex        =   129
            Top             =   9999
            Visible         =   0   'False
            Width           =   6960
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   -2040
            TabIndex        =   127
            Top             =   1545
            Visible         =   0   'False
            Width           =   6960
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   1110
            TabIndex        =   125
            Top             =   9999
            Visible         =   0   'False
            Width           =   6960
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   1035
            TabIndex        =   123
            Top             =   9999
            Visible         =   0   'False
            Width           =   6960
         End
         Begin VB.Label Label55 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00B4BFCC&
            Caption         =   " Click to edit "
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
            Left            =   3450
            TabIndex        =   136
            Top             =   870
            Width           =   1170
         End
         Begin VB.Label Label54 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00B4BFCC&
            Caption         =   " Click to edit "
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
            Left            =   3420
            TabIndex        =   135
            Top             =   1140
            Width           =   1170
         End
         Begin VB.Label Label53 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            Caption         =   " Click to edit "
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
            Left            =   630
            TabIndex        =   134
            Top             =   585
            Width           =   6855
         End
         Begin VB.Label Label52 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
            Caption         =   " Click to edit "
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
            Left            =   630
            TabIndex        =   133
            Top             =   300
            Width           =   6810
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   390
            TabIndex        =   130
            Top             =   1200
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   390
            TabIndex        =   128
            Top             =   885
            Width           =   105
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   405
            TabIndex        =   126
            Top             =   555
            Width           =   90
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   390
            TabIndex        =   124
            Top             =   240
            Width           =   105
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   210
            Index           =   4
            Left            =   600
            Top             =   330
            Width           =   6885
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   225
            Index           =   5
            Left            =   600
            Top             =   600
            Width           =   6855
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   255
            Index           =   6
            Left            =   600
            Top             =   870
            Width           =   6915
         End
         Begin VB.Shape Shape0 
            BackColor       =   &H00B4BFCC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00B4BFCC&
            Height          =   270
            Index           =   7
            Left            =   600
            Top             =   1155
            Width           =   6915
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00A0ACBA&
         Caption         =   "Format"
         ForeColor       =   &H00656D76&
         Height          =   1125
         Left            =   30
         TabIndex        =   115
         Top             =   15
         Width           =   5280
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   315
            ItemData        =   "Main.frx":0181
            Left            =   810
            List            =   "Main.frx":0183
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   195
            Width           =   1905
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "ABCD"
            ForeColor       =   &H00656D76&
            Height          =   390
            Left            =   2775
            TabIndex        =   138
            Top             =   570
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00A0ACBA&
            Caption         =   "ABCD"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   2790
            TabIndex        =   137
            Top             =   315
            Width           =   720
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00B4BFCC&
            ForeColor       =   &H00656D76&
            Height          =   285
            Left            =   810
            MaxLength       =   3
            TabIndex        =   118
            Top             =   540
            Width           =   1905
         End
         Begin VB.CheckBox Check20 
            BackColor       =   &H00B4BFCC&
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
            ForeColor       =   &H00656D76&
            Height          =   225
            Left            =   1110
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   855
            Width           =   270
         End
         Begin VB.CheckBox Check19 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4BFCC&
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
            ForeColor       =   &H00656D76&
            Height          =   225
            Left            =   825
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   855
            Width           =   270
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   465
            TabIndex        =   121
            Top             =   585
            Width           =   285
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   435
            TabIndex        =   120
            Top             =   285
            Width           =   330
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shape"
            ForeColor       =   &H00656D76&
            Height          =   195
            Left            =   345
            TabIndex        =   119
            Top             =   840
            Width           =   450
         End
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackColor       =   &H00A0ACBA&
         Caption         =   "Label58"
         Height          =   195
         Left            =   2520
         TabIndex        =   142
         Top             =   4590
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label57 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   7050
         TabIndex        =   140
         Top             =   2895
         Width           =   570
      End
      Begin VB.Label Label56 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4BFCC&
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
         ForeColor       =   &H00656D76&
         Height          =   255
         Left            =   6060
         TabIndex        =   139
         Top             =   2865
         Width           =   825
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   23
         Left            =   6015
         Top             =   2835
         Width           =   870
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H00B4BFCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B4BFCC&
         Height          =   330
         Index           =   24
         Left            =   6990
         Top             =   2850
         Width           =   675
      End
   End
   Begin VB.Label Label74 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   285
      Left            =   7335
      TabIndex        =   176
      Top             =   135
      Width           =   390
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
      ForeColor       =   &H00656D76&
      Height          =   285
      Left            =   180
      TabIndex        =   95
      Top             =   135
      Width           =   810
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage/Export"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   1590
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
      ForeColor       =   &H00656D76&
      Height          =   285
      Left            =   6225
      TabIndex        =   3
      Top             =   1170
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   285
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00656D76&
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   135
      Width           =   585
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B4BFCC&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00A0ACBA&
      Height          =   315
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
Dim str() As String, wholestr As String, outputs As String, goly As Long
Dim showcnt As Integer, deltachange As Integer, current As Integer, strs() As String, heightdata As Integer, special() As Integer, specialinfo() As Integer, stats() As Boolean
Sub NewMessage(Content As String, Color As Long, Optional ClearList As Boolean = False, Optional ClearOnly = False)
'    current = -1
    If (ClearOnly And Not ClearList) Then
        RaiseSysErr "Clear message list only and do not clear message list were both turned on.", "Create/PageSettings/NewEvent"
        Exit Sub
    End If
    If ClearList Then
        MsgContentList.Clear
        MsgColorList.Clear
        MsgTypeList.Clear
        If Message.Caption <> "" Then Message.Caption = Message.Caption & translate("(Expired)")
        If ClearOnly Then Exit Sub
    End If
    MsgContentList.AddItem Content
    MsgColorList.AddItem Color
    Select Case Color
        Case vbBlack: MsgTypeList.AddItem translate("[Info]")
        Case vbBlue: MsgTypeList.AddItem translate("[Warning]")
        Case vbRed: MsgTypeList.AddItem translate("[Error]")
    End Select
 '   showcnt = 49
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

Private Sub Check21_Click()
    If Check21.Value = 1 Then Label55.Enabled = True Else Label55.Enabled = False
    If Check21.Value = 1 Then Label55.Caption = translate("Click to edit") Else Label55.Caption = translate("Disabled")
    If Check21.Value = 1 Then Option3.Caption = "ABCD" Else Option3.Caption = "ABC"
    If Check21.Value = 1 Then Option4.Caption = "ABCD" Else Option4.Caption = "ABC"
    If Check21.Value = 1 Then Option4.Value = True Else Option3.Value = True
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

Private Sub Combo3_Change()
    Label65.Caption = Text15.Text & Combo3.Text & Combo5.Text
End Sub

Private Sub Combo3_Click()
    Label65.Caption = Text15.Text & Combo3.Text & Combo5.Text
End Sub

Private Sub Combo5_Change()
    Label65.Caption = Text15.Text & Combo3.Text & Combo5.Text
End Sub

Private Sub Combo5_Click()
    Label65.Caption = Text15.Text & Combo3.Text & Combo5.Text
End Sub

Private Sub Dir1_Change()
    On Error GoTo err
    File1.Path = Dir1.Path
    Exit Sub
err:
    NewMessage translate("[SysErr]") & err.Description, vbRed
End Sub

Private Sub Dir2_Change()
    On Error GoTo err
    File2.Path = Dir2.Path
    Exit Sub
err:
    NewMessage translate("[SysErr]") & err.Description, vbRed
End Sub

Private Sub Drive1_Change()
    On Error GoTo err
    Dir1.Path = Drive1.Drive
    Exit Sub
err:
    NewMessage translate("[SysErr]") & err.Description, vbRed
End Sub

Private Sub Drive2_Change()
    On Error GoTo err
    Dir2.Path = Drive2.Drive
    Exit Sub
err:
    NewMessage translate("[SysErr]") & err.Description, vbRed
End Sub

Private Sub Form_Load()
    current = -1
    Dim i As Long
    FontCombo.Clear
    Integrated.Show
    Integrated.WinMode = 1
    Integrated.InitWindow
    For i = 1 To Screen.FontCount
        Integrated.Message.Caption = translate("Loading Fonts(") & i & "/" & Screen.FontCount & ")"
        DoEvents
        FontCombo.AddItem Screen.Fonts(i)
        Combo4.AddItem Screen.Fonts(i)
    Next
    Label1_Click
    Integrated.Message.Caption = translate("Translating...")
    translatecontrol Me.Name
    If EnableTranslation = 1 Then
        
MainFrm.Label27.AutoSize = True
MainFrm.Label27.Left = 60
DoEvents
MainFrm.Text4.Left = MainFrm.Label27.Left + MainFrm.Label27.Width + 10
MainFrm.Text4.Top = MainFrm.Text4.Top + 60
Check17.Top = Check17.Top - 30
MainFrm.Label27.Top = MainFrm.Label27.Top + 60
Label5.Alignment = 1
Label37.Height = Label36.Height
End If
    Unload Integrated
    
'    Shape1.Left = Label1.Left
'    Shape1.Width = Label1.Width
    PreviewButton_Click
End Sub

Private Sub Label17_Click()
    Frame5.Visible = False
End Sub

Private Sub Label19_Click()
    Frame5.Visible = True
End Sub

Private Sub Label2_Click()
    Tools.Visible = True
    Shape1.Left = Label2.Left - 120
    Shape1.Width = Label2.Width + 240
    Frame1.Visible = False
    Frame2.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    List1.Visible = False
    Label14.Visible = False
    Label13.Visible = False
    Frame8.Visible = False
    Manage.Visible = False
    InsText.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = False
    General.Visible = False
    Merge.Visible = False
    Copyright.Visible = False
    SaveLoad.Visible = False
    ABCD.Visible = False
End Sub

Private Sub Label22_Click()
    AnswerLine.Visible = Not False
    InsText.Visible = Not True
    InsPic.Visible = False
    Blk.Visible = False
    ABCD.Visible = False
End Sub

Private Sub Label24_Click()
    If Not IsNumeric(Text3.Text) Or Not IsNumeric(Text4.Text) Then
        NewMessage translate("Invaild Format."), vbRed, True
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
        NewMessage translate("The size you've input is too large"), vbRed
        On Error Resume Next
        Preview.SystemCall = SystemCallFlag
        Unload Preview
    End If
End Sub

Private Sub Label25_Click()
    If Not IsNumeric(Text3.Text) Or Not IsNumeric(Text4.Text) Then
        NewMessage translate("Invaild Format."), vbRed, True
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
        Preview.SystemCall = SystemCallFlag
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
    Preview.SystemCall = SystemCallFlag
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

Private Sub Label3_Click()
    Tools.Visible = False
    Shape1.Left = Label3.Left - 100
    Shape1.Width = Label3.Width + 200
    Frame1.Visible = False
    Frame2.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    List1.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Frame8.Visible = False
    Manage.Visible = False
    InsText.Visible = False
    AnswerLine.Visible = False
    General.Visible = False
    InsPic.Visible = False
    Blk.Visible = False
    Merge.Visible = False
    SaveLoad.Visible = False
    Copyright.Visible = True
    ABCD.Visible = False
End Sub

Private Sub Label30_Click()
    InsText.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = True
    Frame8.Visible = False
    Merge.Visible = False
    SaveLoad.Visible = False
End Sub

Private Sub Label35_Click()
    AnswerLine.Visible = Not True
    InsText.Visible = Not True
    InsPic.Visible = Not False
    Blk.Visible = False
    ABCD.Visible = False
End Sub

Private Sub Label36_Click()
    If Combo2.Text = "" Then
        NewMessage translate("Invaild Format."), vbRed
        Exit Sub
    End If
    InitPreview
    On Error GoTo err
    Preview.Exports.Width = 1
    Preview.Exports.Height = 1
    Preview.Exports.Picture = LoadPicture(File2.Path & "/" & File2.FileName)
    If Left(Combo2.Text, 1) = "0" Then
        Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
    ElseIf Left(Combo2.Text, 1) = "1" Then
        Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
    Else
        Preview.Picture2.PaintPicture Preview.Exports.Picture, ((RightMargin - LeftMargin) / 2) + LeftMargin - Preview.Exports.Width / 2, TopMargin
    End If
    Exit Sub
err:
    NewMessage translate("[SysErr]") & err.Description, vbRed
End Sub

Private Sub Label37_Click()
    If Combo2.Text = "" Then
        NewMessage translate("Invaild Format."), vbRed
        Exit Sub
    End If
    InitPreview
    On Error GoTo err
    Preview.Exports.Width = 1
    Preview.Exports.Height = 1
    Preview.Exports.Picture = LoadPicture(File2.Path & "/" & File2.FileName)
    If Left(Combo2.Text, 1) = "0" Then
        Preview.Picture2.PaintPicture Preview.Exports.Picture, LeftMargin, TopMargin
    ElseIf Left(Combo2.Text, 1) = "1" Then
        Preview.Picture2.PaintPicture Preview.Exports.Picture, RightMargin - Preview.Exports.Width, TopMargin
    Else
        Preview.Picture2.PaintPicture Preview.Exports.Picture, ((RightMargin - LeftMargin) / 2) + LeftMargin - Preview.Exports.Width / 2, TopMargin
    End If
    Exit Sub
    Preview.Export.Height = Preview.Exports.Height
    Preview.Export.Width = RightMargin - LeftMargin
    Preview.Export.PaintPicture Preview.Picture2.Image, 0, 0, , , TopMargin, LeftMargin, RightMargin - LeftMargin, Preview.Exports.Height
    Dim usage As Long
    usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
    If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
    SavePicture Preview.Export.Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
    SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
    List1.AddItem usage + 1
    NewMessage translate("Success with tracknumber") & usage + 1, vbBlack
    On Error Resume Next
    Preview.SystemCall = SystemCallFlag
    Unload Preview
    Exit Sub
err:
    NewMessage translate("[SysErr]") & err.Description, vbRed
End Sub

Private Sub Label38_Click()
    InsText.Visible = True
    Frame8.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = False
    InsPic.Visible = False
    Blk.Visible = True
    InsText.Visible = False
    ABCD.Visible = False
End Sub

Private Sub Label42_Click()
    If Not IsNumeric(Text6.Text) Or Not IsNumeric(Text7.Text) Or Val(Text6.Text) <> Int(Val(Text6.Text)) Or Val(Text7.Text) <> Int(Val(Text7.Text)) Then
        NewMessage translate("Invaild Format"), vbRed
        Exit Sub
    End If
    Dim i As Long, v1 As Long
    v1 = Val(Text6.Text)
    If Check18.Value = 1 Then v1 = v1 * TwipsPerCM
    v1 = v1 * Val(Text7.Text)
    List1.AddItem "BT" & v1
    NewMessage translate("Completed on ") & Now, vbBlack
'    InitPreview
'    Preview.Picture2.Height = v1
'    Preview.Picture2.Width = RightMargin - LeftMargin
'    Dim usage As Long
'    For i = 1 To Val(Text7.Text)
'        If i = 1 Then
'            usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
'            If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
'            SavePicture Preview.Picture2.Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
'            SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
'        End If
'        List1.AddItem usage + 1
'    Next
'    On Error Resume Next
'    Unload Preview
End Sub

Private Sub Label43_Click()
    InsText.Visible = True
    Frame8.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = False
    InsPic.Visible = False
    Blk.Visible = False
    InsText.Visible = False
    ABCD.Visible = True
End Sub

Private Sub Label51_Click()
    Text2.Text = InputText(Text2.Text)
End Sub

Private Sub Label52_Click()
    Text8.Text = InputText(Text8.Text, False)
End Sub

Private Sub Label53_Click()
    Text10.Text = InputText(Text10.Text, False)
End Sub

Private Sub Label54_Click()
    Text11.Text = InputText(Text11.Text, False)
End Sub

Private Sub Label55_Click()
    Text12.Text = InputText(Text12.Text, False)
End Sub

Private Sub Label56_Click()
    ReDim stats(233) 'Safe StatReg
    
    Dim i As Long, l As Long, c As String, r As String, v0 As Long, delta As Long, maxheight As Long, totwidth As Long, srr As String
    If Combo4.Text = "" Or Not IsNumeric(Text9.Text) Then
        NewMessage translate("Invaild Format."), vbRed
        Exit Sub
    End If
    If Option3.Value Then v0 = 4 Else v0 = 2
    InitPreview
    With Me.Temp
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
    End With
    With Preview.Picture2
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
        .CurrentY = TopMargin
        .CurrentX = LeftMargin
        srr = Text8.Text
        Text8.Text = "A. " & srr
        deltachange = 0
        l = Len(Text8.Text)
        For i = 1 To l
            DoEvents
            'Debug.Print deltachange
            .CurrentY = TopMargin + deltachange
            c = Mid(Text8.Text, i, 1)
            If c = "^" Then
                r = ""
                i = i + 1
                While Mid(Text8.Text, i, 1) <> "^" And i <= l
                    r = r & Mid(Text8.Text, i, 1)
                    i = i + 1
                Wend
                RegisterStat r
                
                GoTo nexti
            End If
            Temp.Caption = c
            Debug.Print "*" & .CurrentY
            Preview.Picture2.Print c;
            maxheight = Max(maxheight, Temp.Height)
            If (v0 = 2 And Preview.Picture2.CurrentX - LeftMargin > (RightMargin - LeftMargin) / v0) Or (v0 = 4 And Preview.Picture2.CurrentX - LeftMargin > (RightMargin - LeftMargin) / 2) Then 'NOTE THAT!!
                NewMessage translate("The input A is too large that we can't process that."), vbRed
                On Error Resume Next
                Preview.SystemCall = SystemCallFlag
                Unload Preview
                Text8.Text = srr
                Exit Sub
            End If
nexti:
        Next
        Text8.Text = srr
    End With
    With Preview.Picture2
        'Preview.Picture2.Line ((RightMargin - LeftMargin) / 2, 0)-((RightMargin - LeftMargin) / 2, Preview.Picture2.Height)
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
        .CurrentY = TopMargin
        .CurrentX = (RightMargin - LeftMargin) / v0 + LeftMargin
        srr = Text8.Text
        Text8.Text = "B. " & Text10.Text
        deltachange = 0
        Debug.Print Text8.Text
        Debug.Print Len(Text8.Text)
        l = Len(Text8.Text)
        For i = 1 To l
            DoEvents
            'Debug.Print deltachange
            .CurrentY = TopMargin + deltachange
            c = Mid(Text8.Text, i, 1)
            Debug.Print c
            If c = "^" Then
                r = ""
                i = i + 1
                While Mid(Text8.Text, i, 1) <> "^" And i <= l
                    r = r & Mid(Text8.Text, i, 1)
                    i = i + 1
                Wend
                RegisterStat r
                
                GoTo nextir
            End If
            Temp.Caption = c
            maxheight = Max(maxheight, Temp.Height)
            Debug.Print "*#" & .CurrentX
            Preview.Picture2.Print c;
            Debug.Print c
            If (v0 = 2 And Preview.Picture2.CurrentX > RightMargin) Or (v0 = 4 And Preview.Picture2.CurrentX > (RightMargin - LeftMargin) / 2 + LeftMargin) Then 'NOTE THAT!!
                NewMessage translate("The input B is too large that we can't process that."), vbRed
                Debug.Print v0
                Debug.Print .CurrentX
                On Error Resume Next
                Preview.SystemCall = SystemCallFlag
                Unload Preview
                Text8.Text = srr
                Debug.Print PageWidth, RightMargin - LeftMargin
                Exit Sub
            End If
nextir:
        Next
        Text8.Text = srr
    End With
    With Preview.Picture2
        'Preview.Picture2.Line ((RightMargin - LeftMargin) / 2, 0)-((RightMargin - LeftMargin) / 2, Preview.Picture2.Height)
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
        If Option4.Value Then .CurrentY = TopMargin + maxheight Else .CurrentY = TopMargin
        If Option4.Value Then .CurrentX = LeftMargin Else .CurrentX = (RightMargin - LeftMargin) / v0 * 2 + LeftMargin
        srr = Text8.Text
        Text8.Text = "C. " & Text11.Text
        deltachange = 0
        Debug.Print Text8.Text
        Debug.Print Len(Text8.Text)
        l = Len(Text8.Text)
        For i = 1 To l
            DoEvents
            'Debug.Print deltachange
            .CurrentY = TopMargin + deltachange
            If Option4.Value Then .CurrentY = .CurrentY + maxheight
            c = Mid(Text8.Text, i, 1)
            Debug.Print c
            If c = "^" Then
                r = ""
                i = i + 1
                While Mid(Text8.Text, i, 1) <> "^" And i <= l
                    r = r & Mid(Text8.Text, i, 1)
                    i = i + 1
                Wend
                RegisterStat r
                
                GoTo nextirr
            End If
            Temp.Caption = c
            Debug.Print "*#" & .CurrentX
            Preview.Picture2.Print c;
            Debug.Print c
            If (v0 = 2 And Preview.Picture2.CurrentX > (RightMargin - LeftMargin) / 2) Or (v0 = 4 And Preview.Picture2.CurrentX > (RightMargin - LeftMargin) / 4 * 3 + LeftMargin) Then 'NOTE THAT!!
                NewMessage translate("The input C is too large that we can't process that."), vbRed
                Debug.Print v0
                Debug.Print .CurrentX
                On Error Resume Next
                Preview.SystemCall = SystemCallFlag
                Unload Preview
                Text8.Text = srr
                Debug.Print PageWidth, RightMargin - LeftMargin
                Exit Sub
            End If
nextirr:
        Next
        Text8.Text = srr
    End With
    If Check21.Value = 1 Then
        With Preview.Picture2
            'Preview.Picture2.Line ((RightMargin - LeftMargin) / 2, 0)-((RightMargin - LeftMargin) / 2, Preview.Picture2.Height)
            .FontSize = Text9.Text
            .FontName = Combo4.Text
            If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
            If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
            If Option4.Value Then .CurrentY = TopMargin + maxheight Else .CurrentY = TopMargin
            If Option4.Value Then .CurrentX = (RightMargin - LeftMargin) / 2 + LeftMargin Else .CurrentX = (RightMargin - LeftMargin) / v0 * 3 + LeftMargin
            srr = Text8.Text
            Text8.Text = "D. " & Text12.Text
            deltachange = 0
            Debug.Print Text8.Text
            Debug.Print Len(Text8.Text)
            l = Len(Text8.Text)
            For i = 1 To l
                DoEvents
                'Debug.Print deltachange
                .CurrentY = TopMargin + deltachange
                If Option4.Value Then .CurrentY = .CurrentY + maxheight
                c = Mid(Text8.Text, i, 1)
                Debug.Print c
                If c = "^" Then
                    r = ""
                    i = i + 1
                    While Mid(Text8.Text, i, 1) <> "^" And i <= l
                        r = r & Mid(Text8.Text, i, 1)
                        i = i + 1
                    Wend
                    RegisterStat r
                    
                    GoTo nexti233
                End If
                Temp.Caption = c
                'maxheight = Max(maxheight, Temp.Height)
                Debug.Print "*#" & .CurrentX
                Preview.Picture2.Print c;
                Debug.Print c
                If (v0 = 2 And Preview.Picture2.CurrentX > RightMargin) Or (v0 = 4 And Preview.Picture2.CurrentX > RightMargin) Then 'NOTE THAT!!
                    NewMessage translate("The input D is too large that we can't process that."), vbRed
                    Debug.Print v0
                    Debug.Print .CurrentX
                    On Error Resume Next
                    Preview.SystemCall = SystemCallFlag
                    Unload Preview
                    Text8.Text = srr
                    Debug.Print PageWidth, RightMargin - LeftMargin
                    Exit Sub
                End If
nexti233:
            Next
            Text8.Text = srr
        End With
    End If
End Sub

Private Sub Label57_Click()
        ReDim stats(233) 'Safe StatReg
    
    Dim i As Long, c As String, r As String, l As Long, v0 As Long, delta As Long, maxheight As Long, totwidth As Long, srr As String, maxheightr As Long
    If Combo4.Text = "" Or Not IsNumeric(Text9.Text) Then
        NewMessage translate("Invaild Format."), vbRed
        Exit Sub
    End If
    If Option3.Value Then v0 = 4 Else v0 = 2
    InitPreview
    With Me.Temp
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
    End With
    With Preview.Picture2
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
        .CurrentY = TopMargin
        .CurrentX = LeftMargin
        srr = Text8.Text
        Text8.Text = "A. " & srr
        deltachange = 0
        l = Len(Text8.Text)
        For i = 1 To l
            DoEvents
            'Debug.Print deltachange
            .CurrentY = TopMargin + deltachange
            c = Mid(Text8.Text, i, 1)
            If c = "^" Then
                r = ""
                i = i + 1
                While Mid(Text8.Text, i, 1) <> "^" And i <= l
                    r = r & Mid(Text8.Text, i, 1)
                    i = i + 1
                Wend
                RegisterStat r
                
                GoTo nexti
            End If
            Temp.Caption = c
            Debug.Print "*" & .CurrentY
            Preview.Picture2.Print c;
            maxheight = Max(maxheight, Temp.Height)
            If (v0 = 2 And Preview.Picture2.CurrentX - LeftMargin > (RightMargin - LeftMargin) / v0) Or (v0 = 4 And Preview.Picture2.CurrentX - LeftMargin > (RightMargin - LeftMargin) / 2) Then 'NOTE THAT!!
                NewMessage translate("The input A is too large that we can't process that."), vbRed
                On Error Resume Next
                Preview.SystemCall = SystemCallFlag
                Unload Preview
                Text8.Text = srr
                Exit Sub
            End If
nexti:
        Next
        Text8.Text = srr
    End With
    With Preview.Picture2
        'Preview.Picture2.Line ((RightMargin - LeftMargin) / 2, 0)-((RightMargin - LeftMargin) / 2, Preview.Picture2.Height)
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
        .CurrentY = TopMargin
        .CurrentX = (RightMargin - LeftMargin) / v0 + LeftMargin
        srr = Text8.Text
        Text8.Text = "B. " & Text10.Text
        deltachange = 0
        Debug.Print Text8.Text
        Debug.Print Len(Text8.Text)
        l = Len(Text8.Text)
        For i = 1 To l
            DoEvents
            'Debug.Print deltachange
            .CurrentY = TopMargin + deltachange
            c = Mid(Text8.Text, i, 1)
            Debug.Print c
            If c = "^" Then
                r = ""
                i = i + 1
                While Mid(Text8.Text, i, 1) <> "^" And i <= l
                    r = r & Mid(Text8.Text, i, 1)
                    i = i + 1
                Wend
                RegisterStat r
                
                GoTo nextir
            End If
            Temp.Caption = c
            maxheight = Max(maxheight, Temp.Height)
            Debug.Print "*#" & .CurrentX
            Preview.Picture2.Print c;
            Debug.Print c
            If (v0 = 2 And Preview.Picture2.CurrentX > RightMargin) Or (v0 = 4 And Preview.Picture2.CurrentX > (RightMargin - LeftMargin) / 2 + LeftMargin) Then 'NOTE THAT!!
                NewMessage translate("The input B is too large that we can't process that."), vbRed
                Debug.Print v0
                Debug.Print .CurrentX
                On Error Resume Next
                Preview.SystemCall = SystemCallFlag
                Unload Preview
                Text8.Text = srr
                Debug.Print PageWidth, RightMargin - LeftMargin
                Exit Sub
            End If
nextir:
        Next
        Text8.Text = srr
    End With
    With Preview.Picture2
        'Preview.Picture2.Line ((RightMargin - LeftMargin) / 2, 0)-((RightMargin - LeftMargin) / 2, Preview.Picture2.Height)
        .FontSize = Text9.Text
        .FontName = Combo4.Text
        If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
        If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
        If Option4.Value Then .CurrentY = TopMargin + maxheight Else .CurrentY = TopMargin
        If Option4.Value Then .CurrentX = LeftMargin Else .CurrentX = (RightMargin - LeftMargin) / v0 * 2 + LeftMargin
        srr = Text8.Text
        Text8.Text = "C. " & Text11.Text
        deltachange = 0
        Debug.Print Text8.Text
        Debug.Print Len(Text8.Text)
        l = Len(Text8.Text)
        For i = 1 To l
            DoEvents
            'Debug.Print deltachange
            .CurrentY = TopMargin + deltachange
            If Option4.Value Then .CurrentY = .CurrentY + maxheight
            c = Mid(Text8.Text, i, 1)
            Debug.Print c
            If c = "^" Then
                r = ""
                i = i + 1
                While Mid(Text8.Text, i, 1) <> "^" And i <= l
                    r = r & Mid(Text8.Text, i, 1)
                    i = i + 1
                Wend
                RegisterStat r
                
                GoTo nextirr
            End If
            Temp.Caption = c
            maxheightr = Max(maxheightr, Temp.Height)
            Debug.Print "*#" & .CurrentX
            Preview.Picture2.Print c;
            Debug.Print c
            If (v0 = 2 And Preview.Picture2.CurrentX > (RightMargin - LeftMargin) / 2) Or (v0 = 4 And Preview.Picture2.CurrentX > (RightMargin - LeftMargin) / 4 * 3 + LeftMargin) Then 'NOTE THAT!!
                NewMessage translate("The input C is too large that we can't process that."), vbRed
                Debug.Print v0
                Debug.Print .CurrentX
                On Error Resume Next
                Preview.SystemCall = SystemCallFlag
                Unload Preview
                Text8.Text = srr
                Debug.Print PageWidth, RightMargin - LeftMargin
                Exit Sub
            End If
nextirr:
        Next
        Text8.Text = srr
    End With
    If Check21.Value = 1 Then
        With Preview.Picture2
            'Preview.Picture2.Line ((RightMargin - LeftMargin) / 2, 0)-((RightMargin - LeftMargin) / 2, Preview.Picture2.Height)
            .FontSize = Text9.Text
            .FontName = Combo4.Text
            If Check19.Value = 1 Then .FontBold = True Else .FontBold = False
            If Check20.Value = 1 Then .FontItalic = True Else .FontItalic = False
            If Option4.Value Then .CurrentY = TopMargin + maxheight Else .CurrentY = TopMargin
            If Option4.Value Then .CurrentX = (RightMargin - LeftMargin) / 2 + LeftMargin Else .CurrentX = (RightMargin - LeftMargin) / v0 * 3 + LeftMargin
            srr = Text8.Text
            Text8.Text = "D. " & Text12.Text
            deltachange = 0
            Debug.Print Text8.Text
            Debug.Print Len(Text8.Text)
            l = Len(Text8.Text)
            For i = 1 To l
                DoEvents
                'Debug.Print deltachange
                .CurrentY = TopMargin + deltachange
                If Option4.Value Then .CurrentY = .CurrentY + maxheight
                c = Mid(Text8.Text, i, 1)
                Debug.Print c
                If c = "^" Then
                    r = ""
                    i = i + 1
                    While Mid(Text8.Text, i, 1) <> "^" And i <= l
                        r = r & Mid(Text8.Text, i, 1)
                        i = i + 1
                    Wend
                    RegisterStat r
                    
                    GoTo nexti233
                End If
                Temp.Caption = c
                maxheightr = Max(maxheightr, Temp.Height)
                Debug.Print "*#" & .CurrentX
                Preview.Picture2.Print c;
                Debug.Print c
                If (v0 = 2 And Preview.Picture2.CurrentX > RightMargin) Or (v0 = 4 And Preview.Picture2.CurrentX > RightMargin) Then 'NOTE THAT!!
                    NewMessage translate("The input D is too large that we can't process that."), vbRed
                    Debug.Print v0
                    Debug.Print .CurrentX
                    On Error Resume Next
                    Preview.SystemCall = SystemCallFlag
                    Unload Preview
                    Text8.Text = srr
                    Debug.Print PageWidth, RightMargin - LeftMargin
                    Exit Sub
                End If
nexti233:
            Next
            Text8.Text = srr
        End With
    End If
    Dim usage As Long
    Preview.Exports.Height = maxheight
    Preview.Exports.Width = RightMargin - LeftMargin
    Preview.Exports.PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, TopMargin, RightMargin - LeftMargin, maxheight
    usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
    If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
    SavePicture Preview.Exports.Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
    SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
    List1.AddItem usage + 1
    If Option4.Value Then
        usage = usage + 1
        Preview.Exports.PaintPicture Preview.Picture2.Image, 0, 0, , , LeftMargin, TopMargin + maxheight, RightMargin - LeftMargin, maxheightr
        SavePicture Preview.Exports.Image, App.Path & "\Cache\" & usage + 1 & ".jpg"
        SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
        List1.AddItem usage + 1
    End If
    On Error Resume Next
    Preview.SystemCall = SystemCallFlag
    Unload Preview
End Sub

Private Sub Label59_Click()
    Label15_Click
    Dim Y As Long
    Preview.Exporter.Visible = True
    Y = TopMargin
    Preview.Exporter.Cls
    Preview.Exporter.Height = PageHeight
    Preview.Exporter.Width = PageWidth
    Dim usage As Long, cnt As Long, cntrec As Long, rrr As String, vvv As Long
    vvv = goly
    Debug.Print vvv, Y
    While Y <= vvv
        Debug.Print "-->" & Y
        cnt = cnt + 1
        Preview.Exporter.Cls
        Preview.Exporter.PaintPicture Preview.Picture2.Image, TopMargin, LeftMargin, , , LeftMargin, Y, RightMargin - LeftMargin, Min(BotMargin - TopMargin, Preview.Picture2.Height - Y)
        If Check23.Value = 1 Then Line (LeftMargin, TopMargin)-(RightMargin, TopMargin)
        If Check22.Value = 1 Then Line (LeftMargin, BotMargin)-(RightMargin, BotMargin)
        Preview.Exporter.CurrentY = BotMargin + 20
        On Error Resume Next
        Preview.Exporter.FontName = "宋体"
        Preview.Exporter.FontSize = 10
        Preview.Exporter.CurrentX = PageWidth / 2 - Label65.Width / 2
        Preview.Exporter.Print Label65.Caption
'        SavePicture Preview.Exporter.Image, App.Path & "\Result\" & usage + 1 & "\" & cnt & ".jpg"
        Y = Y + BotMargin - TopMargin
    Wend
    cntrec = cnt
    cnt = 0
    Y = TopMargin
    usage = GetSetting("FreeExam", "Create", "GenerateCnt", 1000)
    SaveSetting "FreeExam", "Create", "GenerateCnt", usage + 1
    If Dir(App.Path & "\Result", vbDirectory) = "" Then MkDir App.Path & "\Result"
    If Dir(App.Path & "\Result\" & usage + 1, vbDirectory) = "" Then MkDir App.Path & "\Result\" & usage + 1
    While Y <= vvv
        Debug.Print "-->" & Y
        cnt = cnt + 1
        Preview.Exporter.Cls
        Preview.Exporter.PaintPicture Preview.Picture2.Image, TopMargin, LeftMargin, , , LeftMargin, Y, RightMargin - LeftMargin, Min(BotMargin - TopMargin, Preview.Picture2.Height - Y)
        If Check22.Value = 1 Then Preview.Exporter.Line (LeftMargin, BotMargin)-(RightMargin, BotMargin)
        If Check23.Value = 1 Then Preview.Exporter.Line (LeftMargin, TopMargin)-(RightMargin, TopMargin)
        Preview.Exporter.CurrentY = BotMargin + 20
        On Error Resume Next
        Preview.Exporter.FontName = "宋体"
        Preview.Exporter.FontSize = 10
        Preview.Exporter.CurrentX = PageWidth / 2 - Label65.Width / 2
        rrr = Label65.Caption
        rrr = Replace(rrr, "PAGENUMBER", cnt)
        rrr = Replace(rrr, "TOTALPAGE", cntrec)
        Preview.Exporter.Print rrr
        
        Temp.FontName = "宋体"
        Temp.FontSize = 10
        Temp.Caption = Text13.Text
        Preview.Exporter.CurrentY = TopMargin - Temp.Height
        Preview.Exporter.CurrentX = PageWidth / 2 - Temp.Width / 2
        Preview.Exporter.Print Text13.Text
        
        SavePicture Preview.Exporter.Image, App.Path & "\Result\" & usage + 1 & "\" & cnt & ".jpg"
        Y = Y + BotMargin - TopMargin
    Wend
    NewMessage translate("Generated and saved in the following path: "), vbBlack
    NewMessage App.Path & "\Result\" & usage + 1, vbBlack
    DoEvents
    On Error Resume Next
    Preview.SystemCall = SystemCallFlag
    Unload Preview
End Sub

Private Sub Label60_Click()
    InsText.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = False
    Frame8.Visible = False
    Merge.Visible = True
    SaveLoad.Visible = False

End Sub

Private Sub Label63_Click()
    InsText.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = False
    Frame8.Visible = True
    Merge.Visible = False
    SaveLoad.Visible = False
End Sub

Private Sub Label64_Click()
    InsText.Visible = False
    AnswerLine.Visible = False
    LogMgr.Visible = False
    Frame8.Visible = False
    Merge.Visible = False
    SaveLoad.Visible = True
End Sub

Private Sub Label66_Click()
    InitPreview
    Dim usage As Long
    usage = GetSetting("FreeExam", "Create", "LoadIDUsage", 1000)
    If Dir(App.Path & "\Saves", vbDirectory) = "" Then MkDir App.Path & "\Saves"
    If Dir(App.Path & "\Saves\" & usage + 1, vbDirectory) = "" Then MkDir App.Path & "\Saves\" & usage + 1
    SaveSetting "FreeExam", "Create", "LoadIDUsage", usage + 1
    Dim i As Long
    Open App.Path & "\Saves\" & usage + 1 & "\Config.txt" For Output As #1
        For i = 0 To List1.ListCount - 1
            List1.ListIndex = i
            If Left(List1.Text, 1) = "B" Or Left(List1.Text, 1) = "P" Then
                Write #1, List1.Text
                GoTo nexti
            End If
            On Error Resume Next
            Preview.Export.Picture = LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg")
            SavePicture Preview.Export.Picture, App.Path & "\Saves\" & usage + 1 & "\" & i & ".jpg"
            Write #1, i
nexti:
        Next
    Close #1
    NewMessage translate("Completed. LoadID=") & usage + 1, vbBlack
    On Error Resume Next
    Preview.SystemCall = SystemCallFlag
    Unload Preview
End Sub

Private Sub Label69_Click()
    If Dir(App.Path & "\Saves\" & Text14.Text, vbDirectory) = "" Or Dir(App.Path & "\Saves\" & Text14.Text & "\Config.txt") = "" Then
        NewMessage translate("This LoadID and/or its configuration not found"), vbBlack
        Exit Sub
    End If
    Dim v As Long, usage As Long, rrr
    InitPreview
    Open App.Path & "\Saves\" & Text14.Text & "\Config.txt" For Input As #1
        While Not EOF(1)
            Input #1, rrr
            If Left(rrr, 1) = "B" Or Left(rrr, 1) = "P" Then
                List1.AddItem rrr
                GoTo nextiii
            End If
            v = Val(rrr)
            On Error Resume Next
            Preview.Exports.Picture = LoadPicture(App.Path & "\Saves\" & Text14.Text & "\" & v & ".jpg")
            usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
            If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
            SavePicture Preview.Exports.Picture, App.Path & "\Cache\" & usage + 1 & ".jpg"
            SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
            List1.AddItem usage + 1
nextiii:
        Wend
    Close #1
    NewMessage translate("Completed."), vbBlack
    On Error Resume Next
    Preview.SystemCall = SystemCallFlag
    Unload Preview
End Sub

Private Sub Label73_Click()
    Dim ret As Long
    ret = Shell("mspaint.exe", vbNormalFocus)
    If ret = 0 Then
        NewMessage translate("Unable to start the program"), vbRed
    Else
        NewMessage translate("Success. PID=") & ret, vbBlack
    End If
End Sub

Private Sub Label74_Click()
    ExitProgram
End Sub

Private Sub Label75_Click()
    Dim ret As Long
    ret = Shell("notepad.exe", vbNormalFocus)
    If ret = 0 Then
        NewMessage translate("Unable to start the program"), vbRed
    Else
        NewMessage translate("Success. PID=") & ret, vbBlack
    End If
End Sub

Private Sub Label76_Click()
    Dim ret As Long
    ret = Shell("calc.exe", vbNormalFocus)
    If ret = 0 Then
        NewMessage translate("Unable to start the program"), vbRed
    Else
        NewMessage translate("Success. PID=") & ret, vbBlack
    End If
End Sub

Private Sub List2_Click()
    Text2.Text = strs(List2.ListIndex)
End Sub

Private Sub ListFormat_Click()
    Dim qqq() As String, i As Long
    Text5.Text = ""
    qqq = Split(ListFormat.Text, ",")
    Text5.Text = Text5.Text & translate("FontName: ") & qqq(0) & vbCrLf
    Text5.Text = Text5.Text & translate("FontSize: ") & qqq(1) & vbCrLf
    Text5.Text = Text5.Text & translate("Bold: ") & qqq(2) & vbCrLf
    Text5.Text = Text5.Text & translate("italic: ") & qqq(3) & vbCrLf
    Text5.Text = Text5.Text & translate("Alignment: ") & qqq(4) & vbCrLf
    Text5.Text = Text5.Text & translate("* For italic/Bold: 1 is true, 0 is false")
End Sub

Private Sub ListPage_Click()
    Dim qqq() As String, i As Long
    Text5.Text = ""
    qqq = Split(ListPage.Text, ",")
    Text5.Text = Text5.Text & translate("FontName: ") & qqq(0) & vbCrLf
    Text5.Text = Text5.Text & translate("FontSize: ") & qqq(1) & vbCrLf
    Text5.Text = Text5.Text & translate("Bold: ") & qqq(2) & vbCrLf
    Text5.Text = Text5.Text & translate("italic: ") & qqq(3) & vbCrLf
    Text5.Text = Text5.Text & translate("Alignment: ") & qqq(4) & vbCrLf
    Text5.Text = Text5.Text & translate("Text: ") & qqq(5) & vbCrLf
    Text5.Text = Text5.Text & translate("DisabledImage: ") & qqq(6) & vbCrLf
    If qqq(6) = "False" Then
        Text5.Text = Text5.Text & translate("Drive: ") & qqq(7) & vbCrLf
        Text5.Text = Text5.Text & translate("Path: ") & qqq(8) & vbCrLf
        Text5.Text = Text5.Text & translate("File: ") & qqq(9) & vbCrLf
        Text5.Text = Text5.Text & translate("Position: ") & qqq(10) & vbCrLf
    Else
        Text5.Text = Text5.Text & translate("[Image Information Unavailable]") & vbCrLf
    End If
    Text5.Text = Text5.Text & translate("* For italic/Bold: 1 is true, 0 is false")
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
    InsPic.Visible = False
    Blk.Visible = False
    ABCD.Visible = False
End Sub

Private Sub Text15_Change()
    Label65.Caption = Text15.Text & Combo3.Text & Combo5.Text
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
        Message.Caption = translate("No new messages.")
        Message.ForeColor = vbWhite
        showcnt = ShowCntPerMsg - 1
        GoTo rrr
    End If
    If current >= MsgContentList.ListCount Then
        Message.Caption = translate("No new messages.")
        Message.ForeColor = vbWhite
        showcnt = ShowCntPerMsg - 1
        GoTo rrr
    End If
    If showcnt = ShowCntPerMsg Then
        If MsgContentList.ListCount = 0 Then
            ProgressBar.Width = 15
            Message.Caption = ""
            Exit Sub
        End If
        If current + 1 >= MsgContentList.ListCount Then
            Message.Caption = translate("No new messages.")
            Message.ForeColor = vbWhite
            showcnt = ShowCntPerMsg - 1
            GoTo rrr
        End If
        showcnt = 0
        current = current + 1
        MsgContentList.ListIndex = current
        MsgColorList.ListIndex = current
        MsgTypeList.ListIndex = current
        Message.Caption = MsgTypeList.Text & MsgContentList.Text
        Message.ForeColor = ReverseColor(MsgColorList.Text)
    End If
rrr:
    ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
'    Message.Caption = Message.Caption & "(" & current + 1 & "/" & MsgTypeList.ListCount & ")"
End Sub


Private Sub Label1_Click()
    Tools.Visible = False
    Shape1.Left = Label1.Left - 120
    Shape1.Width = Label1.Width + 240
    Frame1.Visible = Not False
    Frame2.Visible = Not False
    Label10.Visible = Not False
    Label11.Visible = Not False
    List1.Visible = Not True
    Label14.Visible = Not True
    Label13.Visible = Not True
    Frame8.Visible = False
    Manage.Visible = Not True
    InsText.Visible = True
    AnswerLine.Visible = False
    LogMgr.Visible = False
    General.Visible = True
    Merge.Visible = False
    Copyright.Visible = False
    SaveLoad.Visible = False
    ABCD.Visible = False
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
    If StatName = "d" Then
        Temp.FontStrikethru = Not Temp.FontStrikethru
        Preview.Picture2.FontStrikethru = Not Preview.Picture2.FontStrikethru
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
        Temp.FontSize = Temp.FontSize * 0.59
        deltachange = deltachange + heightdata - Temp.Height
        'delta = delta + heightdata - Temp.Height
        Preview.Picture2.FontSize = Preview.Picture2.FontSize * 0.59
        stats(1) = True
    End If
    If StatName = "sd" Then
        Temp.AutoSize = True
        deltachange = deltachange - (heightdata - Temp.Height)
        'delta = delta - (heightdata - Temp.Height)
        Temp.FontSize = Temp.FontSize / 0.59
        Preview.Picture2.FontSize = Preview.Picture2.FontSize / 0.59
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
        NewMessage translate("Invaild Format."), vbRed, True
        Exit Sub
    End If
    For i = bound To 0 Step -1
        If str(i) <> "" Then Exit For Else bound = bound - 1
    Next
    For i = 0 To bound
        If str(i) <> "" Then Exit For Else start = i + 1
    Next
    If bound - start + 1 < 1 Or Text2.Text = "" Then
        NewMessage translate("Nothing can be previewed."), vbRed, True
        Exit Sub
    End If
    If Frame5.Visible = False And Combo1.Text = "" Then
        NewMessage translate("You have not selected the position of the image."), vbRed
        Exit Sub
    End If
    If Check17.Value = 1 Then Label29_Click
    InitPreview
    If Frame5.Visible = False Then
        On Error GoTo err
        reced = False
        Preview.Exports.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
        If Preview.Exports.Height > BotMargin - TopMargin Or Preview.Exports.Width > RightMargin - LeftMargin Then
            NewMessage translate("Your input is so large that we can't process it."), vbRed
            On Error Resume Next
            Preview.SystemCall = SystemCallFlag
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
        NewMessage translate("[TYPE=RUNTIME_ERROR][ERRORID=") & err.Number & translate("][ERRDESC.=") & err.Description & "]", vbRed
        On Error Resume Next
        Preview.SystemCall = SystemCallFlag
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
                        NewMessage translate("Auto split line is unsupportted for alignment mode 1 or 2."), vbRed
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
                        Preview.SystemCall = SystemCallFlag
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
                            NewMessage translate("The input will be split into multi parts"), vbBlue
                            NewMessage translate("select parts that you want to preview in the list."), vbBlue
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
                    NewMessage translate("The target size is too large, we are unable to process it."), vbRed
                    Text2.Text = wholestr
                    On Error Resume Next
                    Preview.SystemCall = SystemCallFlag
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
'    Preview.VScroll1.Max = Preview.Picture2.Height
    Text2.Text = wholestr
    Preview.Picture2.Top = 0
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
        NewMessage translate("Invaild Format."), vbRed, True
        Exit Sub
    End If
    For i = bound To 0 Step -1
        If str(i) <> "" Then Exit For Else bound = bound - 1
    Next
    For i = 0 To bound
        If str(i) <> "" Then Exit For Else start = i + 1
    Next
    If bound - start + 1 < 1 Or Text2.Text = "" Then
        NewMessage translate("Nothing can be previewed."), vbRed, True
        Exit Sub
    End If
    If Frame5.Visible = False And Combo1.Text = "" Then
        NewMessage translate("You have not selected the position of the image."), vbRed
        Exit Sub
    End If
    If Check17.Value = 1 Then Label29_Click
    InitPreview
    If Frame5.Visible = False Then
        On Error GoTo err
        reced = False
        Preview.Exports.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
        If Preview.Exports.Height > BotMargin - TopMargin Or Preview.Exports.Width > RightMargin - LeftMargin Then
            NewMessage translate("Your input is so large that we can't process it."), vbRed
            On Error Resume Next
            Preview.SystemCall = SystemCallFlag
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
        NewMessage translate("[TYPE=RUNTIME_ERROR][ERRORID=") & err.Number & translate("][ERRDESC.=") & err.Description & "]", vbRed
        On Error Resume Next
        Preview.SystemCall = SystemCallFlag
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
                    NewMessage translate("The input will be split into multi parts"), vbBlue
                    NewMessage translate("select parts that you want to preview in the list."), vbBlue
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
                        NewMessage translate("Auto split line is unsupportted for alignment mode 1 or 2."), vbRed
                        Text2.Text = wholestr
                        On Error Resume Next
                        Preview.SystemCall = SystemCallFlag
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
                            NewMessage translate("The input will be split into multi parts"), vbBlue
                            NewMessage translate("select parts that you want to preview in the list."), vbBlue
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
                    NewMessage translate("The target size is too large, we are unable to process it."), vbRed
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
                    Preview.SystemCall = SystemCallFlag
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
                    Debug.Print translate(" to ") & RightMargin
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
    Preview.SystemCall = SystemCallFlag
    Unload Preview
   
'    wholestr = str(0)
'    For i = 1 To UBound(str)
'        wholestr = wholestr & vbCrLf & str(i)
'    Next
    Text2.Text = wholestr
End Sub

Private Sub Label12_Click()
    Tools.Visible = False
    Shape1.Left = Label12.Left - 100
    Shape1.Width = Label12.Width + 200
    Frame1.Visible = False
    Frame2.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    List1.Visible = True
    Label13.Visible = True
    Label14.Visible = True
    Frame8.Visible = True
    Manage.Visible = True
    InsText.Visible = False
    AnswerLine.Visible = False
    General.Visible = False
    InsPic.Visible = False
    Blk.Visible = False
    Merge.Visible = False
    Copyright.Visible = False
    SaveLoad.Visible = False
    ABCD.Visible = False
End Sub

Private Sub Label13_Click()
    On Error GoTo err
    InitPreview
    Preview.Picture2.AutoSize = True
    Preview.Picture2.Picture = LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg")
    Preview.Picture2.AutoSize = False
err:
    Preview.NewMessage translate("Image which tracknumber=") & List1.Text & translate(" not found"), vbBlue
End Sub

Private Sub Label14_Click()
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub Label15_Click()
    Dim i As Long
    On Error Resume Next
    Dim X As Long, cnt As Long, pos As Long, lll As Long
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
        Preview.Export.Width = 1
        Preview.Export.Height = 1
        Preview.Export.Visible = True
        Preview.Export.Picture = LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg")
        If lll + Preview.Export.Height > BotMargin - TopMargin Then
            X = X + BotMargin - TopMargin - lll
            lll = 0
        End If
        If X > Preview.Picture2.Height Then
            Preview.Picture2.Height = Preview.Picture2.Height + PageWidth
        End If
        Preview.Picture2.Top = -X
        DoEvents
        Debug.Print "CurrentI" & i
        Preview.Picture2.PaintPicture LoadPicture(App.Path & "\Cache\" & List1.Text & ".jpg"), LeftMargin, X
        lll = lll + Preview.Export.Height
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
    Preview.Picture2.Top = 0
    goly = X
End Sub

Private Sub Timer2_Timer()
End Sub
