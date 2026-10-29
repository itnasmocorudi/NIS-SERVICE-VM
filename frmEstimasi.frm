VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEstimasi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Estimasi Service"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14760
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   14760
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   6840
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   15000
      TabIndex        =   70
      Top             =   6480
      Visible         =   0   'False
      Width           =   7575
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   10335
      Left            =   13200
      TabIndex        =   95
      Top             =   0
      Visible         =   0   'False
      Width           =   12855
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   8535
         Left            =   720
         TabIndex        =   96
         Top             =   120
         Width           =   11175
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   9
            Left            =   10080
            MaxLength       =   4
            TabIndex        =   113
            Top             =   1080
            Width           =   885
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   112
            Top             =   1410
            Width           =   1245
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   15
            Left            =   6960
            MaxLength       =   25
            TabIndex        =   111
            Top             =   2730
            Width           =   2490
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   14
            Left            =   6960
            MaxLength       =   25
            TabIndex        =   110
            Top             =   2400
            Width           =   2490
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   12
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   109
            Top             =   1740
            Width           =   4005
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   13
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   108
            Top             =   2070
            Width           =   2490
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   10
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   107
            Top             =   1410
            Width           =   2610
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   106
            Top             =   2385
            Width           =   1965
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   105
            Top             =   2055
            Width           =   1965
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   104
            Top             =   1725
            Width           =   2685
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   103
            Top             =   1725
            Width           =   1005
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   102
            Top             =   1065
            Width           =   4385
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   101
            Top             =   1395
            Width           =   4485
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   100
            Top             =   2720
            Width           =   1965
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   8
            Left            =   6960
            MaxLength       =   15
            TabIndex        =   99
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   0
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   97
            Top             =   240
            Width           =   2130
         End
         Begin RichTextLib.RichTextBox txtReq 
            Height          =   1575
            Left            =   120
            TabIndex        =   98
            Top             =   3600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   2778
            _Version        =   393217
            MaxLength       =   499
            TextRTF         =   $"frmEstimasi.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtKet 
            Height          =   1575
            Left            =   5640
            TabIndex        =   114
            Top             =   3600
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   2778
            _Version        =   393217
            MaxLength       =   199
            TextRTF         =   $"frmEstimasi.frx":007E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ucXPButton.XPButton cmdPrint 
            Height          =   495
            Left            =   8400
            TabIndex        =   115
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&p"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":00FC
            PICN            =   "frmEstimasi.frx":0118
            PICH            =   "frmEstimasi.frx":056F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdExit 
            Height          =   495
            Left            =   9720
            TabIndex        =   116
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&x"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":0A59
            PICN            =   "frmEstimasi.frx":0A75
            PICH            =   "frmEstimasi.frx":0F49
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdDelete 
            Height          =   495
            Left            =   3840
            TabIndex        =   149
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&d"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":1494
            PICN            =   "frmEstimasi.frx":14B0
            PICH            =   "frmEstimasi.frx":1A38
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdSave 
            Height          =   495
            Left            =   2640
            TabIndex        =   150
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&s"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":2074
            PICN            =   "frmEstimasi.frx":2090
            PICH            =   "frmEstimasi.frx":25A3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdCreate 
            Height          =   495
            Left            =   2640
            TabIndex        =   151
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&r"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":2B45
            PICN            =   "frmEstimasi.frx":2B61
            PICH            =   "frmEstimasi.frx":3096
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdCancel 
            Height          =   495
            Left            =   3840
            TabIndex        =   152
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&c"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":3689
            PICN            =   "frmEstimasi.frx":36A5
            PICH            =   "frmEstimasi.frx":3BE9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdAdd 
            Height          =   495
            Left            =   240
            TabIndex        =   153
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&a"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":41E0
            PICN            =   "frmEstimasi.frx":41FC
            PICH            =   "frmEstimasi.frx":4648
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdEdit 
            Height          =   495
            Left            =   1440
            TabIndex        =   154
            Top             =   5760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&e"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEstimasi.frx":4B10
            PICN            =   "frmEstimasi.frx":4B2C
            PICH            =   "frmEstimasi.frx":4FD7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   1
            Left            =   5490
            Top             =   1065
            Width           =   60
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000080FF&
            X1              =   7680
            X2              =   11040
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Kendaraan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   5835
            TabIndex        =   147
            Top             =   840
            Width           =   1560
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   6840
            TabIndex        =   146
            Top             =   2730
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NoMesin"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   5835
            TabIndex        =   145
            Top             =   2730
            Width           =   810
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   6840
            TabIndex        =   144
            Top             =   2400
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NoRangka"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   5835
            TabIndex        =   143
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   6840
            TabIndex        =   142
            Top             =   2070
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Warna"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   5835
            TabIndex        =   141
            Top             =   2070
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   5835
            TabIndex        =   140
            Top             =   1410
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   6840
            TabIndex        =   139
            Top             =   1410
            Width           =   60
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000080FF&
            X1              =   1755
            X2              =   5520
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1680
            TabIndex        =   138
            Top             =   2730
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CDMA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   137
            Top             =   2730
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   136
            Top             =   2400
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GSM"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   135
            Top             =   2400
            Width           =   405
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   1680
            TabIndex        =   134
            Top             =   2070
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PSTN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   1080
            TabIndex        =   133
            Top             =   2070
            Width           =   495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kota"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   132
            Top             =   1740
            Width           =   435
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telp"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   131
            Top             =   2070
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   130
            Top             =   2070
            Width           =   60
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   129
            Top             =   1740
            Width           =   60
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Pelanggan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   120
            TabIndex        =   128
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   127
            Top             =   1410
            Width           =   690
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   960
            TabIndex        =   126
            Top             =   1410
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   3
            Left            =   9600
            Top             =   1410
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   960
            TabIndex        =   125
            Top             =   1080
            Width           =   60
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   124
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pos  :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   3960
            TabIndex        =   123
            Top             =   1800
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Polisi                                            Tahun :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   5835
            TabIndex        =   122
            Top             =   1080
            Width           =   4200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   6840
            TabIndex        =   121
            Top             =   1080
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Rujukan  :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   120
            Top             =   255
            Width           =   1290
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   0
            Left            =   3600
            Top             =   240
            Width           =   60
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000080FF&
            X1              =   2400
            X2              =   5520
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Permintaan / Keluhan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   120
            TabIndex        =   119
            Top             =   3360
            Width           =   2145
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnosa Problem"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   5640
            TabIndex        =   118
            Top             =   3360
            Width           =   1770
         End
         Begin VB.Line Line9 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   10920
            Y1              =   5640
            Y2              =   5640
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   10920
            Y1              =   6360
            Y2              =   6360
         End
         Begin VB.Line Line11 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   11040
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line12 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   11040
            Y1              =   680
            Y2              =   680
         End
         Begin VB.Label Ket 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: Rujukan Heavy Repair  ::."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   5775
            TabIndex        =   117
            Top             =   240
            Width           =   5300
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   5760
            Shape           =   4  'Rounded Rectangle
            Top             =   180
            Width           =   5295
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9975
      Left            =   0
      Picture         =   "frmEstimasi.frx":54FC
      ScaleHeight     =   9975
      ScaleWidth      =   480
      TabIndex        =   148
      Top             =   0
      Width           =   475
   End
   Begin VB.Frame frmeWaktu 
      BackColor       =   &H00C0FFFF&
      Height          =   3375
      Left            =   16920
      TabIndex        =   50
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid GridWaktu 
         Height          =   2625
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4630
         _Version        =   393216
         Cols            =   6
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hari H : Tanggal Customer Walk In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   240
         TabIndex        =   52
         Top             =   2880
         Width           =   4425
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   480
      TabIndex        =   33
      Top             =   480
      Width           =   10635
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   915
         TabIndex        =   1
         Top             =   735
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   915
         TabIndex        =   2
         Top             =   1035
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   915
         TabIndex        =   3
         Top             =   1335
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   915
         TabIndex        =   5
         Top             =   1935
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   6270
         TabIndex        =   7
         Top             =   435
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   6270
         TabIndex        =   8
         Top             =   735
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   7920
         TabIndex        =   9
         Top             =   435
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   6270
         TabIndex        =   34
         Top             =   1035
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   9330
         TabIndex        =   35
         Top             =   720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   9240
         TabIndex        =   36
         Top             =   735
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   915
         TabIndex        =   37
         Top             =   450
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   915
         TabIndex        =   4
         Top             =   1635
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   3705
         TabIndex        =   38
         Top             =   450
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   9720
         TabIndex        =   11
         Top             =   1335
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   6270
         TabIndex        =   10
         Top             =   1335
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   915
         TabIndex        =   6
         Top             =   2235
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   9720
         TabIndex        =   65
         Top             =   2235
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   29
         Left            =   6270
         TabIndex        =   67
         Top             =   2280
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   9720
         TabIndex        =   75
         Top             =   1680
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   30
         Left            =   6270
         TabIndex        =   76
         Top             =   1680
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   270
         Left            =   10440
         TabIndex        =   92
         Top             =   435
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Customer && Kendaraan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3360
         TabIndex        =   93
         Top             =   45
         Width           =   3675
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1920
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   8280
         X2              =   10440
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   8280
         X2              =   10440
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Selesai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   8745
         TabIndex        =   79
         Top             =   1725
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Mulai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   8880
         TabIndex        =   78
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Selesai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4980
         TabIndex        =   77
         Top             =   1725
         Width           =   1110
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5445
         TabIndex        =   68
         Top             =   2325
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9360
         TabIndex        =   66
         Top             =   2280
         Width           =   195
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5520
         TabIndex        =   59
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   48
         Top             =   750
         Width           =   405
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   47
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   46
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   45
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th. Mobil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5475
         TabIndex        =   44
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5700
         TabIndex        =   43
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7320
         TabIndex        =   42
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   495
         Width           =   660
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   40
         Top             =   2250
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   39
         Top             =   1650
         Width           =   675
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   0
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   10575
      End
   End
   Begin VB.CommandButton tblSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan Estimasi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9480
      Width           =   1770
   End
   Begin VB.CommandButton cmdPilih 
      BackColor       =   &H00FFFF80&
      Caption         =   "No &Polisi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   75
      Width           =   1050
   End
   Begin VB.CommandButton cmdMech 
      BackColor       =   &H008080FF&
      Caption         =   "Mechanic Tersedia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   1320
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.CommandButton cmdWaktu 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Waktu Terpakai"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   480
      Width           =   2010
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   6405
      Left            =   480
      TabIndex        =   53
      Top             =   3000
      Width           =   12690
      Begin VB.CheckBox chSublet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Part"
         Height          =   255
         Left            =   11880
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid DGrid 
         Height          =   2745
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4842
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   19
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   3900
         TabIndex        =   54
         Top             =   360
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   18
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   650
         TabIndex        =   55
         Top             =   360
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   20
         Left            =   4000
         TabIndex        =   14
         Top             =   360
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99,99"
         PromptChar      =   ">"
      End
      Begin MSFlexGridLib.MSFlexGrid PartGrid 
         Height          =   2265
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   3995
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   21
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   22
         Left            =   3360
         TabIndex        =   17
         Top             =   3720
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   3240
         TabIndex        =   56
         Top             =   3720
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid subletGrid 
         Height          =   2745
         Left            =   6000
         TabIndex        =   25
         Top             =   720
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   4842
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   34
         Left            =   6000
         TabIndex        =   20
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   35
         Left            =   9165
         TabIndex        =   21
         Top             =   360
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   9060
         TabIndex        =   80
         Top             =   360
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   36
         Left            =   9600
         TabIndex        =   22
         Top             =   360
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   37
         Left            =   10440
         TabIndex        =   23
         Top             =   360
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
         Height          =   405
         Index           =   23
         Left            =   9600
         TabIndex        =   82
         Top             =   3720
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
         Height          =   405
         Index           =   24
         Left            =   9600
         TabIndex        =   83
         Top             =   4140
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   405
         Index           =   25
         Left            =   9600
         TabIndex        =   26
         Top             =   5025
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   405
         Index           =   26
         Left            =   9600
         TabIndex        =   84
         Top             =   5640
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   405
         Index           =   28
         Left            =   9600
         TabIndex        =   85
         Top             =   4560
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   38
         Left            =   4200
         TabIndex        =   18
         Top             =   3720
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99,99"
         PromptChar      =   ">"
      End
      Begin VB.Label lblFOP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6000
         TabIndex        =   156
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8940
         TabIndex        =   91
         Top             =   5745
         Width           =   480
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBLET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8880
         TabIndex        =   90
         Top             =   5130
         Width           =   540
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PART"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9030
         TabIndex        =   89
         Top             =   4245
         Width           =   390
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "_________________  +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9480
         TabIndex        =   88
         Top             =   5235
         Width           =   2940
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JASA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   9045
         TabIndex        =   87
         Top             =   3825
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OLI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9165
         TabIndex        =   86
         Top             =   4665
         Width           =   255
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sublet Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6000
         TabIndex        =   81
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   3480
         Width           =   765
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   " Keluhan && Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   150
         Width           =   1845
      End
   End
   Begin VB.CommandButton tblToOrder 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Buat &Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2040
      Width           =   930
   End
   Begin VB.CommandButton TblCetak 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak Estimasi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Width           =   2010
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9480
      Width           =   1770
   End
   Begin VB.CommandButton TblToBooking 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Buat &Booking "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2040
      Width           =   930
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3960
      TabIndex        =   49
      Top             =   67
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   60
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   503
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "??-9999-CCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   345
      Index           =   27
      Left            =   6000
      TabIndex        =   63
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999999999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   31
      Left            =   17265
      TabIndex        =   71
      Top             =   4395
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   714
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   32
      Left            =   17280
      TabIndex        =   72
      Top             =   5355
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   714
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   33
      Left            =   960
      TabIndex        =   27
      Top             =   9465
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Rujukan Heavy Repair"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   960
      Width           =   2010
   End
   Begin VB.Label JualNRM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7800
      TabIndex        =   155
      Top             =   0
      Width           =   60
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9855
      Left            =   0
      Picture         =   "frmEstimasi.frx":7698
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   74
      Top             =   9555
      Width           =   405
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PART"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   17160
      TabIndex        =   73
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Urut"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5280
      TabIndex        =   64
      Top             =   105
      Width           =   615
   End
End
Attribute VB_Name = "frmEstimasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 15 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Option Explicit

Dim rsTemp As ADODB.Recordset
Dim rsFOP As ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim fOther As Boolean
Dim stkPart As Integer
Dim namaPart As String
Dim satPart As String
Dim HrgPart As Double
Dim oliPart As Boolean
Dim genuineNon As Boolean
Dim baris As Integer
Dim barispart As Integer
Dim barisSublet As Integer
Dim Jrate As Double
Dim gNoRangka As String
Dim nWKT As String
Dim nJRate As Double
Dim i As Integer
Dim simpan As Boolean
Dim hrgLabour As Double
Dim pesan As String
Const MaxBrs = 30

Dim estimasiSublet As String

Private nEst As String
Private nRujuk As String

Public noOrder As String

Private tagihanFOP  As String
Private kodeMDL     As String
Private daftarSBE   As String

Public Property Let passingXX(VALUES As String)
    nEst = VALUES
End Property

Private Sub hapusForm()
1        On Error GoTo hapusForm_Error

2         gNoRangka = ""
3         For i = 1 To 37
4             Select Case i
                  Case 0
5                     Text(i).Mask = ""
6                     Text(i) = ""
7                     Text(i).Mask = IIf(cmdPilih.Caption = "No &Polisi", "??-9999-???", "CCCCCCCCCCCCCCCCCCCCCCCCC")
8                 Case 6, 7, 8
9                     Text(i).Mask = ""
10                    Text(i) = ""
11                    Text(i).Mask = "9999-9999999999"
12                Case 9
13                    Text(i).Mask = ""
14                    Text(i) = 0
15                    Text(i).Mask = "9999"
16                Case 14, 30
17                    Text(i).Mask = ""
18                    Text(i) = ">>->>->>>>"
19                    Text(i).Mask = "99-99-9999"
20                Case 15, 16
21                    Text(i).Mask = ""
22                    Text(i) = ">>:>>"
23                    Text(i).Mask = "99:99"
24                Case 20
25                    Text(i).Mask = ""
26                    Text(i) = 0
27                    Text(i).Mask = "99,99"
28                Case 22, 23, 24, 25, 26, 28, 31, 32, 35, 37
29                    Text(i) = 0
30                Case Else
31                    Text(i) = ""
32            End Select
33        Next
          
          ' ATUR GRID JOB
34        baris = 1
35        DGrid.Rows = 31
36        DGrid.cols = 6
37        DGrid.Clear
38        DGrid.FormatString = " | Tipe | Job Order | Rate | Disc | "
39        DGrid.ColWidth(0) = 0
40        DGrid.ColWidth(1) = 500
41        DGrid.ColWidth(2) = 3000
42        DGrid.ColWidth(3) = 600
43        DGrid.ColWidth(4) = 600
44        DGrid.ColWidth(5) = 0
          
          
          ' ATUR GRID PART
45        barispart = 1
46        partGrid.Rows = 31
47        partGrid.cols = 10
48        partGrid.Clear
49        partGrid.FormatString = " | No. Part | Nama Part |  | Qty | Harga | | | | Disc"
50        partGrid.ColWidth(0) = 0
51        partGrid.ColWidth(1) = 1200
52        partGrid.ColWidth(2) = 2000
53        partGrid.ColWidth(3) = 0
54        partGrid.ColWidth(4) = 500
55        partGrid.ColWidth(5) = 1200
56        partGrid.ColWidth(6) = 0
57        partGrid.ColWidth(7) = 0
58        partGrid.ColWidth(8) = 0
59        partGrid.ColWidth(9) = 600
          
          ' ATUR GRID SUBLET
60        barisSublet = 1
61        subletGrid.Rows = 31
62        subletGrid.cols = 6
63        subletGrid.Clear
64        subletGrid.FormatString = " | Sublet | Qty | Satuan | Harga | Part "
65        subletGrid.ColWidth(0) = 0
66        subletGrid.ColWidth(1) = 3000
67        subletGrid.ColWidth(2) = 600
68        subletGrid.ColWidth(3) = 1000
69        subletGrid.ColWidth(4) = 1200
70        subletGrid.ColWidth(5) = 0
          
          'Text(27) = isiNoUrut
71        Text(17) = gSA
72        nJRate = 0
73        isJualNRM = False
74        JualNRM.Caption = ""
75       On Error GoTo 0
76       Exit Sub

hapusForm_Error:

77        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmEstimasi")
End Sub

Private Function isiNoUrut() As String
1        On Error GoTo isiNoUrut_Error

2         Set rsTemp = objO.FetchData("SELECT noest FROM ESTIMASI WITH(NOLOCK) WHERE left(noest,2) = '" & Right(year(isnow), 2) & "' AND SUBSTRING(noest,3,2) = '" & Format(Month(isnow), "0#") & "' ORDER BY noest DESC")
3         If Not rsTemp.EOF Then
4             isiNoUrut = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(CDbl(Right(rsTemp("NOEST"), 4)) + 1, "0###")
5         Else
6             isiNoUrut = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
7         End If

8        On Error GoTo 0
9        Exit Function

isiNoUrut_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiNoUrut of Form frmEstimasi")
End Function

Private Sub PanggilCustomer(No As String, kode As Integer)
          
1        On Error GoTo PanggilCustomer_Error

2         If kode = 1 Then
3             Set rsTemp = objO.FetchData("SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI = '" & No & "'")
4         ElseIf kode = 2 Then
5             Set rsTemp = objO.FetchData("SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & No & "'")
6         End If
7         If Not rsTemp.EOF Then
8             Text(1) = rsTemp("FS_KD_PLG")
9             Text(2) = rsTemp("FS_KD_WIL")
10            Text(3) = Left(NulSID(rsTemp("CUSTOMER")), 50)
11            Text(4) = Left(Trim(Nul(rsTemp("CUST_ADRES"))), 60)
12            Text(5) = rsTemp("KOTA")
13            Text(6) = rsTemp("TELPON")
14            Text(7) = rsTemp("GSM")
15            Text(8) = rsTemp("CDMA")
16            Text(9) = rsTemp("FS_TH_MOBIL")
17            Text(10) = rsTemp("MDL")
18            Text(11) = rsTemp("MODEL")
19                        Call CexChassis_Discon(Text(11))
20                        If isOnCarClass = False Then
21                            MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
22                        End If
23            Text(12) = rsTemp("FMDL")
24            Text(13) = rsTemp("COLOR")
25            gNoRangka = rsTemp("NO_CHASSIS")
26            If Log.freepart = 1 And NulSID(NulSID(gNoRangka)) <> "" Then
27                 tagihanFOP = ""
28                 kodeMDL = objO.FetchData_String("select Mdl from SingleIDOri..UnitCustomer where norangka = '" & NulSID(gNoRangka) & "'")
29                 If CekUnitTerlibatFOP(NulSID(gNoRangka), kodeMDL, tagihanFOP) Then
30                     lblFOP.Caption = "Free Part + Oli : " & tagihanFOP
31                     sSQL = "select distinct FS_ITEM from [" & dbReplikasiHO & "]..HO_FOP_PROGRAM where mdl = '" & NulSID(kodeMDL) & "' and NIK = '" & Mid(NulSID(NulSID(gNoRangka)), 10, 1) & "'"
32                    Set rsFOP = objO.FetchData(sSQL)
33                    daftarSBE = ""
34                    If Not rsFOP.EOF Then
35                       rsFOP.MoveFirst
36                       Do While Not rsFOP.EOF
37                           daftarSBE = daftarSBE & " - " & NulSID(rsFOP!fs_item) & vbCrLf
38                           rsFOP.MoveNext
39                       Loop
40                    End If
41                    MsgBox "Unit customer berikut terlibat FREE OIL PART " & tagihanFOP & " untuk SB kelipatan " & vbCrLf & daftarSBE, vbInformation, "Unit terlibat Free Service + Part"
                      
42                 Else
43                     lblFOP.Caption = ""
44                 End If
              
45            End If
46            Text(17) = gSA
47        End If

48       On Error GoTo 0
49       Exit Sub

PanggilCustomer_Error:

50        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilCustomer of Form frmEstimasi")
End Sub

Private Sub panggilDataEstimasi(No As String, kode As Integer)
      Dim rsHlp As ADODB.Recordset
      Dim rsHlp2 As ADODB.Recordset
          
1        On Error GoTo panggilDataEstimasi_Error

2         If kode = 1 Then        ' No Polisi
3             Set rsTemp = objO.FetchData("SELECT * FROM ESTIMASI WITH(NOLOCK) WHERE NO_POLISI = '" & Trim(No) & "' AND NOEST = '" & Trim(nEst) & "' ORDER BY NOEST DESC")
4         ElseIf kode = 2 Then    ' No Rangka
5             Set rsTemp = objO.FetchData("SELECT * FROM ESTIMASI WITH(NOLOCK) WHERE NO_RANGKA = '" & Trim(No) & "' AND NOEST = '" & Trim(nEst) & "' ORDER BY NOEST DESC")
6         End If
7         If Not rsTemp.EOF Then
8             rsTemp.MoveLast
9             Text(27) = rsTemp("NOEST")
10            Text(1) = rsTemp("FS_KD_PLG")
11            Text(2) = rsTemp("FS_KD_WIL")
12            Text(3) = rsTemp("CUSTOMER")
13            Text(4) = rsTemp("CUST_ADRES")
14            Text(5) = rsTemp("KOTA")
15            Text(6) = rsTemp("TELPON")
16            Text(7) = rsTemp("GSM")
17            Text(8) = rsTemp("CDMA")
18            Text(9) = rsTemp("FS_TH_MOBIL")
19            Text(10) = rsTemp("MDL")
20            Text(11) = rsTemp("MODEL")
21                        Call CexChassis_Discon(Text(11))
22                        If isOnCarClass = False Then
23                            MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
24                        End If
25            Text(12) = rsTemp("FMDL")
26            Text(13) = rsTemp("COLOR")
27            Text(17) = rsTemp("SA")
28            Text(23) = Format(rsTemp("TOTLABOUR"), "#0,#")
29            Text(24) = Format(rsTemp("TOTPART_G"), "#0,#")
30            Text(28) = Format(rsTemp("TOTOLI_G"), "#0,#")
31            Text(25) = Format(rsTemp("TOTSUBLET"), "#0,#")
32            Text(26) = Format(rsTemp("TOTAL"), "#0,#")
33            gNoRangka = rsTemp("NO_RANGKA")
34            Text(33) = Trim(rsTemp("NOTE"))
              
              ' ISI JOB ORDER
35            nJRate = 0
               
36            Set rsHlp = objO.FetchData("SELECT * FROM ESTJOB WITH(NOLOCK) WHERE NOEST = '" & rsTemp("NOEST") & "' ORDER BY KLPKERJA")
37            If Not rsHlp.EOF Then
38                i = 1
39                Do While Not rsHlp.EOF
40                    With DGrid
41                        .TextMatrix(i, 1) = rsHlp("KLPKERJA")
42                        .TextMatrix(i, 2) = rsHlp("REP")
43                        .TextMatrix(i, 3) = rsHlp("RTE")
44                        .TextMatrix(i, 4) = rsHlp("disc")
45                        .TextMatrix(i, 5) = rsHlp("FOTHER")
46                    End With
47                    nJRate = nJRate + CDbl(rsHlp("RTE"))
48                    rsHlp.MoveNext
49                    baris = baris + 1
50                    i = i + 1
51                Loop
52            End If
              
53            Text(15) = rsTemp("JAM")
54            Text(16) = rsTemp("JAMSELESAI")
55            Text(14) = TglDDMMYYYY(rsTemp("TANGGAL"))
              
              ' ISI PART
               
              ' AND B.TIPE = 'O'
56            Set rsHlp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI, TIPE, GENUINE FROM ESTPART A WITH(NOLOCK) LEFT JOIN T_OLI B WITH(NOLOCK) ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI WHERE NOEST = '" & rsTemp("NOEST") & "' ORDER BY KODEITEM")
57            If Not rsHlp.EOF Then
58                i = 1
59                Do While Not rsHlp.EOF
60                    With partGrid
61                        .TextMatrix(i, 1) = rsHlp("KODEITEM")
62                        .TextMatrix(i, 2) = rsHlp("NAMABARANG")
63                        .TextMatrix(i, 3) = rsHlp("KODESATUAN")
64                        .TextMatrix(i, 4) = rsHlp("QTY")
65                        .TextMatrix(i, 6) = IIf(rsHlp("TIPE") = "O", True, False)
66                        .TextMatrix(i, 7) = IIf(rsHlp("GENUINE") = 1, True, False)
67                        .TextMatrix(i, 9) = rsHlp("disc")
                          
                               
68                            Set rsHlp2 = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH(NOLOCK) WHERE KODEITEM = '" & rsHlp("KODEITEM") & "'")
69                            If Not rsHlp2.EOF Then
70                                .TextMatrix(i, 5) = Format(rsHlp2("HARGAJUALA"), "#,0#")
71                            Else
72                                .TextMatrix(i, 5) = Format(rsHlp("HARGA"), "#,0#")
73                            End If
74                            Call setClose(rsHlp2)
                              
75                    End With
76                    rsHlp.MoveNext
77                    i = i + 1
78                    barispart = barispart + 1
79                Loop
80            End If
              
              ' ISI SUBLET
               
81            Set rsHlp = objO.FetchData("SELECT * FROM ESTSUBLET WITH(NOLOCK) WHERE NOEST = '" & rsTemp("NOEST") & "' ORDER BY PART")
82            If Not rsHlp.EOF Then
83                i = 1
84                Do While Not rsHlp.EOF
85                    With subletGrid
86                        .TextMatrix(i, 1) = Trim(rsHlp("SUBLET"))
87                        .TextMatrix(i, 2) = CDbl(rsHlp("QTY"))
88                        .TextMatrix(i, 3) = Trim(rsHlp("SATUAN"))
89                        .TextMatrix(i, 4) = Format(CDbl(rsHlp("HARGA")), "#,0#") '* ((100 + gUpSublet) / 100)
90                        .TextMatrix(i, 5) = IIf(Trim(rsHlp("PART")) = 0, Unchecked, Checked)
91                    End With
92                    rsHlp.MoveNext
93                    i = i + 1
94                    barisSublet = barisSublet + 1
95                Loop
96            End If
97        End If

98       On Error GoTo 0
99       Exit Sub

panggilDataEstimasi_Error:

100       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilDataEstimasi of Form frmEstimasi")
End Sub

Private Sub TampilWaktu(Tggl As Date)
      Dim rsWaktu As ADODB.Recordset
      Dim nStall As String
      Dim i As Integer

1        On Error GoTo TampilWaktu_Error

2         GridWaktu.TextMatrix(0, 0) = "..."
3         GridWaktu.TextMatrix(0, 1) = "STALL"
          
          'FUNGSI MENGHASILKAN WAKTU TERSEDIA BERDASARKAN TGL TTN (HARI H)
4         If DateDiff("d", isnow, Tggl) >= 2 Then
5             Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl, "mm/dd/yyyy") & "'")
6             GridWaktu.TextMatrix(0, 2) = "H - 2"
7             GridWaktu.TextMatrix(0, 3) = "H - 1"
8             GridWaktu.TextMatrix(0, 4) = "  H  "
9             GridWaktu.TextMatrix(0, 5) = "H + 1"
10        ElseIf DateDiff("d", isnow, Tggl) = 1 Then
11            Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 1, "mm/dd/yyyy") & "'")
12            GridWaktu.TextMatrix(0, 2) = "H - 1"
13            GridWaktu.TextMatrix(0, 3) = "  H  "
14            GridWaktu.TextMatrix(0, 4) = "H + 1"
15            GridWaktu.TextMatrix(0, 5) = "H + 2"
16        ElseIf DateDiff("d", isnow, Tggl) <= 0 Then
17            Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 2, "mm/dd/yyyy") & "'")
18            GridWaktu.TextMatrix(0, 2) = "  H  "
19            GridWaktu.TextMatrix(0, 3) = "H + 1"
20            GridWaktu.TextMatrix(0, 4) = "H + 2"
21            GridWaktu.TextMatrix(0, 5) = "H + 3"
22        End If
          
          'FUNSGI UNTUK MENAMPILKAN WAKTU TERSEDIA H-2 SAMPAI H+1
          'Set rsWaktu = objO.FetchData("SELECT NOSTALL, MAX(ISNULL([H-2],'')) as 'H-2', MAX(ISNULL([H-1],'')) as 'H-1', MAX(ISNULL([H],'')) AS 'H', MAX(ISNULL([H+1],'')) AS 'H+1' " _
          '                          & "FROM STALLWKF GROUP BY NoStall ORDER BY CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
23        Set rsWaktu = objO.FetchData("SELECT * FROM STALLWKF WITH(NOLOCK) ORDER BY LEFT(NOSTALL,1),CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
24        If Not rsWaktu.EOF Then
25            i = 1
26            GridWaktu.Rows = rsWaktu.RecordCount + 1
              
27            GridWaktu.ColWidth(0) = 200
28            GridWaktu.ColWidth(1) = 700
29            GridWaktu.ColWidth(2) = 1200
30            GridWaktu.ColWidth(3) = 1200
31            GridWaktu.ColWidth(4) = 1200
32            GridWaktu.ColWidth(5) = 1200
33            nStall = ""
              
34            Do While Not rsWaktu.EOF
35                If Trim(rsWaktu("NOSTALL")) = nStall Then
36                    GridWaktu.TextMatrix(i, 1) = ""
37                    GridWaktu.TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
38                    GridWaktu.TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
39                    GridWaktu.TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
40                    GridWaktu.TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
41                Else
42                    GridWaktu.TextMatrix(i, 0) = "*"
43                    GridWaktu.TextMatrix(i, 1) = Trim(rsWaktu("NOSTALL"))
44                    GridWaktu.TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
45                    GridWaktu.TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
46                    GridWaktu.TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
47                    GridWaktu.TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
48                    nStall = Trim(rsWaktu("NOSTALL"))
49                End If
50                i = i + 1
51                rsWaktu.MoveNext
52            Loop
53        End If

54       On Error GoTo 0
55       Exit Sub

TampilWaktu_Error:

56        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TampilWaktu of Form frmEstimasi")
End Sub

Function createPart() As Boolean

1        On Error GoTo createPart_Error

2         createPart = False
          
3         If rsPart.State Then rsPart.Close
4         Set rsPart = Nothing
5         With rsPart
6             .Fields.Append "KODEITEM", adVarChar, 50
7             .Fields.Append "NAMAITEM", adVarChar, 200
8             .Fields.Append "SATITEM", adVarChar, 30
9             .Fields.Append "QTY", adInteger
10            .Fields.Append "HRG", adDouble
11            .Fields.Append "GANTI", adBoolean
12            .Fields.Append "BAWA", adBoolean
13            .Fields.Append "OLI", adBoolean
14            .Fields.Append "GENUINE", adBoolean
15            .Fields.Append "nID", adDouble
16            .Fields.Append "NAMABYSA", adVarChar, 50
17            .Open
18        End With
19        If rsPart.State Then createPart = True
20        Set rsTmpPart = rsPart
21        Exit Function

22       On Error GoTo 0
23       Exit Function

createPart_Error:

24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createPart of Form frmEstimasi")
          
End Function

Private Sub HeadColumn(frm As Object)
1        On Error GoTo HeadColumn_Error

2         With frm
3             .partGrid.Columns(0).Caption = "KODE PART"
4             .partGrid.Columns(0).DataField = "KODEITEM"
5             .partGrid.Columns(0).Width = 1500
6             .partGrid.Columns(1).Caption = "NAMA PART"
7             .partGrid.Columns(1).DataField = "NAMAITEM"
8             .partGrid.Columns(1).Width = 2500
9             .partGrid.Columns(2).Caption = "SATUAN PART"
10            .partGrid.Columns(2).DataField = "SATITEM"
11            .partGrid.Columns(2).Width = 500
12            .partGrid.Columns(3).Caption = "QTY"
13            .partGrid.Columns(3).DataField = "QTY"
14            .partGrid.Columns(3).Width = 800
15            .partGrid.Columns(4).Caption = "GANTI"
16            .partGrid.Columns(4).DataField = "GANTI"
17            .partGrid.Columns(4).Alignment = dbgCenter
18            .partGrid.Columns(5).Caption = "PART SENDIRI"
19            .partGrid.Columns(5).DataField = "BAWA"
20            .partGrid.Columns(5).Alignment = dbgCenter
21        End With

22       On Error GoTo 0
23       Exit Sub

HeadColumn_Error:

24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure HeadColumn of Form frmEstimasi")
End Sub

Private Sub simpanEstimasi()

1        On Error GoTo simpanEstimasi_Error

2         ClearsSql
3         gsSql1(0) = "NOEST"
4         gsSql2(0) = Text(27)
5         gsSql1(1) = "FS_KD_PLG"
6         gsSql2(1) = Text(1)
7         gsSql1(2) = "FS_KD_WIL"
8         gsSql2(2) = Text(2)
9         gsSql1(3) = "CUSTOMER"
10        gsSql2(3) = Text(3)
11        gsSql1(4) = "CUST_ADRES"
12        gsSql2(4) = Text(4)
13        gsSql1(5) = "KOTA"
14        gsSql2(5) = Text(5)
15        gsSql1(6) = "TELPON"
16        gsSql2(6) = Text(6)
17        gsSql1(7) = "GSM"
18        gsSql2(7) = Text(7)
19        gsSql1(8) = "CDMA"
20        gsSql2(8) = Text(8)
21        gsSql1(9) = "FS_TH_MOBIL"
22        gsSql2(9) = Text(9)
23        gsSql1(10) = "MDL"
24        gsSql2(10) = Text(10)
25        gsSql1(11) = "MODEL"
26        gsSql2(11) = Text(11)
27        gsSql1(12) = "FMDL"
28        gsSql2(12) = Text(12)
29        gsSql1(13) = "COLOR"
30        gsSql2(13) = Text(13)
31        gsSql1(14) = "TANGGAL"
32        gsSql2(14) = Format(TglYYYYMMDD(Text(14)), "MM/DD/YYYY")
33        gsSql1(15) = "JAM"
34        gsSql2(15) = Text(15)
35        gsSql1(16) = "JAMSELESAI"
36        gsSql2(16) = Text(16)
37        gsSql1(17) = "SA"
38        gsSql2(17) = Text(17)
39        gsSql1(18) = "TOTLABOUR"
40        gsSql2(18) = CDbl(Text(23))
41        gsSql1(19) = "TOTPART_G"
42        gsSql2(19) = CDbl(Text(24))
43        gsSql1(20) = "TOTPART_NG"
44        gsSql2(20) = CDbl(Text(31))
45        gsSql1(21) = "TOTSUBLET"
46        gsSql2(21) = CDbl(Text(25))
47        gsSql1(22) = "TOTOLI_G"
48        gsSql2(22) = CDbl(Text(28))
49        gsSql1(23) = "TOTOLI_NG"
50        gsSql2(23) = CDbl(Text(32))
51        gsSql1(24) = "TOTAL"
52        gsSql2(24) = CDbl(Text(26))
53        gsSql1(25) = "STATUS"
54        gsSql2(25) = ""
55        gsSql1(26) = "NOTE"
56        gsSql2(26) = Text(33)
          
57        If cmdPilih.Caption = "No &Polisi" Then
58            gsSql1(27) = "NO_POLISI"
            ' ==================================================================================
            ' 28 fEB 2017 | THANIA | revisi jika inputan no polisi hanya 2 blok depan / belakang atau hanya angka tanpa "-"
            ' ==================================================================================
59            If Right(Replace(Replace(Text(0), ">", ""), " ", ""), 1) = "-" Then
60                gsSql2(27) = Mid(Replace(Replace(Text(0), ">", ""), " ", ""), 1, Len(Replace(Replace(Text(0), ">", ""), " ", "")) - 1)
61            ElseIf Left(Replace(Replace(Text(0), ">", ""), " ", ""), 1) = "-" Then
62                gsSql2(27) = Mid(Replace(Replace(Text(0), ">", ""), " ", ""), 2, Len(Replace(Replace(Text(0), ">", ""), " ", "")) - 1)
63            Else
64                gsSql2(27) = Replace(Replace(Replace(Text(0), ">", ""), " ", ""), "--", "")
65            End If
            ' ==================================================================================
            
66            gsSql1(28) = "NO_RANGKA"
67            gsSql2(28) = gNoRangka
68        ElseIf cmdPilih.Caption = "No &Rangka" Then
69            gsSql1(27) = "NO_RANGKA"
70            gsSql2(27) = Replace(Replace(Text(0), ">", ""), " ", "")
71        End If
          
72        Set rsTemp = objO.FetchData("SELECT * FROM ESTIMASI WITH(NOLOCK) WHERE NOEST = '" & Text(27) & "'")
73        If Not rsTemp.EOF Then
74            sSQL = GensSql2("ESTIMASI", True) + "WHERE NOEST = '" & Text(27) & "'"
75        Else
76            sSQL = GensSql2("ESTIMASI", False)
77        End If
78        Call objHT.UpdateData_BengkelOnly(sSQL)

79       On Error GoTo 0
80       Exit Sub

simpanEstimasi_Error:

81        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanEstimasi of Form frmEstimasi")
End Sub

Private Sub simpanEstJob()

1        On Error GoTo simpanEstJob_Error

2         Call objHT.DeleteData_BengkelOnly("DELETE ESTJOB WHERE NOEST = '" & Trim(Text(27)) & "'")
3         For i = 1 To baris - 1
4             If DGrid.TextMatrix(i, 1) <> "" Then
                  'Set rsTemp = objO.FetchData("SELECT * FROM ESTJOB WHERE NOEST = '" & Text(27) & "' AND KLPKERJA = '" & DGrid.TextMatrix(i, 1) & "' AND REP = '" & DGrid.TextMatrix(i, 2) & "'")
                  'If Not rsTemp.EOF Then
                  '    Call objHT.UpdateData_BengkelOnly("UPDATE ESTJOB SET RTE = '" & Replace(CDbl(DGrid.TextMatrix(i, 3)), ",", ".") & "', FOTHER = '" & IIf(DGrid.TextMatrix(i, 4), 1, 0) & "' WHERE NOEST = '" & Text(27) & "' AND KLPKERJA = '" & DGrid.TextMatrix(i, 1) & "' AND REP = '" & DGrid.TextMatrix(i, 2) & "'")
                  'Else
5                     Call objHT.InsertData_BengkelOnly("INSERT INTO ESTJOB (NOEST,KLPKERJA,REP,RTE,FOTHER,disc) VALUES ('" & Text(27) & "', '" & DGrid.TextMatrix(i, 1) & "', '" & DGrid.TextMatrix(i, 2) & "', '" & Replace(CDbl(DGrid.TextMatrix(i, 3)), ",", ".") & "', '" & IIf(DGrid.TextMatrix(i, 5), 1, 0) & "'," & k2t(DGrid.TextMatrix(i, 4)) & ")")
                  'End If
6             End If
7         Next

8        On Error GoTo 0
9        Exit Sub

simpanEstJob_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanEstJob of Form frmEstimasi")
End Sub

Private Sub simpanEstPart()
   On Error GoTo simpanEstPart_Error

    Call objHT.DeleteData_BengkelOnly("DELETE ESTPART WHERE NOEST = '" & Trim(Text(27)) & "'")
    For i = 1 To barispart - 1
        If partGrid.TextMatrix(i, 1) <> "" Or partGrid.TextMatrix(i, 2) <> "" Then
'            Set rsTemp = objO.FetchData("SELECT * FROM ESTPART WHERE NOEST = '" & Text(27) & "' AND KODEITEM = '" & partGrid.TextMatrix(i, 1) & "'")
            'If Not rsTemp.EOF Then
                'Call objHT.UpdateData_BengkelOnly("UPDATE ESTPART SET NAMABARANG = '" & partGrid.TextMatrix(i, 2) & "', KODESATUAN = '" & partGrid.TextMatrix(i, 3) & "', QTY = '" & CDbl(partGrid.TextMatrix(i, 4)) & "', HARGA = " & CDbl(partGrid.TextMatrix(i, 5)) & " WHERE NOEST = '" & Text(27) & "' AND KODEITEM = '" & partGrid.TextMatrix(i, 1) & "'")
            'Else
                Call objHT.InsertData_BengkelOnly("INSERT INTO ESTPART (NOEST,KODEITEM,NAMABARANG,KODESATUAN,QTY,harga,disc) VALUES ('" & Text(27) & "', '" & partGrid.TextMatrix(i, 1) & "', '" & partGrid.TextMatrix(i, 2) & "','" & partGrid.TextMatrix(i, 3) & "', '" & CDbl(partGrid.TextMatrix(i, 4)) & "', " & CDbl(partGrid.TextMatrix(i, 5)) & "," & k2t(CDbl(IIf(partGrid.TextMatrix(i, 9) = "", 0, partGrid.TextMatrix(i, 9)))) & ")")
            'End If
        End If
    Next

   On Error GoTo 0
   Exit Sub

simpanEstPart_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanEstPart of Form frmEstimasi")
End Sub

Private Sub simpanEstSublet()
1        On Error GoTo simpanEstSublet_Error

2         Call objHT.DeleteData_BengkelOnly("DELETE ESTSUBLET WHERE NOEST = '" & Trim(Text(27)) & "'")
3         For i = 1 To barisSublet - 1
4             With subletGrid
5                 If .TextMatrix(i, 1) <> "" Then
                      'Set rsTemp = objO.FetchData("SELECT * FROM ESTSUBLET WHERE NOEST = '" & Text(27) & "' AND SUBLET = '" & .TextMatrix(i, 1) & "'")
                      'If Not rsTemp.EOF Then
                      '    Call objHT.UpdateData_BengkelOnly("UPDATE ESTSUBLET SET SUBLET = '" & .TextMatrix(i, 1) & "', SATUAN = '" & .TextMatrix(i, 3) & "', QTY = '" & CDbl(.TextMatrix(i, 2)) & "', HARGA = " & CDbl(.TextMatrix(i, 4)) & ", PART = '" & .TextMatrix(i, 5) & "' WHERE NOEST = '" & Text(27) & "' AND SUBLET = '" & .TextMatrix(i, 1) & "'")
                      'Else
6                         Call objHT.InsertData_BengkelOnly("INSERT INTO ESTSUBLET VALUES ('" & Text(27) & "', '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 3) & "', '" & CDbl(.TextMatrix(i, 2)) & "', " & CDbl(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 5) & "')")
                      'End If
7                 End If
8             End With
9         Next

10       On Error GoTo 0
11       Exit Sub

simpanEstSublet_Error:

12        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanEstSublet of Form frmEstimasi")
End Sub

Private Sub Cetak(txt As String, txt2 As String, txt3 As String, txt4 As String)
      Dim oRs As Object   'New ADODB.Recordset  'Object
      Dim jHlp As Object  'New ADODB.Recordset
      Dim pHlp As Object  'New ADODB.Recordset
      Dim sHlp As Object  'New ADODB.Recordset
      Dim Report As New RptEstimasi
      Dim crxSubReportJob As CRAXDRT.Report
      Dim crxSubReportPart As CRAXDRT.Report
      Dim crxSubReportSublet As CRAXDRT.Report
1        On Error GoTo Cetak_Error

2     Screen.MousePointer = vbHourglass
          
3         Set oRs = objO.FetchData(txt)
4         Report.database.SetDataSource oRs
5         Report.txtKota.SetText gKotaCbg
6         Report.txtTerbilang.SetText Trim(Angka2Huruf(Round(Text(26)), 1))
          
7         Set jHlp = objO.FetchData(txt2)
8         Set crxSubReportJob = Report.Subreport1.OpenSubreport
9         crxSubReportJob.database.SetDataSource jHlp
          
10        Set pHlp = objO.FetchData(txt3)
11        Set crxSubReportPart = Report.Subreport2.OpenSubreport
12        crxSubReportPart.database.SetDataSource pHlp
          
13        Set sHlp = objO.FetchData(txt4)
14        Set crxSubReportSublet = Report.Subreport3.OpenSubreport
15        crxSubReportSublet.database.SetDataSource sHlp
16        Report.txtNote.SetText Trim(Text(33))
          
17        With CRViewer1
18            .ReportSource = Report
19            .ViewReport
20            .Zoom 100
21            .Visible = True
22            cmdTutup.Visible = True
23        End With
24        cmdTutup.Visible = True
25        Screen.MousePointer = vbDefault

26       On Error GoTo 0
27       Exit Sub

Cetak_Error:

28        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetak of Form frmEstimasi")
End Sub

Private Sub chSublet_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub cmdPilih_Click()
    If cmdPilih.Caption = "No &Polisi" Then
        cmdPilih.Caption = "No &Rangka"
        Text(0).Mask = "CCCCCCCCCCCCCCCCCCCCCCCCC"
    ElseIf cmdPilih.Caption = "No &Rangka" Then
        cmdPilih.Caption = "No &Polisi"
        Text(0).Mask = "??-9999-CCC"
    End If
End Sub

Private Sub cmdPrint_Click()

Dim Report As New RptRujukan
Screen.MousePointer = vbHourglass
    
    Dim oRs As ADODB.Recordset
    Set oRs = objO.FetchData("Select * from Rujukan WITH(NOLOCK) Where NoRujuk = '" & txt(0).Text & "'")
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
        cmdTutup.Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdWaktu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Text(23) <> "" And Text(23) <> ">>->>->>>>" Then
        frmeWaktu.Visible = True
        frmeWaktu.Top = cmdWaktu.Top + cmdWaktu.Height + 100
        frmeWaktu.Left = cmdWaktu.Left - (frmeWaktu.Width - cmdWaktu.Width - 50)
        Call TampilWaktu(Trim(Text(23)))
    Else
        MsgBox "Tanggal booking harus diisi"
        Text(23).SelLength = Len(Text(23))
        Text(23).SetFocus
        Exit Sub
    End If
End Sub

Private Sub DGrid_DblClick()
Dim a As Integer
Dim disc As Single
    If baris > 1 And DGrid.TextMatrix(DGrid.row, 1) <> "" Then
        a = DGrid.row
        disc = 0
        'AMBIL
        With DGrid
            Text(18) = .TextMatrix(a, 1)
            Text(19) = .TextMatrix(a, 2)
            Jrate = .TextMatrix(a, 3)
            fOther = .TextMatrix(a, 5)
            If .TextMatrix(a, 4) <> "" Then
                disc = .TextMatrix(a, 4)
            End If
            If .TextMatrix(a, 4) <> "" Then
'                If .TextMatrix(a, 4) Then
'                    fOther = True
'                    Text(20).Visible = True
'                    Text(20).Mask = ""
'                    Text(20) = Jrate
'                Else
'                    fOther = False
'                    Text(20).Visible = False
'                End If
'            Else
'                fOther = False
'                Text(20).Visible = False
                Text(20).Visible = True
                Text(20).Mask = ""
                Text(20) = .TextMatrix(a, 4)
            End If
        End With
        'Text(40) = Format(CDbl(Text(40)) - hitungPart(partGrid.TextMatrix(A, 3), hrgPart), "#0,#")
        Call cekPartPack(DGrid.TextMatrix(a, 1), DGrid.TextMatrix(a, 2), Text(10), 2, 1, barispart, frmEstimasi)
        
        nJRate = nJRate - Jrate
        If Text(15) <> ">>:>>" And Text(14) <> ">>->>->>>>" Then
            nWKT = hitungWaktu(CDbl(nJRate), Text(15), Text(14))
            Text(16) = Left(nWKT, InStr(nWKT, "|") - 1)
            Text(30) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
        End If
        
        '30/01/2012
         myChassis = gNoRangka
        'hrgLabour = hitungLabour(dGrid.TextMatrix(baris - 1, 3), Text(11))
         hrgLabour = (hitungLabour(DGrid.TextMatrix(a, 3), Text(11)) * (100 - disc) / 100)
        If isJualNRM = True Then
            Me.JualNRM.Caption = nJualNRM
        Else
            Me.JualNRM.Caption = ""
        End If
        'HAPUS
        For i = a To baris
            With DGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
            End With
        Next
        baris = baris - 1
        
        Text(23) = Format(CDbl(Text(23) - (CDbl(hrgLabour))), "#0,#") ' * CDbl(Jrate)
        Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
        
        Text(18).SetFocus
    End If
End Sub

Private Sub dGrid_GotFocus()
      Dim bolehTambah As Boolean
      Dim disc        As Double
          
1         If Trim(Text(19)) <> "" Then
2             For i = 0 To baris
3                 If DGrid.TextMatrix(i, 1) = Trim(Text(19)) Then
4                     bolehTambah = False
5                     Text(19) = ""
6                     Exit Sub
7                 Else
8                     bolehTambah = True
9                 End If
10            Next
11            If Jrate = 0 Then
12                bolehTambah = False
13            End If
              
14            If Text(20) = "" Then
15                disc = 0
16            Else
17                disc = CDbl(Replace(Text(20), ">", "0"))
18            End If
              
19            If bolehTambah Then
                  ' JIKA PEKERJAAN INPUT SENDIRI
20                If fOther Then
21                    Jrate = CDbl(Replace(Replace(Text(20), " ", ""), ">", ""))
22                End If
                  
23                If MaxBrs >= baris Then
24                    With DGrid
25                        .TextMatrix(baris, 1) = Text(18)
26                        .TextMatrix(baris, 2) = Text(19)
27                        .TextMatrix(baris, 3) = Jrate
28                        .TextMatrix(baris, 4) = disc
29                        .TextMatrix(baris, 5) = fOther
30                    End With
31                    baris = baris + 1
                      
                      ' ISI PART PACKddddddddddddddddddddd
32                    If Log.freepart = 1 And tagihanFOP <> "" Then
33                        Call cekFopPart(Text(18), Text(19), Text(10), 1, 1, barispart, frmEstimasi)
34                    Else
35                        Call cekPartPack(Text(18), Text(19), Text(10), 1, 1, barispart, frmEstimasi)
36                    End If
                      
37                    If Log.freepart = 1 Then
38                        Call cekSubletFOP(NulSID(Text(10)), NulSID(Text(19)))
39                    End If
                      ''Call cekPartPack(Text(18), Text(19), Text(10), 1, 1, barispart, frmEstimasi)
                      
40                    myChassis = gNoRangka
41                    Text(23) = Format(CDbl(Text(23)) + (((100 - disc) / 100) * Val(hitungLabour(CDbl(Jrate), Trim(Text(11))))), "#0,0")
42                    If isJualNRM = True Then
43                        Me.JualNRM.Caption = nJualNRM
44                    Else
45                        Me.JualNRM.Caption = ""
46                    End If
47                    nJRate = nJRate + Jrate
48                    If Text(15) <> ">>:>>" And Text(14) <> ">>->>->>>>" Then
49                        nWKT = hitungWaktu(CDbl(nJRate), Text(15), Text(14))
50                        Text(16) = Left(nWKT, InStr(nWKT, "|") - 1)
51                        Text(30) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
52                    End If
      '                Text(16) = hitungWaktu(CDbl(Jrate), Text(15), Text(14))
53                    Text(19) = ""
54                    Text(18) = ""
55                    Text(20) = ">>,>>"
56                    Text(20).Visible = False
57                    Text(18).SetFocus
                      
58                    hrgLabour = Val(hitungLabour(CDbl(Jrate), Trim(Text(11)))) / CDbl(Jrate)
59                    Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
60                    If isJualNRM = True Then
61                        Me.JualNRM.Caption = nJualNRM
62                    Else
63                        Me.JualNRM.Caption = ""
64                    End If
65                Else
66                    MsgBox "Maksimal " & str(MaxBrs) & " !"
67                    Text(19).SetFocus
68                End If
69            End If
70        End If
End Sub

Private Sub Form_Load()
    Text(0).Mask = ""
    Text(0) = ""
    Text(0).Mask = IIf(cmdPilih.Caption = "No &Polisi", "??-9999-CCC", "CCCCCCCCCCCCCCCCCCCCCCCCC")
    Call hapusForm
    cmdWaktu.Visible = False
    cmdMech.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmeWaktu.Visible = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > 100 Then
        CRViewer1.Height = ScaleHeight - 100
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    'asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub partGrid_DblClick()
      Dim a As Integer

1         If barispart > 1 And partGrid.TextMatrix(partGrid.row, 1) <> "" Then
2             a = partGrid.row
              'AMBIL
3             With partGrid
4                 Text(21) = .TextMatrix(a, 1)
5                 namaPart = .TextMatrix(a, 2)
6                 satPart = .TextMatrix(a, 3)
7                 Text(22) = CDbl(.TextMatrix(a, 4))
8                 HrgPart = .TextMatrix(a, 5)
9                 oliPart = .TextMatrix(a, 6)
10                Text(38) = Format(IIf(.TextMatrix(a, 9) = "", 0, .TextMatrix(a, 9)), "#00.00#")
11                genuineNon = IIf(.TextMatrix(a, 7) = "", True, .TextMatrix(a, 7))
12            End With
              
13            If oliPart Then
      '            If genuineNon Then
14                    Text(28) = Format(CDbl(Text(28)) - (((100 - Text(38)) / 100) * hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5))), "#0,#")
      '            Else
      '                Text(32) = Format(CDbl(Text(32)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
      '            End If
15            Else
      '            If genuineNon Then
16                    Text(24) = Format(CDbl(Text(24)) - (((100 - Text(38)) / 100) * hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5))), "#0,#")
      '            Else
      '                Text(31) = Format(CDbl(Text(31)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
      '            End If
17            End If
18            Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
              
              'HAPUS
19            For i = a To barispart
20                With partGrid
21                    .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
22                    .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)       ' KODE PART
23                    .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)       ' NAMA PART
24                    .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)       ' SAT PART
25                    .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)       ' QTY
26                    .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)       ' HRG PART
27                    .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)       ' OLI PART
28                    .TextMatrix(i, 7) = .TextMatrix(i + 1, 7)       ' GENUINE PART/OLI
29                    .TextMatrix(i, 9) = .TextMatrix(i + 1, 9)
30                End With
31            Next
32            barispart = barispart - 1
33            Text(21).SetFocus
34        End If
End Sub

Private Sub PartGrid_GotFocus()
      Dim bolehTambah As Boolean
      Dim disc        As Double
1         If Trim(Text(21)) <> "" And CDbl(Text(22)) > 0 And namaPart <> "" Then
2             For i = 1 To barispart
3                 If partGrid.TextMatrix(i, 1) = Trim(Text(21)) Then
4                     bolehTambah = False
5                     Text(35) = ""
6                     Exit Sub
7                 Else
8                     bolehTambah = True
9                 End If
10            Next
11            If bolehTambah Then
12                If NulSID(Text(38)) = "" Then
13                    disc = 0
14                Else
15                    disc = CDbl(Replace(NulSID(Text(38)), ">", "0"))
16                End If
17                If MaxBrs >= barispart Then
18                    With partGrid
19                        .TextMatrix(barispart, 1) = Text(21)
20                        .TextMatrix(barispart, 2) = namaPart
21                        .TextMatrix(barispart, 3) = satPart
22                        .TextMatrix(barispart, 4) = Text(22)
23                        .TextMatrix(barispart, 5) = HrgPart
24                        .TextMatrix(barispart, 6) = oliPart
25                        .TextMatrix(barispart, 7) = genuineNon
26                        .TextMatrix(barispart, 9) = disc
27                    End With
28                    barispart = barispart + 1
29                    Text(21) = ""
30                    Text(22) = 0
31                    namaPart = ""
32                    If oliPart Then
33                        Text(28) = Format(CDbl(Text(28)) + (((100 - disc) / 100) * hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5))), "#0,0")
34                    Else
35                        Text(24) = Format(CDbl(Text(24)) + (((100 - disc) / 100) * hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5))), "#0,0")
36                    End If
37                    Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
38                    Text(21).SetFocus
39                Else
40                    MsgBox "Maksimal " & str(MaxBrs) & " !"
41                    Text(21).SetFocus
42                End If
43            End If
44        ElseIf CDbl(Text(22)) = 0 And Trim(Text(21)) <> "" Then
45            Text(22).SetFocus
46            Exit Sub
47        End If
End Sub

Private Sub subletGrid_DblClick()
Dim a As Integer

    If barisSublet > 1 And subletGrid.TextMatrix(subletGrid.row, 1) <> "" Then
        a = subletGrid.row
        
        With subletGrid
            Text(34) = Trim(.TextMatrix(a, 1))
            Text(35) = CDbl(.TextMatrix(a, 2))
            Text(36) = Trim(.TextMatrix(a, 3))
            Text(37) = CDbl(.TextMatrix(a, 4)) '* (100 / (100 + gUpSublet))
            chSublet = IIf(CDbl(.TextMatrix(a, 5)) = 0, Unchecked, Checked)
        End With
        
        For i = a To barisSublet
            With subletGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
                .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
            End With
        Next
        Text(25) = Format(CDbl(Text(25)) - ((CDbl(Text(35)) * CDbl(Text(37)))), "#0,#") ' * ((100 + gUpSublet) / 100)
        Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
                
        barisSublet = barisSublet - 1
        Text(34).SetFocus
    End If
End Sub

Private Sub subletGrid_GotFocus()
Dim bolehTambah As Boolean
    
    If Trim(Text(34)) <> "" And Trim(Text(35)) <> "" And Trim(Text(37)) <> "" Then
        For i = 0 To baris
            If subletGrid.TextMatrix(i, 1) = Trim(Text(34)) Then
                bolehTambah = False
                Text(34) = ""
                Exit Sub
            Else
                bolehTambah = True
            End If
        Next
        
        If bolehTambah Then
            If MaxBrs >= barisSublet Then
                With subletGrid
                    .TextMatrix(barisSublet, 1) = Trim(Text(34))                                ' NAMA SUBLET
                    .TextMatrix(barisSublet, 2) = CDbl(Text(35))                                ' QTY
                    .TextMatrix(barisSublet, 3) = Trim(Text(36))                                ' SATUAN
                    .TextMatrix(barisSublet, 4) = Format(CDbl(Text(37)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
                    .TextMatrix(barisSublet, 5) = IIf(chSublet = Checked, 1, 0)                 ' STATUS PART
                End With
                barisSublet = barisSublet + 1
                Text(25) = Format(CDbl(Text(25)) + ((CDbl(Text(35)) * CDbl(Text(37)))), "#0,#") '* ((100 + gUpSublet) / 100)
                Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
                
                Text(34) = ""
                Text(35) = 0
                Text(36) = ""
                Text(37) = 0
                chSublet = Unchecked
                Text(34).SetFocus
            End If
        End If
    End If
End Sub

Private Sub TblCetak_Click()
      Dim SSQL2 As String
      Dim SSQL3 As String
      Dim SSQL4 As String
1        On Error GoTo TblCetak_Click_Error

2         If Text(0) <> ">>->>>>->>>" Then
3             myChassis = gNoRangka
4             If CexChassis_Discon(Trim(Text(11))) = True Then
      '            sSQL = "SELECT NO_ESTIMASI, CUSTOMER, CUST_ADRES, KOTA, TELPON, GSM, CDMA, MODE, FMDL, COLOR, NO_POLISI, " & _
                         "TANGGAL, NO_RANGKA, JAM, JAMSELESAI, SA, ((TOTALLABOUR * (100-" & nDiscNRM & ")) / 100) as TOTALLABOUR, TOTPART,TOTOLI, TOTALSUBLET,TOTAL, STATUS " & _
                         "FROM VIEWESTIMASI A INNER JOIN T_PETUGAS B WITH(NOLOCK) ON A.SA = B.FS_KD_PETUGAS WHERE NO_ESTIMASI = '" & text(27) & "'"
5                 sSQL = "SELECT NO_ESTIMASI, CUSTOMER, CUST_ADRES, KOTA, TELPON, GSM, CDMA, MODE, FMDL, COLOR, NO_POLISI, " & _
                         "TANGGAL, NO_RANGKA, JAM, JAMSELESAI, SA, TOTALLABOUR, TOTPART,TOTOLI, TOTALSUBLET,TOTAL, STATUS " & _
                         "FROM VIEWESTIMASI A INNER JOIN T_PETUGAS B WITH(NOLOCK) ON A.SA = B.FS_KD_PETUGAS WHERE NO_ESTIMASI = '" & Text(27) & "'"

                  SSQL2 = "SELECT NO_ESTIMASI, REP, RTE, ((Price * (100-" & nDiscNRM & ")) / 100) as Price, ((TOTAL * (100-" & nDiscNRM & ")) / 100) as TOTAL FROM VIEWJOBESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
6                 'SSQL2 = "SELECT NO_ESTIMASI, REP, RTE,  Price,  TOTAL FROM VIEWJOBESTIMASI WHERE NO_ESTIMASI = '" & text(27) & "'"
7                 SSQL3 = "SELECT NO_ESTIMASI, KODEITEM, NAMABARANG, QTY, HARGA, TOTAL FROM VIEWPARTESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
8                 SSQL4 = "SELECT NO_ESTIMASI, SUBLET, SATUAN, QTY, HARGA, TOTAL FROM VIEWSUBLETESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
9             Else
      '            sSQL = "SELECT NO_ESTIMASI, CUSTOMER, CUST_ADRES, KOTA, TELPON, GSM, CDMA, MODE, FMDL, COLOR, NO_POLISI, " & _
      '                   "TANGGAL, NO_RANGKA, JAM, JAMSELESAI, SA, ((TOTALLABOUR * (100-" & nDiscNON & ")) / 100) as TOTALLABOUR, TOTPART,TOTOLI, TOTALSUBLET,TOTAL, STATUS " & _
      '                   "FROM VIEWESTIMASI A INNER JOIN T_PETUGAS B ON A.SA = B.FS_KD_PETUGAS WHERE NO_ESTIMASI = '" & text(27) & "'"
10                sSQL = "SELECT NO_ESTIMASI, CUSTOMER, CUST_ADRES, KOTA, TELPON, GSM, CDMA, MODE, FMDL, COLOR, NO_POLISI, " & _
                         "TANGGAL, NO_RANGKA, JAM, JAMSELESAI, SA,TOTALLABOUR, TOTPART,TOTOLI, TOTALSUBLET,TOTAL, STATUS " & _
                         "FROM VIEWESTIMASI A INNER JOIN T_PETUGAS B ON A.SA = B.FS_KD_PETUGAS WHERE NO_ESTIMASI = '" & Text(27) & "'"
                  SSQL2 = "SELECT NO_ESTIMASI, REP, RTE, ((Price * (100-" & nDiscNON & ")) / 100) as Price, ((TOTAL * (100-" & nDiscNON & ")) / 100) as TOTAL FROM VIEWJOBESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
11                'SSQL2 = "SELECT NO_ESTIMASI, REP, RTE, Price, TOTAL FROM VIEWJOBESTIMASI WHERE NO_ESTIMASI = '" & text(27) & "'"
12                SSQL3 = "SELECT NO_ESTIMASI, KODEITEM, NAMABARANG, QTY, HARGA, TOTAL FROM VIEWPARTESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
13                SSQL4 = "SELECT NO_ESTIMASI, SUBLET, SATUAN, QTY, HARGA, TOTAL FROM VIEWSUBLETESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
      '            sSQL = "SELECT * FROM VIEWESTIMASI A INNER JOIN T_PETUGAS B ON A.SA = B.FS_KD_PETUGAS WHERE NO_ESTIMASI = '" & Text(27) & "'"
      '            SSQL2 = "SELECT * FROM VIEWJOBESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
      '            SSQL3 = "SELECT * FROM VIEWPARTESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
      '            SSQL4 = "SELECT * FROM VIEWSUBLETESTIMASI WHERE NO_ESTIMASI = '" & Text(27) & "'"
14            End If
15            Call Cetak(sSQL, SSQL2, SSQL3, SSQL4)
16        Else
17            MsgBox "Lengkapi Data", vbCritical, "KELENGKAPAN"
18            Exit Sub
19        End If

20       On Error GoTo 0
21       Exit Sub

TblCetak_Click_Error:

22        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TblCetak_Click of Form frmEstimasi")
End Sub

Private Sub TblKeluar_Click()
    Layar.Tag = ""
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim pesan As String
Dim rsSearch As ADODB.Recordset
    On Error GoTo vbError
    boleh = True
    pesan = "Lengkapi data berikut " & vbCr
    ' ======================================================================================================================
    ' 28 fEB 2017 | THANIA | revisi jika inputan no polisi hanya 2 blok depan / belakang atau hanya angka tanpa "-"
    ' ======================================================================================================================
    If Replace(Trim(Replace(Replace(Text(0), ">", ""), " ", "")), "--", "") = "" Then
        boleh = False
        pesan = pesan + "No Polisi. "
    End If
    ' ======================================================================================================================
    If Text(3) = "" Then
        boleh = False
        pesan = pesan + "Customer. "
    End If
    If Text(4) = "" Then
        boleh = False
        pesan = pesan + "Alamat. "
    End If
    If Text(5) = "" Then
        boleh = False
        pesan = pesan + "Kota. "
    End If
    If Text(6) = "" And Text(7) = "" And Text(8) = "" Then
        boleh = False
        pesan = pesan + "Telepon. "
    End If
    If Text(9) = "" Then
        boleh = False
        pesan = pesan + "Th Mobil. "
    End If
    If Text(10) = "" Then
        boleh = False
        pesan = pesan + "Model. "
    Else
        If NulSID(Text(10)) <> "" Then
             If Log.RPLHistori = 0 Then
                Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
            Else
                Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
            End If
            If rsSearch.BOF Then
                MsgBox "Kelas kendaraan tidak sesuai master Model. Perbaiki data Model kendaraan di master Customer", vbInformation
                boleh = False
                With frmSIDCustomer
                    If NulSID(gNoRangka) <> "" Then
                        If Left(gNoRangka, 3) <> "" Then
                        .Text(50) = Left(gNoRangka, 3)
                        End If
                        .Text(1) = Mid(gNoRangka, 4, 6)
                        .Text(0) = Mid(gNoRangka, 10, IIf(Len(gNoRangka) > 0, Len(gNoRangka) - 9, 0))
                    End If
                    frmSIDCustomer.show vbModal
                End With
            End If
        End If
    End If
    If Text(14) = ">>->>->>>>" Then
        boleh = False
        pesan = pesan + "Tanggal Estimasi. "
    End If
    If Text(15) = ">>:>>" Then
        boleh = False
        pesan = pesan + "Jam Estimasi. "
    End If
    If baris <= 1 Then
        If MsgBox("Tanpa Job Order ?") = vbNo Then
            boleh = False
            pesan = pesan + "Job Order. "
        End If
    End If
    
    If boleh Then
        Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
        If Text(27) = "" Then
            Text(27) = isiNoUrut
        End If
        Call saveLogMenu(gSA, ComptName, gVersi, "frmEstimasi", "tblSimpan", 0)
        erTranAll = False
        onTR = True
        Call objHT.OpenTransactionBengkel
        
        Call simpanEstimasi
        Call simpanEstJob
        If barispart > 1 Then
            Call simpanEstPart
        Else
            Call objHT.DeleteData_BengkelOnly("DELETE ESTPART WHERE NOEST = '" & Trim(Text(27)) & "'")
        End If
        If barisSublet > 1 Then
            Call simpanEstSublet
        Else
            Call objHT.DeleteData_BengkelOnly("DELETE ESTSUBLET WHERE NOEST = '" & Trim(Text(27)) & "'")
        End If
        
        If Log.DIDX = 1 Then
            If noOrder <> "" Then
                Call objHT.DeleteData_BengkelOnly("UPDATE JOBDTADX SET ESTBYWAC = '" & Trim(Text(27)) & "' WHERE no_wo = '" & noOrder & "'")
            End If
            
        End If
        
        If erTranAll = False Then
            Call objHT.CloseTransactionBengkel
            onTR = False
            Call saveLogMenu(gSA, ComptName, gVersi, "frmEstimasi", "tblSimpan", 1)
            MsgBox "No. Estimasi : " & Trim(Text(27))
            TblCetak_Click
            TblKeluar.SetFocus
        Else
            Call objHT.CloseTransactionBengkel_err
            onTR = False
            Call saveLogMenu(gSA, ComptName, gVersi, "frmEstimasi", "tblSimpan", 2)
            MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
            Exit Sub
        End If
    Else
        MsgBox pesan, vbCritical, "KELENGKAPAN DATA"
        Text(3).SetFocus
        Exit Sub
    End If
    Exit Sub
    
vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frmEstimasi", "tblSimpan", 2)
    MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmEstimasi"), vbInformation, "Gagal Simpan"
    Exit Sub
End Sub

Private Sub TblToBooking_Click()
      Dim nopol As String
      Dim temp As String
      Dim totalLabor As Double

1        On Error GoTo TblToBooking_Click_Error

2         If Text(0) <> ">>->>>>->>>" Then
              ' TAMPILKAN DATA KE FORM BOOKING
3             Layar.Tag = "estBook"
4             With frmBooking
5                 gBantuan = CDbl(baris)
6                 gBantuan2 = CDbl(barispart)
7                 gBantuan4 = CDbl(barisSublet)
                  
8                 If cmdPilih.Caption = "No &Polisi" Then
                      '.text(0) = text(0)
9                     nopol = Replace(Replace(Text(0), ">", ""), " ", "")
10                    gBantuan3 = nopol
11                    .Text(0) = Left(nopol, InStr(nopol, "-") - 1)
12                    .Text(56) = Mid(nopol, InStr(nopol, "-") + 1, InStr((Right(nopol, Len(nopol) - InStr(nopol, "-"))), "-") - 1)
13                    .Text(57) = Mid(Right(nopol, Len(nopol) - InStr(nopol, "-")), InStr(Right(nopol, Len(nopol) - InStr(nopol, "-")), "-") + 1, Len(Right(nopol, Len(nopol) - InStr(nopol, "-"))) - InStr(Right(nopol, Len(nopol) - InStr(nopol, "-")), "-"))
14                    .Text(41) = gNoRangka
15                ElseIf cmdPilih.Caption = "No &Rangka" Then
16                    .Text(41) = Text(0)
17                End If
18                .Text(2) = Text(1)
19                .Text(44) = Text(2)
20                .Text(3) = Text(3)
21                .Text(4) = Text(4)
22                .Text(5) = Text(5)
23                .Text(8) = Text(6)
24                .Text(6) = Text(7)
25                .Text(7) = Text(8)
26                .Text(9) = Text(9)
27                .Text(10) = Text(10)
28                .Text(11) = Text(11)
29                .Text(12) = Text(12)
30                .Text(13) = Text(13)
31                .Text(23) = Text(14)
32                .Text(24) = Text(15)
33                .Text(26) = Text(17)
34                .Text(59) = Text(30)
35                .Text(45).Mask = ""
36                .Text(45) = Text(16)
37                .Text(39) = Text(23)
38                .Text(40) = Format(hitungPartTanpaDisc(True), "#0,0") ''Text(24)
39                .Text(55) = Format(hitungPartTanpaDisc(False), "#0,0") '' Text(28)
40                .Text(52) = Text(25)
41                .Text(53) = Text(26)
42                Set rsTemp = objO.FetchData("SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & gNoRangka & "'")
43                If Not rsTemp.EOF Then
44                    .Text(14) = rsTemp("FS_KD_DEALER")
45                    .Text(15) = rsTemp("FS_NM_DEALER")
46                    .Text(16).Mask = ""
47                    .Text(16) = IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsTemp("DOTAM"), "/", "-")))
48                    .Text(17).Mask = ""
49                    .Text(17) = IIf(Replace(rsTemp("DDATE"), "/", "-") = "01-01-3000", "", IIf(Replace(rsTemp("DDATE"), "/", "-") = "01-01-1900", "", Replace(rsTemp("DDATE"), "/", "-")))
50                    .Text(2) = rsTemp("FS_KD_PLG")
51                    .Text(44) = rsTemp("FS_KD_WIL")
52                End If
                  
                  ' ISI JOB
53                totalLabor = 0
                  Dim nTemp As String
                  Dim pesan As String
54                With .DGrid
55                    For i = 1 To 7
56                        pesan = ""
57                        If DGrid.TextMatrix(i, 1) <> "" Then
ulangi:
58                            nTemp = InputBox("Input Mekanik yang mengerjakan !" & vbCr & "Job : " & DGrid.TextMatrix(i, 1) & pesan, "MEKANIK")
59                            nTemp = UCase(nTemp)
60                            If nTemp = "" Then
61                                temp = MsgBox("Membatalkan booking ?", vbYesNo, "VALIDASI MEKANIK")
62                                If temp = vbNo Then
63                                    GoTo ulangi
64                                Else
65                                    Exit Sub
66                                End If
67                            Else
68                                If Not cekMekanik(nTemp, CDbl(DGrid.TextMatrix(i, 3)), Replace(Text(14), "-", "/"), Trim(Text(15))) Then
69                                    pesan = gBantuan3
70                                    GoTo ulangi
71                                Else
72                                    frmBooking.Text(64) = Trim(nTemp)
73                                    .TextMatrix(i, 4) = DGrid.TextMatrix(i, 1)
74                                    .TextMatrix(i, 1) = DGrid.TextMatrix(i, 2)
75                                    .TextMatrix(i, 2) = DGrid.TextMatrix(i, 3)
76                                    .TextMatrix(i, 3) = DGrid.TextMatrix(i, 4)
77                                    totalLabor = totalLabor + Val(hitungLabour(CDbl(DGrid.TextMatrix(i, 3)), Trim(Text(11))))
                                      
78                                End If
79                            End If
80                        End If
81                    Next
82                End With
83                .Text(39) = Format(totalLabor, "#0,#")
84                .Text(53) = Format(CDbl(totalLabor) + CDbl(Text(24)) + CDbl(Text(28)) + CDbl(Text(25)), "#0,#")
                  ' ISI SUBLET
                  Dim jJ As Integer
                  Dim jM As Integer
                  Dim total As Integer
85                With .subletGrid
86                    For i = 1 To 30
87                        estimasiSublet = "00:00"
88                        If subletGrid.TextMatrix(i, 1) <> "" Then
89                            estimasiSublet = InputBox("Masukan estimasi waktu untuk sublet " & subletGrid.TextMatrix(i, 1) & " (HH:MM) ")
90                            If estimasiSublet <> "" Then
91                                If InStr(estimasiSublet, ":") > 0 Then
92                                    jJ = CInt(Left(Trim(estimasiSublet), InStr(Trim(estimasiSublet), ":") - 1))
93                                    jM = Mid(Trim(estimasiSublet), InStr(Trim(estimasiSublet), ":") + 1, Len(Trim(estimasiSublet)) - InStr(Trim(estimasiSublet), ":"))
                                      
94                                    total = (jJ * 60) + jM
95                                    If total = 0 Then
96                                        estimasiSublet = "00:00"
97                                    ElseIf total > 240 Then
98                                        estimasiSublet = "00:00"
99                                    End If
100                               Else
101                                   estimasiSublet = "00:00"
102                               End If
103                           Else
104                               estimasiSublet = "00:00"
105                           End If
106                       End If
107                       .TextMatrix(i, 0) = subletGrid.TextMatrix(i, 0)
108                       .TextMatrix(i, 1) = subletGrid.TextMatrix(i, 1)
109                       .TextMatrix(i, 2) = subletGrid.TextMatrix(i, 2)
110                       .TextMatrix(i, 3) = subletGrid.TextMatrix(i, 3)
111                       .TextMatrix(i, 4) = subletGrid.TextMatrix(i, 4)
112                       .TextMatrix(i, 5) = subletGrid.TextMatrix(i, 5)
113                       .TextMatrix(i, 6) = estimasiSublet
114                   Next
115               End With
                      
                  ' ISI PART
116               With .partGrid
117                   For i = 1 To 14
118                       .TextMatrix(i, 1) = partGrid.TextMatrix(i, 1)
119                       .TextMatrix(i, 2) = IIf(partGrid.TextMatrix(i, 1) = "", "", partGrid.TextMatrix(i, 2))
120                       .TextMatrix(i, 3) = partGrid.TextMatrix(i, 3)
121                       .TextMatrix(i, 4) = partGrid.TextMatrix(i, 4)
122                       .TextMatrix(i, 5) = partGrid.TextMatrix(i, 5)
123                       .TextMatrix(i, 6) = partGrid.TextMatrix(i, 6)
124                       .TextMatrix(i, 7) = IIf(partGrid.TextMatrix(i, 1) = "", partGrid.TextMatrix(i, 2), "")
125                       .TextMatrix(i, 8) = i
126                   Next
127               End With
      '            .cmdNo.SetFocus
128               Call .panggilWarrantyBok(gNoRangka, Text(10))
129               Call .panggilHistoryBok(gNoRangka, Text(29))
130               If cekPiutang(Trim(Text(1)), Text(10), gNoRangka) Then
131               End If
132               .Text(43) = cekMOU(Trim(Text(10)), gNoRangka)
133               TOChild frmEstimasi.hWnd, frmBooking.hWnd
134               .show
135               .Text(23).SelLength = Len(Text(23))
136               .Text(23).SetFocus
137               .Text(88) = Text(27)
138           End With
139       Else
140           MsgBox "LENGKAPI DATA", vbCritical, "KELENGKAPAN DATA"
141           Exit Sub
142       End If

143      On Error GoTo 0
144      Exit Sub

TblToBooking_Click_Error:

145       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TblToBooking_Click of Form frmEstimasi")
End Sub

Private Sub tblToOrder_Click()
      Dim rsHlp As ADODB.Recordset
      Dim SSQL4 As String
      Dim totalLabor As Double
1        On Error GoTo tblToOrder_Click_Error

2         If Text(0) <> ">>->>>>->>>" Then
              ' TAMPILKAN DATA KE FORM ORDER
3             If Text(14) = TglDDMMYYYY(isnow) Then
4             With frmOrder
              
5                 gBantuan = baris
6                 gBantuan2 = barispart
7                 gBantuan3 = nJRate
                  
8                 Layar.Tag = "estOrder"
                            
              
                  Dim isID As Double
9                 isID = 1
10                If rsPart.State <> adStateClosed Then rsPart.Close
11                If createPart Then
                  
12                    With rsPart
13                            For i = 1 To 30
14                                If partGrid.TextMatrix(i, 1) <> "" Or partGrid.TextMatrix(i, 2) <> "" Then
15                                    .AddNew
16                                    !KodeItem = partGrid.TextMatrix(i, 1)
17                                    !NAMAITEM = IIf(partGrid.TextMatrix(i, 1) = "", "", partGrid.TextMatrix(i, 2))
18                                    !SATITEM = partGrid.TextMatrix(i, 3)
19                                    !qty = partGrid.TextMatrix(i, 4)
20                                    !oli = partGrid.TextMatrix(i, 6)
          '                            !GENUINE = partGrid.TextMatrix(I, 7)
21                                    !hrg = Replace(partGrid.TextMatrix(i, 5), ".", "")
22                                    !ganti = False
23                                    !BAWA = False
24                                    !nID = isID
25                                    !namabySA = IIf(partGrid.TextMatrix(i, 1) = "", partGrid.TextMatrix(i, 2), "")
26                                    .Update
27                                    isID = isID + 1
28                                End If
29                            Next
30                    End With
31                End If
                  
32                If cmdPilih.Caption = "No &Polisi" Then
33                    .Text(14).Mask = ""
34                    .Text(14) = Replace(Text(0), ">", "")
35                    .Text(15) = gNoRangka
36                    .Text(0) = gNoRangka
37                    SSQL4 = "SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI = '" & Trim(Replace(Text(0), ">", "")) & "'"
38                ElseIf cmdPilih.Caption = "No &Rangka" Then
39                    .Text(0) = Text(0)
40                    .Text(15) = Text(0)
41                    SSQL4 = "SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(Replace(Text(0), ">", "")) & "'"
42                End If
43                Call .cmdPilih_Click
                  '.cmdPilih.Caption = "No. Rangka"
44                .Text(1) = Text(1)
45                .Text(2) = Text(2)
46                .Text(3) = Text(3)
47                .Text(4) = Text(4)
48                .Text(5) = Text(5)
49                .Text(6) = Text(6)
50                .Text(7) = Text(7)
51                .Text(8) = Text(8)
52                .Text(9) = Text(9)
53                .Text(10) = Text(10)
54                .Text(11) = Text(11)
55                .Text(12) = Text(12)
56                .Text(13) = Text(13)
57                .Text(56) = Text(23)
58                .Text(57) = Text(24)
59                .Text(59) = Text(28)
60                .Text(60) = Text(25)
61                .Text(61) = Text(26)
                   
62                Set rsHlp = objO.FetchData(SSQL4)
63                If Not rsHlp.EOF Then
64                    .Text(1) = rsHlp("FS_KD_PLG")
65                    .Text(2) = rsHlp("FS_KD_WIL")
66                End If
                  
                  
                  ' ISI JOB
                  Dim nTemp As String
67                totalLabor = 0
68                With .DGrid
69                    For i = 1 To 7
70                        If Me.DGrid.TextMatrix(i, 1) <> "" Then
ulangi:
71                            nTemp = InputBox("Input Mekanik yang mengerjakan !" & vbCr & "Job : " & DGrid.TextMatrix(i, 2) & pesan, "MEKANIK")
72                            nTemp = UCase(nTemp)
73                            If nTemp = "" Then
74                                If MsgBox("Membatalkan Order ?", vbYesNo, "VALIDASI MEKANIK") = vbNo Then
75                                    GoTo ulangi
76                                Else
                                      'Call batalMekanik("")
77                                    Unload frmOrder
78                                    Exit Sub
79                                End If
80                            Else
81                                If Not cekMekanik(nTemp, CDbl(DGrid.TextMatrix(i, 3)), Replace(Text(14), "-", "/"), Trim(Text(15))) Then
82                                    pesan = gBantuan3
83                                    GoTo ulangi
84                                Else
85                                    myChassis = gNoRangka
86                                    frmOrder.Text(50) = Trim(nTemp)
87                                    .TextMatrix(i, 4) = Trim(DGrid.TextMatrix(i, 1))
88                                    .TextMatrix(i, 1) = Trim(DGrid.TextMatrix(i, 2))
89                                    .TextMatrix(i, 2) = Trim(DGrid.TextMatrix(i, 3))
90                                    .TextMatrix(i, 3) = Trim(DGrid.TextMatrix(i, 4))
91                                    .TextMatrix(i, 5) = Val(hitungLabour(CDbl(DGrid.TextMatrix(i, 3)), Trim(Text(11))))     'frmOrder.\
92                                    totalLabor = totalLabor + Val(hitungLabour(CDbl(DGrid.TextMatrix(i, 3)), Trim(Text(11))))
93                                    .TextMatrix(i, 6) = Trim(nTemp)
94                                    If isJualNRM = True Then
95                                        Me.JualNRM.Caption = nJualNRM
96                                    Else
97                                        Me.JualNRM.Caption = ""
98                                    End If
99                                End If
100                           End If
101                       End If
102                   Next
103               End With
                  
                  ' ISI SUBLET
104               With subletGrid
105                   For i = 1 To barisSublet
106                       If .TextMatrix(i, 1) <> "" Then
107                           Call objO.InsertData("INSERT INTO TEMPJOBSUBLET (NO_CHASSIS, NAMABARANG, KODESATUAN, QTY, FLAG, NO_POLISI) VALUES  " _
                                           & "('" & Trim(gNoRangka) & "', '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 3) & "', '" & CDbl(.TextMatrix(i, 2)) & "', '" & .TextMatrix(i, 5) & "', '" & Replace(Trim(Text(0)), ">", "") & "')")
108                       End If
109                   Next
110               End With
111               .Text(56) = totalLabor
112               .Text(61) = Format(CDbl(.Text(56)) + CDbl(.Text(57)) + CDbl(.Text(58)) + CDbl(.Text(59)) + CDbl(.Text(60)) + CDbl(.Text(77)), "#0,#")
113               Call HeadColumn(frmOrder)
114               .SSTab1.Enabled = True
                  '.cmdPilih.SetFocus
115               .Text(1).SetFocus
116               .Text(149) = Text(27)
117               .show
118           End With
119           Else
120               MsgBox "Tanggal tidak sesuai dengan Tanggal hari ini !" & vbCr & "Ubah Tanggal"
121               Text(14).SetFocus
122           End If
123       Else
124           MsgBox "LENGKAPI DATA", vbCritical, "KELENGKAPAN"
125           Exit Sub
126       End If

127      On Error GoTo 0
128      Exit Sub

tblToOrder_Click_Error:

129       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblToOrder_Click of Form frmEstimasi")
End Sub


Private Sub Text_Change(Index As Integer)
    If Index = 14 Or Index = 21 Then
        Select Case Index
            Case 14
                If CekTgl(Text(Index)) Then
                    If Text(15) <> ">>:>>" And Text(14) <> ">>->>->>>>" Then
                        nWKT = hitungWaktu(CDbl(nJRate), Text(15), Text(14))
                        Text(16) = Left(nWKT, InStr(nWKT, "|") - 1)
                        Text(30) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
                    End If
                End If
            Case 21
                If Text(Index) = "" Then
                    Text(Index + 1) = 0
                End If
        End Select
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                If cmdPilih.Caption = "No &Polisi" Then
                    gMainKeyCari = Replace(Trim(Text(Index)), "-", "")
                    gSearchOption = "ESTPOLISI"
                    frmCari.show (vbModal)
                ElseIf cmdPilih.Caption = "No &Rangka" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "ESTRANGKA"
                    frmCari.show (vbModal)
                End If
            Case 10
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 13
                If Trim(Text(10)) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "WARNA"
                    gBantuan = Trim(Text(10))
                    frmCariSPK.show (vbModal)
                Else
                    MsgBox "Isi Model Kendaraan ! "
                    Text(10).SetFocus
                    Exit Sub
                End If
            Case 18
                If Text(10) = "" Then
                    MsgBox "Isi model kendaraan !!!"
                    Text(10).SelLength = 10
                    Text(10).SetFocus
                    Exit Sub
                Else
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.show (vbModal)
                End If
            Case 19
                If Text(18) <> "" Then
                    If Text(15) <> ">>:>>" Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "JOBLIST2"
                        gBantuan = Text(10)
                        gBantuan2 = Text(18)
                        frmCari.show (vbModal)
                    Else
                        MsgBox "Isi Tanggal & Jam Service"
                        Text(14).SetFocus
                        Exit Sub
                    End If
                Else
                    MsgBox "Isi Tipe Job Order !!!"
                    Text(18).SelLength = 3
                    Text(18).SetFocus
                    Exit Sub
                End If
            Case 21
                If Text(10) = "" Then
                    MsgBox "Isi model kendaraan !!!"
                    Text(10).SelLength = 10
                    Text(10).SetFocus
                    Exit Sub
                Else
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PART"
                    frmCariAcess.show (vbModal)
                End If
            Case 34
                gMainKeyCari = Text(Index)
                gSearchOption = "SUBLET"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    Select Case Index
        Case 25, 37, 6, 7, 8, 9
            KeyAscii = CekNumerik(KeyAscii)
            simpan = False
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset
      Dim n As String
1         If Frame2.Visible = True Then Exit Sub
2         Select Case Index
              Case 0
3                 If cmdPilih.Caption = "No &Polisi" Then
4                     n = Replace(Text(Index), ">", "")
                      
                      ' ==================================================================================
                      ' 03 fEB 2017 | THANIA | revisi jika inputan no polisi hanya 2 blok depan / belakang atau hanya angka tanpa "-"
                      ' ==================================================================================
5                     If Right(Trim(n), 1) = "-" Then
6                         n = Left(n, Len(n) - 1)
7                     End If
8                     If Left(n, 1) = "-" Then
9                         n = Right(n, Len(n) - 1)
10                    End If
                      ' ==================================================================================
11                    If n = "--" Then Exit Sub
12                    Set rsSearch = objO.FetchData("SELECT * FROM ESTIMASI WITH(NOLOCK) WHERE NO_POLISI = '" & Replace(Text(Index), ">", "") & "'")
13                    Call hapusForm
14                    gNoRangka = ""
15                    If Not rsSearch.BOF Then
16                        If MsgBox("Input Estimasi baru ?", vbYesNo, "Estimasi") = vbYes Then
17                            Call PanggilCustomer(n, 1)
18                        Else
19                            Call panggilDataEstimasi(n, 1)
20                        End If
21                    Else
22                        Call PanggilCustomer(n, 1)
23                    End If
24                ElseIf cmdPilih.Caption = "No &Rangka" Then
25                    gNoRangka = Replace(Text(Index), ">", "")
26                    Set rsSearch = objO.FetchData("SELECT * FROM ESTIMASI WITH(NOLOCK) WHERE NO_RANGKA = '" & Replace(Text(Index), ">", "") & "'")
27                    Call hapusForm
28                    If Not rsSearch.BOF Then
29                        If MsgBox("Input Estimasi baru ?", vbYesNo, "Estimasi") = vbYes Then
30                            Call PanggilCustomer(gNoRangka, 2)
31                        Else
32                            Call panggilDataEstimasi(gNoRangka, 2)
33                        End If
34                    Else
35                        Call PanggilCustomer(gNoRangka, 2)
36                    End If
37                End If
38                If NulSID(Text(10)) <> "" Then
39                     If Log.RPLHistori = 0 Then
40                        Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
41                    Else
42                        Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
43                    End If
44                    If rsSearch.BOF Then
45                        MsgBox "Kelas kendaraan tidak sesuai master Model. Perbaiki data Model kendaraan di master Customer", vbInformation
lanjutKeMasterCustomer:
46                        With frmSIDCustomer
47                            If NulSID(gNoRangka) <> "" Then
48                                If Left(gNoRangka, 3) <> "" Then
49                                .Text(50) = Left(gNoRangka, 3)
50                                End If
51                                .Text(1) = Mid(gNoRangka, 4, 6)
52                                .Text(0) = Mid(gNoRangka, 10, IIf(Len(gNoRangka) > 0, Len(gNoRangka) - 9, 0))
53                            End If
54                            frmSIDCustomer.show vbModal
55                        End With
56                    Else
57                        If NulSID(rsSearch!model) <> NulSID(Text(11)) Then
58                            MsgBox "Kelas kendaraan tidak sesuai master Model. Perbaiki data Model kendaraan di master Customer", vbInformation
59                            GoTo lanjutKeMasterCustomer
60                        End If
61                    End If
62                End If
63            Case 9
64                If Text(Index) = "" Then
65                    Text(Index) = 0
66                    Text(Index).SetFocus
67                    Exit Sub
68                ElseIf CDbl(Text(Index)) < year(isnow) - 50 Or CDbl(Text(Index)) > (year(isnow) + 1) Then
69                    MsgBox "Tahun kendaraan salah. " & vbCr & " Isi kembali, dalam kisaran sekarang sampai 50 th kebelakang"
70                    Text(Index) = 0
71                    Text(Index).SelLength = 1
72                    Text(Index).SetFocus
73                    Exit Sub
74                End If
75            Case 10
76                If Text(Index) <> "" Then
                      
77                    If Log.RPLHistori = 0 Then
78                        Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(Index) & "'")
79                    Else
80                        Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(Index) & "'")
81                    End If
82                    If rsSearch.BOF Then
83                        Cancel = True
84                        gMainKeyCari = Text(Index)
85                        gSearchOption = "MDL"
86                        frmCari.show (vbModal)
87                        gGroupModel = ""
88                    Else
89                        Text(10) = rsSearch("MDL")
90                        Text(12) = rsSearch("Desc")
91                        Text(11) = rsSearch("MODEL")
92                        Call CexChassis_Discon(Text(11))
93                        If isOnCarClass = False Then
94                            MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
95                        End If
96                        gGroupModel = rsSearch("FS_GRP_MDL")
97                    End If
98                End If
99            Case 14
100               If CekTgl(Text(Index)) Then
101                   If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) < 0 Then
102                       Text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
103                       cmdWaktu.Visible = True
104                   End If
105               ElseIf Not CekTgl(Text(Index)) Then
106                   Text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
107                   cmdWaktu.Visible = True
108               End If
109               If Text(15) <> ">>:>>" And Text(14) <> ">>->>->>>>" Then
110                   nWKT = hitungWaktu(CDbl(nJRate), Text(15), Text(14))
111                   Text(16) = Left(nWKT, InStr(nWKT, "|") - 1)
112                   Text(30) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
113               End If
114           Case 15
115               If Not CekJam(Text(Index)) Then
116                   Text(Index) = Format(isnow, "HH:NN")
117               Else
118                   If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(23), "DD/MM/YYYY")) = 0 Then
119                       If DateDiff("s", Format(isnow, "HH:NN"), Format(Text(Index), "HH:NN")) < 0 Then
120                           Text(Index) = Format(isnow, "HH:NN")
121                       End If
122                   End If
123               End If
124               If Text(15) <> ">>:>>" And Text(14) <> ">>->>->>>>" Then
125                   nWKT = hitungWaktu(CDbl(nJRate), Text(15), Text(14))
126                   Text(16) = Left(nWKT, InStr(nWKT, "|") - 1)
127                   Text(30) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
128               End If
129           Case 18
130               If Text(Index) <> "" Then
131                   Set rsSearch = objO.FetchData("Select DISTINCT JOB, REMARK from HO_FLATE WITH(NOLOCK) where job = '" & Text(Index) & "' AND FS_LOG_DEL = ''")
132                   If rsSearch.BOF Then
133                       Cancel = True
134                       gMainKeyCari = Text(Index)
135                       gSearchOption = "JENIS_KERJA"
136                       frmCari.show (vbModal)
137                   End If
138               End If
139           Case 19
140               If Text(19) <> "" And Text(18) <> "" Then
141                   Set rsSearch = objO.FetchData("select DISTINCT FS_ITEM, RATE, KETERANGAN from HO_FLATE WITH(NOLOCK) where job = '" & Text(18) & "' and fs_item = '" & Text(Index) & "' AND mdl = '" & Text(10) & "' and FS_LOG_DEL = ''")
142                   If rsSearch.BOF Then
143                       Cancel = True
144                       gMainKeyCari = Text(Index)
145                       gSearchOption = "JOBLIST2"
146                       gBantuan = Text(10)
147                       gBantuan2 = Text(18)
148                       frmCari.show (vbModal)
149                       If fOther Then
150                           Text(20).SelLength = Len(Text(20))
151                       End If
152                   Else
153                       If Trim(rsSearch("FS_ITEM")) = "OTHER" Then
154                           fOther = True
155                           Text(19) = ""
156                           Text(19).Mask = ">CCCCCCCCCCCCCCCCCCCCCCCCC"
157                           Text(20).Visible = True
158                           Text(20).Enabled = True
159                           Text(19).SelLength = Len(Text(19))
160                           Text(19).SetFocus
161                       Else
162                           fOther = False
163                           Text(20).Visible = False
164                           Jrate = CDbl(rsSearch("Rate"))
165                       End If
166                   End If
167               End If
168           Case 21
169               If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
170                   If Text(Index) <> "" Then
                      'PART
171                       Set rsSearch = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH(NOLOCK) LEFT JOIN intrn B WITH(NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
172                       If Not rsSearch.EOF Then
173                           stkPart = CDbl(rsSearch("QTYSTOCK"))
174                           namaPart = rsSearch("NAMABARANG")
175                           HrgPart = Val(rsSearch("HARGAJUALA"))
176                           satPart = rsSearch("KODESATUAN")
177                           Set rsTemp = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI, TIPE, GENUINE FROM T_OLI WITH(NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "'")
178                           If Not rsTemp.EOF Then
179                               oliPart = IIf(rsTemp("TIPE") = "O", True, False)
180                               genuineNon = IIf(rsTemp("GENUINE") = 1, True, False)
181                           Else
182                               oliPart = False
183                               genuineNon = True
184                           End If
185                       Else
186                           gMainKeyCari = Text(Index)
187                           Text(Index) = ""
188                           gSearchOption = "PART"
189                           frmCariAcess.show (vbModal)
190                       End If
191                   End If
192               Else
193                   If Text(Index) <> "" Then
                      'PART
194                       Set rsSearch = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH(NOLOCK) LEFT JOIN intrn B WITH(NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
195                       If Not rsSearch.EOF Then
196                           stkPart = CDbl(rsSearch("QTYSTOCK"))
197                           namaPart = rsSearch("NAMABARANG")
198                           HrgPart = Val(rsSearch("HARGAJUALA"))
199                           satPart = rsSearch("KODESATUAN")
200                           Set rsTemp = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI, TIPE, GENUINE FROM T_OLI WITH(NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "'")
201                           If Not rsTemp.EOF Then
202                               oliPart = IIf(rsTemp("TIPE") = "O", True, False)
203                               genuineNon = IIf(rsTemp("GENUINE") = 1, True, False)
204                           Else
205                               oliPart = False
206                               genuineNon = True
207                           End If
208                       Else
209                           gMainKeyCari = Text(Index)
210                           Text(Index) = ""
211                           gSearchOption = "PART"
212                           frmCariAcess.show (vbModal)
213                           Set rsTemp = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI, TIPE, GENUINE FROM T_OLI WITH(NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "'")
214                           If Not rsTemp.EOF Then
215                               oliPart = IIf(rsTemp("TIPE") = "O", True, False)
216                               genuineNon = IIf(rsTemp("GENUINE") = 1, True, False)
217                           Else
218                               oliPart = False
219                               genuineNon = True
220                           End If
221                            stkPart = gStok
222                           namaPart = gNamaPart
223                           satPart = gSatPart
224                           HrgPart = gHargaJual
225                       End If
226                   End If
227               End If
                  
228           Case 22
                  
229           Case 25
230               If CDbl(Text(Index)) = 0 Then
231                   Text(Index).SetFocus
232               End If
233       End Select
           
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Command1_Click()
1        On Error GoTo Command1_Click_Error

2         Frame2.Visible = True
3         Frame2.Left = 480
4         Frame2.Top = 0
5         Call LoadFirstRujukan

6        On Error GoTo 0
7        Exit Sub

Command1_Click_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Command1_Click of Form frmEstimasi")
End Sub

Private Sub LoadFirstRujukan()
1        On Error GoTo LoadFirstRujukan_Error

2          For i = 1 To 15
3             txt(i).BackColor = fColor
4             txt(i).Text = ""
5          Next i
6          txt(0).Text = ""
7          txtReq.Text = ""
8          txtKet.Text = ""
9          txtKet.BackColor = fColor
10         txtReq.BackColor = fColor
11         With Me
12            .cmdAdd.Enabled = True
13            .cmdCreate.Visible = False
14            .cmdEdit.Enabled = False
15            .cmdSave.Visible = True
16            .cmdSave.Enabled = False
17            .cmdCancel.Visible = False
18            .cmdDelete.Visible = False
19            .cmdDelete.Enabled = False
20         End With
21         cmdPrint.Enabled = True
22         cmdPrint.Visible = False
23         cmdExit.Enabled = True

24       On Error GoTo 0
25       Exit Sub

LoadFirstRujukan_Error:

26        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure LoadFirstRujukan of Form frmEstimasi")
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If txt(Index).BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Dim rsFind As ADODB.Recordset
        Dim ft As New zSearch
        Select Case Index
            Case 0
                ShowForm = "Rujukan"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.rs
                    nRujuk = rsFind.Fields(0).value
                    txt(0).Text = rsFind.Fields(0).value
                    txt(1).Text = rsFind.Fields(1).value
                    txt(2).Text = rsFind.Fields(2).value
                    txt(3).Text = rsFind.Fields(3).value
                    txt(4).Text = rsFind.Fields(4).value
                    txt(5).Text = rsFind.Fields(5).value
                    txt(6).Text = rsFind.Fields(6).value
                    txt(7).Text = rsFind.Fields(7).value
                    txt(8).Text = rsFind.Fields(8).value
                    txt(9).Text = rsFind.Fields(9).value
                    txt(10).Text = rsFind.Fields(10).value
                    txt(11).Text = rsFind.Fields(11).value
                    txt(12).Text = rsFind.Fields(12).value
                    txt(13).Text = rsFind.Fields(13).value
                    txt(14).Text = rsFind.Fields(14).value
                    txt(15).Text = rsFind.Fields(15).value

                    txtReq.Text = rsFind.Fields(16).value
                    txtKet.Text = rsFind.Fields(17).value
                    With Me
                        .cmdAdd.Enabled = True
                        .cmdCreate.Visible = False
                        .cmdEdit.Enabled = True
                        .cmdSave.Visible = True
                        .cmdSave.Enabled = False
                        .cmdCancel.Visible = False
                        .cmdDelete.Visible = True
                        .cmdDelete.Enabled = True
                        .cmdPrint.Visible = True
                    End With
                End If
            Case 10         '" & DB_bengkel & "..Model
                ShowForm = "Model"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.rs
                    txt(10).Text = rsFind.Fields(0).value
                    txt(11).Text = rsFind.Fields(1).value
                    txt(12).Text = rsFind.Fields(2).value
                End If
            Case 1         'Customer
                ShowForm = "Custdat"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.rs
                    txt(1).Text = rsFind.Fields(0).value
                    txt(2).Text = rsFind.Fields(1).value
                    txt(3).Text = rsFind.Fields(2).value
                    txt(4).Text = ""
                    txt(5).Text = rsFind.Fields(4).value
                    txt(6).Text = ""
                    txt(7).Text = rsFind.Fields(3).value
                    txt(8).Text = rsFind.Fields(5).value
                    txt(9).Text = rsFind.Fields(36).value
                    txt(10).Text = rsFind.Fields(9).value
                    txt(11).Text = rsFind.Fields(10).value
                    txt(12).Text = rsFind.Fields(11).value
                    txt(13).Text = rsFind.Fields(12).value
                    txt(14).Text = rsFind.Fields(8).value
                    txt(15).Text = rsFind.Fields(6).value
                End If
        End Select
    End If
End Sub

Private Sub CmdAdd_Click()
     For i = 1 To 15
        txt(i).BackColor = tColor
        txt(i).Text = ""
     Next i
     txt(11).BackColor = fColor
     txt(12).BackColor = fColor
     txt(0).BackColor = fColor
     txt(0).Text = ""
     txtReq.Text = ""
     txtKet.Text = ""
     txtKet.BackColor = tColor
     txtReq.BackColor = tColor
     With Me
        .cmdAdd.Enabled = False
        .cmdCreate.Visible = True
        .cmdCreate.Enabled = True
        .cmdEdit.Enabled = False
        .cmdSave.Visible = False
        .cmdCancel.Visible = True
        .cmdCancel.Enabled = True
        .cmdDelete.Visible = False
     End With
     cmdPrint.Visible = False
     cmdExit.Enabled = False
End Sub
Private Sub CmdDelete_Click()
    If MsgBox("Batalkan Rujukan Nomor : " & nRujuk & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call objO.UpdateData("Update Rujukan Set UserDel = '" & gSA & "', TglDel = Getdate() Where NoRujuk = '" & nRujuk & "'")
        Call LoadFirstRujukan
    End If
End Sub
Private Sub cmdCancel_Click()
    If txt(0).Text = "" Then
        If MsgBox("Batalkan Penginputan Data Rujukan ?", vbQuestion + vbYesNo) = vbYes Then
            Call LoadFirstRujukan
            txt(0).BackColor = tColor
        End If
    Else
        If MsgBox("Batalkan Update Data Rujukan No : " & nRujuk & " ?", vbQuestion + vbYesNo) = vbYes Then
            Call LoadFirstRujukan
            txt(0).BackColor = tColor
        End If
    End If
End Sub
Private Sub cmdSave_Click()
    If MsgBox("Update Data Rujukan Nomor : " & txt(0).Text & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call objO.InsertData("Update Rujukan Set Nama='" & txt(1).Text & "',Alamat='" & txt(2).Text & "',Kota = '" & txt(3).Text & "' " & _
                             ",TelpP='" & txt(4).Text & "',TelpG='" & txt(5).Text & "',TelpC='" & txt(6).Text & "',Pos='" & txt(7).Text & "',NoPolisi='" & txt(8).Text & "' " & _
                             ",Tahun='" & txt(9).Text & "',model='" & txt(10).Text & "',tipe='" & txt(11).Text & "',namaModel='" & txt(12).Text & "',Warna='" & txt(13).Text & "' " & _
                             ",noRangka='" & txt(14).Text & "',NoMesin='" & txt(15).Text & "',Permintaan='" & txtReq.Text & "',Catatan='" & txtKet.Text & "' " & _
                             ",UserLog='" & gSA & "',TglLog=getdate() Where NoRujuk = '" & nRujuk & "'")
        For i = 1 To 15
           txt(i).BackColor = fColor
        Next i
        txt(0).BackColor = tColor
        txt(0).Text = nRujuk
        txtKet.BackColor = fColor
        txtReq.BackColor = fColor
        With Me
            .cmdAdd.Enabled = True
            .cmdCreate.Visible = False
            .cmdEdit.Enabled = True
            .cmdSave.Visible = True
            .cmdSave.Enabled = False
            .cmdCancel.Visible = False
            .cmdDelete.Visible = True
            .cmdDelete.Enabled = True
        End With
        cmdPrint.Visible = True
        cmdPrint.Enabled = True
        cmdExit.Enabled = True
    End If
End Sub

Private Sub cmdCreate_Click()
    If MsgBox("Simpan Data Rujukan ?", vbQuestion + vbYesNo) = vbYes Then
        nRujuk = Format(isnow, "YYMMDDHHNNSS")
        Call objO.InsertData("Insert into Rujukan Values ('" & nRujuk & "','" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "','" & txt(4).Text & "','" & txt(5).Text & "','" & txt(6).Text & "','" & txt(7).Text & "','" & txt(8).Text & "','" & txt(9).Text & "','" & txt(10).Text & "','" & txt(11).Text & "','" & txt(12).Text & "','" & txt(13).Text & "','" & txt(14).Text & "','" & txt(15).Text & "','" & txtReq.Text & "','" & txtKet.Text & "','" & gSA & "',getdate(),'','01/01/1900')")
        For i = 1 To 15
           txt(i).BackColor = fColor
        Next i
        txt(0).BackColor = tColor
        txt(0).Text = nRujuk
        txtKet.BackColor = fColor
        txtReq.BackColor = fColor
        With Me
            .cmdAdd.Enabled = True
            .cmdCreate.Visible = False
            .cmdEdit.Enabled = True
            .cmdSave.Visible = True
            .cmdSave.Enabled = False
            .cmdCancel.Visible = False
            .cmdDelete.Visible = True
            .cmdDelete.Enabled = True
        End With
        cmdPrint.Visible = True
        cmdExit.Enabled = True
    End If
End Sub
Private Sub cmdEdit_Click()
     For i = 1 To 15
        txt(i).BackColor = tColor
     Next i
     txt(11).BackColor = fColor
     txt(12).BackColor = fColor
     txt(0).BackColor = fColor
     txtKet.BackColor = tColor
     txtReq.BackColor = tColor
    With Me
        .cmdAdd.Enabled = False
        .cmdCreate.Visible = False
        .cmdEdit.Enabled = False
        .cmdSave.Visible = True
        .cmdSave.Enabled = True
        .cmdCancel.Visible = True
        .cmdCancel.Enabled = True
        .cmdDelete.Visible = False
        .cmdExit.Enabled = False
        .cmdPrint.Enabled = False
    End With
End Sub

Private Sub cmdExit_Click()
    Frame2.Visible = False
End Sub

Private Sub txtReq_GotFocus()
    If txtReq.BackColor = fColor Then Sendkeys ("{TAB}")
End Sub
Private Sub txtKet_GotFocus()
    If txtReq.BackColor = fColor Then Sendkeys ("{TAB}")
End Sub




Private Sub cekSubletFOP(kodeModel As String, pekerjaan As String)
          Dim rsSublet As ADODB.Recordset
1         sSQL = "select distinct c.KDSUBLET,c.NAMASUBLET,a.QTY,c.HARGAJUAL,c.kodesatuan"
2         sSQL = sSQL & " from [" & dbReplikasiHO & "]..HO_FOP_Program a"
3         sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..HO_SUBLETPACK b on a.ITEM_PART = b.KDSUBLET and b.FS_LOG_DEL = ''"
4         sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..MSTSUBLET c  on c.KDSUBLET = a.ITEM_PART and c.FS_LOG_DEL = ''"
5         sSQL = sSQL & " where a.FS_LOG_DEL = '' and a.MDL = '" & kodeModel & "' and a.FS_ITEM = '" & pekerjaan & "'"
6         Set rsSublet = objO.FetchData(sSQL)
7         If rsSublet.RecordCount > 0 And tagihanFOP <> "" Then
8             GoTo loopSubletTerkait
9         Else
10            sSQL = "select c.KDSUBLET,c.NAMASUBLET,b.QTY,c.HARGAJUAL,c.kodesatuan"
11            sSQL = sSQL & " from [" & dbReplikasiHO & "]..HO_SUBLETPACK b "
12            sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..MSTSUBLET c  on c.KDSUBLET = b.KDSUBLET and c.FS_LOG_DEL = ''"
13            sSQL = sSQL & " where b.FS_LOG_DEL = '' and b.MDL = '" & kodeModel & "' and b.FS_ITEM = '" & pekerjaan & "'"
14            Set rsSublet = objO.FetchData(sSQL)
15            If rsSublet.RecordCount > 0 Then
16                GoTo loopSubletTerkait
17            Else
18                Exit Sub
19            End If
20        End If
loopSubletTerkait:
21        rsSublet.MoveFirst
          Dim isBoleh     As Boolean
22        Do While Not rsSublet.EOF

23            If MaxBrs >= barisSublet Then
24                isBoleh = True
25                For i = 0 To barisSublet
26                    If NulSID(subletGrid.TextMatrix(i, 1)) = NulSID(rsSublet!namaSublet) Then
27                        isBoleh = False
28                        GoTo lanjutBoleh
                          ''isBoleh = True
29                    End If
30                Next
lanjutBoleh:
31                If isBoleh Then
32                    With subletGrid
33                        .TextMatrix(barisSublet, 1) = NulSID(rsSublet!namaSublet)                   ' NAMA SUBLET
34                        .TextMatrix(barisSublet, 2) = CDbl(NulSID(rsSublet!qty))                    ' QTY
35                        .TextMatrix(barisSublet, 3) = NulSID(rsSublet!KODESATUAN)                                            ' SATUAN
36                        .TextMatrix(barisSublet, 4) = Format(CDbl(NulSID(rsSublet!hargaJual)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
37                        .TextMatrix(barisSublet, 5) = 1                ' STATUS PART
38                    End With
39                    barisSublet = barisSublet + 1
40                    Text(25) = Format(CDbl(Text(25)) + (CDbl(NulSID(rsSublet!qty) * NulSID(rsSublet!hargaJual))), "#0,#") '* ((100 + gUpSublet) / 100)
41                    Text(26) = Format(CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(28)) + CDbl(Text(31)) + CDbl(Text(32)), "#0,#")
42                End If
43            End If
44            rsSublet.MoveNext
45        Loop
          
End Sub

Private Function hitungPartTanpaDisc(part As Boolean) As Double
    Dim nilai As Double
    Dim oli As Boolean
     
    nilai = 0
    For i = 1 To barispart - 1
        With partGrid
        
            If .TextMatrix(barispart, 6) = "" Then
                oli = False
            Else
                oli = .TextMatrix(barispart, 6)
            End If
            If part = True And oli = False Then
                nilai = nilai + hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5))
            ElseIf part = False And oli = True Then
                nilai = nilai + hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5))
            End If
        End With
    Next
    hitungPartTanpaDisc = nilai
End Function
