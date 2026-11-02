VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBtlBooking 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pembatalan Booking"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15795
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   15795
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "History Service"
      Height          =   1935
      Left            =   9120
      TabIndex        =   156
      Top             =   1320
      Width           =   3735
      Begin VB.CommandButton cmdHistory 
         BackColor       =   &H00DCFCFB&
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
         Left            =   -120
         Picture         =   "frmBtlBooking.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DCFCFB&
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
         Left            =   1740
         Picture         =   "frmBtlBooking.frx":07BD
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   1320
         Width           =   1935
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   1080
         TabIndex        =   157
         Top             =   980
         Width           =   975
         _ExtentX        =   1720
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
         Index           =   23
         Left            =   1080
         TabIndex        =   158
         Top             =   705
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   22
         Left            =   1080
         TabIndex        =   159
         Top             =   435
         Width           =   1245
         _ExtentX        =   2196
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   240
         TabIndex        =   167
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM Akhir"
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
         TabIndex        =   166
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label16 
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
         Left            =   240
         TabIndex        =   165
         Top             =   480
         Width           =   570
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   3000
         X2              =   3480
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   3000
         X2              =   3480
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   600
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   600
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Service"
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
         Left            =   840
         TabIndex        =   164
         Top             =   45
         Width           =   1620
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Line HO"
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
         Height          =   240
         Left            =   420
         TabIndex        =   163
         Top             =   1695
         Width           =   1050
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On LOCAL"
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
         Height          =   240
         Left            =   2340
         TabIndex        =   162
         Top             =   1695
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   0
      Picture         =   "frmBtlBooking.frx":0F7A
      ScaleHeight     =   10695
      ScaleWidth      =   495
      TabIndex        =   152
      Top             =   0
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6100
      LargeChange     =   350
      Left            =   12480
      Max             =   4500
      SmallChange     =   350
      TabIndex        =   88
      Top             =   3380
      Width           =   495
   End
   Begin VB.PictureBox frmBtlBook 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   3360
      ScaleHeight     =   2115
      ScaleWidth      =   5640
      TabIndex        =   3
      Top             =   3000
      Width           =   5700
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   5415
      End
      Begin ucXPButton.XPButton cmdOK 
         Height          =   495
         Left            =   1560
         TabIndex        =   154
         Top             =   1440
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&O"
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
         MICON           =   "frmBtlBooking.frx":2EDC
         PICN            =   "frmBtlBooking.frx":2EF8
         PICH            =   "frmBtlBooking.frx":3399
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
         Left            =   2760
         TabIndex        =   155
         Top             =   1440
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
         MICON           =   "frmBtlBooking.frx":38B3
         PICN            =   "frmBtlBooking.frx":38CF
         PICH            =   "frmBtlBooking.frx":3E13
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   5520
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   5520
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: Alasan Pembatalan  ::."
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
         Left            =   1440
         TabIndex        =   153
         Top             =   165
         Width           =   2745
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   480
      ScaleHeight     =   6075
      ScaleWidth      =   12195
      TabIndex        =   4
      Top             =   3360
      Width           =   12255
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Input Booking"
         Height          =   10335
         Left            =   0
         TabIndex        =   5
         Top             =   -4200
         Width           =   11955
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
            Height          =   375
            Left            =   9720
            Style           =   1  'Graphical
            TabIndex        =   134
            Top             =   9840
            Width           =   2055
         End
         Begin VB.CommandButton TblBatal 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Batal"
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
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   9840
            Width           =   1455
         End
         Begin VB.CommandButton TblSimpan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Simpan"
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
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   11160
            Width           =   1890
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   5625
            Begin VB.CheckBox chkRtn 
               BackColor       =   &H00FFFFFF&
               Caption         =   "RTJ"
               Enabled         =   0   'False
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
               TabIndex        =   140
               Top             =   975
               Width           =   735
            End
            Begin VB.CheckBox chkThs 
               BackColor       =   &H00FFFFFF&
               Caption         =   "THS"
               Enabled         =   0   'False
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
               Left            =   840
               TabIndex        =   139
               Top             =   975
               Width           =   735
            End
            Begin VB.CheckBox chkAmbil 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ambil"
               Enabled         =   0   'False
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
               Left            =   3840
               TabIndex        =   138
               Top             =   975
               Width           =   735
            End
            Begin VB.CheckBox chEM 
               BackColor       =   &H00FFFFFF&
               Caption         =   "EM"
               Enabled         =   0   'False
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
               Left            =   1800
               TabIndex        =   137
               Top             =   975
               Width           =   735
            End
            Begin VB.CheckBox chSBNP 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SB NP"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2520
               TabIndex        =   136
               Top             =   975
               Width           =   855
            End
            Begin VB.CheckBox chkKirim 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Kirim"
               Enabled         =   0   'False
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
               Left            =   4680
               TabIndex        =   135
               Top             =   975
               Width           =   735
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   25
               Left            =   1260
               TabIndex        =   79
               Top             =   0
               Width           =   1200
               _ExtentX        =   2117
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
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   26
               Left            =   1260
               TabIndex        =   80
               Top             =   315
               Width           =   675
               _ExtentX        =   1191
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
               Index           =   27
               Left            =   1260
               TabIndex        =   81
               Top             =   630
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
               Mask            =   ">CC"
               PromptChar      =   ">"
            End
            Begin VB.Shape Shape3 
               Height          =   315
               Left            =   0
               Top             =   960
               Width           =   5535
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Booking"
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
               Left            =   60
               TabIndex        =   87
               Top             =   45
               Width           =   1170
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jam Booking"
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
               Left            =   345
               TabIndex        =   86
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Via"
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
               Left            =   1005
               TabIndex        =   85
               Top             =   675
               Width           =   210
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "D [ Datang ] / T [Telepon] / S [Sms]"
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
               Left            =   1710
               TabIndex        =   84
               Top             =   675
               Width           =   2550
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "(DD-MM-YYYY)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2535
               TabIndex        =   83
               Top             =   15
               Width           =   1140
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "(HH:MM)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2025
               TabIndex        =   82
               Top             =   375
               Width           =   1140
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4560
            Left            =   120
            TabIndex        =   54
            Top             =   1560
            Width           =   5640
            Begin VB.TextBox Text2 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   59
               Top             =   1440
               Width           =   4455
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   58
               Top             =   1680
               Width           =   4455
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   57
               Top             =   1920
               Width           =   4455
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   56
               Top             =   2160
               Width           =   4455
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   55
               Top             =   2400
               Width           =   4455
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   28
               Left            =   1095
               TabIndex        =   60
               Top             =   75
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Height          =   240
               Index           =   29
               Left            =   1095
               TabIndex        =   61
               Top             =   300
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Height          =   240
               Index           =   30
               Left            =   1095
               TabIndex        =   62
               Top             =   525
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Height          =   240
               Index           =   31
               Left            =   1095
               TabIndex        =   63
               Top             =   750
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Height          =   240
               Index           =   32
               Left            =   1095
               TabIndex        =   64
               Top             =   975
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Height          =   315
               Index           =   60
               Left            =   120
               TabIndex        =   65
               Top             =   840
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   33
               Left            =   1095
               TabIndex        =   66
               Top             =   2880
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Index           =   34
               Left            =   1095
               TabIndex        =   67
               Top             =   3165
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Index           =   35
               Left            =   1095
               TabIndex        =   68
               Top             =   3480
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Index           =   36
               Left            =   1095
               TabIndex        =   69
               Top             =   3765
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Height          =   315
               Index           =   38
               Left            =   2280
               TabIndex        =   70
               Top             =   4200
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   39
               Left            =   4440
               TabIndex        =   71
               Top             =   4200
               Width           =   975
               _ExtentX        =   1720
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
               Index           =   37
               Left            =   1095
               TabIndex        =   72
               Top             =   4200
               Width           =   1200
               _ExtentX        =   2117
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
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Permintaan Customer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   77
               Top             =   120
               Width           =   945
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Keluhan Customer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   76
               Top             =   1440
               Width           =   945
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
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
               Left            =   360
               TabIndex        =   75
               Top             =   2925
               Width           =   585
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Estimasi Selesai  "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   74
               Top             =   4125
               Width           =   720
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
               Left            =   4125
               TabIndex        =   73
               Top             =   4245
               Width           =   195
            End
         End
         Begin VB.CommandButton cmdCetak 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cetak"
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
            Left            =   9600
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   11160
            Width           =   1890
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   5880
            Left            =   5760
            TabIndex        =   6
            Top             =   360
            Width           =   6015
            Begin VB.Frame Frame9 
               BackColor       =   &H00FFFFFF&
               Caption         =   "                                           Kondisi ketika terjadi Keluhan                                          "
               Height          =   2535
               Left            =   0
               TabIndex        =   17
               Top             =   120
               Width           =   6015
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   40
                  Left            =   885
                  TabIndex        =   18
                  Top             =   360
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin MSMask.MaskEdBox MaskEdBox4 
                  Height          =   270
                  Left            =   3285
                  TabIndex        =   19
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
                  Index           =   41
                  Left            =   885
                  TabIndex        =   20
                  Top             =   720
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin MSMask.MaskEdBox MaskEdBox7 
                  Height          =   270
                  Left            =   3285
                  TabIndex        =   21
                  Top             =   720
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
                  Index           =   42
                  Left            =   885
                  TabIndex        =   22
                  Top             =   1080
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin MSMask.MaskEdBox MaskEdBox8 
                  Height          =   270
                  Left            =   3285
                  TabIndex        =   23
                  Top             =   1080
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
                  Index           =   43
                  Left            =   885
                  TabIndex        =   24
                  Top             =   1440
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin MSMask.MaskEdBox MaskEdBox9 
                  Height          =   270
                  Left            =   3285
                  TabIndex        =   25
                  Top             =   1440
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
                  Index           =   44
                  Left            =   885
                  TabIndex        =   26
                  Top             =   1800
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin MSMask.MaskEdBox MaskEdBox10 
                  Height          =   270
                  Left            =   3285
                  TabIndex        =   27
                  Top             =   1800
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
                  Index           =   45
                  Left            =   885
                  TabIndex        =   28
                  Top             =   2160
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin MSMask.MaskEdBox MaskEdBox11 
                  Height          =   270
                  Left            =   3285
                  TabIndex        =   29
                  Top             =   2160
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
                  Index           =   46
                  Left            =   3480
                  TabIndex        =   30
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
                  Index           =   47
                  Left            =   3480
                  TabIndex        =   31
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
                  Index           =   48
                  Left            =   3480
                  TabIndex        =   32
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
                  Index           =   49
                  Left            =   3480
                  TabIndex        =   33
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
                  Index           =   50
                  Left            =   3480
                  TabIndex        =   34
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
                  Index           =   51
                  Left            =   3480
                  TabIndex        =   35
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin VB.Label Label44 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Waktu"
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
                  Left            =   330
                  TabIndex        =   40
                  Top             =   405
                  Width           =   465
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Frekuensi"
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
                  Left            =   105
                  TabIndex        =   39
                  Top             =   765
                  Width           =   690
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lokasi"
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
                  Left            =   360
                  TabIndex        =   38
                  Top             =   1125
                  Width           =   435
               End
               Begin VB.Label Label41 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lampu Peringatan"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   -15
                  TabIndex        =   37
                  Top             =   1485
                  Width           =   810
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lainnya"
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
                  TabIndex        =   36
                  Top             =   2040
                  Width           =   555
               End
            End
            Begin VB.ComboBox cboPelanggan 
               BackColor       =   &H00C0E0FF&
               Height          =   315
               ItemData        =   "frmBtlBooking.frx":440A
               Left            =   840
               List            =   "frmBtlBooking.frx":440C
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   3360
               Width           =   1215
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00FFFFFF&
               Height          =   735
               Left            =   240
               TabIndex        =   7
               Top             =   4080
               Width           =   2175
               Begin VB.OptionButton opSvc 
                  BackColor       =   &H00BFE8EA&
                  Caption         =   "Ditunggu"
                  Enabled         =   0   'False
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
                  Left            =   120
                  TabIndex        =   9
                  Top             =   420
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton opSvc 
                  BackColor       =   &H00BFE8EA&
                  Caption         =   "Ditinggal"
                  Enabled         =   0   'False
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
                  Left            =   1080
                  TabIndex        =   8
                  Top             =   420
                  Width           =   975
               End
               Begin VB.Label Label32 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "SERVICE"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   600
                  TabIndex        =   10
                  Top             =   120
                  Width           =   1065
               End
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   55
               Left            =   3495
               TabIndex        =   12
               Top             =   3900
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   56
               Left            =   3495
               TabIndex        =   13
               Top             =   4200
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   58
               Left            =   3495
               TabIndex        =   14
               Top             =   4860
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   59
               Left            =   3495
               TabIndex        =   15
               Top             =   5460
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   57
               Left            =   3495
               TabIndex        =   16
               Top             =   4500
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   270
               Index           =   54
               Left            =   2040
               TabIndex        =   41
               Top             =   3345
               Width           =   3885
               _ExtentX        =   6853
               _ExtentY        =   476
               _Version        =   393216
               BackColor       =   12640511
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
               Index           =   52
               Left            =   840
               TabIndex        =   42
               Top             =   2760
               Width           =   5100
               _ExtentX        =   8996
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
               Index           =   53
               Left            =   840
               TabIndex        =   43
               Top             =   3030
               Width           =   5100
               _ExtentX        =   8996
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
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
            Begin VB.Label Label49 
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
               Left            =   3075
               TabIndex        =   52
               Top             =   4560
               Width           =   255
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
               Left            =   2955
               TabIndex        =   51
               Top             =   3960
               Width           =   375
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "________________________  +"
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
               Left            =   3525
               TabIndex        =   50
               Top             =   5175
               Width           =   2370
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
               Left            =   2940
               TabIndex        =   49
               Top             =   4260
               Width           =   390
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
               Left            =   2790
               TabIndex        =   48
               Top             =   4920
               Width           =   540
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
               Left            =   2850
               TabIndex        =   47
               Top             =   5520
               Width           =   480
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Karakter"
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
               Left            =   135
               TabIndex        =   46
               Top             =   3405
               Width           =   615
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saran"
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
               Left            =   330
               TabIndex        =   45
               Top             =   3075
               Width           =   420
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pesan Khusus"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   44
               Top             =   2640
               Width           =   615
            End
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   3495
            Left            =   75
            TabIndex        =   141
            Top             =   6240
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6165
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   617
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Job Order && Sublet"
            TabPicture(0)   =   "frmBtlBooking.frx":440E
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "cmdStall"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame5"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Part Order"
            TabPicture(1)   =   "frmBtlBooking.frx":442A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame10"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   3100
               Left            =   30
               TabIndex        =   146
               Top             =   360
               Width           =   11655
               Begin MSFlexGridLib.MSFlexGrid DGrid 
                  Height          =   2625
                  Left            =   120
                  TabIndex        =   147
                  Top             =   360
                  Width           =   5505
                  _ExtentX        =   9710
                  _ExtentY        =   4630
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  ScrollBars      =   1
                  SelectionMode   =   1
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   68
                  Left            =   1920
                  TabIndex        =   148
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
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
               Begin MSFlexGridLib.MSFlexGrid subletGrid 
                  Height          =   2625
                  Left            =   5640
                  TabIndex        =   149
                  Top             =   360
                  Width           =   5985
                  _ExtentX        =   10557
                  _ExtentY        =   4630
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Job Order"
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
                  Left            =   120
                  TabIndex        =   151
                  Top             =   120
                  Width           =   720
               End
               Begin VB.Label Label56 
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
                  Left            =   5640
                  TabIndex        =   150
                  Top             =   120
                  Width           =   915
               End
            End
            Begin VB.CommandButton cmdStall 
               BackColor       =   &H00DCFCFB&
               Caption         =   "S&tall"
               Height          =   615
               Left            =   10680
               Style           =   1  'Graphical
               TabIndex        =   145
               Top             =   360
               Width           =   930
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00BFE8EA&
               BorderStyle     =   0  'None
               Height          =   3105
               Left            =   -74970
               TabIndex        =   142
               Top             =   360
               Width           =   11655
               Begin MSFlexGridLib.MSFlexGrid PartGrid 
                  Height          =   2745
                  Left            =   120
                  TabIndex        =   143
                  Top             =   240
                  Width           =   5865
                  _ExtentX        =   10345
                  _ExtentY        =   4842
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Order Part"
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
                  Left            =   120
                  TabIndex        =   144
                  Top             =   0
                  Width           =   765
               End
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Data Pelanggan"
      Height          =   2895
      Left            =   480
      TabIndex        =   89
      Top             =   360
      Width           =   8565
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   915
         TabIndex        =   90
         Top             =   705
         Width           =   3300
         _ExtentX        =   5821
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
         Index           =   4
         Left            =   915
         TabIndex        =   91
         Top             =   1005
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   40
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
         TabIndex        =   92
         Top             =   1305
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         TabIndex        =   93
         Top             =   1905
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   18
         Left            =   5430
         TabIndex        =   94
         Top             =   2490
         Width           =   1185
         _ExtentX        =   2090
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   19
         Left            =   7245
         TabIndex        =   95
         Top             =   2505
         Width           =   1245
         _ExtentX        =   2196
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   5430
         TabIndex        =   96
         Top             =   405
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   14
         Left            =   5430
         TabIndex        =   97
         Top             =   1305
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   13
         Left            =   5430
         TabIndex        =   98
         Top             =   1005
         Width           =   3060
         _ExtentX        =   5398
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
         Index           =   12
         Left            =   7695
         TabIndex        =   99
         Top             =   705
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   5430
         TabIndex        =   100
         Top             =   2205
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   1
         Left            =   915
         TabIndex        =   101
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
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
         Index           =   16
         Left            =   5430
         TabIndex        =   102
         Top             =   1905
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   915
         TabIndex        =   103
         Top             =   2205
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   6
         Left            =   915
         TabIndex        =   104
         Top             =   1605
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   15
         Left            =   5430
         TabIndex        =   105
         Top             =   1605
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   9
         Left            =   915
         TabIndex        =   106
         Top             =   2505
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   2
         Left            =   2475
         TabIndex        =   107
         Top             =   420
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
         Index           =   11
         Left            =   5430
         TabIndex        =   108
         Top             =   705
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   61
         Left            =   2865
         TabIndex        =   109
         Top             =   2490
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
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
         Left            =   2280
         TabIndex        =   132
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
         X1              =   6240
         X2              =   8400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   6240
         X2              =   8400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label10 
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
         TabIndex        =   124
         Top             =   720
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
         Left            =   390
         TabIndex        =   123
         Top             =   1020
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
         TabIndex        =   122
         Top             =   1320
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
         Left            =   270
         TabIndex        =   121
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOTAM"
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
         Left            =   4740
         TabIndex        =   120
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDTE."
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
         Left            =   6720
         TabIndex        =   119
         Top             =   2550
         Width           =   450
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
         Left            =   4635
         TabIndex        =   118
         Top             =   450
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
         Left            =   4860
         TabIndex        =   117
         Top             =   750
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
         Left            =   4800
         TabIndex        =   116
         Top             =   1350
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
         TabIndex        =   115
         Top             =   450
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
         Left            =   150
         TabIndex        =   114
         Top             =   2220
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
         TabIndex        =   113
         Top             =   1620
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   4695
         TabIndex        =   112
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rangka"
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
         Left            =   4440
         TabIndex        =   111
         Top             =   1650
         Width           =   840
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOU"
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
         Left            =   450
         TabIndex        =   110
         Top             =   2550
         Width           =   345
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   8535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   20
      Left            =   10320
      TabIndex        =   125
      Top             =   570
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   21
      Left            =   10320
      TabIndex        =   126
      Top             =   900
      Width           =   675
      _ExtentX        =   1191
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
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   53
      Width           =   1410
      _ExtentX        =   2487
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
      Mask            =   "CCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3000
      TabIndex        =   130
      Top             =   60
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmBtlBooking.frx":4446
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Booking"
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
      Left            =   600
      TabIndex        =   131
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Order"
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
      Left            =   9150
      TabIndex        =   129
      Top             =   615
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "(HH:MM)"
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
      Left            =   11085
      TabIndex        =   128
      Top             =   945
      Width           =   630
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
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
      Left            =   11520
      TabIndex        =   127
      Top             =   600
      Width           =   1050
   End
   Begin VB.Image tndBatal 
      Height          =   375
      Left            =   3480
      Picture         =   "frmBtlBooking.frx":2522A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmBtlBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 15 April 2009 / 21 Feb 2012
'Created Coding  : Thania / Acep
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim objBooking As New BClsBooking
Dim bKondisi(6) As String
Dim gNoRangka As String
Dim nopol As String
Dim baris As Integer
Dim barisSublet As Integer
Dim barispart As Integer
Dim nBooking As String
Dim noBookingAppOnline As String
Dim nobookingdidx As String


Private Sub hapusForm()
' *** FUNGSI UNTUK MEMBERSIHKAN TEXTBOX PADA FORM ***

    For i = 0 To 61
        Select Case i
            Case 0
                Text(i) = ""
                Text(i).Mask = "CCCCCCCCCC"
            Case 6, 7, 8
                Text(i) = ""
                Text(i).Mask = "9999-9999999999"
            Case 10
                Text(i) = "0"
                Text(i).Mask = "9999"
            Case 18, 19, 20, 22, 25, 37
                Text(i) = ">>->>->>>>"
                Text(i).Mask = "99-99-9999"
            Case 21, 26, 38, 60
                Text(i) = ">>:>>"
                Text(i).Mask = "99:99"
            Case 24, 55, 56, 57, 58, 59
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    Text1 = ""
    VScroll1.value = 0
    
    For i = 0 To 4
        Text2(i) = ""
    Next
    
    For i = 1 To 7
        dGrid.TextMatrix(i, 1) = ""
        dGrid.TextMatrix(i, 2) = ""
        dGrid.TextMatrix(i, 3) = ""
        dGrid.TextMatrix(i, 4) = ""
        partGrid.TextMatrix(i, 1) = ""
        partGrid.TextMatrix(i, 2) = ""
        partGrid.TextMatrix(i, 3) = ""
    Next
    
    For i = 1 To 30
        subletGrid.TextMatrix(i, 0) = ""
        subletGrid.TextMatrix(i, 1) = ""
        subletGrid.TextMatrix(i, 2) = ""
        subletGrid.TextMatrix(i, 3) = ""
        subletGrid.TextMatrix(i, 4) = ""
        subletGrid.TextMatrix(i, 5) = ""
    Next
    
    For i = 0 To 5
        bKondisi(i) = ""
    Next
    
    cmdHistory.Enabled = False
    Command1.Enabled = False
    chkRtn = Unchecked
    chkThs = Unchecked
    VScroll1.value = 0
    
    For i = 46 To 51
        Text(i).Visible = False
    Next
    
    ' DEFAULT DOTAM & DDATE
    Text(18) = "01-01-3000"
    Text(19) = "01-01-3000"
    gNoRangka = ""
    Text(20) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
    Text(21) = Format(isnow, "HH:NN")
    tndBatal.Visible = False
    TblBatal.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    frmBtlBook.Visible = False
End Sub

Private Sub cmdHistory_Click()
'    If gNoRangka <> "" Then
'        gTag = "BTLBOOKING"
'        frmHistory.Text(0) = gNoRangka
'        frmHistory.Text(1) = Text(3)
'        frmHistory.Text(2) = noPol
'        frmHistory.Show (vbModal)
'    End If
    If gNoRangka = "" Then Exit Sub
    If Log.LogHOConn = 0 Then
        '' 2015-09-08 | Konek HO ''
        If MsgBox("Gagal Koneksi ke Server HO !" & vbCrLf & "Apakah ingin menyambungkan koneksi kembali?", vbYesNo) = vbYes Then
            If Get_ConnectionProfil_Bengkel99 Then
                Log.LogHOConn = 1
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
'        MsgBox "Gagal Koneksi ke Server HO !", vbInformation + vbOKOnly, "History On Line HO"
'        Exit Sub
    End If
    gTag = "ONLINE"
    dRangka = gNoRangka
    nData(0) = gNoRangka
    nData(1) = Text(3)
    nData(2) = nopol
    Dim fH As New FrmHistory_new
    fH.show (vbModal)
End Sub
Private Sub Command1_Click()
    If Replace(Replace(Text(15), ">", ""), " ", "") = "" Then Exit Sub
    gTag = "LOCAL"
    dRangka = gNoRangka
    nData(0) = gNoRangka
    nData(1) = Text(3)
    nData(2) = nopol
    Dim fH As New FrmHistory_new
    fH.show (vbModal)
End Sub

Private Sub cmdOK_Click()
1     If Text(0) = "" Then
2         Call cmdCancel_Click
3         Exit Sub
4     End If
      Dim rsLama As ADODB.Recordset
      Dim rsAct As ADODB.Recordset
      Dim hlp As String
      Dim nTotAct As Integer
      Dim evenID As String

5         On Error GoTo vbError

          Dim rsSO As ADODB.Recordset
6         If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "DD/MM/YYYY") Then
7             Set rsSO = objO.FetchDataACC("Select NOSO FROM SOORD WITH(NOLOCK) Where NoSO = '" & Trim(Text(0)) & "' and ((IsNull(QtyKirimBL,0) + IsNull(QtyKirimBI,0) + IsNull(QtyKeluar,0)) <> 0 ) AND ([TO] <> 0 Or BO not in ('','CANCEL') or IsNull(QtyReserve,0) <> 0)")
8             If rsSO.RecordCount > 0 Then
9                 MsgBox "Part sudah dilakukan Pre-Picking oleh Partman !" & Chr(13) & "Batalkan Pre-Picking terlebih dahulu (via Partman)", vbInformation + vbOKOnly, "Pre-Picking"
10                Exit Sub
11            End If

12        End If
13        If Log.DiDxP1 = 1 Or Log.p2Didx = 1 Then
      '      If Get_ConnectionProfil_Bengkel99 = False Then
      '          MsgBox "Koneksi HO terputus , proses tidak dapat dilanjutkan.", vbInformation
      '          Exit Sub
      '      End If
          
14            nobookingdidx = ""
15            If noBookingAppOnline <> "" Then
      '            nobookingdidx = objO.FetchDataHO_String("SELECT booking_ID from didx..bookingonline_didx WHERE nobookingapps = '" & noBookingAppOnline & "'")
      '            evenID = objO.FetchDataHO_String("select event_id from didx..bookingonline_didx_toSend where nobookingapps = '" & NulSID(noBookingAppOnline) & "'")
              
16                nobookingdidx = objO.FetchData_String("SELECT booking_ID from [" & dbReplikasiHO & "]..p2_bookingonline_didx WHERE nobookingapps = '" & noBookingAppOnline & "'")
17                evenID = objO.FetchData_String("select event_id from [" & dbReplikasiHO & "]..p2_bookingonline_didx where nobookingapps = '" & NulSID(noBookingAppOnline) & "'")
              
18            End If
19        End If
20        Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlBooking", "cmdOk", 0)
          
21        On Error GoTo vbError
22        If Text1 <> "" Then
23            Call objHT.OpenTransactionBengkel
24            erTranAll = False
              ' PROSES PEMBATALAN BOOKING
25            If Log.helpJpcb = 1 Then
                  Dim rsHlpJPCB As ADODB.Recordset
                  Dim rsLogJPCB As ADODB.Recordset
                  Dim rsTemp As ADODB.Recordset
26                Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM STALL WHERE KETERANGAN LIKE '%|" & Trim(Text(0)) & "|%' AND TANGGAL = '" & Format(TglYYYYMMDD(Text(25)), "MM/DD/YYYY") & "'")
                  
27                Set rsLogJPCB = objHT.FetchDataBengkel("SELECT * FROM PROFILBENGKEL.DBO.LOGJPCB")
28                Set rsHlpJPCB = objHT.FetchDataBengkel("SELECT TOP 1 * FROM HELPJPCB WHERE KETERANGAN LIKE '%|" & Trim(Text(0)) & "|%' AND TANGGAL = '" & Format(TglYYYYMMDD(Text(25)), "MM/DD/YYYY") & "'")
29                If rsHlpJPCB.EOF Then
30                     If rsLogJPCB.RecordCount > 0 Then
31                        rsLogJPCB.MoveFirst
32                        Do While Not rsLogJPCB.EOF
                              'lakukan pengisian data help jpcb berdasarkan mech
                              ''2016-06-15 no booking yang ada distall yang diinsert
                              
33                           If Not rsTemp.EOF Then Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(Text(0)) & "','T','S','" & UCase(Trim(rsTemp("nostall"))) & "','" & Format(rsTemp("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsTemp("JAMMULAI")) & "','" & Trim(rsTemp("JAMSELESAI")) & "'," & Replace(rsTemp("rate"), ",", ".") & ",'" & Trim(rsTemp("mech")) & "','','" & Trim(rsTemp!keterangan) & "','" & Trim(rsTemp("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
34                           rsLogJPCB.MoveNext
35                        Loop
36                     End If
37                 Else
38                    If rsLogJPCB.RecordCount > 0 Then
39                        rsLogJPCB.MoveFirst
40                        Do While Not rsLogJPCB.EOF
                              'lakukan pengisian data help jpcb berdasarkan mech
41                           If Not rsTemp.EOF Then Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(Text(0)) & "','T','S','" & UCase(Trim(rsTemp("nostall"))) & "','" & Format(rsTemp("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsTemp("JAMMULAI")) & "','" & Trim(rsTemp("JAMSELESAI")) & "'," & Replace(rsTemp("rate"), ",", ".") & ",'" & Trim(rsTemp("mech")) & "','','" & Trim(rsTemp!keterangan) & "','" & Trim(rsTemp("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
42                           rsLogJPCB.MoveNext
43                        Loop
44                     End If
      '                rsHlpJPCB.MoveFirst
      '                Do While Not rsHlpJPCB.EOF
      '                    If rsLogJPCB.RecordCount > 0 Then
      '                        rsLogJPCB.MoveFirst
      '                        Do While Not rsLogJPCB.EOF
      '                            If (rsHlpJPCB("STATUS") = "F" Or rsHlpJPCB!tipeUpdate = "CREATE") And rsHlpJPCB!KOMP = rsLogJPCB!KOMP Then
      '                                Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry]) VALUES ('" & Trim(text(0)) & "','T','S','" & UCase(Trim(rsTempStall("nostall"))) & "','" & Format(rsTempStall("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsTempStall("JAMMULAI")) & "','" & Trim(rsTempStall("JAMSELESAI")) & "'," & Replace(rsTempStall("rate"), ",", ".") & ",'" & Trim(rsTempStall("mech")) & "','','" & Trim(rsTempStall!keterangan) & "','" & Trim(rsTempStall("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE())")
      '                            End If
      '                            rsLogJPCB.MoveNext
      '                        Loop
      '                    End If
      '                    rsHlpJPCB.MoveNext
      '                Loop
45                End If
46            End If
47            Set rsLama = objO.FetchData("SELECT * FROM BOOK WITH(NOLOCK) WHERE NOFAK = '" & Text(0) & "'")
48            Call objHT.UpdateData_BengkelOnly("UPDATE BOOK SET DELIV = 'BTL', ALASAN = '" & Text1 & "', FS_LOG_DEL = '" & gSA & "', FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY") & "' WHERE NOFAK = '" & Text(0) & "'")
49            Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'X' WHERE KETERANGAN LIKE '%|" & Trim(Text(0)) & "|%' AND TANGGAL = '" & Format(TglYYYYMMDD(Text(25)), "MM/DD/YYYY") & "'")
50            Call objHT.InsertData_BengkelOnly("INSERT INTO NOTEPART (NOSO, TIPE) VALUES ('" & Trim(Text(0)) & "', 'BB')")
              
              ''20210302 | req mb thania | update diagnose
              ''---------------------------------------------------------------
51            Call objHT.InsertData_BengkelOnly("update diagdta set nofak = '' where nofak = '" & NulSID(Text(0)) & "'")
52            Call objHT.InsertData_BengkelOnly("update diagapproval set nofak = '' where nofak = '" & NulSID(Text(0)) & "'")
              
53            If Not rsLama.EOF Then
54                nTotAct = 0
55                For i = 1 To 7
56                    hlp = "REP" & i
57                    If rsLama(hlp) <> "" Then
                       '   If Log.RPLHistori = 0 Then
58                            Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH(NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'")
                         ' Else
                         '     Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM [" & dbReplikasiHO & "]..ACTRATE WITH(NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'")
                         ' End If
59                        If Not rsAct.EOF Then
60                            nTotAct = nTotAct + CDbl(rsAct("RATE"))
61                        End If
62                    End If
63                Next
64                Call createMekanik
                  'Call decresRestTime(rsLama("MECH"), CDbl(nTotAct), 1, Format(rsLama("DATE"), "MM/DD/YYYY"))
65            End If
66            Call updateAntrian(nopol, Text(13), Text(3), Format(Text(25), "DD/MM/YYYY"), Text(26), 3, Trim(Text(0)))
              
67            If Log.DiDxP1 = 1 Then
68                If evenID <> "" Then
      '                Call objO.InsertDataHO("INSERT INTO didx..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & noBookingAppOnline & "','" & noBookingAppOnline & "','','" & Right(gDb, 2) & "','CANCELLED','0','" & evenID & "',getdate())")
      '                Call objO.InsertDataHO("UPDATE didx..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = '" & NulSID(Text1.Text) & "' WHERE event_id = '" & evenID & "' ")
69                    Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & noBookingAppOnline & "','" & noBookingAppOnline & "','','" & Right(gDb, 2) & "','CANCELLED','0','" & evenID & "',getdate())")
70                    Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = '" & NulSID(Text1.Text) & "' WHERE event_id = '" & evenID & "' ")
71                Else
      '                Call objO.InsertDataHO("INSERT INTO didx..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & Right(gDb, 2) & "." & NulSID(Text(0)) & "','" & Right(gDb, 2) & "." & NulSID(Text(0)) & "','','" & Right(gDb, 2) & "','CANCELLED','0','',getdate())")
      '                Call objO.InsertDataHO("UPDATE didx..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = '" & NulSID(Text1.Text) & "' WHERE booking_id = '" & Right(gDb, 2) & "." & NulSID(Text(0)) & "' ")
                  
72                    Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & Right(gDb, 2) & "." & NulSID(Text(0)) & "','" & Right(gDb, 2) & "." & NulSID(Text(0)) & "','','" & Right(gDb, 2) & "','CANCELLED','0','',getdate())")
73                    Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = '" & NulSID(Text1.Text) & "' WHERE booking_id = '" & Right(gDb, 2) & "." & NulSID(Text(0)) & "' ")
74                End If
75            End If
              
76            If Log.p2Didx = 1 Then
77                If evenID <> "" Then
78                    Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & noBookingAppOnline & "','" & noBookingAppOnline & "','','" & Right(gDb, 2) & "','CANCELLED','0','" & evenID & "',getdate())")
79                    Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..p2_bookingonline_didx SET booking_status = 'CANCELLED',cancellation_reason = '" & NulSID(Text1.Text) & "' WHERE event_id = '" & evenID & "' ")
80                Else
                
81                    Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & Right(gDb, 2) & "." & NulSID(Text(0)) & "','" & Right(gDb, 2) & "." & NulSID(Text(0)) & "','','" & Right(gDb, 2) & "','CANCELLED','0','',getdate())")
82                    Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..p2_bookingonline_didx SET booking_status = 'CANCELLED',cancellation_reason = '" & NulSID(Text1.Text) & "' WHERE booking_id = '" & Right(gDb, 2) & "." & NulSID(Text(0)) & "' ")
83                End If
84            End If
85            If erTranAll = False Then
86                Call objHT.CloseTransactionBengkel
87                onTR = False
88                Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlBooking", "cmdOk", 1)
89            Else
90                Call objHT.CloseTransactionBengkel_err
91                onTR = False
92                Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlBooking", "cmdOk", 2)
93                MsgBox "Proses pembatalan gagal." & vbCr & "Ulangi pembatalan Booking ", vbInformation, "Gagal Proses"
94                Exit Sub
95            End If
                  
96        Else
97            MsgBox "Isi Alasan Pembatalan !!! "
98            Text1.SetFocus
99            Exit Sub
100       End If
101       frmBtlBook.Visible = False
102       Call hapusForm
103       Exit Sub
          
vbError:
104       Call objHT.CloseTransactionAll_Err
105       onTR = False
106       Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlBooking", "cmdOk", 2)
107       MsgBox "Proses pembatalan gagal." & vbCr & "Ulangi pembatalan Booking ", vbInformation, "Gagal Proses"
108       Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub Form_Load()
    dGrid.Rows = 11
    dGrid.cols = 5
    dGrid.ColWidth(0) = 0
    dGrid.ColWidth(1) = 3500
    dGrid.ColWidth(2) = 800
    dGrid.ColWidth(3) = 0
    dGrid.ColWidth(4) = 0
    dGrid.ColAlignment(1) = flexAlignLeftCenter
    dGrid.TextMatrix(0, 1) = "Keluhan & Order Booking"
    
    partGrid.Rows = 15
    partGrid.cols = 7
    partGrid.ColWidth(0) = 0
    partGrid.ColWidth(1) = 1500
    partGrid.ColWidth(2) = 2500
    partGrid.ColWidth(3) = 0
    partGrid.ColWidth(4) = 800
    partGrid.ColWidth(5) = 0
    partGrid.ColWidth(6) = 0
    partGrid.ColAlignment(1) = flexAlignLeftCenter
    partGrid.TextMatrix(0, 1) = "Order Part Booking"
    
    ' ATUR GRID SUBLET
    barisSublet = 1
    subletGrid.Rows = 31
    subletGrid.cols = 6
    subletGrid.Clear
    subletGrid.FormatString = " | Sublet | Qty | Satuan | Harga | Part "
    subletGrid.ColWidth(0) = 0
    subletGrid.ColWidth(1) = 2500
    subletGrid.ColWidth(2) = 600
    subletGrid.ColWidth(3) = 1000
    subletGrid.ColWidth(4) = 1200
    subletGrid.ColWidth(5) = 0
    subletGrid.TextMatrix(0, 1) = "Order Sublet"

    Call hapusForm
    frmBtlBook.Visible = False
    tndBatal.Visible = False
    TblBatal.Enabled = True
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub tblBatal_Click()
    If Log.isFill = 1 Then
        Dim rsBookPart As ADODB.Recordset
        Set rsBookPart = objO.FetchDataACC("select NOSO as NOFAK, KodeItem, Keterangan as NamaBarang, SatuanNS as KodeSatuan, Qty, 'KURANG' as Flag, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID  from SOORD WITH(NOLOCK) where  noso = '" & nBooking & "'")
        If rsBookPart.RecordCount < 0 Then Set rsBookPart = objBooking.searchBookingPart(nBooking)
        If rsBookPart.RecordCount > 0 Then
            For a = 1 To 7
                If Trim(dGrid.TextMatrix(a, 1)) <> "" Then
                    Dim nCexPP As String
                    Dim rsPP As ADODB.Recordset
                    
                    'Set rsPP = objO.FetchData("Select Item_Part from HO_PartPack Where MDL = '" & Text(11) & "' and FS_Item = '" & dGrid.TextMatrix(a, 1) & "' and JOB = '" & dGrid.TextMatrix(a, 4) & "'")
                    Set rsPP = objO.FetchDataACC("Select KodeItem as Item_Part from SOORD WITH(NOLOCK) Where NOSPK = '" & nBooking & "'")
                    
                    If rsPP.RecordCount > 0 Then
                        nCexPP = "Select NOSO FROM SOORD WITH(NOLOCK) Where NoSPK = '" & nBooking & "' and ("
                        rsPP.MoveFirst
                        Do Until rsPP.EOF
                            nCexPP = nCexPP & " KodeItem like '%" & rsPP.Fields(0).value & "' OR "
                            rsPP.MoveNext
                        Loop
                        nCexPP = Left(nCexPP, Len(nCexPP) - 3) & ")"
                        nCexPP = nCexPP & " AND ((IsNull(QtyKirimBL,0) + IsNull(QtyKirimBI,0) + IsNull(QtyKeluar,0)) <> 0 or IsNull(QtyReserve,0) <> 0) AND ([TO] <> 0 Or BO not in ('','CANCEL') )"
                        Dim rsPSO As ADODB.Recordset
                        Set rsPSO = objO.FetchDataACC(nCexPP)
                        If rsPSO.RecordCount > 0 Then
                            MsgBox "Paket Part sudah di proses oleh partman !," & Chr(13) & "Booking tidak dapat dibatalkan" & Chr(13) & Chr(13) & "Lakukan pembatalan proses part oleh Partman, untuk membatalkan Booking"
                            Exit Sub
                        End If
                    End If
                    
                End If
            Next
        End If
        frmBtlBook.Visible = True
    Else
        If MsgBox("Yakin Membatalkan booking ?", vbYesNo, "Konfirmasi") = vbYes Then
            frmBtlBook.Visible = True
'            Call objO.UpdateData("UPDATE BOOK set DELIV='BTL',PAID='D' where NO_POLISI ='" & noPol & "' and DELIV='BOK'")
'            Call hapusForm
'            Text(0) = ""
'            Text(56) = ""
'            Text(57) = ""
        End If
    End If
    
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub Text_Change(Index As Integer)
    If Index = 40 Or Index = 41 Or Index = 42 Or Index = 43 Or Index = 44 Or Index = 45 Then
        If Text(0) <> "" Then
            Select Case Index
                Case 40
                    Me.Text(46).Visible = IIf(Me.Text(40) = "LAINNYA", True, False)
                Case 41
                    Me.Text(47).Visible = IIf(Me.Text(41) = "LAINNYA", True, False)
                Case 42
                    Me.Text(48).Visible = IIf(Me.Text(42) = "LAINNYA", True, False)
                Case 43
                    Me.Text(49).Visible = IIf(Me.Text(43) = "LAINNYA", True, False)
                Case 44
                    Me.Text(50).Visible = IIf(Me.Text(44) = "LAINNYA", True, False)
                Case 45
                    Me.Text(51).Visible = IIf(Me.Text(45) = "LAINNYA", True, False)
            End Select
        End If
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
    Select Case Index
        Case 23, 46
            VScroll1.value = 0
        Case 32
            VScroll1.value = 4000
        Case 60
            VScroll1.value = 1000
        Case 63, 37
            VScroll1.value = 2000
    End Select
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
                gMainKeyCari = Text(Index)
                gSearchOption = "BTLBOOK"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
    KeyAscii = UpCase(KeyAscii)
    Select Case Index
        Case 52, 23
            KeyAscii = CekNumerik(KeyAscii)
            simpan = False
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim nIndex As String
    
    Select Case Index
        Case 0
            nBooking = ""
            If Text(Index) <> "" Then
                nBooking = Trim(Text(Index))
                If Not objBooking.isBookingExist(Trim(Text(Index))) Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "BTLBOOK"
                    frmCari.show (vbModal)
                Else
                    nIndex = Text(Index)
                    Call hapusForm
                    Text(Index) = nIndex
                    Call panggilBooking(Text(Index))
                End If
            End If
    End Select
End Sub

Private Sub panggilBooking(nomorBooking As String)
      Dim rsBooK As ADODB.Recordset
      Dim hlp As String
          
1         Set rsBooK = objBooking.searchBooking(nomorBooking)
2         If Not rsBooK.EOF Then
3             If rsBooK("DELIV") = "BTL" Then tndBatal.Visible = True
4             If tndBatal.Visible Then TblBatal.Enabled = False
5             noBookingAppOnline = ""
6             If NulSID(rsBooK("KODETIPE")) <> "" Then
                  '' isi nobookingapp dan kodeprogram
7                 If InStr(1, NulSID(rsBooK("KODETIPE")), "#|#") > 0 Then
8                     noBookingAppOnline = Left(NulSID(rsBooK("KODETIPE")), InStr(1, NulSID(rsBooK("KODETIPE")), "#|#") - 1)
9                 End If
10            End If
11            Text(1) = rsBooK("FS_KD_PLG")
12            Text(2) = rsBooK("FS_KD_WIL")
13            Text(3) = rsBooK("CUSTOMER")
14            Text(4) = rsBooK("CUST_ADRES")
15            Text(5) = rsBooK("KOTA")
16            Text(6) = rsBooK("TELPON")
17            Text(7) = rsBooK("GSM")
18            Text(8) = rsBooK("CDMA")
19            Text(9) = cekMOU(rsBooK("MDL"), rsBooK("NO_CHASSIS"))
20            Text(10) = rsBooK("FS_TH_MOBIL")
21            Text(11) = rsBooK("MDL")
22            Text(12) = rsBooK("MODEL")
23            Text(13) = rsBooK("FMDL")
24            Text(14) = rsBooK("COLOR")
25            Text(15) = rsBooK("NO_CHASSIS")
26            Text(16) = rsBooK("FS_KD_DEALER")
27            Text(17) = rsBooK("FS_NM_DEALER")
28            Text(18) = TglDDMMYYYY(NulTgl(rsBooK("DOTAM")))
29            Text(19) = TglDDMMYYYY(NulTgl(rsBooK("DDATE")))
30            Text(20) = TglDDMMYYYY(NulTgl(rsBooK("DATE_BOOK")))
31            Text(21) = NulTgl(Trim(rsBooK("TIME_D")))
              
32            Text(25) = TglDDMMYYYY(rsBooK("DATE"))
33            Text(26) = Trim(rsBooK("TIME"))
34            Text(27) = rsBooK("PAID")
35            gNoRangka = rsBooK("NO_CHASSIS")
36            nopol = rsBooK("NO_POLISI")
              
              ' PERMINTAAN CUSTOMER
37            For i = 1 To 5
38                hlp = "REQ" & i
39                Text(27 + i) = rsBooK(hlp)
40            Next
              
              ' KELUHAN CUSTOMER
41            For i = 1 To 5
42                hlp = "CMPL" & i
43                Text2(-1 + i) = rsBooK(hlp)
44            Next
              
              ' CATATAN
45            For i = 1 To 4
46                hlp = "CATATAN" & i
47                Text(32 + i) = rsBooK(hlp)
48            Next
              
49            Text(37) = TglDDMMYYYY(rsBooK("ESTDATE"))
50            Text(38) = Trim(rsBooK("ESTTIME"))
51            Text(39) = gSA
52            chkRtn = IIf(rsBooK("FN_RETURN") = "1", Checked, Unchecked)
53            chkThs = IIf(rsBooK("THS") = "1", Checked, Unchecked)
54            chkAmbil = IIf(rsBooK("AMBIL") = "1", Checked, Unchecked)
55            chkKirim = IIf(rsBooK("KIRIM") = "1", Checked, Unchecked)
56            chEM = IIf(rsBooK("CODE") = "1", Checked, Unchecked)
57            chSBNP = IIf(rsBooK("KD_K") = "1", Checked, Unchecked)
              
              ' TAMPIL CONDITION
              Dim kCond As String
              Dim kBant As String
              Dim rsTmp As ADODB.Recordset
58            Set rsTmp = objO.FetchData("SELECT * FROM BOOK WITH(NOLOCK) WHERE NOFAK = '" & rsBooK("NOFAK") & "' ")
59            If Not rsTmp.EOF Then
60                For i = 1 To 6
61                    kBant = ""
62                    kCond = ""
63                    hlp = "COND" & i
64                    If rsTmp(hlp) <> "" Then
65                        If Left(rsTmp(hlp), 3) = "L| " Then
66                            kBant = Mid(rsTmp(hlp), 4, Len(rsTmp(hlp)) - 3)
67                            kCond = "LAINNYA"
68                        Else
69                            kCond = Trim(rsTmp(hlp))
70                            kBant = ""
71                        End If
72                    End If
73                    Text(39 + i) = kCond
74                    Text(45 + i) = kBant
75                    bKondisi(i - 1) = kBant
76                Next
77            End If
              
78            Text(52) = rsBooK("REMARK1")
79            Text(53) = rsBooK("SUGGEST")
80            If rsBooK("KARPLG") <> "" Then
81                Text(54) = Mid(rsBooK("KARPLG"), (InStr(rsBooK("KARPLG"), "|") + 2), Len(rsBooK("KARPLG")) - (InStr(rsBooK("KARPLG"), "|") + 1))
82                cboPelanggan = Left(rsBooK("KARPLG"), InStr(rsBooK("KARPLG"), " |"))
83            End If
              
84            Text(55) = NulA(Format(rsBooK("TOTLABOUR"), "#0,#"))
85            Text(56) = NulA(Format(rsBooK("TOTPART"), "#0,#"))
86            Text(57) = NulA(Format(rsBooK("TOTSUBLET"), "#0,#"))
87            Text(58) = NulA(Format(rsBooK("TOTOLI"), "#0,#"))
88            Text(59) = Format(CDbl(Text(55) + CDbl(Text(56)) + CDbl(Text(57))) + CDbl(Text(58)), "#0,#")
              
89            Call panggilHistory(rsBooK("NO_CHASSIS"), rsBooK("NO_POLISI"))
90            Call isiJobOrder(rsBooK)
91            If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "DD/MM/YYYY") Then
92                Call isiPartOrder(nomorBooking)
93            Else
94                Call isiPartOrderNonFillRate(nomorBooking)
95            End If
96            Call isiSubletOrder(nomorBooking)
              
97        End If
End Sub

Private Sub isiSubletOrder(noBook As String)
Dim rsSearch As ADODB.Recordset
    
    Set rsSearch = objO.FetchData("SELECT * FROM BOOK_SUBLET WITH(NOLOCK) WHERE NOFAK = '" & Trim(noBook) & "'")
    If Not rsSearch.EOF Then
        i = 1
        Do While Not rsSearch.EOF
            subletGrid.TextMatrix(i, 1) = Trim(rsSearch("NAMABARANG"))
            subletGrid.TextMatrix(i, 2) = CDbl(rsSearch("QTY"))
            subletGrid.TextMatrix(i, 3) = Trim(rsSearch("KODESATUAN"))
            subletGrid.TextMatrix(i, 4) = Format(CDbl(rsSearch("HARGA")), "#0,##")
            subletGrid.TextMatrix(i, 5) = Trim(rsSearch("FLAG"))
            i = i + 1
            rsSearch.MoveNext
        Loop
    End If
End Sub

Private Sub isiJobOrder(rsHlp2 As ADODB.Recordset)
    
    dGrid.TextMatrix(1, 1) = rsHlp2("REP1")
    dGrid.TextMatrix(2, 1) = rsHlp2("REP2")
    dGrid.TextMatrix(3, 1) = rsHlp2("REP3")
    dGrid.TextMatrix(4, 1) = rsHlp2("REP4")
    dGrid.TextMatrix(5, 1) = rsHlp2("REP5")
    dGrid.TextMatrix(6, 1) = rsHlp2("REP6")
    dGrid.TextMatrix(7, 1) = rsHlp2("REP7")
    
    dGrid.TextMatrix(1, 2) = IIf(dGrid.TextMatrix(1, 1) <> "", rsHlp2("RTE1"), "")
    dGrid.TextMatrix(2, 2) = IIf(dGrid.TextMatrix(2, 1) <> "", rsHlp2("RTE2"), "")
    dGrid.TextMatrix(3, 2) = IIf(dGrid.TextMatrix(3, 1) <> "", rsHlp2("RTE3"), "")
    dGrid.TextMatrix(4, 2) = IIf(dGrid.TextMatrix(4, 1) <> "", rsHlp2("RTE4"), "")
    dGrid.TextMatrix(5, 2) = IIf(dGrid.TextMatrix(5, 1) <> "", rsHlp2("RTE5"), "")
    dGrid.TextMatrix(6, 2) = IIf(dGrid.TextMatrix(6, 1) <> "", rsHlp2("RTE6"), "")
    dGrid.TextMatrix(7, 2) = IIf(dGrid.TextMatrix(7, 1) <> "", rsHlp2("RTE7"), "")
    
    Dim rsTemp As ADODB.Recordset
    baris = 1
    nJRate = 0
    For j = 1 To 7
        If Trim(dGrid.TextMatrix(j, 1)) <> "" Then
            Set rsTemp = objO.FetchData("SELECT rate,job FROM HO_flate WITH(NOLOCK) WHERE FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND MDL = '" & rsHlp2("MDL") & "' and FS_LOG_DEL = ''")
            If Not rsTemp.EOF Then
                dGrid.TextMatrix(j, 2) = rsTemp("RATE")
                dGrid.TextMatrix(j, 3) = ""
                dGrid.TextMatrix(j, 4) = rsTemp("JOB")
            Else
                dGrid.TextMatrix(j, 3) = True '"fOther"
            End If
            nJRate = nJRate + CDbl(dGrid.TextMatrix(j, 2))
            baris = baris + 1
        End If
    Next
End Sub

Private Sub isiPartOrder(numberBooking As String)
Dim rsBookPart As ADODB.Recordset
Dim i As Integer
    Set rsBookPart = objO.FetchDataACC("select NOSO as NOFAK, KodeItem, Keterangan as NamaBarang, SatuanNS as KodeSatuan, Qty, 'KURANG' as Flag, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID  from SOORD WITH(NOLOCK) where  noso = '" & numberBooking & "'")
    If rsBookPart.RecordCount = 0 Then Set rsBookPart = objBooking.searchBookingPart(numberBooking)
    If Not rsBookPart.EOF Then
        i = 1
        Do While Not rsBookPart.EOF
            With partGrid
                .TextMatrix(i, 1) = rsBookPart("KODEITEM")
                .TextMatrix(i, 2) = rsBookPart("NAMABARANG")
                .TextMatrix(i, 3) = rsBookPart("QTY")
            End With
            i = i + 1
            rsBookPart.MoveNext
        Loop
    End If
End Sub

Private Sub isiPartOrderNonFillRate(numberBooking As String)
Dim rsBookPart As ADODB.Recordset
Dim i As Integer
    Set rsBookPart = objO.FetchData("select * from BOOK_PART WITH(NOLOCK) where  NOFAK = '" & numberBooking & "'")
    If rsBookPart.RecordCount = 0 Then Set rsBookPart = objBooking.searchBookingPart(numberBooking)
    If Not rsBookPart.EOF Then
        i = 1
        Do While Not rsBookPart.EOF
            With partGrid
                .TextMatrix(i, 1) = rsBookPart("KODEITEM")
                    
                .TextMatrix(i, 2) = rsBookPart("NAMABARANG")
                .TextMatrix(i, 3) = rsBookPart("QTY")
            End With
            i = i + 1
            rsBookPart.MoveNext
        Loop
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub VScroll1_Change()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 200
    End If
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 200
    End If
End Sub

Private Sub panggilHistory(noRgk As String, noPols As String)
Dim rsHist As ADODB.Recordset

    If noRgk = "" Then
        If Log.RPLHistori = 0 Then
            Set rsHist = objO.FetchData("SELECT * FROM History WITH(NOLOCK) WHERE no_polisi = '" & Trim(noPols) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC")
        Else
            Set rsHist = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..Historynew WITH(NOLOCK) WHERE no_polisi = '" & Trim(noPols) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC")
        End If
    Else
        If Log.RPLHistori = 0 Then
            Set rsHist = objO.FetchData("SELECT * FROM History WITH(NOLOCK) WHERE no_chassis = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC")
        Else
            Set rsHist = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..Historynew WITH(NOLOCK) WHERE no_chassis = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC")
        End If
    End If
    If Not rsHist.EOF Then
        Text(22) = Format(Trim(rsHist("DATE")), "dd-mm-yyyy")
        Text(23) = Nul(Trim(rsHist("FS_NM_DEALER")))
        Text(24) = Format(CDbl(rsHist("KM")), "#,#")
        If DateDiff("d", Format(rsHist("DATE"), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) > 15 Then
            chkRtn.Enabled = False
        Else
            chkRtn.Enabled = True
        End If
        cmdHistory.Enabled = True
        Command1.Enabled = True
    Else
        chkRtn.Enabled = False
        Text(24) = 0
        'cmdHistory.Enabled = False
    End If
    cmdHistory.Enabled = True
    Command1.Enabled = True
End Sub
