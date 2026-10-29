VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBtlInv 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pembatalan Invoice"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   0
      Picture         =   "frmBtlInv.frx":0000
      ScaleHeight     =   10695
      ScaleWidth      =   495
      TabIndex        =   112
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   9000
      Width           =   2535
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6135
      LargeChange     =   800
      Left            =   12840
      Max             =   2600
      SmallChange     =   600
      TabIndex        =   76
      Top             =   2880
      Width           =   495
   End
   Begin VB.PictureBox frmBtlInvoice 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   2880
      ScaleHeight     =   2115
      ScaleWidth      =   5595
      TabIndex        =   102
      Top             =   2880
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   5355
      End
      Begin ucXPButton.XPButton cmdOK 
         Height          =   495
         Left            =   1530
         TabIndex        =   113
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
         MICON           =   "frmBtlInv.frx":1D91
         PICN            =   "frmBtlInv.frx":1DAD
         PICH            =   "frmBtlInv.frx":224E
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
         TabIndex        =   114
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
         MICON           =   "frmBtlInv.frx":2768
         PICN            =   "frmBtlInv.frx":2784
         PICH            =   "frmBtlInv.frx":2CC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         TabIndex        =   115
         Top             =   165
         Width           =   2745
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   5520
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   5520
         Y1              =   2040
         Y2              =   2040
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   360
      TabIndex        =   43
      Top             =   480
      Width           =   8685
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   930
         TabIndex        =   44
         Top             =   510
         Width           =   3405
         _ExtentX        =   6006
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   930
         TabIndex        =   45
         Top             =   800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   930
         TabIndex        =   46
         Top             =   1110
         Width           =   3405
         _ExtentX        =   6006
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   930
         TabIndex        =   47
         Top             =   1410
         Width           =   1845
         _ExtentX        =   3254
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   10
         Left            =   5670
         TabIndex        =   48
         Top             =   210
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   5670
         TabIndex        =   49
         Top             =   510
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   15
         Left            =   5670
         TabIndex        =   50
         Top             =   1110
         Width           =   2355
         _ExtentX        =   4154
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   14
         Left            =   5670
         TabIndex        =   51
         Top             =   810
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   13
         Left            =   7650
         TabIndex        =   52
         Top             =   510
         Width           =   795
         _ExtentX        =   1402
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   53
         Top             =   210
         Width           =   585
         _ExtentX        =   1032
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   54
         Top             =   210
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   5670
         TabIndex        =   55
         Top             =   1410
         Width           =   1410
         _ExtentX        =   2487
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   17
         Left            =   5670
         TabIndex        =   56
         Top             =   1695
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   11
         Left            =   7230
         TabIndex        =   57
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   930
         TabIndex        =   58
         Top             =   1700
         Width           =   1845
         _ExtentX        =   3254
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   930
         TabIndex        =   59
         Top             =   2000
         Width           =   1845
         _ExtentX        =   3254
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   18
         Left            =   5670
         TabIndex        =   60
         Top             =   1995
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   41
         Left            =   3120
         TabIndex        =   61
         Top             =   210
         Width           =   1185
         _ExtentX        =   2090
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
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine."
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
         TabIndex        =   75
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA"
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
         TabIndex        =   74
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM"
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
         Left            =   465
         TabIndex        =   73
         Top             =   1740
         Width           =   315
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO TAM"
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
         Left            =   6495
         TabIndex        =   72
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame."
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
         Left            =   5010
         TabIndex        =   71
         Top             =   1740
         Width           =   510
      End
      Begin VB.Label Label4 
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
         Left            =   4875
         TabIndex        =   70
         Top             =   1455
         Width           =   645
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
         Left            =   135
         TabIndex        =   69
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label5 
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
         Left            =   5040
         TabIndex        =   68
         Top             =   1125
         Width           =   480
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
         Left            =   5100
         TabIndex        =   67
         Top             =   525
         Width           =   420
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
         Left            =   4875
         TabIndex        =   66
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN"
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
         Left            =   405
         TabIndex        =   65
         Top             =   1455
         Width           =   375
      End
      Begin VB.Label Label23 
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
         Left            =   450
         TabIndex        =   64
         Top             =   1155
         Width           =   330
      End
      Begin VB.Label Label24 
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
         Left            =   285
         TabIndex        =   63
         Top             =   855
         Width           =   495
      End
      Begin VB.Label Label27 
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
         Left            =   375
         TabIndex        =   62
         Top             =   555
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   9000
      TabIndex        =   77
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   3480
         TabIndex        =   116
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Regular"
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
         Index           =   0
         Left            =   840
         TabIndex        =   80
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Program"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   79
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warranty"
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
         Index           =   2
         Left            =   2880
         TabIndex        =   78
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   20
         Left            =   2640
         TabIndex        =   81
         Top             =   1605
         Width           =   1050
         _ExtentX        =   1852
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
         Mask            =   "CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   19
         Left            =   2640
         TabIndex        =   82
         Top             =   1320
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   35
         Left            =   840
         TabIndex        =   83
         Top             =   600
         Width           =   1290
         _ExtentX        =   2275
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   36
         Left            =   840
         TabIndex        =   84
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   37
         Left            =   840
         TabIndex        =   85
         Top             =   1320
         Width           =   930
         _ExtentX        =   1640
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   38
         Left            =   840
         TabIndex        =   86
         Top             =   1600
         Width           =   690
         _ExtentX        =   1217
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   40
         Left            =   840
         TabIndex        =   87
         Top             =   1920
         Width           =   1050
         _ExtentX        =   1852
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   43
         Left            =   2640
         TabIndex        =   103
         Top             =   1920
         Width           =   1050
         _ExtentX        =   1852
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
         Mask            =   "CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   31
         Left            =   2640
         TabIndex        =   106
         Top             =   2520
         Width           =   1050
         _ExtentX        =   1852
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label lblPDS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3600
         TabIndex        =   108
         Top             =   1320
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. MOU"
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
         Left            =   1920
         TabIndex        =   107
         Top             =   2565
         Width           =   645
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. WO"
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
         Left            =   1995
         TabIndex        =   104
         Top             =   1965
         Width           =   570
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Inv"
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
         Left            =   2025
         TabIndex        =   93
         Top             =   1650
         Width           =   540
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
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
         Left            =   2175
         TabIndex        =   92
         Top             =   1365
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe WO"
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
         TabIndex        =   91
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM"
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
         Left            =   525
         TabIndex        =   90
         Top             =   1365
         Width           =   210
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mech"
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
         TabIndex        =   89
         Top             =   1645
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         TabIndex        =   88
         Top             =   1965
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   480
      ScaleHeight     =   6075
      ScaleWidth      =   12675
      TabIndex        =   2
      Top             =   2880
      Width           =   12735
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8535
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12375
         Begin VB.CommandButton TblBatal 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Batal"
            Height          =   375
            Left            =   10320
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   8040
            Width           =   1935
         End
         Begin VB.TextBox txtTotalRp 
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1485
            TabIndex        =   105
            Top             =   7080
            Width           =   10815
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Height          =   540
            Left            =   120
            TabIndex        =   26
            Top             =   7440
            Width           =   12165
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   32
               Left            =   1920
               TabIndex        =   27
               Top             =   165
               Width           =   3540
               _ExtentX        =   6244
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   33
               Left            =   11355
               TabIndex        =   28
               Top             =   165
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
               Mask            =   ">CCCC"
               PromptChar      =   ">"
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Suplay Slip/Keterangan"
               Height          =   255
               Left            =   0
               TabIndex        =   30
               Top             =   180
               Width           =   1755
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRC"
               Height          =   195
               Left            =   10875
               TabIndex        =   29
               Top             =   210
               Width           =   330
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Left            =   10320
            TabIndex        =   19
            Top             =   4800
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   26
               Left            =   915
               TabIndex        =   20
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
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
               Mask            =   "999"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   27
               Left            =   720
               TabIndex        =   21
               Top             =   480
               Width           =   1260
               _ExtentX        =   2223
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
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   28
               Left            =   720
               TabIndex        =   22
               Top             =   795
               Width           =   1260
               _ExtentX        =   2223
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
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material"
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
               Left            =   75
               TabIndex        =   25
               Top             =   840
               Width           =   570
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sublet"
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
               Left            =   195
               TabIndex        =   24
               Top             =   525
               Width           =   450
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               Left            =   195
               TabIndex        =   23
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10320
            TabIndex        =   12
            Top             =   2640
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   23
               Left            =   915
               TabIndex        =   13
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
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
               Mask            =   "999"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   24
               Left            =   495
               TabIndex        =   14
               Top             =   480
               Width           =   1500
               _ExtentX        =   2646
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
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   25
               Left            =   480
               TabIndex        =   15
               Top             =   840
               Width           =   1500
               _ExtentX        =   2646
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
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Oli"
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
               TabIndex        =   18
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Part"
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
               TabIndex        =   17
               Top             =   525
               Width           =   300
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               TabIndex        =   16
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10320
            TabIndex        =   5
            Top             =   240
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   21
               Left            =   915
               TabIndex        =   6
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
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
               Mask            =   "999"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   22
               Left            =   720
               TabIndex        =   7
               Top             =   480
               Width           =   1260
               _ExtentX        =   2223
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
               Format          =   "#,##0.00"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   39
               Left            =   720
               TabIndex        =   8
               Top             =   840
               Width           =   1260
               _ExtentX        =   2223
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
               Format          =   "#,##0.00"
               PromptChar      =   ">"
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Body R"
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
               TabIndex        =   11
               Top             =   885
               Width           =   510
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Labour"
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
               TabIndex        =   10
               Top             =   525
               Width           =   495
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               TabIndex        =   9
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.CommandButton cmdCetak 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cetak"
            Height          =   390
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   8040
            Visible         =   0   'False
            Width           =   1890
         End
         Begin TrueOleDBGrid70.TDBGrid JobGrid 
            Height          =   2295
            Left            =   0
            TabIndex        =   31
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4048
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).DataField=   ""
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).DataField=   ""
            Columns(8).NumberFormat=   "Standard"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).DataField=   ""
            Columns(9).NumberFormat=   "Standard"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   4
            Columns(10)._MaxComboItems=   5
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8196"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=8196"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(45)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(48)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(49)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(52)=   "Column(10)._ColStyle=8196"
            Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "JOB ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
            _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.locked=-1"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.locked=-1"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.locked=-1"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=28,.parent=13,.locked=-1"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=78,.parent=13"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=32,.parent=13,.locked=-1"
            _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
            _StyleDefs(82)  =   "Named:id=33:Normal"
            _StyleDefs(83)  =   ":id=33,.parent=0"
            _StyleDefs(84)  =   "Named:id=34:Heading"
            _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(86)  =   ":id=34,.wraptext=-1"
            _StyleDefs(87)  =   "Named:id=35:Footing"
            _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(89)  =   "Named:id=36:Selected"
            _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(91)  =   "Named:id=37:Caption"
            _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(93)  =   "Named:id=38:HighlightRow"
            _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(95)  =   "Named:id=39:EvenRow"
            _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(97)  =   "Named:id=40:OddRow"
            _StyleDefs(98)  =   ":id=40,.parent=33"
            _StyleDefs(99)  =   "Named:id=41:RecordSelector"
            _StyleDefs(100) =   ":id=41,.parent=34"
            _StyleDefs(101) =   "Named:id=42:FilterBar"
            _StyleDefs(102) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid subletGrid 
            Height          =   1815
            Left            =   0
            TabIndex        =   33
            Top             =   4800
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3201
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).DataField=   ""
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   4
            Columns(7)._MaxComboItems=   5
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "SUBLET ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
            _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=66,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(70)  =   "Named:id=33:Normal"
            _StyleDefs(71)  =   ":id=33,.parent=0"
            _StyleDefs(72)  =   "Named:id=34:Heading"
            _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(74)  =   ":id=34,.wraptext=-1"
            _StyleDefs(75)  =   "Named:id=35:Footing"
            _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   "Named:id=36:Selected"
            _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=37:Caption"
            _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(81)  =   "Named:id=38:HighlightRow"
            _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(83)  =   "Named:id=39:EvenRow"
            _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(85)  =   "Named:id=40:OddRow"
            _StyleDefs(86)  =   ":id=40,.parent=33"
            _StyleDefs(87)  =   "Named:id=41:RecordSelector"
            _StyleDefs(88)  =   ":id=41,.parent=34"
            _StyleDefs(89)  =   "Named:id=42:FilterBar"
            _StyleDefs(90)  =   ":id=42,.parent=33"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   30
            Left            =   1485
            TabIndex        =   34
            Top             =   6735
            Width           =   5655
            _ExtentX        =   9975
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
         Begin MSMask.MaskEdBox text 
            Height          =   405
            Index           =   29
            Left            =   10365
            TabIndex        =   35
            Top             =   6600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   42
            Left            =   10995
            TabIndex        =   36
            Top             =   6120
            Width           =   1020
            _ExtentX        =   1799
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
            Format          =   "#,##0.00##"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   270
            Left            =   12045
            TabIndex        =   37
            Top             =   6120
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   44
            Left            =   9000
            TabIndex        =   111
            Top             =   2640
            Width           =   1290
            _ExtentX        =   2275
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
            Mask            =   "CCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   2055
            Left            =   0
            TabIndex        =   32
            Top             =   2640
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3625
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   4
            Columns(6)._MaxComboItems=   5
            Columns(6).DataField=   ""
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "PART ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(66)  =   "Named:id=33:Normal"
            _StyleDefs(67)  =   ":id=33,.parent=0"
            _StyleDefs(68)  =   "Named:id=34:Heading"
            _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   ":id=34,.wraptext=-1"
            _StyleDefs(71)  =   "Named:id=35:Footing"
            _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   "Named:id=36:Selected"
            _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=37:Caption"
            _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(77)  =   "Named:id=38:HighlightRow"
            _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=39:EvenRow"
            _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(81)  =   "Named:id=40:OddRow"
            _StyleDefs(82)  =   ":id=40,.parent=33"
            _StyleDefs(83)  =   "Named:id=41:RecordSelector"
            _StyleDefs(84)  =   ":id=41,.parent=34"
            _StyleDefs(85)  =   "Named:id=42:FilterBar"
            _StyleDefs(86)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "__________________    +"
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
            Left            =   10320
            TabIndex        =   42
            Top             =   6360
            Width           =   1920
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8550
            TabIndex        =   41
            Top             =   6652
            Width           =   1530
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Terbilang (Rupiah)"
            Height          =   195
            Left            =   60
            TabIndex        =   40
            Top             =   7200
            Width           =   1305
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
            Height          =   255
            Left            =   405
            TabIndex        =   39
            Top             =   6780
            Width           =   945
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Materai"
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
            Left            =   10350
            TabIndex        =   38
            Top             =   6165
            Width           =   540
         End
      End
   End
   Begin VB.CheckBox ChkCtkRate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rate tercetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11760
      TabIndex        =   95
      Top             =   120
      Width           =   1380
   End
   Begin VB.CheckBox chCtkDisc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discount Tidak tercetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9600
      TabIndex        =   94
      Top             =   120
      Width           =   2100
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2600
      TabIndex        =   96
      Top             =   120
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   105
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
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
      Mask            =   ">9999999999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   97
      Top             =   90
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
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   34
      Left            =   10560
      TabIndex        =   98
      Top             =   720
      Width           =   1050
      _ExtentX        =   1852
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
      Mask            =   "CCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmBtlInv.frx":32BF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image tndBatal 
      Height          =   375
      Left            =   2880
      Picture         =   "frmBtlInv.frx":1E129
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Inv"
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
      Left            =   630
      TabIndex        =   101
      Top             =   150
      Width           =   540
   End
   Begin VB.Label LblReturn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Job"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6000
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
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
      Left            =   7575
      TabIndex        =   99
      Top             =   135
      Width           =   570
   End
End
Attribute VB_Name = "frmBtlInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 18 April 2009 / 21 Feb 2012
'Created Coding  : Thania / Acep
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsPart As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsSublet As ADODB.Recordset
Dim rsMappingCustomer As ADODB.Recordset

Dim subJob As Double
Dim gNoWo As String
Dim tipeByr As String
Dim PRC As Boolean
Dim gKdPlgDealer As String
Dim woStat As String
Dim nDuoInv As Boolean
Dim objBilling As New BClsBilling
Dim vNoRetur As String
Dim vTipeCustomer As String
Dim WarnError As Boolean
Dim pesanIT As String

Dim vPiutangGL_Int As String
Dim vPenjualanGL_Int As String
Dim vPpnGL_Int As String
Dim vMeteraiGL_Int As String

Dim vPiutangGL_Ext As String
Dim vPenjualanGL_Ext As String
Dim vPpnGL_Ext As String
Dim vMeteraiGL_Ext As String

Dim vPiutangGL_AFFNONPPN As String
Dim vPenjualanGL_AFFNONPPN As String
Dim vPpnGL_AFFNONPPN As String
Dim vMeteraiGL_AFFNONPPN As String

Dim vPiutangGL_AFFPPN As String
Dim vPenjualanGL_AFFPPN As String
Dim vPpnGL_AFFPPN As String
Dim vMeteraiGL_AFFPPN As String

Dim fakturMaterai1 As Double
Dim fakturMaterai2 As Double

Dim vPiutangGL As String
Dim vPenjualanGL As String
Dim vPenjSubletGL As String
Dim vPenjPartGL As String           ' kode akun gl
Dim vPenjOlieGL As String           ' kode akun gl
Dim vPenjAccsGL As String           ' kode akun gl
Dim vPpnGL As String
Dim vPpnPartGL As String
Dim vPpnAccsGL As String
Dim vMeteraiGL As String
Dim vDelIndex As String
Dim vHPPSublet As String
Dim vPDPSublet As String
    

' UNTUK INVOICE KEDUA
Dim gKdPlgDealer2 As String
Dim gKodePlg2 As String
Dim gInvCust2 As String
Dim gInvAlmt2 As String
Dim gInvKota2 As String
Dim gInvNPWP2 As String
Dim vTglInv As Date

Dim noFaktur1 As String
Dim noFaktur2 As String
Dim support As New clsFormDealer
Dim nSQLo(5) As String


' UNTUK RETUR INVOICE KEDUA
Dim XPiutangGL_Int As String
Dim XPenjualanGL_Int As String
Dim XPpnGL_Int As String
Dim XMeteraiGL_Int As String

Dim XPiutangGL_Ext As String
Dim XPenjualanGL_Ext As String
Dim XPpnGL_Ext As String
Dim XMeteraiGL_Ext As String

Dim XPiutangGL_AFFNONPPN As String
Dim XPenjualanGL_AFFNONPPN As String
Dim XPpnGL_AFFNONPPN As String
Dim XMeteraiGL_AFFNONPPN As String

Dim XPiutangGL_AFFPPN As String
Dim XPenjualanGL_AFFPPN As String
Dim XPpnGL_AFFPPN As String
Dim XMeteraiGL_AFFPPN As String

Dim metInvB As Double
Public tglWO       As Date
Private isInvoiceC  As Boolean
Dim rsJobDetail     As ADODB.Recordset
'' 20190625DEV1
Dim isTerlibatFreeSvcHO     As Boolean

Dim nilaiPPnInv     As Double
Dim nilaiDppInv     As Double
Dim nilaiTotalInv   As Double



Private Sub hapusForm()
1        On Error GoTo hapusForm_Error

2         For i = 0 To 44
3             Select Case i
                  Case 1, 11
4                     text(i) = "99-99-9999"
5                 Case 10, 21, 22, 23, 24, 25, 26, 27, 28, 29, 39
6                     text(i) = 0
7                 Case 31
8                 Case Else
9                     text(i) = ""
10            End Select
11        Next
          
12        opWO(0).Enabled = False
13        opWO(1).Enabled = False
14        opWO(2).Enabled = False
          
15        If rsJob.State Then rsJob.Close
16        Set rsJob = Nothing
          
17        If rsPart.State Then rsPart.Close
18        Set rsPart = Nothing
          
19        subJob = 0
20        gNoWo = ""
21        tipeByr = ""
22        gKdPlgDealer = ""
23        woStat = ""
24        nDuoInv = False
25        vNoRetur = ""
26        vTipeCustomer = ""
27        pesanIT = ""
28        gKodePlg = ""
29        gInvNo = ""
30        gNoInvoice = ""
31        gInvTgl = ""
32        gInvCust = ""
33        gInvAlmt = ""
34        gInvKota = ""
35        gInvNPWP = ""
          
36        gKdPlgDealer2 = ""
37        gKodePlg2 = ""
38        gInvCust2 = ""
39        gInvAlmt2 = ""
40        gInvKota2 = ""
41        gInvNPWP2 = ""
          
42        noFaktur1 = ""
43        noFaktur2 = ""
44        TblBatal.Enabled = False
45        tglWO = "01-01-1900"
          
46        nilaiPPnInv = 10
47        nilaiDppInv = 100
48        nilaiTotalInv = nilaiDppInv + nilaiPPnInv
          
         'On Error GoTo 0
49       Exit Sub

hapusForm_Error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
50    varErlError = Erl
51    varErrNumber = err.Number
      '' varErlError,varErrNumber
52        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure hapusForm of Form frmBtlInv")
End Sub

Private Sub doCreatePart()
    
   On Error GoTo doCreatePart_Error
    If rsPart.State Then rsPart.Close
    Set rsPart = Nothing
    With rsPart
        'If .State Then .Close
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "NAMAITEM", adVarChar, 200
        .Fields.Append "QTY", adInteger
        .Fields.Append "HARGA", adDouble
        .Fields.Append "OLI", adBoolean
        .Fields.Append "EXTDISC", adDouble
        'If Log.freepart = 1 Then
            .Fields.Append "FREEC", adBoolean
        'End If
        .Open
    End With

   'On Error GoTo 0
   Exit Sub

doCreatePart_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure doCreatePart of Form frmBtlInv")
End Sub

Private Sub doHeadPart()
          
1        On Error GoTo doHeadPart_Error

2         With partGrid
3             .Columns(0).Width = 1800
4             .Columns(0).Caption = "KODE PART"
5             .Columns(0).DataField = "KODEITEM"
6             .Columns(1).Width = 2500
7             .Columns(1).Caption = "NAMA PART"
8             .Columns(1).DataField = "NAMAITEM"
9             .Columns(2).Width = 800
10            .Columns(2).Caption = "QTY"
11            .Columns(2).DataField = "QTY"
12            .Columns(3).Visible = False
13            .Columns(3).Caption = "HARGA"
14            .Columns(3).DataField = "HRG"
15            .Columns(4).Visible = False
16            .Columns(4).Caption = "OLI"
17            .Columns(4).DataField = "OLI"
              '.Columns(5).Visible = False
18            If Log.freepart = 1 Then
                  
19                .Columns(6).Caption = "CUST"
20                .Columns(6).DataField = "FREEC"
21                .Columns(6).ValueItems.Presentation = dbgCheckBox
22                .Columns(6).Locked = True
23                If isInvoiceC Then
24                    .Columns(6).Visible = True
25                Else
26                    .Columns(6).Visible = False
27                End If
28            Else
29                .Columns(6).Visible = False
30            End If
31        End With

         'On Error GoTo 0
32       Exit Sub

doHeadPart_Error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
33    varErlError = Erl
34    varErrNumber = err.Number
      '' varErlError,varErrNumber
35        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure doHeadPart of Form frmBtlInv")
End Sub

Private Sub doCreateJob()
          
1        On Error GoTo doCreateJob_Error
2         If rsJob.State Then rsJob.Close
          'Set rsJob = Nothing
3         With rsJob
              'If .State Then .Close
4             .Fields.Append "JOB", adVarChar, 50
5             .Fields.Append "RATE", adDouble
6             .Fields.Append "ACTRT", adDouble
7             .Fields.Append "GRP", adVarChar, 6
8             .Fields.Append "MECH", adVarChar, 10
9             .Fields.Append "ADDJOB", adBoolean
10            .Fields.Append "HRG", adDouble
11            .Fields.Append "DISC", adDouble
12            .Fields.Append "DISCRP", adDouble
13            .Fields.Append "STLDISC", adDouble
14            .Fields.Append "FREEC", adBoolean
15            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(tglWO, "DD/MM/YYYY")) >= 0 Then
16                .Fields.Append "IDPROGRAM", adVarChar, 20
17                .Fields.Append "PROGRAM", adVarChar, 50
18            End If
19            .Open
20        End With

         'On Error GoTo 0
21       Exit Sub

doCreateJob_Error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
22    varErlError = Erl
23    varErrNumber = err.Number
      '' varErlError,varErrNumber
24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure doCreateJob of Form frmBtlInv")
End Sub

Private Sub doHeadJob()
    
   On Error GoTo doHeadJob_Error

        With JobGrid
        .Columns(0).Width = 2500
        .Columns(0).Caption = "JOB ORDER"
        .Columns(0).DataField = "JOB"
        
        .Columns(1).Width = 500
        .Columns(1).Caption = "RATE"
        .Columns(1).DataField = "RATE"
        
        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(tglWO, "DD/MM/YYYY")) >= 0 Then
            .Columns(2).Width = 2000
            .Columns(2).Caption = "NAMA PROGRAM"
            .Columns(2).DataField = "PROGRAM"
            .Columns(2).Visible = True
            
        Else
            .Columns(2).Width = 0
            .Columns(2).Caption = "ACTRATE"
            .Columns(2).DataField = "ACTRT"
            .Columns(2).Visible = False
        End If
        
        .Columns(3).Width = 0
        .Columns(3).Caption = "GROUP"
        .Columns(3).DataField = "GRP"
        .Columns(3).Visible = False
        
        .Columns(4).Width = 0
        .Columns(4).Caption = "MEKANIK"
        .Columns(4).DataField = "MECH"
        .Columns(4).Visible = False
        
        .Columns(5).Width = 0
        .Columns(5).Caption = "Additional Job"
        .Columns(5).DataField = "ADDJOB"
        .Columns(5).Visible = False
        
        .Columns(6).Width = 1200
        .Columns(6).Caption = "HARGA"
        .Columns(6).DataField = "HRG"
               
        .Columns(7).Width = 600
        .Columns(7).Caption = "DISC"
        .Columns(7).DataField = "DISC"
        
        .Columns(8).Width = 1200
        .Columns(8).Caption = "DISC RUPIAH"
        .Columns(8).DataField = "DISCRP"
        
        .Columns(9).Width = 1200
        .Columns(9).Caption = "SUB TOTAL"
        .Columns(9).DataField = "STLDISC"
        .Splits(0).Columns(9).Locked = True
        
        .Columns(10).Width = 800
        .Columns(10).Caption = "Customer"
        .Columns(10).DataField = "FREEC"
        .Columns(10).Alignment = dbgCenter
        .Splits(0).Columns(10).Locked = False
        .Columns(10).Visible = False
        If nDuoInv Then
            .Columns(10).Visible = True
        End If
    End With

   'On Error GoTo 0
   Exit Sub

doHeadJob_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure doHeadJob of Form frmBtlInv")
End Sub

Private Sub doHeadSublet()
    
   On Error GoTo doHeadSublet_Error

    With subletGrid
        .Columns(0).Width = 1500
        .Columns(0).Caption = "No. Nota"
        .Columns(0).DataField = "NONOTA"
        .Splits(0).Columns(0).Locked = True
        .Columns(1).Width = 1800
        .Columns(1).Caption = "KODEITEM"
        .Columns(1).Visible = False
        .Columns(2).Width = 2000
        .Columns(2).Caption = "Pekerjaan"
        .Columns(2).DataField = "NAMABARANG"
        .Splits(0).Columns(2).Locked = True
        .Columns(3).Width = 1500
        .Columns(3).Caption = "Satuan"
        .Columns(3).DataField = "KODESATUAN"
        .Splits(0).Columns(3).Locked = True
        .Columns(4).Width = 800
        .Columns(4).Caption = "Jumlah"
        .Columns(4).DataField = "QTY"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 1500
        .Columns(5).Caption = "Harga"
        .Columns(5).DataField = "HARGA"
        .Splits(0).Columns(5).Locked = True
        
        If Log.standarisasiSublet = 1 Then
            .Columns(6).Width = 800
            .Columns(6).Caption = "DISC"
            .Columns(6).DataField = "diskon"
            .Splits(0).Columns(6).Locked = True
            .Columns(6).Visible = True
        Else
            .Columns(6).Visible = False
        End If
        
        If Log.freepart = 1 Then
            .Columns(7).Width = 800
            .Columns(7).Caption = "Cust"
            .Columns(7).DataField = "freec"
            .Splits(0).Columns(7).Locked = True

            If isInvoiceC Then
                .Columns(7).Visible = True
            Else
                .Columns(7).Visible = False
            End If
        Else
            .Columns(7).Visible = False
        End If
    End With

   'On Error GoTo 0
   Exit Sub

doHeadSublet_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure doHeadSublet of Form frmBtlInv")
End Sub


Private Sub panggilINV(noInv As String)
      Dim rsInvoice As ADODB.Recordset
      Dim rsInvoiceC As ADODB.Recordset
      Dim rsBantu As ADODB.Recordset
      Dim rsHlp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset


1        isInvoiceC = False
2        If Log.freepart = 1 Then
3             Set rsInvoiceC = objO.FetchData("SELECT * FROM T_INVOICE WITH(NOLOCK) WHERE NOINV like '" & NulSID(noInv) & "%C'")
4             If Not rsInvoiceC.EOF Then
5                 isInvoiceC = True
6             End If
7        End If

8        On Error GoTo panggilINV_Error

9         Set rsInvoice = objO.FetchData(" SELECT *, (SELECT GSM FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS GSM, " _
                                      & "(SELECT CDMA FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS CDMA, " _
                                      & "(SELECT DOTAM FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS DOTAM " _
                                      & " FROM T_INVOICE WITH(NOLOCK) WHERE NOINV LIKE '" & NulSID(noInv) & "%'")
10        If Not rsInvoice.EOF Then
              
11            nilaiPPnInv = NulSID(rsInvoice("PPn"))
12            nilaiDppInv = 100
13            nilaiTotalInv = nilaiDppInv + nilaiPPnInv
              
14            text(1) = TglDDMMYYYY(rsInvoice("INVDATE"))
15            text(2) = NulSID(rsInvoice("FS_KD_PLG"))
16            text(3) = NulSID(rsInvoice("FS_KD_WIL"))
17            text(4) = NulSID(rsInvoice("CUSTOMER"))
18            text(5) = Left(NulSID(rsInvoice("CUST_ADRES")), 60)
19            text(6) = NulSID(rsInvoice("KOTA"))
20            text(7) = Replace(NulSID(rsInvoice("TELPON")), "-", "")
21            text(8) = Replace(NulSID(rsInvoice("GSM")), "-", "-")
22            text(9) = Replace(NulSID(rsInvoice("CDMA")), "-", "-")
23            text(10) = NulSID(rsInvoice("FS_TH_MOBIL"))
24            text(11) = TglDDMMYYYY(NulSID(rsInvoice("DOTAM")))
25            vTglInv = TglDDMMYYYY(rsInvoice("INVDATE"))
26            text(12) = NulSID(rsInvoice("MDL"))
27            text(13) = NulSID(rsInvoice("MODEL"))
28            text(14) = NulSID(rsInvoice("FMDL"))
29            text(15) = NulSID(rsInvoice("COLOR"))
30            text(16) = NulSID(rsInvoice("NO_POLISI"))
31            text(17) = NulSID(rsInvoice("NO_CHASSIS"))
32            text(18) = NulSID(rsInvoice("NO_ENGINE"))
33            text(19) = NulSID(rsInvoice("FRONT"))
34            text(20) = NulSID(rsInvoice("NOINV"))
35            text(30) = Left(NulSID(rsInvoice("FS_CATATAN")), 60)
36            text(31) = NulSID(rsInvoice("NOFAK"))
37            text(32) = NulSID(rsInvoice("SLIP"))
38            text(33) = NulSID(rsInvoice("FS_RC"))
39            text(34) = NulSID(rsInvoice("MOU"))
40            text(35) = NulSID(rsInvoice("KODETIPE"))
41            noFaktur1 = NulSID(rsInvoice("NOFP"))
42            noFaktur2 = NulSID(rsInvoice("NOFP2"))
43            nDuoInv = False
44            vTipeCustomer = ""
45            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(NulSID(rsInvoice!Date), "DD/MM/YYYY")) >= 0 Then
                  ''jika tipe warranty akan load data metode lama
                    ''----------------------------------------------------------
46                If typeJob(rsInvoice("TPSVC")) <> "WARRANTY" Then
47                    sSQL = "select * from [t_invoicejobdetail] with (nolock) where no_wo = '" & NulSID(rsInvoice("NO_WO")) & "' and no_chassis = '" & NulSID(rsInvoice("NO_CHASSIS")) & "' and fs_log_del = ''"
48                    Set rsJobDetail = objO.FetchData(sSQL)
49                    n = 0
50                    If Not rsJobDetail.EOF Then
51                        rsJobDetail.MoveFirst
52                        Do While Not rsJobDetail.EOF
53                            If NulSID(rsJobDetail!id_program) <> "" Then
54                                n = 2
55                                If objBilling.isDuoInvoice(NulSID(rsJobDetail!id_program)) Then
56                                    nDuoInv = True
57                                End If
58                            End If
59                            rsJobDetail.MoveNext
60                         Loop
61                    End If
62                Else
63                    n = IIf(NulSID(rsInvoice("TPSVC")) = "P", 1, IIf(NulSID(rsInvoice("TPSVC")) = "W", 2, 0))
64                    opWO(n).value = True
65                End If
66            Else
67                n = IIf(NulSID(rsInvoice("TPSVC")) = "P", 1, IIf(NulSID(rsInvoice("TPSVC")) = "W", 2, 0))
68                opWO(n).value = True
69            End If
      '        n = IIf(NulSID(rsInvoice("TPSVC")) = "P", 1, IIf(NulSID(rsInvoice("TPSVC")) = "W", 2, 0))
      '        opWO(n).value = True
              
              '' 20190625DEV1
70            If Log.lockFreeSVCHO = 1 And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
71                 Set rsHlp = objO.FetchData("SELECT top 1 * FROM [" & dbReplikasiHO & "]..summary WITH (NOLOCK) WHERE TipeTr = 'NK' and tanggal >= '" & Format(Log.tglFreeSVCHO, "yyyy-mm-dd") & "' and NoRangka = '" & Nul(rsInvoice("NO_CHASSIS")) & "' order by  Tanggal desc")
72                 If Not rsHlp.EOF Then
73                    isTerlibatFreeSvcHO = True
74                 End If
75            End If
              

76            If returnJob(LblReturn) Then
77                If typeJob(NulSID(rsInvoice("TPSVC"))) = "PROGRAM" Then
78                    opWO(1).value = True
79                    Set rsHlp = objBilling.searchProgramDetail(rsInvoice("KODETIPE"))
80                    If Not rsHlp.EOF Then
81                        text(35) = rsHlp("ID_PROGRAM")
82                        text(36) = rsHlp("PROGRAM")
83                        text(35).Visible = True
84                        text(36).Visible = True
85                    End If
86                    If objBilling.isDuoInvoice(rsInvoice("KODETIPE")) Then
87                        nDuoInv = True
                          'RTJ WO BERKEMBANG
                          '' 20190625DEV1
88                        If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
89                            Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
90                        Else
91                            Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
92                        End If
                          ''Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
93                        If Not rsHlp.EOF Then
94                            gKdPlgDealer2 = NulSID(rsHlp("KDPLG"))
95                            gKodePlg2 = NulSID(rsHlp("KDPLG"))
96                            gInvCust2 = NulSID(rsHlp("FS_NM_DEALER"))
97                            gInvAlmt2 = NulSID(rsHlp("FS_ALAMAT"))
98                            gInvKota2 = NulSID(rsHlp("FS_KOTA"))
99                            gInvNPWP2 = NulSID(rsHlp("FS_NPWP"))
                              
100                           If NulSID(rsHlp("INTERNAL")) = 1 Then
101                               If NulSID(gNpwp) = NulSID(gInvNPWP2) Then
102                                   vTipeCustomer = "AFFNONPPN"
103                               Else
104                                   vTipeCustomer = "AFFPPN"
105                               End If
106                           Else
107                               vTipeCustomer = "EXT"
108                           End If
109                       End If
110                       Call giveDiscount("100", text(21), False)
111                       text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
112                   End If
113               Else
                      ' HANYA RTJ
114                   Call giveDiscount("100", text(21), False)
115                   vTipeCustomer = "INT"
116               End If
117           Else
118               If typeJob(NulSID(rsInvoice("TPSVC"))) = "WARRANTY" Then
119                   opWO(2).value = True
                       
120                   Set rsHlp = objBilling.searchWarrantyDetail(rsInvoice("KODETIPE"))
121                   If Not rsHlp.EOF Then
122                       text(35) = rsHlp("ID_WARRANTY")
123                       text(36) = rsHlp("WARRANTY")
124                       text(35).Visible = True
125                       text(36).Visible = True
126                   End If
127                   Call giveDiscount("0", text(21), True)
128                   text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
129                   Set rsHlp = objO.FetchData("SELECT * FROM T_PROFIL WITH(NOLOCK) ")
130                   If Not rsHlp.EOF Then
131                       gKdPlgDealer = NulSID(rsHlp("FN_HOKDPLG"))
132                       gKodePlg = NulSID(rsHlp("FN_HOKDPLG"))
133                       gInvCust = NulSID(rsHlp("FN_HO"))
134                       gInvAlmt = NulSID(rsHlp("FN_HOADRES"))
135                       gInvKota = NulSID(rsHlp("FN_HOKOTA"))
136                       gInvNPWP = NulSID(rsHlp("FN_HONPWP"))
137                       vTipeCustomer = "AFFPPN"
138                   End If
139               ElseIf typeJob(NulSID(rsInvoice("TPSVC"))) = "PROGRAM" Then
140                   opWO(1).value = True
141                   Set rsHlp = objBilling.searchProgramDetail(rsInvoice("KODETIPE"))
142                   If Not rsHlp.EOF Then
143                       text(35) = rsHlp("ID_PROGRAM")
144                       text(36) = rsHlp("PROGRAM")
145                       text(35).Visible = True
146                       text(36).Visible = True
147                   End If
148                   If objBilling.isDuoInvoice(rsInvoice("KODETIPE")) Then
149                       nDuoInv = True
                          'SBI WO BERKEMBANG
150                       If rsInvoice("FS_RC") = "R001" Then
                              '' 20190625DEV1
151                           If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
152                               Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
153                           Else
154                               Set rsHlp = objBilling.takeInvoiceName(rsInvoice("NO_CHASSIS"))
155                           End If
                              ''Set rsHlp = objBilling.takeInvoiceName(rsInvoice("NO_CHASSIS"))
156                           If Not rsHlp.EOF Then
157                               gKdPlgDealer2 = NulSID(rsHlp("KDPLG"))
158                               gKodePlg2 = NulSID(rsHlp("KDPLG"))
159                               gInvCust2 = NulSID(rsHlp("FS_NM_DEALER"))
160                               gInvAlmt2 = NulSID(rsHlp("FS_ALAMAT"))
161                               gInvKota2 = NulSID(rsHlp("FS_KOTA"))
162                               gInvNPWP2 = NulSID(rsHlp("FS_NPWP"))
                                  
163                               If NulSID(rsHlp("INTERNAL")) = 1 Then
164                                   If NulSID(gNpwp) = NulSID(gInvNPWP2) Then
165                                       vTipeCustomer = "AFFNONPPN"
166                                   Else
167                                       vTipeCustomer = "AFFPPN"
168                                   End If
169                               Else
170                                   vTipeCustomer = "EXT"
171                               End If
172                           End If
173                           Call giveDiscount("100", text(21), False)
174                           text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
175                       Else
                              '' 20190625DEV1
176                           If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
177                               Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
178                           Else
179                               Set rsHlp = objBilling.takeInvoiceName(rsInvoice("NO_CHASSIS"))
180                           End If
                              ''Set rsHlp = objBilling.takeInvoiceName(rsInvoice("NO_CHASSIS"))
181                           If Not rsHlp.EOF Then
182                               gKdPlgDealer2 = Nul(NulSID(rsHlp("KDPLG")))
183                               gKodePlg2 = Nul(NulSID(rsHlp("KDPLG")))
184                               gInvCust2 = Nul(NulSID(rsHlp("FS_NM_DEALER")))
185                               gInvAlmt2 = Nul(NulSID(rsHlp("FS_ALAMAT")))
186                               gInvKota2 = Nul(NulSID(rsHlp("FS_KOTA")))
187                               gInvNPWP2 = Nul(NulSID(rsHlp("FS_NPWP")))
                                  
188                               If NulSID(rsHlp("INTERNAL")) = 1 Then
189                                   If NulSID(gNpwp) = NulSID(gInvNPWP2) Then
190                                       vTipeCustomer = "AFFNONPPN"
191                                   Else
192                                       vTipeCustomer = "AFFPPN"
193                                   End If
194                               Else
195                                   vTipeCustomer = "EXT"
196                               End If
197                           End If
198                           Call giveDiscount("100", text(21), False)
                              '' 20200107 catatan maks 60
199                           text(30) = Left(NulSID("A/N " & text(4) & "  [" & text(2) & "]"), 60)
200                       End If
201                   Else
202                       Call giveDiscount(objBilling.takeDiscountProgram(rsInvoice("KODETIPE")), text(21), False)
203                   End If
204               ElseIf typeJob(rsInvoice("TPSVC")) = "REGULAR" Then
205                   If lblPDS.Visible Then
206                       Call giveDiscount("0", text(21), False)
207                       vTipeCustomer = "INT"
208                   ElseIf rsInvoice("FS_RC") = "R001" Then
                          ' HANYA SBI
209                       woStat = "SBI"
                          '' 20190625DEV1
210                       If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
211                           Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
212                       Else
213                           Set rsHlp = objBilling.takeInvoiceName(rsInvoice("NO_CHASSIS"))
214                       End If
                          ''Set rsHlp = objBilling.takeInvoiceName(rsInvoice("NO_CHASSIS"))
215                       If Not rsHlp.EOF Then
216                           gKdPlgDealer = NulSID(rsHlp("KDPLG"))
217                           gKodePlg = NulSID(rsHlp("KDPLG"))
218                           gInvCust = NulSID(rsHlp("FS_NM_DEALER"))
219                           gInvAlmt = NulSID(rsHlp("FS_ALAMAT"))
220                           gInvKota = NulSID(rsHlp("FS_KOTA"))
221                           gInvNPWP = NulSID(rsHlp("FS_NPWP"))
222                           If NulSID(rsHlp("INTERNAL")) = 1 Then
223                               If NulSID(gNpwp) = NulSID(gInvNPWP) Then
224                                   text(40) = "AFFNONPPN"
225                               Else
226                                   text(40) = "AFFPPN"
227                               End If
228                           Else
229                               text(40) = "EXT"
230                           End If
231                       End If
232                       Call giveDiscount("0", text(21), False)
233                       text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
234                   ElseIf objBilling.isMOUCustomer(gKdCbg, rsInvoice("MDL"), rsInvoice("NO_CHASSIS")) Then
235                       Set rsHlp = objBilling.searchMOUDetail(gKdCbg, rsInvoice("MDL"), rsInvoice("NO_CHASSIS"))
236                       If Not rsHlp.EOF Then
237                           text(34) = rsHlp("NO_MOU")
238                       End If
239                       Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsInvoice("MDL"), rsInvoice("NO_CHASSIS")), text(21), False)
240                   Else
241                       Call giveDiscount("0", text(21), True)
242                   End If
243               End If
244           End If
              
245           If n = 1 Then
                  
246               If Log.RPLHistori = 0 Then
247                   Set rsBantu = objO.FetchData("SELECT PROGRAM AS JENIS FROM HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & NulSID(rsInvoice("KODETIPE")) & "'")
248               Else
249                   Set rsBantu = objO.FetchData("SELECT PROGRAM AS JENIS FROM [" & dbReplikasiHO & "]..HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & NulSID(rsInvoice("KODETIPE")) & "'")
250               End If
251           ElseIf n = 2 Then
                  
252               If Log.RPLHistori = 0 Then
253                   Set rsBantu = objO.FetchData("SELECT WARRANTY AS JENIS FROM HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & NulSID(rsInvoice("KODETIPE")) & "'")
254               Else
255                   Set rsBantu = objO.FetchData("SELECT WARRANTY AS JENIS FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & NulSID(rsInvoice("KODETIPE")) & "'")
256               End If
257           End If
258           If n > 0 Then
259               If Not rsBantu.EOF Then
260                   text(36) = NulSID(rsBantu("JENIS"))
261               End If
262           End If
              
263           text(37) = CDbl(rsInvoice("KM"))
264           text(38) = NulSID(rsInvoice("MECH"))
265           text(40) = NulSID(rsInvoice("STATUS"))
266           text(43) = NulSID(rsInvoice("NO_WO")) 'TglDDMMYYYY(rsInvoice("DATE"))
267           tglWO = TglDDMMYYYY(rsInvoice("DATE"))
268           gNoWo = NulSID(rsInvoice("NO_WO"))
              
              ' PANGGIL JOB ORDER
              Dim hlp As String
              Dim hlp2 As String
              Dim hlp3 As String
              Dim hlp4 As String
              Dim hlp5 As String
              
269           subJob = 0
270           If rsJob.State Then rsJob.Close
271           If rsJob.State = adStateClosed Then doCreateJob
              
              
272           If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(NulSID(rsInvoice!Date), "DD/MM/YYYY")) >= 0 Then
                  ''jika tipe warranty akan load data metode lama
                    ''----------------------------------------------------------
273               If typeJob(rsInvoice("TPSVC")) <> "WARRANTY" Then
                      
274                   Call loadJobDetailMultiple("", NulSID(rsInvoice("NO_CHASSIS")))
                      
275                   GoTo lanjutJobOrder
276               End If
                  
277           End If
278           For i = 1 To 7
279               hlp = "REP" & i
280               hlp2 = "RTE" & i
281               hlp3 = "DISC" & i
282               hlp4 = "OFF" & i
283               hlp5 = "FREEC" & i
                  
284               If NulSID(rsInvoice(hlp)) <> "" Then
285                   With rsJob
286                       .AddNew
287                       !job = NulSID(rsInvoice(hlp))
288                       !Rate = CDbl(rsInvoice(hlp2))
289                       !addJob = False
290                       !hrg = CDbl(rsInvoice(hlp4))
291                       !disc = CDbl(rsInvoice(hlp3))
292                       !DISCRP = CDbl(!hrg * (!disc / 100))
293                       !STLDISC = Round(!hrg * ((100 - !disc) / 100), 0)
294                       !FREEC = IIf(NulSID(rsInvoice(hlp5)) = 1, True, False)
295                       .Update
296                       text(22) = CDbl(text(22)) + CDbl(!STLDISC)
297                       subJob = subJob + IIf(!FREEC, 0, CDbl(!hrg))
298                   End With
299               End If
300           Next
              
              ' DARI ADDJOB
301           Set rsBantu = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND FS_LOG_DEL = '' AND FD_TGL_DEL = '" & Format(rsInvoice("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
302           If Not rsBantu.EOF Then
303               If rsJob.State = adStateClosed Then doCreateJob
304               Do While Not rsBantu.EOF
305                   With rsJob
306                       .AddNew
307                       !job = NulSID(rsBantu("REP"))
308                       !Rate = CDbl(rsBantu("RTE")) 'IIf(nulsid(rsBantu("VALID")) <> "", 0, CDbl(rsBantu("RTE")))
309                       !addJob = True
310                       !hrg = CDbl(rsBantu("OFFN")) 'IIf(nulsid(rsBantu("VALID")) <> "", 0, CDbl(rsBantu("OFFN")))
311                       !disc = CDbl(rsBantu("DISC"))
312                       !DISCRP = CDbl(!hrg * (!disc / 100))
313                       !STLDISC = CDbl(!hrg * ((100 - !disc) / 100)) 'Round(CDbl(!hrg * ((100 - !disc) / 100)), 0)
314                       !FREEC = IIf(NulSID(rsBantu("FREEC")) = 1, True, False)
315                       .Update
316                       text(22) = CDbl(text(22)) + CDbl(!STLDISC)
317                       subJob = subJob + IIf(!FREEC, 0, CDbl(!hrg))
318                   End With
319                   rsBantu.MoveNext
320               Loop
                  'cetakAdd.Visible = True
321           End If
322           If rsBantu.State <> adStateClosed Then rsBantu.Close
lanjutJobOrder:
323           Set JobGrid.DataSource = rsJob
324           Call doHeadJob
              
              ' PANGGIL PART
325           text(24) = 0
326           text(25) = 0
              Dim rsSearchPart As ADODB.Recordset
327           If Log.freepart = 1 Then
328               Set rsBantu = objO.FetchData("SELECT * FROM T_PART WITH(NOLOCK) WHERE NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND NOINV like '" & NulSID(text(0)) & "%'")
329           Else
330               Set rsBantu = objO.FetchData("SELECT * FROM T_PART WITH(NOLOCK) WHERE NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND NOINV = '" & NulSID(rsInvoice("NOINV")) & "'")
331           End If
332           If Not rsBantu.EOF Then
                  Dim partDb As String
                  Dim partFaktur As String
333               Set rsSearchPart = objO.FetchData("SELECT DBASEFK, NOFK FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND NOINV = '" & Left(NulSID(rsInvoice("NOINV")), IIf(InStr(NulSID(rsInvoice("NOINV")), "-") > 0, InStr(NulSID(rsInvoice("NOINV")), "-") - 1, Len(NulSID(rsInvoice("NOINV"))))) & "'")
334               If Not rsSearchPart.EOF Then
335                   partDb = Nul(NulSID(rsSearchPart("DBASEFK")))
336                   partFaktur = Nul(NulSID(rsSearchPart("NOFK")))
337               End If
                  
338               If partDb <> "" Then
339                   Set rsSearchPart = objO.FetchDataACC("SELECT KODEJURNAL FROM [" & partDb & "]..SIHDR WITH(NOLOCK) WHERE NOFAKTUR = '" & NulSID(partFaktur) & "'")
340                   If Not rsSearchPart.EOF Then
341                       text(44) = NulSID(rsSearchPart("KODEJURNAL"))
342                   End If
343               End If
                  
344               If rsPart.State Then rsPart.Close
345               If rsPart.State = adStateClosed Then doCreatePart
346               Do While Not rsBantu.EOF
347                   With rsPart
348                       .AddNew
349                       !KodeItem = NulSID(rsBantu("KODEITEM"))
350                       !NAMAITEM = NulSID(rsBantu("KETERANGAN"))
351                       !qty = CDbl(rsBantu("QTY"))
352                       !Harga = CDbl(rsBantu("HARGA"))
353                       !oli = IIf(NulSID(rsBantu("OLI")), True, False)
354                       !EXTDISC = CDbl(rsBantu("FN_EXT_DISKON"))
355                       If Log.freepart = 1 Then
356                           If Right(NulSID(rsBantu("noinv")), 1) = "C" Then
357                               !FREEC = False
358                               GoTo lanjutLooping
359                           Else
360                               !FREEC = True
361                           End If
362                       Else
363                           !FREEC = True
364                       End If
365                       .Update
366                       If Not !oli Then
367                           text(24) = CDbl(text(24)) + CDbl(!Harga)
368                       Else
369                           text(25) = CDbl(text(25)) + CDbl(!Harga)
370                       End If
371                   End With
lanjutLooping:
372                   rsBantu.MoveNext
373               Loop
374           End If
375           If rsBantu.State <> adStateClosed Then rsBantu.Close
376           Set partGrid.DataSource = rsPart
377           Call doHeadPart
378           DoEvents
379           partGrid.Refresh
              ' PANGGIL SUBLET
              
              
380           If Log.freepart = 1 Then
381               Set rsSublet = objO.FetchData("SELECT a.noinv,a.NONOTA, a.KODEITEM, a.NAMABARANG, a.KODESATUAN, a.QTY, 'TRUE' AS VALID, a.HARGA AS HARGA,ISNULL(D.DISCbilling,0) diskon , case when right(noinv,1) = 'C' then 'FALSE' ELSE 'TRUE' END as FREEC FROM HIST_SUBLET A WITH(NOLOCK) LEFT JOIN [vo_discsublet] D with (nolock) on a.NO_WO = d.NO_WO and a.NO_CHASSIS = d.NO_CHASSIS and d.kodeitem = A.KODEITEM AND D.FS_LOG_DEL = ''  WHERE a.NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND a.NOINV like '" & NulSID(text(0)) & "%'")
382           Else
383               If Log.standarisasiSublet = 1 Then
384                   Set rsSublet = objO.FetchData("SELECT a.noinv,a.NONOTA, a.KODEITEM, a.NAMABARANG, a.KODESATUAN, a.QTY, 'TRUE' AS VALID, a.HARGA AS HARGA,ISNULL(D.DISCbilling,0) diskon FROM HIST_SUBLET A WITH(NOLOCK) LEFT JOIN [vo_discsublet] D with (nolock) on a.NO_WO = d.NO_WO and a.NO_CHASSIS = d.NO_CHASSIS and d.kodeitem = A.KODEITEM AND D.FS_LOG_DEL = ''  WHERE a.NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND a.NOINV = '" & NulSID(rsInvoice("NOINV")) & "'")
385               Else
386                   Set rsSublet = objO.FetchData("SELECT noinv,NONOTA, KODEITEM, NAMABARANG, KODESATUAN, QTY, 'TRUE' AS VALID, HARGA AS HARGA, 0 as diskon FROM HIST_SUBLET A WITH(NOLOCK) WHERE NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND NOINV = '" & NulSID(rsInvoice("NOINV")) & "'")
387               End If
388           End If
389           If Not rsSublet.EOF Then
390               Set subletGrid.DataSource = rsSublet
391               Call doHeadSublet
392               DoEvents
393               subletGrid.Refresh
394               rsSublet.MoveFirst
395               Do While Not rsSublet.EOF
      '                If Right(NulSID(rsSublet!noInv), 1) = "C" Then
      '                    rsSublet!freec = False
      '                End If
396                   If Log.freepart = 1 Then
397                       If rsSublet!FREEC Then
398                           text(27) = CDbl(text(27)) + ((CDbl(rsSublet("HARGA")) * CDbl(rsSublet("QTY")))) 'CDbl(Text(27)) + ((CDbl(rsSublet("HARGA")) * CDbl(rsSublet("QTY"))) * ((100 - CDbl(rsSublet!diskon)) / 100))
399                       End If
400                   Else
401                       text(27) = CDbl(text(27)) + ((CDbl(rsSublet("HARGA")) * CDbl(rsSublet("QTY")))) 'CDbl(Text(27)) + ((CDbl(rsSublet("HARGA")) * CDbl(rsSublet("QTY"))) * ((100 - CDbl(rsSublet!diskon)) / 100))
402                   End If
403                   rsSublet.MoveNext
404               Loop
405           End If
              
406           text(39) = CDbl(rsInvoice("BODY"))
407           text(28) = CDbl(rsInvoice("MTRL"))
      '        Do While CDbl(rsInvoice("MET")) = 0
      '            If Not rsInvoice.EOF Then
      '                rsInvoice.MoveNext
      '            End If
      '        Loop
408           text(42) = CDbl(rsInvoice("MET"))
409           metInvB = 0
410           If Right(rsInvoice("NOINV"), 2) = "-A" Then
411               rsInvoice.MoveNext
412               If Right(rsInvoice("NOINV"), 2) = "-B" Then
                     ''2016-03-29 || ketika ada materai nilai nya muncul sebagai nilai di invoice A, harusnya di invoice B
                     '' text(42) = CDbl(rsInvoice("MET"))
413                  metInvB = CDbl(rsInvoice("MET"))
414               End If
415           Else
416               metInvB = CDbl(rsInvoice("MET"))
417           End If
418           text(29) = Round(Koma2Titik(text(42)) + Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
419           txtTotalRp = Angka2Huruf(text(29), 1)
              
420               If OnTopSV = 0 Then
421                   Set rsSearch = objO.FetchDataACC("select KodeLgn, SyaratBayar from Armst WITH(NOLOCK) where KodeLgn = '" & NulSID(text(2)) & "'")
                      'UNTUK FEBUARI - TOP PELANGGAN
422               Else
423                   Set rsSearch = objO.FetchDataACC("select KodeLgn, TOPGRP as SyaratBayar from Armst WITH(NOLOCK) where KodeLgn = '" & NulSID(text(2)) & "'")
424               End If
425               If Not rsSearch.BOF Then
426                   text(41) = rsSearch("SyaratBayar")
427               End If
              
              ''20220802 || penambahan validasi part yang tidak ada di inmst tidak boleh lanjut
              ''--------------------------------------------------------------------------------------------
428           sSQL = "SELECT A.KODEJURNAL, A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc,"
429           sSQL = sSQL & " C.KodeItem"
430           sSQL = sSQL & " FROM SIHDR A WITH (NOLOCK)"
431           sSQL = sSQL & " INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR"
432           sSQL = sSQL & " LEFT join Inmst C WITH (NOLOCK) ON b.KodeItem = c.KodeItem"
433           sSQL = sSQL & " WHERE A.NOSO = '" & Trim(text(43)) & "' and KodeLgn = '" & text(2) & "' AND ISNULL(NOFJDULU,'') = ''"
434           sSQL = sSQL & " AND A.NOFAKTUR NOT IN ("
435           sSQL = sSQL & "     SELECT DISTINCT ISNULL(NOFJDULU,'') AS NOFJDULU"
436           sSQL = sSQL & "     FROM SIHDR A INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR"
437           sSQL = sSQL & "     WHERE A.NOSO = '" & Trim(text(43)) & "' and KodeLgn = '" & text(2) & "'"
438           sSQL = sSQL & " ) and ISNULL(C.KodeItem,'') = ''"
439           Set rsHlp = objO.FetchDataACC(sSQL)
440           If Not rsHlp.EOF Then
441               rsHlp.MoveFirst
442               Do While Not rsHlp.EOF
443                   pesanError = pesanError & "Kode item : " & NulSID(rsHlp!KodeItem) & " | Nama Item : " & NulSID(rsHlp!keterangan) & vbCrLf
444                   rsHlp.MoveNext
445               Loop
446               MsgBox "Proses tidak dapat dilanjutkan, mohon lengkapi daftar part berikut dimaster Part :" & vbCrLf & pesanError, vbInformation
                            
447               TblBatal.Enabled = False
                   
448           End If
              
449       End If
450       rsInvoice.Close

         'On Error GoTo 0
451      Exit Sub

panggilINV_Error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
452   varErlError = Erl
453   varErrNumber = err.Number
      '' varErlError,varErrNumber
454       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure panggilINV of Form frmBtlInv")
End Sub

Private Function returnJob(labelReturn As Object) As Boolean
1        On Error GoTo returnJob_Error

2         returnJob = False
3         If labelReturn.Visible Then
4             returnJob = True
5         End If

6        'On Error GoTo 0
7        Exit Function

returnJob_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure returnJob of Form frmBtlInv")
End Function

Private Function typeJob(statusJob As String) As String
1        On Error GoTo typeJob_Error

2         If statusJob = "W" Then
3             typeJob = "WARRANTY"
4         ElseIf statusJob = "P" Then
5             typeJob = "PROGRAM"
6         Else
7             typeJob = "REGULAR"
8         End If

9        'On Error GoTo 0
10       Exit Function

typeJob_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure typeJob of Form frmBtlInv")
End Function

Private Function giveDiscount(discountValue As String, discountText As Object, statusText As Boolean)
1        On Error GoTo giveDiscount_Error

2         discountText = discountValue
3         discountText.Enabled = statusText

4        'On Error GoTo 0
5        Exit Function

giveDiscount_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
6         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure giveDiscount of Form frmBtlInv")
End Function

Private Sub CetakFormKecil()
          Dim rpt As New RptInvRetur
          
1        On Error GoTo CetakFormKecil_Error

2         rpt.Label48 = "No Urut Retur : " & vNoRetur
3         rpt.Label39 = gSecmanager
4         rpt.Label40 = "KEPALA BENGKEL"
5         rpt.Label2 = "NPWP : " & gNpwp
6         rpt.Label3 = "PKP  : " & gPkp
7         rpt.Label44 = "TELP : " & gTelp
8         rpt.Label18 = "BIAYA JASA  : " & AddSpace(Format(text(22), "#,##0"), 13, 3, " ")
9         rpt.Label19 = "SUBLET      : " & AddSpace(Format(text(27), "#,##0"), 13, 3, " ")
10        rpt.Label20 = "PARTS       : " & AddSpace(Format(text(24), "#,##0"), 13, 3, " ")
11        rpt.Label21 = "OIL/GRERASE : " & AddSpace(Format(text(25), "#,##0"), 13, 3, " ")
12        rpt.Label22 = "MATERIAL    : " & AddSpace(Format(text(28), "#,##0"), 13, 3, " ")
      '    rpt.Label23 = "SUBTOTAL    : " & AddSpace(Format(.Text(6), "#,##0"), 13, 3, " ")
13        rpt.Label24 = "METERAI     : " & AddSpace(Format(text(42), "#,##0"), 13, 3, " ")
14        rpt.Label25 = "GRAND TOTAL : " & AddSpace(Format(text(29), "#,##0"), 13, 3, " ")
          
15        rpt.Label26 = AddSpace("URAIAN BIAYA JASA", 25, 1, " ") + _
                      AddSpace("MKN", 4, 3, " ") + _
                      AddSpace("RUPIAH", 12, 3, " ")
16        If rsJob.State <> adStateClosed Then
17            rsJob.MoveFirst
18            Do While Not rsJob.EOF
19                rpt.Label33 = AddSpace(rsJob("JOB"), 25, 1, " ") + _
                                AddSpace(Format(rsJob("HRG"), "#,##0"), 12, 3, " ")
20                rsJob.MoveNext
21            Loop
22        End If
          
23        rpt.Label8 = "NO. WO. : " & text(43)
24        rpt.Label9 = "FRAME : " & text(17)
25        rpt.Label10 = "E/G : " & text(18)
26        rpt.Label11 = "NAMA : " & text(4)
27        rpt.Label12 = "ALAMAT : " & text(5)
28        rpt.Label13 = "KOTA : " & text(6)
29        rpt.Label14 = text(14) & "(" & text(13) & ")"
30        rpt.Label15 = "SA : " & text(19)
31        rpt.Label16 = "OPT : " & gSA
          
32        rpt.Label47 = "Kode Plgn  : " & text(2)
33        rpt.Label4 = "No. Retur  : " & gNoInvoice & gKodeRetur & "/" & text(40)
34        rpt.Label5 = "Tanggal    : " & isnow
35        rpt.Label6 = "Jam        : " & vJam
36        rpt.Label7 = "No. Polisi : " & text(16)
          
37        rpt.Label36 = Text1.text
38        rpt.Label37 = text(32)
39        rpt.Label38 = text(30)
          
40        rpt.Label45 = "Bank     : " & gBank
41        rpt.Label46 = "Acc. No. : " & gAcc

42        rpt.Printer.PaperWidth = 1440 * 8.5
43        rpt.Printer.PaperHeight = 1440 * 5.5
44        rpt.Printer.Orientation = 1
              
45        If gPreviewCetak = True Then
46            rpt.PrintReport False
47            Unload rpt
48        Else
49            rpt.show vbModal
50        End If

51       'On Error GoTo 0
52       Exit Sub

CetakFormKecil_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
53        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure CetakFormKecil of Form frmBtlInv")
End Sub

Private Sub Cetak()
          Dim rpt As New RptFromRetur
          Dim gInvUrai(8) As String
          Dim gInvRt(8) As String
          Dim gInvRp(8) As String
          Dim gInvRup(8) As String
          Dim gcat As String
          Dim gDisc As String
          Dim gDiscount(8) As String
          Dim gRateAdd As Double
          Dim gRpAdd As Double
          Dim gRupAdd As Double
          Dim adaAddJob As Boolean
          
1        On Error GoTo Cetak_Error

2         gKodePlg = IIf(gKodePlg = "", text(2), gKodePlg)
3         gInvNo = text(0) & IIf(nDuoInv, "-A", "") & "/" & text(40)
4         gNoInvoice = text(0) & IIf(nDuoInv, "-A", "")
5         gInvTgl = TglDDMMYYYY(isnow) 'isnow
6         gInvCust = IIf(gInvCust = "", text(4), gInvCust)
7         gInvAlmt = IIf(gInvAlmt = "", text(5), gInvAlmt)
8         gInvKota = IIf(gInvKota = "", text(6) & "-" & text(7), gInvKota) 'vNoTelp
9         gInvNPWP = IIf(gInvNPWP = "", "", gInvNPWP) 'Text(14)
          
10        gInvNoPolisi = text(16)
11        gInvModel = text(14) & " (" & text(13) & ")"
12        gInvNoMesin = text(18)
13        gInvNoChas = text(17)
14        gInvNoOrder = text(43)
15        gInvSA = text(19)
16        gInvStatus = text(40)
17        gInvMech = text(38)
18        gInvPrint = vJam
19        gInvBilling = gSA

20        For i = 1 To 7
21            gInvRt(i) = ""
22            gInvUrai(i) = ""
23            gInvRup(i) = ""
24        Next
25        gDisc = CDbl(text(21))
26        gRateAdd = 0
27        gRupAdd = 0
28        gRpAdd = 0
29        adaAddJob = False
          
30        If rsJob.State <> adStateClosed Then
31            If Not rsJob.EOF Then
32                rsJob.MoveFirst
33                i = 1
34                Do While Not rsJob.EOF
35                    If Not rsJob("ADDJOB") Then
36                        gInvUrai(i) = rsJob("JOB")
37                        gInvRt(i) = rsJob("RATE")
38                        gInvRp(i) = Format(CDbl(rsJob("HRG")), "#,#0")
39                        gInvRup(i) = Format(CDbl(rsJob("HRG")) * ((100 - rsJob("DISC")) / 100), "#,#0")
40                        gDiscount(i) = CDbl(rsJob("DISC"))
41                        i = i + 1
42                    Else
43                        adaAddJob = True
44                        gRateAdd = gRateAdd + rsJob("RATE")
45                        gRpAdd = gRpAdd + CDbl(rsJob("HRG"))
46                        gRupAdd = gRupAdd + (CDbl(rsJob("HRG")) * (100 - rsJob("DISC")) / 100)
                          'gDiscount(i) = CDbl(rsJob("DISC"))
47                    End If
48                    rsJob.MoveNext
49                Loop
50            End If
51        End If

52        gGInvRp = Format(text(22), "#,##0")
53        gInvTRp1 = Format(text(22), "#,##0")
54        gInvTRp2 = Format(text(27), "#,##0")
55        gInvTRp3 = Format(text(24), "#,##0")
56        gInvTRp4 = Format(text(25), "#,##0")
57        gInvTRp5 = Format(text(28), "#,##0")
          'gInvTRp6 = Format(.Text(6), "#,##0")
58        gInvTRp7 = Format(text(42), "#,##0")
59        gGInvTRp = Format(text(29), "#,##0")
60        gInvTerbilang = txtTotalRp
61        gInvSS = text(32)
62        gInvCatatan = text(30)

63        gInvSM = gSecmanager
64        gcat = ""
          
65        rpt.Field67.DataValue = vNoRetur
66        rpt.Field66.DataValue = gNoInvoice & gKodeRetur & "/" & text(40)
67        rpt.Field65.DataValue = gKodePlg
68        rpt.Field39.DataValue = gInvNo
69        rpt.Field40.DataValue = gInvTgl

70        rpt.Field1.DataValue = gInvCust
71        rpt.Field4.DataValue = gInvAlmt
72        rpt.Field28.DataValue = gInvKota
73        rpt.Field7.DataValue = gInvNPWP

74        rpt.Field2.DataValue = gInvNoPolisi
75        rpt.Field5.DataValue = gInvModel
76        rpt.Field12.DataValue = gInvNoMesin
77        rpt.Field14.DataValue = gInvNoChas
          
78        rpt.Field29.DataValue = gInvNoOrder
79        rpt.Field30.DataValue = gInvSA

80        rpt.Field31.DataValue = gInvStatus
81        rpt.Field32.DataValue = gInvMech
82        rpt.Field68.DataValue = gcat

83        rpt.Field33.DataValue = gInvPrint
84        rpt.Field38.DataValue = gInvBilling
          
85        rpt.Field16.DataValue = gInvUrai(1)
86        rpt.Field17.DataValue = gInvUrai(2)
87        rpt.Field18.DataValue = gInvUrai(3)
88        rpt.Field19.DataValue = gInvUrai(4)
89        rpt.Field20.DataValue = gInvUrai(5)
90        rpt.Field21.DataValue = gInvUrai(6)
91        rpt.Field22.DataValue = gInvUrai(7)
          
92        rpt.Field41.DataValue = gInvRt(1)
93        rpt.Field42.DataValue = gInvRt(2)
94        rpt.Field43.DataValue = gInvRt(3)
95        rpt.Field44.DataValue = gInvRt(4)
96        rpt.Field45.DataValue = gInvRt(5)
97        rpt.Field46.DataValue = gInvRt(6)
98        rpt.Field47.DataValue = gInvRt(7)
          
99        rpt.Field48.DataValue = gInvRup(1)
100       rpt.Field49.DataValue = gInvRup(2)
101       rpt.Field50.DataValue = gInvRup(3)
102       rpt.Field51.DataValue = gInvRup(4)
103       rpt.Field52.DataValue = gInvRup(5)
104       rpt.Field53.DataValue = gInvRup(6)
105       rpt.Field54.DataValue = gInvRup(7)
          
106       rpt.Field55.DataValue = gInvTRp1
107       rpt.Field56.DataValue = gInvTRp2
108       rpt.Field57.DataValue = gInvTRp3
109       rpt.Field58.DataValue = gInvTRp4
110       rpt.Field59.DataValue = gInvTRp5
111       rpt.Field60.DataValue = gInvTRp6
112       rpt.Field61.DataValue = gInvTRp7
          
113       rpt.Field23.DataValue = gGInvRp
114       rpt.Field27.DataValue = gGInvTRp
115       rpt.Field25.DataValue = gInvTerbilang
          
116       rpt.Field62.DataValue = gInvSS
117       rpt.Field63.DataValue = gInvCatatan
118       rpt.Field64.DataValue = gInvSM
          '9=A4
          'kwarto

      '    rpt.Printer.PaperSize = 9
119       rpt.Printer.PaperSize = 256
120       rpt.Printer.PaperHeight = 1440 * 11
121       rpt.Printer.PaperWidth = 1440 * 8.5
          
122       rpt.Printer.Orientation = 1
          
123       If gPreviewCetak = True Then
124           rpt.PrintReport False
125           Unload rpt
126       Else
127           rpt.show 'vbModal
128       End If

129      'On Error GoTo 0
130      Exit Sub

Cetak_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
131       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cetak of Form frmBtlInv")
End Sub

Private Sub Cetak2()
          Dim rpt As New RptFromRetur1
          Dim gInvUrai(8) As String
          Dim gInvRt(8) As String
          Dim gInvRp(8) As String
          Dim gInvRup(8) As String
          Dim gcat As String
          Dim gDisc As String
          Dim gDiscount(8) As String
          Dim gRateAdd As Double
          Dim gRpAdd As Double
          Dim gRupAdd As Double
          Dim adaAddJob As Boolean
          
1        On Error GoTo Cetak2_Error
         
         '' 20190911 nilai terbilang diambil dari nilai t_invoice
         '' --------------------------------------------------------------
2        subJob = objO.FetchData_Double("SELECT total FROM t_invoice WITH (NOLOCK) WHERE noinv = '" & NulSID(text(20)) & "-B'")
3         If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
4           If Trim(gNpwp) = Trim(gInvNPWP2) Then
               '' tanpa ppn
5              subJob = subJob * nilaiDppInv / nilaiTotalInv
6           End If
7         End If

8         gKodePlg = IIf(gKodePlg2 = "", text(2), gKodePlg2)
9         gInvNo = text(20) & IIf(nDuoInv, "-B", "") & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
10        gNoInvoice = text(0) & IIf(nDuoInv, "-B", "")
11        gInvTgl = vTglInv 'TglDDMMYYYY(isnow) '
12        gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
13        gInvAlmt = IIf(gInvAlmt2 = "", text(5), gInvAlmt2)
14        gInvKota = IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota) 'vNoTelp
15        gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
16        gInvNoPolisi = text(16)
17        gInvModel = text(14) & " (" & text(13) & ")"
18        gInvNoMesin = text(18)
19        gInvNoChas = text(17)
20        gInvNoOrder = text(43)
21        gInvSA = text(19)
22        gInvStatus = text(40)
23        gInvMech = text(38)
24        gInvPrint = vJam
25        gInvBilling = gSA
          
         
26        For i = 1 To 7
27            gInvRt(i) = ""
28            gInvUrai(i) = ""
29            gInvRup(i) = ""
30        Next
31        gDisc = CDbl(text(21))
32        gRateAdd = 0
33        gRupAdd = 0
34        gRpAdd = 0
35        adaAddJob = False
          
36        If rsJob.State <> adStateClosed Then
37            rsJob.MoveFirst
38            i = 1
39            Do While Not rsJob.EOF
40                If Not rsJob("ADDJOB") Then
41                    gInvUrai(i) = rsJob("JOB")
42                    gInvRt(i) = rsJob("RATE")
43                    gInvRp(i) = Format(CDbl(rsJob("HRG")), "#,#0")
44                    gInvRup(i) = Format(CDbl(rsJob("HRG")) * ((100 - rsJob("DISC")) / 100), "#,#0")
45                    gDiscount(i) = CDbl(rsJob("DISC"))
46                    i = i + 1
47                Else
48                    adaAddJob = True
49                    gRateAdd = gRateAdd + rsJob("RATE")
50                    gRpAdd = gRpAdd + CDbl(rsJob("HRG"))
51                    gRupAdd = gRupAdd + (CDbl(rsJob("HRG")) * (100 - rsJob("DISC")) / 100)
                      'gDiscount(i) = CDbl(rsJob("DISC"))
52                End If
53                rsJob.MoveNext
54            Loop
55        End If

56        gGInvRp = Format(subJob, "#,##0")
57        gInvTRp1 = 0    'Format(Text(22), "#,##0")
58        gInvTRp2 = 0    'Format(Text(27), "#,##0")
59        gInvTRp3 = 0    'Format(Text(24), "#,##0")
60        gInvTRp4 = 0    'Format(Text(25), "#,##0")
61        gInvTRp5 = 0    'Format(Text(28), "#,##0")
          'gInvTRp6 = Format(.Text(6), "#,##0")
62        gInvTRp7 = 0    'Format(Text(42), "#,##0")
63        gGInvTRp = Format(subJob, "#,##0")
64        gInvTerbilang = Angka2Huruf(SpyRound(subJob), 1)
65        gInvSS = text(32)
66        gInvCatatan = text(30)

67        gInvSM = gSecmanager
68        gcat = ""
          
69        rpt.Field67.DataValue = vNoRetur
70        rpt.Field66.DataValue = gNoInvoice & gKodeRetur & "/" & text(40)
71        rpt.Field65.DataValue = gKodePlg
72        rpt.Field39.DataValue = gInvNo
73        rpt.Field40.DataValue = gInvTgl

74        rpt.Field1.DataValue = gInvCust
75        rpt.Field4.DataValue = gInvAlmt
76        rpt.Field28.DataValue = gInvKota
77        rpt.Field7.DataValue = gInvNPWP

78        rpt.Field2.DataValue = gInvNoPolisi
79        rpt.Field5.DataValue = gInvModel
80        rpt.Field12.DataValue = gInvNoMesin
81        rpt.Field14.DataValue = gInvNoChas
          
82        rpt.Field29.DataValue = gInvNoOrder
83        rpt.Field30.DataValue = gInvSA

84        rpt.Field31.DataValue = gInvStatus
85        rpt.Field32.DataValue = gInvMech
86        rpt.Field68.DataValue = gcat

87        rpt.Field33.DataValue = gInvPrint
88        rpt.Field38.DataValue = gInvBilling
          
89        rpt.Field16.DataValue = gInvUrai(1)
90        rpt.Field17.DataValue = gInvUrai(2)
91        rpt.Field18.DataValue = gInvUrai(3)
92        rpt.Field19.DataValue = gInvUrai(4)
93        rpt.Field20.DataValue = gInvUrai(5)
94        rpt.Field21.DataValue = gInvUrai(6)
95        rpt.Field22.DataValue = gInvUrai(7)
          
96        rpt.Field41.DataValue = gInvRt(1)
97        rpt.Field42.DataValue = gInvRt(2)
98        rpt.Field43.DataValue = gInvRt(3)
99        rpt.Field44.DataValue = gInvRt(4)
100       rpt.Field45.DataValue = gInvRt(5)
101       rpt.Field46.DataValue = gInvRt(6)
102       rpt.Field47.DataValue = gInvRt(7)
          
103       rpt.Field48.DataValue = gInvRup(1)
104       rpt.Field49.DataValue = gInvRup(2)
105       rpt.Field50.DataValue = gInvRup(3)
106       rpt.Field51.DataValue = gInvRup(4)
107       rpt.Field52.DataValue = gInvRup(5)
108       rpt.Field53.DataValue = gInvRup(6)
109       rpt.Field54.DataValue = gInvRup(7)
          
110       rpt.Field55.DataValue = gInvTRp1
111       rpt.Field56.DataValue = gInvTRp2
112       rpt.Field57.DataValue = gInvTRp3
113       rpt.Field58.DataValue = gInvTRp4
114       rpt.Field59.DataValue = gInvTRp5
115       rpt.Field60.DataValue = gInvTRp6
116       rpt.Field61.DataValue = gInvTRp7
          
117       rpt.Field23.DataValue = gGInvRp
118       rpt.Field27.DataValue = gGInvTRp
119       rpt.Field25.DataValue = gInvTerbilang
          
120       rpt.Field62.DataValue = gInvSS
121       rpt.Field63.DataValue = gInvCatatan
122       rpt.Field64.DataValue = gInvSM
          '9=A4
          'kwarto

      '    rpt.Printer.PaperSize = 9
123       rpt.Printer.PaperSize = 256
124       rpt.Printer.PaperHeight = 1440 * 11
125       rpt.Printer.PaperWidth = 1440 * 8.5
          
126       rpt.Printer.Orientation = 1
          
127       If gPreviewCetak = True Then
128           rpt.PrintReport False
129           Unload rpt
130       Else
131           rpt.show 'vbModal
132       End If

         'On Error GoTo 0
133      Exit Sub

Cetak2_Error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
134   varErlError = Erl
135   varErrNumber = err.Number
      '' varErlError,varErrNumber
136       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure Cetak2 of Form frmBtlInv")
End Sub

Private Sub cmdCancel_Click()
    frmBtlInvoice.Visible = False
End Sub

Private Sub cmdOK_Click()
      Dim NoRetur As Double
      Dim boleh As Boolean
      Dim pesan As String
      Dim rsSearch As ADODB.Recordset
      Dim strSql As String
      Dim isWarrantyHO As Boolean

1         On Error GoTo vbError
2         boleh = True
3         pesan = ""
4         If Trim(Text1) = "" Then
5             boleh = False
6             pesan = pesan + "Alasan Pembatalan harus diisi. "
7         End If
          
8         Set rsSearch = objO.FetchData("SELECT FS_KD_RETUR FROM T_INVOICE WITH(NOLOCK) WHERE NOINV LIKE '" & Trim(text(0)) & "%'")
9         If Not rsSearch.EOF Then
10            If Trim(rsSearch("FS_KD_RETUR")) <> "" Then
11                boleh = False
12                pesan = pesan + "Invoice ini sudah diretur. "
13            End If
14        End If
           
15        Set rsSearch = objO.FetchData("SELECT NO_WO FROM JOBDTA WITH(NOLOCK) WHERE NOINV = '" & Trim(text(0)) & "'")
16        If rsSearch.EOF Then
17            boleh = False
18            pesan = pesan + "Invoice ini tidak terdaftar. "
19        End If
          
20        If Log.addMaterai = 1 Then
              Dim nilaiMAT As Double
21            nilaiMAT = objO.FetchData_Double("select isnull(sum(materai),0) as Jumlah from [FAKTUR" & Right(gDb, 2) & "]..TransaksiKuitansi where NoInv Like 'E-" & NulSID(text(0)) & "%' and TipeTr = 'SV'")
22            If nilaiMAT > 0 Then
23                Text1.text = ""
24                MsgBox "Terdapat Kwitansi dan VA atas Inv " & objO.FetchData_String("select NoInv as Jumlah from [FAKTUR" & Right(gDb, 2) & "]..TransaksiKuitansi where NoInv Like 'E-" & NulSID(text(0)) & "%' and TipeTr = 'SV'") & " yang belum dibatalkan," & vbCrLf & "Harap hubungi Admin untuk melakukan pembatalan Kwitansi dan VA.", vbInformation
25                Exit Sub
26            End If
27        End If

          '20/12/2012  Cust Non Axtif
28        vKodeLgn = text(2).text
          Dim rsNon As ADODB.Recordset
29        Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) WHERE kodeLgn = '" & Trim(vKodeLgn) & "' and nonaktifflag = 0")
30        If rsNon.RecordCount = 0 Then
31            boleh = False
32            pesan = "Kode Langgan : " & vKodeLgn & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut"
33        End If
          'Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WHERE kodeLgn = '" & Trim(gKdPlgDealer2) & "' and nonaktifflag = 0")
          'If rsNon.RecordCount = 0 Then
          '    boleh = False
          '    pesan = "Kode Langgan : " & gKdPlgDealer2 & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut"
          'End If
          
34        If boleh Then
35            nSQLo(0) = ""
36            nSQLo(1) = ""
37            nSQLo(2) = ""
38            nSQLo(3) = ""
39            nSQLo(4) = ""
40            nSQLo(5) = ""
41            isWarrantyHO = False
42            Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlInv", "cmdOk", 0)
43            Call gTglServerBilling
              
                  
              'Cek di AR, ambil delindexnya untuk hapus trans & Adjust
              
44            On Error GoTo vbError
45            Call objHT.OpenTransactionAll
46            erTranAll = False
              
47            sSQL = "Select * from T_PROFIL WITH(NOLOCK)"
48            Set rsSearch = objO.FetchData("Select FN_NO_RET from T_PROFIL WITH(NOLOCK)")
49            If Not rsSearch.BOF Then
50                NoRetur = Format(rsSearch("FN_NO_RET"), "00000000")
51            End If
              
              
              '-----------------------------------------------------
              '- 2015-04-10 | RUDI | PROSES RETUR CLAIM WARRANTY   -
              '-----------------------------------------------------
              
52            If opWO(2).value = True Then
53                Call objHT.UpdateData_Bengkel("UPDATE CLAIMWARRANTY SET STATUS = 'BTL' WHERE NO_WO = '" & Trim(text(43)) & "' AND NOINV = '" & Trim(text(0)) & "'")
                  'lakukan proses transfer HO
54                strSql = "UPDATE CLAIMWARRANTY SET STATUS = 'BTL' WHERE NO_WO = '" & Trim(text(43)) & "' AND NOINV = '" & Trim(text(0)) & "'"
55                If Log.LogHOConn = 1 Then
56                    If ConHO = False Then
57                        Call objHT.InsertData_Bengkel(strSql)
58                    Else
59                        If Get_ConnectionProfil_Bengkel99 = False Then
60                            sSQL = Replace(strSql, "'", "~")
61                            strSql = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                              "Values ( 'CLAIM','" & Trim(text(0)) & "', 'Retur', 0,'01/01/1900', '" & sSQL & "')"
62                            Call objHT.InsertData_Bengkel(strSql)
63                        Else
                              '2015-09-07 | ganti dengan koneksi HO
64                            isWarrantyHO = True
                              'Call objHT.InsertData_Ho(strSql)
65                        End If
66                    End If
67                Else
68                    sSQL = Replace(strSql, "'", "~")
69                    strSql = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                      "Values ( 'CLAIM','" & Trim(text(0)) & "', 'Retur', 0,'01/01/1900', '" & sSQL & "')"
70                    Call objHT.InsertData_Bengkel(strSql)
71                End If
72            End If
              
73            If Log.freepart = 1 Then
74              strSql = "UPDATE ClaimFOP SET STATUS = 'BTL' WHERE NO_WO = '" & Trim(text(43)) & "' AND NOINV = '" & Trim(text(0)) & "-C'"
75              Call objHT.InsertData_Bengkel(strSql)
76              strSql = "UPDATE [" & dbReplikasiHO & "]..ClaimFOP SET STATUS = 'BTL' WHERE NO_WO = '" & Trim(text(43)) & "' AND NOINV = '" & Trim(text(0)) & "-C'"
77              Call objHT.InsertData_Bengkel(strSql)
78            End If
              
79            Call objHT.InsertData_Bengkel("UPDATE CHASSINGFORM SET TANGGALINVOICE = NULL WHERE NOWO = '" & Trim(text(43)) & "'")
              
              'PERUBAHAN PENYERAGAMAN AKUN
              'TANGGAL 04-09-2014
              'RUDI KRISTANTO
              
80            If erTranAll = False Then
81                If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(tglWO, "DD/MM/YYYY")) >= 0 Then
82                    Call simpanJobDetail
83                End If
84            End If
              ''Call simpanJobDetail
              
85            If gConnecAAS = True Then
86                If Log.isAkun = 1 Then
87                    Call SimpanACCT_invoice ''Call SimpanAAS2
88                Else
89                    Call SimpanAASNonPenyeragamanAkun
90                End If
                  
91            End If
92            If erTranAll = True Then GoTo vbError
              
93            vJam = Format(isnow, "HH:NN")
94            sSQL = "update T_INVOICE set " _
                   & " FS_NO_RETUR= NOINV+'" & gKodeRetur & "' , " _
                   & " FS_LOG_DEL='" & gSA & "', " _
                   & " ALASAN_INV='" & Text1.text & "', " _
                   & " FD_RETUR='" & TglYYYYMMDD(Format(isnow, "DD/MM/YYYY")) & " " & vJam & "', " _
                   & " DELIV = 'BTL' " _
                   & " where NOINV LIKE '" & text(0) & "%'"
95            Call objHT.UpdateData_Bengkel(sSQL)
96            nSQLo(5) = Replace(sSQL, "T_INVOICE", "T_INVOICEnew") & " And CABANG = '" & Log.LogBengkel & "'"
97            If Log.RPLHistori = 1 Then
98                nSQLo(5) = Replace(sSQL, "T_INVOICE", "[" & dbReplikasiHO & "]..T_INVOICEnew") & " And CABANG = '" & Log.LogBengkel & "'"
99                Call objHT.UpdateData_Bengkel(nSQLo(5))
100           End If
101           vNoRetur = Format(NoRetur, "00000000")
102           sSQL = "update T_INVOICE set FS_KD_RETUR = '" & vNoRetur & "'" _
                   & " where NOINV LIKE '" & text(0) & "%'"
103           Call objHT.UpdateData_Bengkel(sSQL)
104           nSQLo(0) = Replace(sSQL, "T_INVOICE", "T_INVOICEnew") & " And CABANG = '" & Log.LogBengkel & "'"
105           If Log.RPLHistori = 1 Then
106               nSQLo(0) = Replace(nSQLo(0), "T_INVOICEnew", "[" & dbReplikasiHO & "]..T_INVOICEnew") & " And CABANG = '" & Log.LogBengkel & "'"
107               Call objHT.UpdateData_Bengkel(nSQLo(0))
108           End If
109           If Val(NoRetur) >= 99999999 Then
110               NoRetur = 0
111           Else
112               NoRetur = NoRetur + 1
113           End If
114           sSQL = "update T_PROFIL set FN_NO_RET = " & NoRetur
115           Call objHT.UpdateData_Bengkel(sSQL)
          
116           sSQL = "update JOBDTA set noinv='', INVDATE = '" & gTglDefault & "', RPR_COND='S' where NOINV = '" & text(0) & "'"
117           Call objHT.UpdateData_Bengkel(sSQL)
118           nSQLo(1) = sSQL & " And FS_NM_DEALER = '" & Trim(gNmCbg) & "'"
119           If Log.RPLHistori = 1 Then
120               nSQLo(1) = Replace(sSQL, "JOBDTA", "[" & dbReplikasiHO & "]..JOBDTA") & " And FS_NM_DEALER = '" & Trim(gNmCbg) & "'"
121               Call objHT.UpdateData_Bengkel(nSQLo(1))
122           End If
              
              ' UPDATE NO INV DI DIAGDTA
123           Call objHT.UpdateData_Bengkel("UPDATE DIAGDTA SET NOINV = '' WHERE NOINV = '" & text(0) & "'")
124           Call objHT.UpdateData_Bengkel("UPDATE DIAGSUBLET SET NOINV = '' WHERE NOINV = '" & text(0) & "'")
              
      '        sSQL = "DELETE HISTORY where NOINV='" & Text(0) & "' and NO_WO = '" & gNoWo & "'"
125           sSQL = "UPDATE HISTORY SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "', FS_LOG_DEL = '" & Trim(gSA) & "', ALASAN_DEL = '" & Trim(Text1) & "' where NOINV LIKE '" & text(0) & "%' and NO_WO = '" & gNoWo & "'"
126           Call objHT.UpdateData_Bengkel(sSQL)
127           nSQLo(2) = Replace(sSQL, "HISTORY", "HISTORYnew") & " And FS_NM_DEALER = '" & Trim(gNmCbg) & "'"
              
128           If Log.RPLHistori = 1 Then
129               nSQLo(2) = Replace(nSQLo(2), "HISTORYnew", "[" & dbReplikasiHO & "]..HISTORYnew") & " And FS_NM_DEALER = '" & Trim(gNmCbg) & "'"
130               Call objHT.UpdateData_Bengkel(nSQLo(2))
131           End If
              
132           sSQL = "UPDATE HIST_JOBADD SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "', FS_LOG_DEL = '" & Trim(gSA) & "' where NO_CHASSIS = '" & Trim(text(17)) & "' and NO_WO = '" & gNoWo & "'  AND NO_INV like '" & Trim(text(0)) & "%'"
133           Call objHT.UpdateData_Bengkel(sSQL)
134           nSQLo(3) = Replace(sSQL, "HIST_JOBADD", "HIST_JOBADDnew") & " And CABANG = '" & Log.LogBengkel & "'"
135           If Log.RPLHistori = 1 Then
136               nSQLo(3) = Replace(sSQL, "HIST_JOBADD", "[" & dbReplikasiHO & "]..HIST_JOBADDnew") & " And CABANG = '" & Log.LogBengkel & "'"
137               Call objHT.UpdateData_Bengkel(nSQLo(3))
138           End If
      '        sSQL = "DELETE T_PART where NOINV='" & Text(0) & "' and NO_WO = '" & gNoWo & "'"
139           sSQL = "UPDATE T_PART SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "', FS_LOG_DEL = '" & Trim(gSA) & "' where NOINV LIKE '" & text(0) & "%' and NO_WO = '" & gNoWo & "'"
140           Call objHT.UpdateData_Bengkel(sSQL)
141           nSQLo(4) = Replace(sSQL, "T_PART", "T_PARTnew") & " And CABANG = '" & Log.LogBengkel & "'"
142           If Log.RPLHistori = 1 Then
143               nSQLo(4) = Replace(sSQL, "T_PART", "[" & dbReplikasiHO & "]..T_PARTnew") & " And CABANG = '" & Log.LogBengkel & "'"
144               Call objHT.UpdateData_Bengkel(nSQLo(4))
145           End If
              'sSQL = "UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%|" & gNoWo & "|%'"
146           sSQL = "UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%" & Trim(text(16)) & " |" & Trim(text(43)) & "%' AND FLAG = 'I'"
147           Call objHT.UpdateData_Bengkel(sSQL)
              
148           If PRC Then
149               Call objHT.DeleteData_Bengkel("DELETE M_ESTIMASI WHERE NORANGKA = '" & Trim(text(17)) & "' AND NOWO = '" & text(43) & "' AND NOINV LIKE '" & text(20) & "%'")
150           End If
              
151           If Trim(text(31)) <> "" And tipeByr = "T" Then
                  
152               sSQL = "INSERT INTO MOUTRN (TANGGAL,NO_MOU,NO_WO,NOINV,NO_CHASSIS,NO_POLISI,FS_KD_DEALER,TOTAL,STATUS,FS_LOG_DEL,FD_TGL_DEL) VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & text(31) & "', '" & gNoWo & "', '" & text(0) & "', '" & text(17) & "', '" & text(16) & "', '" & gKdCbg & "', " & (CDbl(text(29)) * -1) & ", 'U','','')"
153               If Log.RPLHistori = 1 Then
154                   sSQL = Replace(sSQL, "MOUTRN", "[" & dbReplikasiHO & "]..MOUTRN")
155               End If
156               Call objHT.InsertData_Bengkel(sSQL)
157               sSQL = "UPDATE MOUHDR SET HUTANG = HUTANG - " & CDbl(text(29)) & " WHERE NO_MOU = '" & text(31) & "'"
158               If Log.RPLHistori = 1 Then
159                   sSQL = Replace(sSQL, "MOUHDR", "[" & dbReplikasiHO & "]..MOUHDR")
160               End If
161               Call objHT.UpdateData_Bengkel(sSQL)
162           End If
              
              Dim rsHlp As ADODB.Recordset
163           If Log.fakturMaterai = 1 Then
164               Set rsHlp = objHT.FetchDataBengkel("SELECT * FROM [Faktur" & Right(gDb, 2) & "]..[PemakaianMaterai] WHERE noinv like '" & text(0) & "%' and materai > 0 ")
165               If Not rsHlp.EOF Then
166                   rsHlp.MoveFirst
167                   Do While Not rsHlp.EOF
168                       sSQL = "INSERT INTO [t_invoiceMaterai]([NOINV],[NOIJIN],[NILAI]) VALUES ('" & NulSID(rsHlp!noInv) & "R','" & NulSID(rsHlp!NOIJIN) & "','" & CDbl(NulSID(rsHlp!materai)) * -1 & "')"
169                       Call objHT.InsertData_Bengkel(sSQL)
                          
170                       sSQL = "INSERT INTO [Faktur" & Right(gDb, 2) & "]..[PemakaianMaterai]([NoInv],[TglInv],[Materai],[UserEntry],[TglEntry],NOijin) VALUES ('" & NulSID(rsHlp!noInv) & "R',CONVERT(varchar(10),GETDATE(),1),'" & CDbl(NulSID(rsHlp!materai)) * -1 & "','" & gSA & "',getdate(),'" & NulSID(rsHlp!NOIJIN) & "')"
171                       Call objHT.InsertData_Bengkel(sSQL)
                          
172                       sSQL = "UPDATE [Faktur" & Right(gDb, 2) & "]..[FPKonterMat] SET SALDO = SALDO + " & CDbl(NulSID(rsHlp!materai)) & ",TERPAKAI = TERPAKAI - " & CDbl(NulSID(rsHlp!materai)) & " WHERE tahun = year(getdate()) and aktiv = 'Y'"
173                       Call objHT.InsertData_Bengkel(sSQL)
174                       rsHlp.MoveNext
175                   Loop
176               End If
177           End If
              
              
178           If tipeByr = "T" And gVerifikator <> "" Then
179               Call objHT.UpdateData_Bengkel("INSERT INTO PIUTANGTRN VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & text(43) & "', '" & text(20) & "', '" & text(17) & "', '" & text(16) & "', " & (CDbl(text(29)) * -1) & ", 'U', '" & gVerifikator & "')")
                   
180               Set rsHlp = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE FS_KD_PETUGAS = '" & gVerifikator & "'")
181               If Not rsHlp.EOF Then
182                   If rsHlp("TIPE") = "KACAB" Then
183                       Call objHT.UpdateData_Bengkel("UPDATE T_PROFIL SET HUTANGKACAB = HUTANGKACAB - " & CDbl(text(29)) & "")
184                   ElseIf rsHlp("TIPE") = "KABENG" Then
185                       Call objHT.UpdateData_Bengkel("UPDATE T_PROFIL SET HUTANGKABENG = HUTANGKABENG - " & CDbl(text(29)) & "")
186                   End If
187               End If
188           End If
              
189           Call finishService(Trim(text(43)), Trim(text(16)), Trim(text(12)), Trim(namaModel(text(12))), Trim(text(4)), Format(Trim(text(1)), "MM/DD/YYYY"), Format(isnow, "HH:NN"), 2)
              
190           If erTranAll = False Then
191               Call objHT.CloseTransactionAll
                  
192               onTR = False
193               Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlInv", "cmdOk", 1)
194               If gUkuranRetur Then
195                   Call Cetak
196                   If nDuoInv Then
197                       Call Cetak2
198                   End If
199                   If isInvoiceC Then
200                       Call Cetak3
201                   End If
202               Else
203                   Call CetakFormKecil
204               End If
205           Else
206               Call objHT.CloseTransactionAll_Err
207               onTR = False
208               Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlInv", "cmdOk", 2)
209               MsgBox "Ada Error, Pembatalan Invoice dibatalkan, Batalkan invoice kembali / Hubungi EDP !" & vbCr & pesanIT
210               Exit Sub
211           End If
                          
              ' EXPORT KE HO
212           If isWarrantyHO = True Then
213               If Get_ConnectionProfil_Bengkel99 = False Then
214                   sSQL = Replace(strSql, "'", "~")
215                   strSql = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                      "Values ( 'CLAIM','" & Trim(text(0)) & "', 'Retur', 0,'01/01/1900', '" & sSQL & "')"
216                   Call objO.InsertData(strSql)
217               Else
                      '2015-09-07 | ganti dengan koneksi HO
218                   isWarrantyHO = True
219                   Call objO.InsertDataHO(strSql)
220               End If
221           End If
              
222           If Log.RPLHistori = 0 Then Call hapusHO
223           Unload Me
224       Else
225           MsgBox pesan, vbCritical, "Pembatalan Invoice"
226           Text1.SetFocus
227       End If
228       Exit Sub
          
vbError:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
229   varErlError = Erl
230   varErrNumber = err.Number
      '' varErlError,varErrNumber
231       Call objHT.CloseTransactionAll_Err
232       onTR = False
233       Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlInv", "cmdOk", 2)
234       MsgBox "Ada Error, Pembatalan Invoice dibatalkan, Batalkan invoice kembali / Hubungi EDP !" & vbCr & pesanIT & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure CmdOK_Click of Form frmBtlInv")
235       Exit Sub
End Sub

Private Sub SimpanAASNonPenyeragamanAkun()
         
          Dim SsqlNis As String
          
          Dim vSyaratByr As String
          'Dim vJmlHari As Double
          Dim vTglSkr As Date
          Dim vTglTempo As Date
          Dim vTPenjualan As Double
          Dim vTPenjSublet As Double
          Dim vTPenjNonMet As Double
          Dim vTPPN As Double
          Dim vTotal As Double
          Dim vTotalNonPPn As Double
          Dim vTMeterai As Double
          Dim vKeterangan As String
          
          Dim vPiutangGL As String
          Dim vPenjualanGL As String
          Dim vPenjSubletGL As String
          Dim vPpnGL As String
          Dim vMeteraiGL As String
          Dim vHPPSublet As String
          Dim vPDPSublet As String
          
          Dim vDelIndex As String
          Dim vKodeLgn As String
          
          Dim NilaiGLNonPart As Double
          
          'vKodeLgn = ""
1        On Error GoTo SimpanAASNonPenyeragamanAkun_Error

2         vKeterangan = ""
3         vDelIndex = IIf(opWO(2), "TW", gKode) & Left(McDelIndex(7), 8)
4         vTglSkr = gTanggal
5         If Trim(text(41)) = "" Then
6             vTglTempo = gTanggal
7         Else
8             vTglTempo = DateAdd(LCase(Left(text(41), 1)), Val(Right(text(41), Len(text(41)) - 1)), vTglSkr)
9         End If
          
10        If gJurnalPart Then
              '' JURNAL PART TRANSFER KE ACC
11            vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
12            vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaiDppInv / nilaiTotalInv), 0)
13            vTPenjNonMet = Round(Koma2Titik(text(22)) * (nilaiDppInv / nilaiTotalInv), 0)
14            If nDuoInv Then
15                vTMeterai = 0
16            Else
17                vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))                         ' materai
18            End If
              
19            vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaiDppInv / nilaiTotalInv), 0)  ' jasa non ppn + meterai
20            vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaiDppInv / nilaiTotalInv), 0)
21            vTPPN = vTotal - vTotalNonPPn
              
      '        vTPPN = Round((vTotal * nilaidppINv / nilaiTotalInv0), 0)
      '        vTPenjualan = (vTotal - vTPPN) + vTMeterai
              
22        Else
              '' JURNAL PART TDK TRANSFER KE ACC
23            vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
24            vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaiDppInv / nilaiTotalInv), 0)
25            vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0) - vTPenjSublet
26            If nDuoInv Then
27                vTMeterai = 0
28            Else
29                vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))                         ' materai
30            End If
31            vTPenjualan = vTPenjNonMet + vTPenjSublet + vTMeterai
32            vTPPN = Round(vTotal * (nilaiPPnInv / nilaiTotalInv), 0) ''Round(vTotal * (10 / 110), 0)
33        End If
          
34        If opWO(2) Then
              ' WARRANTY
35            vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[ " & text(43) & " - " & text(2) & " ]"
36            vKodeLgn = gKdPlgHo
37        ElseIf woStat = "SBI" Then
              ' SBI
38            vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[ " & text(43) & " - " & text(2) & " ]"
39            vKodeLgn = Trim(gKdPlgDealer)
40        Else
41            vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[WO: " & text(43) & " ]"
42            vKodeLgn = text(2)
43        End If
          
          ' KODE AKUN
          Dim rsAkun As ADODB.Recordset
44        Set rsAkun = objO.FetchData("Select * FROM T_AKUN WITH(NOLOCK)")
45        If Trim(text(40)) = "INT" Then
46            If Not rsAkun.BOF Then
47                vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_INT"))
48                vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_INT"))
49                vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_RSUBLET_INT"))
50                vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_INT"))
51                vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_INT"))
52                vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_INT"))
53                vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
54                If vPenjSubletGL = "" Then
55                    vTPenjNonMet = vTPenjNonMet + vTPenjSublet
56                    vTPenjSublet = 0
57                End If
58            End If
59        ElseIf Trim(text(40)) = "EXT" Then
60            If Not rsAkun.BOF Then
61                vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_EXT"))
62                vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_EXT"))
63                vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_RSUBLET_EXT"))
64                vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_EXT"))
65                vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_EXT"))
66                vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_EXT"))
67                vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
68                If vPenjSubletGL = "" Then
69                    vTPenjNonMet = vTPenjNonMet + vTPenjSublet
70                    vTPenjSublet = 0
71                End If
72            End If
73        ElseIf Trim(text(40)) = "AFFNONPPN" Then
74            If Not rsAkun.BOF Then
75                vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_AFFNONPPN"))
76                vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_AFFNONPPN"))
77                vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_RSUBLET_AFFNONPPN"))
78                vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_AFFNONPPN"))
79                vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_AFFNONPPN"))
80                vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_AFFNONPPN"))
81                vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
82                vTotal = vTotal - vTPPN
83                vTPPN = 0
84                If vPenjSubletGL = "" Then
85                    vTPenjNonMet = vTPenjNonMet + vTPenjSublet
86                    vTPenjSublet = 0
87                End If
88            End If
89        ElseIf Trim(text(40)) = "AFFPPN" Then
90            If Not rsAkun.BOF Then
91                vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_AFFPPN"))
92                vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_AFFPPN"))
93                vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_RSUBLET_AFFPPN"))
94                vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_AFFPPN"))
95                vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_AFFPPN"))
96                vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_AFFPPN"))
97                vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
98                If vPenjSubletGL = "" Then
99                    vTPenjNonMet = vTPenjNonMet + vTPenjSublet
100                   vTPenjSublet = 0
101               End If
102           End If
103       End If
104       rsAkun.Close
          
          
105               Set rsAkun = objO.FetchData("select * FROM T_AKUN WITH(NOLOCK)")
106                   If Not rsAkun.BOF Then
107                       XPiutangGL_Int = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_INT"))
108                       XPenjualanGL_Int = cekKodeAkun(rsAkun("FS_KD_RLABOUR_INT"))
109                       XPpnGL_Int = cekKodeAkun(rsAkun("FS_KD_RPPN_INT"))
110                       XMeteraiGL_Int = cekKodeAkun(rsAkun("FS_KD_RMETERAI_INT"))
111                       XPiutangGL_Ext = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_EXT"))
112                       XPenjualanGL_Ext = cekKodeAkun(rsAkun("FS_KD_RLABOUR_EXT"))
113                       XPpnGL_Ext = cekKodeAkun(rsAkun("FS_KD_RPPN_EXT"))
114                       XMeteraiGL_Ext = cekKodeAkun(rsAkun("FS_KD_RMETERAI_EXT"))
115                       XPiutangGL_AFFNONPPN = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_AFFNONPPN"))
116                       XPenjualanGL_AFFNONPPN = cekKodeAkun(rsAkun("FS_KD_RLABOUR_AFFNONPPN"))
117                       XPpnGL_AFFNONPPN = cekKodeAkun(rsAkun("FS_KD_RPPN_AFFNONPPN"))
118                       XMeteraiGL_AFFNONPPN = cekKodeAkun(rsAkun("FS_KD_RMETERAI_AFFNONPPN"))
119                       XPiutangGL_AFFPPN = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_AFFPPN"))
120                       XPenjualanGL_AFFPPN = cekKodeAkun(rsAkun("FS_KD_RLABOUR_AFFPPN"))
121                       XPpnGL_AFFPPN = cekKodeAkun(rsAkun("FS_KD_RPPN_AFFPPN"))
122                       XMeteraiGL_AFFPPN = cekKodeAkun(rsAkun("FS_KD_RMETERAI_AFFPPN"))
123                   End If
124               rsAkun.Close


          
125       If vPiutangGL = "" Or vPenjualanGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Or vMeteraiGL = "" Then
126           pesanIT = pesanIT + "Kode GL belum lengkap !"
127           WarnError = True
128           erTranAll = True
129       End If
130       If Trim(text(40)) <> "AFFNONPPN" And vPpnGL = "" Then
131           pesanIT = pesanIT + "Kode GL belum lengkap !"
132           WarnError = True
133           erTranAll = True
134       End If
135       If Trim(vKodeLgn) = "" Then
136           pesanIT = pesanIT + "Kode Pelanggan belum lengkap !"
137           WarnError = True
138           erTranAll = True
139       End If
          
140       If erTranAll = True Then Exit Sub
          
141               sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                        & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                        & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry) values " _
                        & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & "', 'Q'," _
                        & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                        & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                        & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "'," _
                        & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
                        & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
142               Call objHT.InsertData_ACC(sSQL1)
                  
                  'PPN
143               If vTPPN <> 0 Then
144                   sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                            & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                            & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry) values " _
                            & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & "', 'S'," _
                            & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                            & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                            & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "'," _
                            & " " & SimpanKoma2Titik(vTPPN * -1) & ", '" & gSA & "'," _
                            & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
145                   Call objHT.InsertData_ACC(sSQL1)
146               End If
              
                  'Update Armst Saldo berjalan
147               sSQL1 = "Update Armst set SaldoJalan=SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(text(2)) & "'"
148               Call objHT.UpdateData_ACC(sSQL1)
              'End If
                  
          'Membuat transaksi di GL
              'If vTotal + vTMeterai <> 0 Then
149               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                        & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                        & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
150               Call objHT.InsertData_ACC(sSQL1)
151               NilaiGLNonPart = vTotal + vTMeterai
152               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                        & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                        & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
153               Call objHT.InsertData_ACC(sSQL1)
154               NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
155               If vTPenjSublet <> 0 Then
156                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                            & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                            & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                            & " " & SimpanKoma2Titik(vTPenjSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
157                   Call objHT.InsertData_ACC(sSQL1)
158                   NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
159               End If
                  
160               If vTMeterai <> 0 Then
161                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                            & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                            & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                            & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
162                   Call objHT.InsertData_ACC(sSQL1)
163                   NilaiGLNonPart = NilaiGLNonPart - vTMeterai
164               End If
              
165               If vTPPN <> 0 Then
166                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                            & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                            & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                            & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
167                   Call objHT.InsertData_ACC(sSQL1)
168                   NilaiGLNonPart = NilaiGLNonPart - vTPPN
169               End If
                  
170               If gJurnalPart Then
                      '' NILAI SPAREPART
                      Dim rsSearchPart As ADODB.Recordset
171                   vTPenjPart = 0
172                   If rsPart.State <> adStateClosed Then
                      
173                       rsPart.MoveFirst
174                       Do While Not rsPart.EOF
175                           vTPenjPart = vTPenjPart + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
176                           rsPart.MoveNext
177                       Loop
                          
178                           If NilaiGLNonPart <> vTPenjPart Then
179                               If NilaiGLNonPart - vTPenjPart = 1 Or NilaiGLNonPart - vTPenjPart = -1 Then
180                                   vTPenjPart = NilaiGLNonPart
181                               End If
182                           End If
                              
                              ' AMBIL KODE AKUN PENJUALAN PART
183                           Set rsSearchPart = objO.FetchDataACC("SELECT NoJSales, NoRSales FROM SIJURNAL WITH(NOLOCK) WHERE KODEJURNAL = '" & Trim(text(44)) & "'")
184                           If Not rsSearchPart.EOF Then
185                               vPenjPartGL = Trim(rsSearchPart("NoRSales"))
186                           End If
                              
187                           If Trim(vPenjPartGL) = "" Then
188                               pesanIT = pesanIT + "Kode Akun Part belum lengkap !"
189                               WarnError = True
190                               erTranAll = True
191                           End If
                              
      '''                        Set rsSearchPart = objO.FetchDataACC("SELECT * FROM GLTRN WHERE NOBUKTI = '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "' AND KODEGL = '" & Trim(vPenjPartGL) & "' AND DELINDEX = '" & Trim(vDelIndex) & "'")
      '''                        If Not rsSearchPart.EOF Then
      '''                            sSQL1 = "UPDATE GLTRN SET JlhDebit = JlhDebit + " & SimpanKoma2Titik(vTPenjPart) & " WHERE NOBUKTI = '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "' AND KODEGL = '" & Trim(vPenjPartGL) & "' AND DELINDEX = '" & Trim(vDelIndex) & "'"
      '''                            Call objHT.UpdateData_ACC(sSQL1)
      '''                        Else
192                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                        & " " & SimpanKoma2Titik(vTPenjPart) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
193                               Call objHT.InsertData_ACC(sSQL1)
      '''                        End If
                              
194                           sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPenjPart) & ", " _
                                    & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjPart) & " " _
                                    & " where KodeGl = '" & vPenjPartGL & "'"
195                           Call objHT.UpdateData_ACC(sSQL1)
                          
196                   End If
197               End If
                  
                  ' AMBIL HPP SUBLET
                  Dim vTPpnS As String
                  Dim vTotalS As Double
                  Dim vTBeliS As Double
                  Dim nonJurnal As Boolean
                  Dim rsSearch As New ADODB.Recordset
                  Dim vKeterangan1 As String
198               nonJurnal = False
                  
199               vTPenjSublet = 0
200               Set rsSearch = objO.FetchData("SELECT PKP,NAMABARANG,NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, NONJURNAL FROM JOBSUBLET a WITH(NOLOCK) INNER JOIN T_SUPPLIER B WITH(NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(text(43)) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
201               If Not rsSearch.EOF Then
202                   Do While Not rsSearch.EOF
203                       nonJurnal = IIf(Trim(rsSearch("NONJURNAL")) = 1, True, False)
204                       If Not nonJurnal Then
205                           vTotalS = Round(CDbl(rsSearch("HPP")))
206                           vTPpnS = Round(vTotalS * (nilaiPPnInv / nilaiTotalInv))
207                           vTBeliS = Round(vTotalS - vTPpnS)
208                           vKeterangan1 = vKeterangan & " SUBLET" '(" & rsSearch("NAMABARANG") & ")"   'No. WO :" & Trim(Text(0))
209                           If Trim(rsSearch("PKP")) = "1" Then
210                               vPpnSGL = ""
211                               vTPpnS = 0
212                               vTotalS = Round(vTBeliS)
213                           End If
214                           vTPenjSublet = vTPenjSublet + vTotalS
215                       End If
216                       rsSearch.MoveNext
217                   Loop
218               End If
                   
                      
                  'If Not nonJurnal Then
219                   If vTPenjSublet <> 0 Then
220                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vHPPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTPenjSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
221                       Call objHT.InsertData_ACC(sSQL1)
                          
222                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPDPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTPenjSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
223                       Call objHT.InsertData_ACC(sSQL1)
                          
                          ' HPP SUBLET TERHADAP PDP
224                       sSQL1 = "update glmst set JbulanIni = JbulanIni -" & (vTPenjSublet) & ", " _
                                & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjSublet) & " " _
                                & " where KodeGl = '" & vHPPSublet & "'"
225                       Call objHT.InsertData_ACC(sSQL1)
                          
226                       sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjSublet) & ", " _
                                & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjSublet) & " " _
                                & " where KodeGl = '" & vPDPSublet & "'"
227                       Call objHT.InsertData_ACC(sSQL1)
228                   End If
                  'End If
                      
                              
229               sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTotal + vTMeterai) & ", " _
                        & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                        & " where KodeGl = '" & vPiutangGL & "'"
230               Call objHT.UpdateData_ACC(sSQL1)
                  
231               sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                        & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                        & " where KodeGl = '" & vPenjualanGL & "'"
232               Call objHT.UpdateData_ACC(sSQL1)
                  
233               sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjSublet) & ", " _
                        & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjSublet) & " " _
                        & " where KodeGl = '" & vPenjSubletGL & "'"
234               Call objHT.UpdateData_ACC(sSQL1)
                  
235               sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTMeterai) & ", " _
                        & " JDebitBi = JDebitBi +" & vTMeterai & " " _
                        & " where KodeGl = '" & vMeteraiGL & "'"
236               Call objHT.UpdateData_ACC(sSQL1)
                  
237               sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPPN) & ", " _
                        & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPPN) & " " _
                        & " where KodeGl = '" & vPpnGL & "'"
238               Call objHT.UpdateData_ACC(sSQL1)
                      
                  
              'End If
              
              
239           If nDuoInv Then

240               vTotal = subJob
241               vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0)  'dikurangi ppn
242               vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))
243               vTPenjualan = vTPenjNonMet + vTMeterai          'non ppn + meterai
244               vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          ' ppn
                  'vTPenjualan = (vTotal - vTPPN) + vTMeterai
                  
                  
                  ' UNTUK RETUR INVOICE KE 2
245               If vPiutangGL = "" Or vPenjualanGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Or vMeteraiGL = "" Then
246                   pesanIT = pesanIT + "Kode GL belum lengkap !"
247                   WarnError = True
248                   erTranAll = True
249               End If
250               If Trim(vTipeCustomer) <> "AFFNONPPN" And vPpnGL = "" Then
251                   pesanIT = pesanIT + "Kode GL belum lengkap !"
252                   WarnError = True
253                   erTranAll = True
254               End If
255               If Trim(vKodeLgn) = "" Then
256                   pesanIT = pesanIT + "Kode Pelanggan belum lengkap !"
257                   WarnError = True
258                   erTranAll = True
259               End If
                  
                  
      '''''            Set rsAkun = objO.FetchData("select * FROM T_AKUN")
      '''''            If Trim(vTipeCustomer) = "INT" Then
      '''''                If Not rsAkun.BOF Then
      '''''                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_INT"))
      '''''                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_INT"))
      '''''                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_INT"))
      '''''                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_INT"))
      '''''                End If
      '''''            ElseIf Trim(vTipeCustomer) = "EXT" Then
      '''''                If Not rsAkun.BOF Then
      '''''                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_EXT"))
      '''''                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_EXT"))
      '''''                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_EXT"))
      '''''                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_EXT"))
      '''''                End If
      '''''            ElseIf Trim(vTipeCustomer) = "AFFNONPPN" Then
      '''''                If Not rsAkun.BOF Then
      '''''                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_AFFNONPPN"))
      '''''                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_AFFNONPPN"))
      '''''                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_AFFNONPPN"))
      '''''                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_AFFNONPPN"))
      '''''                    vTotal = vTotal - vTPPN
      '''''                    vTPPN = 0
      '''''                End If
      '''''            ElseIf Trim(vTipeCustomer) = "AFFPPN" Then
      '''''                If Not rsAkun.BOF Then
      '''''                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_RPIUTANG_AFFPPN"))
      '''''                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_RLABOUR_AFFPPN"))
      '''''                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_RPPN_AFFPPN"))
      '''''                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_RMETERAI_AFFPPN"))
      '''''                End If
      '''''            End If
      '''''            rsAkun.Close
                  
                  
          
260               If Trim(vTipeCustomer) = "INT" Then
                      'If Not rsAkun.BOF Then
261                       vPiutangGL = XPiutangGL_Int
262                       vPenjualanGL = XPenjualanGL_Int
263                       vPpnGL = XPpnGL_Int
264                       vMeteraiGL = XMeteraiGL_Int
                      'End If
265               ElseIf Trim(vTipeCustomer) = "EXT" Then
                      'If Not rsAkun.BOF Then
266                       vPiutangGL = XPiutangGL_Ext
267                       vPenjualanGL = XPenjualanGL_Ext
268                       vPpnGL = XPpnGL_Ext
269                       vMeteraiGL = XMeteraiGL_Ext
                      'End If
270               ElseIf Trim(vTipeCustomer) = "AFFNONPPN" Then
                      'If Not rsAkun.BOF Then
271                       vPiutangGL = XPiutangGL_AFFNONPPN
272                       vPenjualanGL = XPenjualanGL_AFFNONPPN
273                       vPpnGL = XPpnGL_AFFNONPPN
274                       vMeteraiGL = XMeteraiGL_AFFNONPPN
275                       vTotal = vTotal - vTPPN
276                       vTPPN = 0
                      'End If
277               ElseIf Trim(vTipeCustomer) = "AFFPPN" Then
                      'If Not rsAkun.BOF Then
278                       vPiutangGL = XPiutangGL_AFFPPN
279                       vPenjualanGL = XPenjualanGL_AFFPPN
280                       vPpnGL = XPpnGL_AFFPPN
281                       vMeteraiGL = XMeteraiGL_AFFPPN
                      'End If
282               End If


                  ' BUAT ARTRN
283                   sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                            & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                            & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry) values " _
                            & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & "', 'Q'," _
                            & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                            & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                            & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "'," _
                            & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
                            & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
284                   Call objHT.InsertData_ACC(sSQL1)
                      
                      'PPN
285                   If vTPPN <> 0 Then
286                       sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                                & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                                & " JumlahTrn, Kodeuser, DelIndex, Kodesumber,nowo, nopolisi, tglEntry) values " _
                                & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & "', 'S'," _
                                & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                                & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                                & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "'," _
                                & " " & SimpanKoma2Titik(vTPPN * -1) & ", '" & gSA & "'," _
                                & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
287                       Call objHT.InsertData_ACC(sSQL1)
288                   End If
                  
                      'Update Armst Saldo berjalan
289                   sSQL1 = "Update Armst set SaldoJalan=SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(text(2)) & "'"
290                   Call objHT.UpdateData_ACC(sSQL1)
                  'End If
                  
                  'Membuat transaksi di GL
291                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                            & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                            & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                            & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
292                   Call objHT.InsertData_ACC(sSQL1)
                      
293                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                            & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                            & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                            & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
294                   Call objHT.InsertData_ACC(sSQL1)
                      
295                   If vTMeterai <> 0 Then
296                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
297                       Call objHT.InsertData_ACC(sSQL1)
298                   End If
                      
299                   If vTPPN <> 0 Then
300                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
301                       Call objHT.InsertData_ACC(sSQL1)
302                   End If
                  
303                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & (vTotal + vTMeterai) & ", " _
                            & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                            & " where KodeGl = '" & vPiutangGL & "'"
304                   Call objHT.UpdateData_ACC(sSQL1)
                      
305                   sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                            & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                            & " where KodeGl = '" & vPenjualanGL & "'"
306                   Call objHT.UpdateData_ACC(sSQL1)
                      
307                   sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTMeterai) & ", " _
                            & " JDebitBi = JDebitBi +" & vTMeterai & " " _
                            & " where KodeGl = '" & vMeteraiGL & "'"
308                   Call objHT.UpdateData_ACC(sSQL1)
                      
309                   sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPPN) & ", " _
                            & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPPN) & " " _
                            & " where KodeGl = '" & vPpnGL & "'"
310                   Call objHT.UpdateData_ACC(sSQL1)
                  'End If
311           End If
312       Exit Sub

313      'On Error GoTo 0
314      Exit Sub

SimpanAASNonPenyeragamanAkun_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
315       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure SimpanAASNonPenyeragamanAkun of Form frmBtlInv")
End Sub

Private Sub SimpanAAS2()
    Dim SsqlNis As String
    Dim isSudahDikurangi As Boolean
   On Error GoTo SimpanAAS2_Error

    isSudahDikurangi = False
  
    Dim vSyaratByr As String
    'Dim vJmlHari As Double
    Dim vTglSkr As Date
    Dim vTglTempo As Date
    Dim vTPenjualan As Double
    Dim vTPenjSublet As Double
    Dim vTPenjNonMet As Double
    Dim vTPPN As Double
    
    Dim vTPembSublet As Double
    
    Dim vJSPPN As Double                ' NILAI PPN JASA + SUBLET
    Dim vPaPPN As Double                ' NILAI PPN Part
    Dim vOlPPN As Double                ' NILAI PPN Oli
    Dim vAcPPN As Double                ' NILAI PPN Accs
    
    Dim vTotal As Double
    Dim vTotalNonPPn As Double
    Dim vTMeterai As Double
    Dim vKeterangan As String
    
    Dim vPiutangGL As String
    Dim vPenjualanGL As String
    Dim vPenjSubletGL As String
    Dim vPpnGL As String
    Dim vMeteraiGL As String
    Dim vHPPSublet As String
    Dim vPDPSublet As String
    
    Dim vDelIndex As String
    Dim vKodeLgn As String
    
    Dim NilaiGLNonPart As Double
    
    Dim totalKreditTemp As Double
    Dim totalDebitTemp As Double
    
    Dim adaSelisih As Boolean
    Dim kreditSelisih As Boolean
    
    kreditSelisih = True
    adaSelisih = False
    
    vKodeLgn = ""
    vKeterangan = ""
    vDelIndex = IIf(opWO(2), "TW", gKode) & Left(McDelIndex(7), 8)
    vTglSkr = gTanggal
    If Trim(text(41)) = "" Then
        vTglTempo = gTanggal
    Else
        vTglTempo = DateAdd(LCase(Left(text(41), 1)), Val(Right(text(41), Len(text(41)) - 1)), vTglSkr)
    End If
    
    
   
    
    If gJurnalPart Then
        '' JURNAL PART TRANSFER KE ACC
        vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
        vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaiDppInv / nilaiTotalInv), 0)
        vTPenjNonMet = Round(Koma2Titik(text(22)) * (nilaiDppInv / nilaiTotalInv), 0)
        If Log.fakturMaterai = 1 Then
            vTMeterai = Koma2Titik(text(42))
        Else
            If nDuoInv Then
               Meterai = 0
            Else
               vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))                         ' materai
            End If
        End If
      
  
      vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaiDppInv / nilaiTotalInv), 0)  ' jasa non ppn + meterai
      vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaiDppInv / nilaiTotalInv), 0)
      vTPPN = vTotal - vTotalNonPPn
  '' REVISI THANIA 09/05/2014
      vJSPPN = vTotal - Round(((Koma2Titik(text(24)) + Koma2Titik(text(25)))), 0) - vTPenjSublet - vTPenjNonMet
  
'        vTPPN = Round((vTotal * 10 / 110), 0)
'        vTPenjualan = (vTotal - vTPPN) + vTMeterai
  
    Else
  '' JURNAL PART TDK TRANSFER KE ACC
      vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
      vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaiDppInv / nilaiTotalInv), 0)
      vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0) - vTPenjSublet
      
      If Log.fakturMaterai = 1 Then
            vTMeterai = Koma2Titik(text(42))
      Else
        If nDuoInv Then
            vTMeterai = 0
        Else
            vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))                         ' materai
        End If
      End If
     
      vTPenjualan = vTPenjNonMet + vTPenjSublet + vTMeterai
      vTPPN = Round(vTotal * (nilaiPPnInv / nilaiTotalInv), 0)
    End If
    
    If opWO(2) Then
          ' WARRANTY
          vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[ " & text(43) & " - " & text(2) & " ]"
          vKodeLgn = gKdPlgHo
    ElseIf woStat = "SBI" Then
          ' SBI
          vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[ " & text(43) & " - " & text(2) & " ]"
          vKodeLgn = Trim(gKdPlgDealer)
    Else
          vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[WO: " & text(43) & " ]"
          vKodeLgn = text(2)
    End If

    ' KODE AKUN
    
      Dim isDyna As Boolean
      isDyna = IIf(InStr(1, UCase(text(14).text), "DYNA") > 0, True, False)
      Dim isCust As String
      Dim isExAflAR As Boolean
      Dim isExAflSI As Boolean
      Dim isExIntSI As Boolean
      Dim nKodeSi As String
  
      nKodeSi = Trim(vKodeLgn)
      isCust = Trim(text(40))
      isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(nKodeSi)  'cex dari ARMST apakah AFL / not AFL
      If isExAflAR = True Then  'AFL
          isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
          If isExAflSI = True Then
              isCust = "AFFPPN"
          Else
              MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
              GoTo vbError
          End If
      Else
          If isCust = "AFFPPN" Then  'cex jika tidak sesuai antara ARMST dan M_Asuransi
              MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
              GoTo vbError
          End If
          If isCust = "INT" Then
              isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah INT terdaftar
              If isExIntSI = False Then
                  MsgBox "Data Customer '" & nKodeSi & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
                  GoTo vbError
              End If
          End If
      End If
  
  'AKUNBARU
  Dim rsGL As ADODB.Recordset
  Select Case isCust
      Case "EXT"
          Set rsGL = objO.Get_AkunGL_NEW("OTHERS", isDyna, True)
      Case Else  'INT/AFL
          Set rsGL = objO.Get_AkunGL_NEW(nKodeSi, isDyna, True)
  End Select
      
'''piutang gr
'''  penj bengkel
'''  penj sublet
'''  penj Part
'''  penj oli
'''  penj Accs
'''  PPn Jasa
'''  PPn Part + oli
'''  PPn Accs
'''  Materai
'''HPP sublet
''' Persediaan sublet
    
    If Not rsGL.BOF Then
      vPiutangGL = cekKodeAkun(rsGL.Fields(0).value)
      vPenjualanGL = cekKodeAkun(rsGL.Fields(1).value)
      vPenjSubletGL = cekKodeAkun(rsGL.Fields(2).value)
      vPenjPartGL = cekKodeAkun(rsGL.Fields(3).value)
      vPenjOlieGL = cekKodeAkun(rsGL.Fields(4).value)
      vPenjAccsGL = cekKodeAkun(rsGL.Fields(5).value)
      vPpnGL = cekKodeAkun(rsGL.Fields(6).value)
      vPpnPartGL = cekKodeAkun(rsGL.Fields(7).value)
      vPpnAccsGL = cekKodeAkun(rsGL.Fields(8).value)
      vMeteraiGL = cekKodeAkun(rsGL.Fields(9).value)
      vHPPSublet = cekKodeAkun(rsGL.Fields(10).value)
      vPDPSublet = cekKodeAkun(rsGL.Fields(11).value)
''                If Trim(vPenjSubletGL) = "" Then
''                    vTPenjNonMet = vTPenjNonMet + vTPenjSublet
''                    vTPenjSublet = 0
''                End If
    End If

    If rsGL.State <> adStateClosed Then rsGL.Close



    If vPiutangGL = "" Or vPenjualanGL = "" Or vPenjSubletGL = "" Or vPenjPartGL = "" Or vPenjOlieGL = "" Or vPenjAccsGL = "" _
 Or vPpnGL = "" Or vPpnPartGL = "" Or vPpnAccsGL = "" Or vMeteraiGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Then
      pesanIT = pesanIT + "Kode GL belum lengkap !"
      WarnError = True
      erTranAll = True
    End If
    
    If Trim(text(40)) <> "AFFNONPPN" And vPpnGL = "" Then
      pesanIT = pesanIT + "Kode GL belum lengkap !"
      WarnError = True
      erTranAll = True
    End If
    If Trim(vKodeLgn) = "" Then
      pesanIT = pesanIT + "Kode Pelanggan belum lengkap !"
      WarnError = True
      erTranAll = True
    End If
    
    '-------------------- PROSES PENGECEKAN SELISIH --------------------
    totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTotal + vTMeterai))
    NilaiGLNonPart = vTotal + vTMeterai
    
    totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjNonMet))
    NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
    
    If vTPenjSublet <> 0 Then
      totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjSublet))
      NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
    End If
    
    If vTMeterai <> 0 Then
      totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTMeterai))
      NilaiGLNonPart = NilaiGLNonPart - vTMeterai
    End If

    If (vTPenjNonMet + vTPenjSublet) <> 0 Then
      totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)))
      NilaiGLNonPart = NilaiGLNonPart - (Round(((vTPenjNonMet + vTPenjSublet) / 10), 0))
    End If
    
    If gJurnalPart Then
  Dim SumPartTemp As Double
  Dim SumOlieTemp As Double
  Dim SumAccsTemp As Double
  Dim SumPartFullTemp As Double
  Dim SumOlieFullTemp As Double
  Dim SumAccsFullTemp As Double
  '' NILAI SPAREPART
  vTPenjPart = 0
  Dim rsSearchPartTemp As ADODB.Recordset
  
  If rsPart.State <> adStateClosed Then
      SumPartTemp = 0
      SumOlieTemp = 0
      SumAccsTemp = 0
      SumPartFullTemp = 0
      SumOlieFullTemp = 0
      SumAccsFullTemp = 0
      vPaPPN = 0
      vOlPPN = 0
      vAcPPN = 0
      rsPart.MoveFirst
      Do Until rsPart.EOF
          Select Case Left(UCase(objO.FetchData_String_Acc("select B.Jenis from Inmst A WITH(NOLOCK) Inner Join Profilacc..SCC B WITH(NOLOCK) On A.SCC_Code = B.Scc_Code Where A.KodeItem = '" & Trim(rsPart("KodeItem")) & "'")), 3)
              Case "PAR"
                  SumPartTemp = SumPartTemp + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
                  SumPartFullTemp = SumPartFullTemp + CDbl(rsPart("HARGA"))
              Case "OLI"
                  SumOlieTemp = SumOlieTemp + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
                  SumOlieFullTemp = SumOlieFullTemp + CDbl(rsPart("HARGA"))
              Case "ACC"
                  SumAccsTemp = SumAccsTemp + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
                  SumAccsFullTemp = SumAccsFullTemp + CDbl(rsPart("HARGA"))
              End Select
          rsPart.MoveNext
      Loop
    vTPenjPart = SumPartTemp + SumOlieTemp + SumAccsTemp
    
      If NilaiGLNonPart <> vTPenjPart Then
          If NilaiGLNonPart - vTPenjPart = 1 Then
              If SumPartTemp <> 0 Then
                  SumPartTemp = SumPartTemp + 1
              Else
                  If SumOlieTemp <> 0 Then
                      SumOlieTemp = SumOlieTemp + 1
                  Else
                      If SumAccsTemp <> 0 Then
                          SumAccsTemp = SumAccsTemp + 1
                      End If
                  End If
              End If
              vTPenjPart = NilaiGLNonPart
          ElseIf NilaiGLNonPart - vTPenjPart = -1 Then
              If SumPartTemp <> 0 Then
                  SumPartTemp = SumPartTemp - 1
              Else
                  If SumOlieTemp <> 0 Then
                      SumOlieTemp = SumOlieTemp - 1
                  Else
                      If SumAccsTemp <> 0 Then
                          SumAccsTemp = SumAccsTemp - 1
                      End If
                  End If
              End If
              vTPenjPart = NilaiGLNonPart
          End If
      End If
    
      vPaPPN = SumPartFullTemp - SumPartTemp
      vOlPPN = SumOlieFullTemp - SumOlieTemp
      vAcPPN = SumAccsFullTemp - SumAccsTemp
    
      If SumPartTemp <> 0 Then
          totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(SumPartTemp))
      End If
      
      If SumOlieTemp <> 0 Then
          totalDebitTemp = totalDebitTemp + SimpanKoma2Titik(SumOlieTemp)
      End If
      
      If (SumPartTemp + SumOlieTemp) <> 0 Then
          totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0)))
      End If
      
      If SumAccsTemp <> 0 Then
          totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(SumAccsTemp))
          
          totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(Round((vAcPPN), 0)))
      End If
  End If
    End If
    
    Dim vTPpnSTemp As String
    Dim vTotalSTemp As Double
    Dim vTBeliSTemp As Double
    Dim nonJurnalTemp As Boolean
    Dim rsSearchTemp As New ADODB.Recordset
    Dim vKeterangan1Temp As String
    
    nonJurnalTemp = False
          
    vTPembSublet = 0
    vTPenjSublet = 0
    Set rsSearchTemp = objO.FetchData("SELECT PKP,NAMABARANG,NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, NONJURNAL FROM JOBSUBLET a WITH(NOLOCK) INNER JOIN T_SUPPLIER B WITH(NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(text(43)) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
    If Not rsSearchTemp.EOF Then
        Do While Not rsSearchTemp.EOF
            nonJurnalTemp = IIf(Trim(rsSearchTemp("NONJURNAL")) = 1, True, False)
            If Not nonJurnalTemp Then
                vTotalSTemp = Round(CDbl(rsSearchTemp("HPP")))
                vTPpnSTemp = Round(vTotalSTemp * (nilaiPPnInv / nilaiTotalInv))
                vTBeliSTemp = Round(vTotalSTemp - vTPpnSTemp)
               ' vKeterangan1Temp = vKeterangan & " SUBLET" '(" & rsSearch("NAMABARANG") & ")"   'No. WO :" & Trim(Text(0))
                If Trim(rsSearchTemp("PKP")) = "1" Then
                    vPpnSGL = ""
                    vTPpnSTemp = 0
                    vTotalSTemp = Round(vTotalSTemp)
                End If
                vTPenjSublet = vTPenjSublet + vTotalSTemp
            End If
            rsSearchTemp.MoveNext
        Loop
    End If
    
    If vTPenjSublet <> 0 Then
      totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjSublet))
      totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTPenjSublet))
    End If
    
    If nDuoInv Then
        ' UNTUK INVOICE KEDUA -B KE DEALER PENJUAL
        ' MEMBUAT TRANSAKSI DUO INV DI AR
        vTotal = subJob
        vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0)  'dikurangi ppn
        vTMeterai = Koma2Titik(metInvB) ''Koma2Titik(text(42))               'materai
        vTPenjualan = vTPenjNonMet + vTMeterai         'non ppn + meterai
        ''2015-10-19 || AKTIF AFFPPN
        If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
            If Trim(gNpwp) = Trim(gInvNPWP2) Then
                vTotal = vTPenjNonMet
                vTPPN = 0     'Round(vTotal * (10 / 110), 0)
            Else
                vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
            End If
        Else
            vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
        End If
        'vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          'ppn
        'vTPenjualan = (vTotal - vTPPN) + vTMeterai
        vKeterangan = gInsCust2 & ". AN/ " & vKeterangan
        
        ' MEMBUAT TRANSAKSI DUO INV DI GL
        
        totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTotal + vTMeterai))
        
        totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjNonMet))
        
        If vTMeterai <> 0 Then
            totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTMeterai))
        End If
        
        If vTPPN <> 0 Then
            totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPPN))
        End If
    End If
    
    If totalDebitTemp <> totalKreditTemp Then
      If (totalDebitTemp - totalKreditTemp) = -1 Or (totalDebitTemp - totalKreditTemp) = 1 Then
          If (totalDebitTemp - totalKreditTemp) = -1 Then
              kreditSelisih = False
          End If
          adaSelisih = True
      End If
    End If
    
    '-------------------- AKHIR PENGECEKAN SELISIH ---------------------
    
    If gJurnalPart Then
      '' JURNAL PART TRANSFER KE ACC
      vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
      vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaiDppInv / nilaiTotalInv), 0)
      vTPenjNonMet = Round(Koma2Titik(text(22)) * (nilaiDppInv / nilaiTotalInv), 0)
      If Log.fakturMaterai = 1 Then
        vTMeterai = Koma2Titik(text(42))
      Else
        If nDuoInv Then
            vTMeterai = 0
        Else
            vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42)) ' materai
        End If
      End If
        
  
  
      vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaiDppInv / nilaiTotalInv), 0)  ' jasa non ppn + meterai
      vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaiDppInv / nilaiTotalInv), 0)
      vTPPN = vTotal - vTotalNonPPn
      '' REVISI THANIA 09/05/2014
      vJSPPN = vTotal - Round(((Koma2Titik(text(24)) + Koma2Titik(text(25)))), 0) - vTPenjSublet - vTPenjNonMet
  
'        vTPPN = Round((vTotal * 10 / 110), 0)
'        vTPenjualan = (vTotal - vTPPN) + vTMeterai
  
    Else
  '' JURNAL PART TDK TRANSFER KE ACC
      vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
      vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaiDppInv / nilaiTotalInv), 0)
      vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0) - vTPenjSublet
    If Log.fakturMaterai = 1 Then
        vTMeterai = Koma2Titik(text(42))
    Else
        If nDuoInv Then
            vTMeterai = 0
        Else
            vTMeterai = Koma2Titik(metInvB)  '' Koma2Titik(Text(42))                         ' materai
        End If
    End If
  
      vTPenjualan = vTPenjNonMet + vTPenjSublet + vTMeterai
      vTPPN = Round(vTotal * (nilaiPPnInv / nilaiTotalInv), 0)
    End If
    
    If erTranAll = True Then GoTo vbError ' Exit Sub
    
      sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
            & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
            & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
            & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & "', 'Q'," _
            & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
            & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
            & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "'," _
            & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
            & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
      Call objHT.InsertData_ACC(sSQL1)
      
      'PPN
      If vTPPN <> 0 Then
          sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & "', 'S'," _
                & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "'," _
                & " " & SimpanKoma2Titik(Round(vTPPN, 0) * -1) & ", '" & gSA & "'," _
                & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
          Call objHT.InsertData_ACC(sSQL1)
      End If
  
      'Update Armst Saldo berjalan
      sSQL1 = "Update Armst set SaldoJalan=SaldoJalan + " & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(text(2)) & "'"
      Call objHT.UpdateData_ACC(sSQL1)
  'End If
      
    'Membuat transaksi di GL
  'If vTotal + vTMeterai <> 0 Then
      'Piutang
      sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
            & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
            & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
            & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
      Call objHT.InsertData_ACC(sSQL1)
      NilaiGLNonPart = vTotal + vTMeterai
      'Penj Jasa
      sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
            & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
            & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
            & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
      Call objHT.InsertData_ACC(sSQL1)
      NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
      'Penj Sublet
      If vTPenjSublet <> 0 Then
          sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik(vTPenjSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
          Call objHT.InsertData_ACC(sSQL1)
          NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
      End If
      'Materai
      If vTMeterai <> 0 Then
          sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
          Call objHT.InsertData_ACC(sSQL1)
          NilaiGLNonPart = NilaiGLNonPart - vTMeterai
      End If
      'PPn Jasa + Sublet
       If (vTPenjNonMet + vTPenjSublet) <> 0 Then
'                sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
'                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
'                      & " ('" & Trim(gKodeFaktur) & Trim(Text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
'                      & " " & SimpanKoma2Titik(Round(((vTPenjNonMet + vTPenjSublet) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
'                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                
          ' REVISI THANIA 09/05/2014
           'sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik(vJSPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                
          ' REVISI ACEP 10/06/2014
           sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
          
          If adaSelisih Then
              If kreditSelisih Then
                  sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              Else
                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              End If
              isSudahDikurangi = True
          End If
          Call objHT.InsertData_ACC(sSQL1)
          NilaiGLNonPart = NilaiGLNonPart - (Round(((vTPenjNonMet + vTPenjSublet) / 10), 0))
      End If
      
      If gJurnalPart Then
          '' NILAI SPAREPART
              Dim SumPart As Double
              Dim SumOlie As Double
              Dim SumAccs As Double
              Dim SumPartFull As Double
              Dim SumOlieFull As Double
              Dim SumAccsFull As Double
              '' NILAI SPAREPART
              vTPenjPart = 0
              Dim rsSearchPart As ADODB.Recordset
              If rsPart.State <> adStateClosed Then
                  SumPart = 0
                  SumOlie = 0
                  SumAccs = 0
                  SumPartFull = 0
                  SumOlieFull = 0
                  SumAccsFull = 0
                  vPaPPN = 0
                  vOlPPN = 0
                  vAcPPN = 0
                  rsPart.MoveFirst
                  Do Until rsPart.EOF
                      Select Case Left(UCase(objO.FetchData_String_Acc("select B.Jenis from Inmst A WITH(NOLOCK) Inner Join Profilacc..SCC B WITH(NOLOCK) On A.SCC_Code = B.Scc_Code Where A.KodeItem = '" & Trim(rsPart("KodeItem")) & "'")), 3)
                          Case "PAR"
                              SumPart = SumPart + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
                              SumPartFull = SumPartFull + CDbl(rsPart("HARGA"))
                          Case "OLI"
                              SumOlie = SumOlie + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
                              SumOlieFull = SumOlieFull + CDbl(rsPart("HARGA"))
                          Case "ACC"
                              SumAccs = SumAccs + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
                              SumAccsFull = SumAccsFull + CDbl(rsPart("HARGA"))
                      End Select
                      rsPart.MoveNext
                  Loop
                  vTPenjPart = SumPart + SumOlie + SumAccs
                  
                  If NilaiGLNonPart <> vTPenjPart Then
                      If NilaiGLNonPart - vTPenjPart = 1 Then
                          If SumPart <> 0 Then
                              SumPart = SumPart + 1
                          Else
                              If SumOlie <> 0 Then
                                  SumOlie = SumOlie + 1
                              Else
                                  If SumAccs <> 0 Then
                                      SumAccs = SumAccs + 1
                                  End If
                              End If
                          End If
                          vTPenjPart = NilaiGLNonPart
                      ElseIf NilaiGLNonPart - vTPenjPart = -1 Then
                          If SumPart <> 0 Then
                              SumPart = SumPart - 1
                          Else
                              If SumOlie <> 0 Then
                                  SumOlie = SumOlie - 1
                              Else
                                  If SumAccs <> 0 Then
                                      SumAccs = SumAccs - 1
                                  End If
                              End If
                          End If
                          vTPenjPart = NilaiGLNonPart
                      End If
                  End If
          
          
                  vPaPPN = SumPartFull - SumPart
                  vOlPPN = SumOlieFull - SumOlie
                  vAcPPN = SumAccsFull - SumAccs
          
'''
              If SumPart <> 0 Then
                  sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                        & " " & SimpanKoma2Titik(SumPart) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                  sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumPart) & ", " _
                        & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumPart) & " " _
                        & " where KodeGl = '" & vPenjPartGL & "'"
                        Call objHT.UpdateData_ACC(sSQL1)
              End If
              If SumOlie <> 0 Then
                  sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjOlieGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                        & " " & SimpanKoma2Titik(SumOlie) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                  sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumOlie) & ", " _
                        & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumOlie) & " " _
                        & " where KodeGl = '" & vPenjOlieGL & "'"
                        Call objHT.UpdateData_ACC(sSQL1)
              End If
              
                      If (SumPart + SumOlie) <> 0 Then
'                                sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
'                                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
'                                      & " ('" & Trim(gKodeFaktur) & Trim(Text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
'                                      & " " & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
'                                      & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(Text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
'                                      Call objHT.InsertData_ACC(sSQL1)
'                                sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
'                                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & " " _
'                                      & " where KodeGl = '" & vPpnPartGL & "'"
'                                      Call objHT.InsertData_ACC(sSQL1)

                          sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                              If adaSelisih Then
                                  If isSudahDikurangi = False Then
                                      If kreditSelisih Then
                                          sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                          & " " & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) + 1) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                          & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                                      Else
                                          sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                          & " " & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) - 1) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                          & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                                      End If
                                  End If
                              End If
                                  
                          Call objHT.InsertData_ACC(sSQL1)
                          'sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _

                          'sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(((SumPart + SumOlie) / 10)) & ", " _

                          sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
                                & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0)) & " " _
                                & " where KodeGl = '" & vPpnPartGL & "'"
                          
                          If adaSelisih Then
                              If isSudahDikurangi = False Then
                                  If kreditSelisih Then
                                      sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0) + 1) & ", " _
                                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) + 1) & " " _
                                      & " where KodeGl = '" & vPpnPartGL & "'"
                                  Else
                                      sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0) - 1) & ", " _
                                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) - 1) & " " _
                                      & " where KodeGl = '" & vPpnPartGL & "'"
                                  End If
                              End If
                          End If
                              
                          Call objHT.InsertData_ACC(sSQL1)
                      End If
              
              If SumAccs <> 0 Then
                  sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                        & " " & SimpanKoma2Titik(SumAccs) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                  sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumAccs) & ", " _
                        & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumAccs) & " " _
                        & " where KodeGl = '" & vPenjAccsGL & "'"
                        Call objHT.UpdateData_ACC(sSQL1)
                        
'                        sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
'                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
'                              & " ('" & Trim(gKodeFaktur) & Trim(Text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
'                              & " " & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
'                              & " '" & vDelIndex & "','" & vKeterangan & "- PPn Accs" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(Text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
'                              Call objHT.InsertData_ACC(sSQL1)
'                        sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", " _
'                              & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & " " _
'                              & " where KodeGl = '" & vPpnAccsGL & "'"
'                              Call objHT.InsertData_ACC(sSQL1)
                        
                  sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                        & " " & SimpanKoma2Titik(Round(vAcPPN, 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                        & " '" & vDelIndex & "','" & vKeterangan & "- PPn Accs" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                  sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", " _
                        & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(Round(vAcPPN, 0)) & " " _
                        & " where KodeGl = '" & vPpnAccsGL & "'"
                        Call objHT.InsertData_ACC(sSQL1)

              End If
              
          End If
      End If
      
      ' AMBIL HPP SUBLET
      Dim vTPpnS As String
      Dim vTotalS As Double
      Dim vTBeliS As Double
      Dim nonJurnal As Boolean
      Dim rsSearch As New ADODB.Recordset
      Dim vKeterangan1 As String
      
      nonJurnal = False
      vTPembSublet = 0
     ' vTPenjSublet = 0
      Set rsSearch = objO.FetchData("SELECT PKP,NAMABARANG,NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, NONJURNAL FROM JOBSUBLET a WITH(NOLOCK) INNER JOIN T_SUPPLIER B WITH(NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(text(43)) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
      If Not rsSearch.EOF Then
          Do While Not rsSearch.EOF
                                  
              nonJurnal = IIf(Trim(rsSearch("NONJURNAL")) = 1, True, False)
              If Not nonJurnal Then
                  vTotalS = Round(CDbl(rsSearch("HPP")))
                  vTPpnS = Round(vTotalS * (nilaiPPnInv / nilaiTotalInv), 0)
                  vTBeliS = Round(vTotalS - vTPpnS, 0)
                  vKeterangan1 = vKeterangan & " SUBLET" '(" & rsSearch("NAMABARANG") & ")"   'No. WO :" & Trim(Text(0))
                  If Trim(rsSearch("PKP")) = "1" Then
                      vPpnSGL = ""
                      vTPpnS = 0
                      vTotalS = Round(vTBeliS, 0)
                  End If
                  'vTPenjSublet = vTPenjSublet + vTotalS
                  vTPembSublet = vTPembSublet + vTotalS
              End If
              rsSearch.MoveNext
          Loop
      End If
      
          
      'If Not nonJurnal Then
          If vTPembSublet <> 0 Then
              sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                    & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                    & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vHPPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                    & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                    & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              Call objHT.InsertData_ACC(sSQL1)
              
              sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                    & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                    & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPDPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                    & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                    & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              Call objHT.InsertData_ACC(sSQL1)
              
              ' HPP SUBLET TERHADAP PDP
              sSQL1 = "update glmst set JbulanIni = JbulanIni -" & (vTPembSublet) & ", " _
                    & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                    & " where KodeGl = '" & vHPPSublet & "'"
              Call objHT.InsertData_ACC(sSQL1)
              
              sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPembSublet) & ", " _
                    & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                    & " where KodeGl = '" & vPDPSublet & "'"
              Call objHT.InsertData_ACC(sSQL1)
          End If
      'End If
          
                  
      sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTotal + vTMeterai) & ", " _
            & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
            & " where KodeGl = '" & vPiutangGL & "'"
      Call objHT.UpdateData_ACC(sSQL1)
      
      sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
            & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
            & " where KodeGl = '" & vPenjualanGL & "'"
      Call objHT.UpdateData_ACC(sSQL1)
      
      sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjSublet) & ", " _
            & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjSublet) & " " _
            & " where KodeGl = '" & vPenjSubletGL & "'"
      Call objHT.UpdateData_ACC(sSQL1)
      
      sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTMeterai) & ", " _
            & " JDebitBi = JDebitBi +" & vTMeterai & " " _
            & " where KodeGl = '" & vMeteraiGL & "'"
      Call objHT.UpdateData_ACC(sSQL1)
      
'            sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPPN) & ", " _
'                  & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPPN) & " " _
'                  & " where KodeGl = '" & vPpnGL & "'"
'            Call objHT.UpdateData_ACC(sSQL1)



  sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", " _
  & " JDebitBi = JDebitBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & " " _
  & " where KodeGl = '" & vPpnGL & "'"
  
  If adaSelisih Then
      If kreditSelisih Then
          sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & ", " _
          & " JDebitBi = JDebitBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & " " _
          & " where KodeGl = '" & vPpnGL & "'"
      Else
          sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & ", " _
          & " JDebitBi = JDebitBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & " " _
          & " where KodeGl = '" & vPpnGL & "'"
      End If
  End If
  
  Call objHT.InsertData_ACC(sSQL1)
          
      
  'End If
  
  
  If nDuoInv Then
     '' INV B
      vTotal = subJob
      vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0)  'dikurangi ppn
      vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))
      vTPenjualan = vTPenjNonMet + vTMeterai          'non ppn + meterai
      ''2015-10-19 || AKTIF AFFPPN
      If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
          If Trim(gNpwp) = Trim(gInvNPWP2) Then
              vTotal = vTPenjNonMet
              vTPPN = 0     'Round(vTotal * (10 / 110), 0)
          Else
              vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
          End If
      Else
          vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
      End If
      'vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          ' ppn
      'vTPenjualan = (vTotal - vTPPN) + vTMeterai
      
      
      ' UNTUK RETUR INVOICE KE 2
      If vPiutangGL = "" Or vPenjualanGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Or vMeteraiGL = "" Then
          pesanIT = pesanIT + "Kode GL belum lengkap !"
          WarnError = True
          erTranAll = True
      End If
      If Trim(vTipeCustomer) <> "AFFNONPPN" And vPpnGL = "" Then
          pesanIT = pesanIT + "Kode GL belum lengkap !"
          WarnError = True
          erTranAll = True
      End If
      If Trim(vKodeLgn) = "" Then
          pesanIT = pesanIT + "Kode Pelanggan belum lengkap !"
          WarnError = True
          erTranAll = True
      End If
      
  
          If nDuoInv = True Then
              nKodeSi = Trim(gKdPlgDealer2)
              isCust = Trim(text(40))
              isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(nKodeSi)  'cex dari ARMST apakah AFL / not AFL
              If isExAflAR = True Then  'AFL
                  isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
                  If isExAflSI = True Then
                      isCust = "AFFPPN"
                  Else
                      MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
                      GoTo vbError
                  End If
              Else
                  If isCust = "AFFPPN" Then  'cex jika tidak sesuai antara ARMST dan M_Asuransi
                      MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
                      GoTo vbError
                  End If
                  If isCust = "INT" Then
                      isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah INT terdaftar
                      If isExIntSI = False Then
                          MsgBox "Data Customer '" & nKodeSi & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
                          GoTo vbError
                      End If
                  End If
              End If
  
              Dim rsGL2 As ADODB.Recordset
              Select Case isCust
                  Case "EXT"
                      Set rsGL2 = objO.Get_AkunGL_NEW("OTHERS", isDyna, False)
                  Case Else  'INT/AFL
                      Set rsGL2 = objO.Get_AkunGL_NEW(nKodeSi, isDyna, False)
              End Select

              If Not rsGL2.BOF Then
                  'vPiutangGL = cekKodeAkun(rsGL2.Fields(0).value)
                  'vPenjualanGL = cekKodeAkun(rsGL2.Fields(1).value)
                  'vPpnGL = cekKodeAkun(rsGL2.Fields(6).value)
                  'vMeteraiGL = cekKodeAkun(rsGL2.Fields(9).value)
                  
                  vPiutangGL = (rsGL2.Fields(0).value)
                  vPenjualanGL = (rsGL2.Fields(1).value)
                  vPpnGL = (rsGL2.Fields(6).value)
                  vMeteraiGL = (rsGL2.Fields(9).value)
              End If
          End If
          

      ' BUAT ARTRN
          sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & "', 'Q'," _
                & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "'," _
                & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
                & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
          Call objHT.InsertData_ACC(sSQL1)
          
          'PPN
          If vTPPN <> 0 Then
              sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                    & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                    & " JumlahTrn, Kodeuser, DelIndex, Kodesumber,nowo, nopolisi, tglEntry,kodejurnal) values " _
                    & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & "', 'S'," _
                    & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                    & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                    & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "'," _
                    & " " & SimpanKoma2Titik(vTPPN * -1) & ", '" & gSA & "'," _
                    & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
              Call objHT.InsertData_ACC(sSQL1)
          End If
      
          'Update Armst Saldo berjalan
          sSQL1 = "Update Armst set SaldoJalan=SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(text(2)) & "'"
          Call objHT.UpdateData_ACC(sSQL1)
      'End If
      
      'Membuat transaksi di GL
          sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
          Call objHT.InsertData_ACC(sSQL1)
          
          sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
          Call objHT.InsertData_ACC(sSQL1)
          
          If vTMeterai <> 0 Then
              sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                    & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                    & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                    & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                    & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              Call objHT.InsertData_ACC(sSQL1)
          End If
          
          If vTPPN <> 0 Then
              sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                    & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                    & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                    & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                    & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              Call objHT.InsertData_ACC(sSQL1)
          End If
      
          sSQL1 = "update glmst set JbulanIni = JbulanIni -" & (vTotal + vTMeterai) & ", " _
                & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                & " where KodeGl = '" & vPiutangGL & "'"
          Call objHT.UpdateData_ACC(sSQL1)
          
          sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                & " where KodeGl = '" & vPenjualanGL & "'"
          Call objHT.UpdateData_ACC(sSQL1)
          
          sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTMeterai) & ", " _
                & " JDebitBi = JDebitBi +" & vTMeterai & " " _
                & " where KodeGl = '" & vMeteraiGL & "'"
          Call objHT.UpdateData_ACC(sSQL1)
          
          sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPPN) & ", " _
                & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPPN) & " " _
                & " where KodeGl = '" & vPpnGL & "'"
          Call objHT.UpdateData_ACC(sSQL1)
      'End If
  End If
    Exit Sub
vbError:
    erTranAll = True
    Exit Sub

   'On Error GoTo 0
   Exit Sub

SimpanAAS2_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure SimpanAAS2 of Form frmBtlInv")
End Sub

Private Sub hapusHO()
1        On Error GoTo SavingBatch

2         If Log.LogHOConn = 1 Then
3             If ConHO = False Then
4                 GoTo SavingBatch
5             Else
6                 If Get_ConnectionProfil_Bengkel99 = False Then
7                     GoTo SavingBatch
8                 Else
9                     errHO = False
10                    Call objHO.OpenTransactionLog_NRM
11                    Call objHO.AddNew_WO(nSQLo(5))
12                    If errHO = False Then
13                        Call objHO.AddNew_WO(nSQLo(0))
14                    End If
15                    If errHO = False Then
16                        If nSQLo(1) <> "" Then Call objHO.AddNew_WO(nSQLo(1))
17                    End If
18                    If errHO = False Then
19                        If nSQLo(2) <> "" Then Call objHO.AddNew_WO(nSQLo(2))
20                    End If
21                    If errHO = False Then
22                        If nSQLo(3) <> "" Then Call objHO.AddNew_WO(nSQLo(3))
23                    End If
24                    If errHO = False Then
25                        If nSQLo(4) <> "" Then Call objHO.AddNew_WO(nSQLo(4))
26                          End If
27                          If errHO = False Then
28                        objHO.CloseTransactionLog_NRM
29                    End If
30                    If errHO = True Then GoTo SavingBatch
31                End If
32            End If
33        Else
34            GoTo SavingBatch
35        End If
36    Exit Sub
SavingBatch:
37        Call Saving_Log_HO("Retur", nSQLo(5))
38        Call Saving_Log_HO("Retur0", nSQLo(0))
39        If nSQLo(1) <> "" Then Call Saving_Log_HO("Retur1", nSQLo(1))
40        If nSQLo(2) <> "" Then Call Saving_Log_HO("Retur2", nSQLo(2))
41        If nSQLo(3) <> "" Then Call Saving_Log_HO("Retur3", nSQLo(3))
42        If nSQLo(4) <> "" Then Call Saving_Log_HO("Retur4", nSQLo(4))
43        Exit Sub

         'On Error GoTo 0
44       Exit Sub

hapusHO_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
45        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure hapusHO of Form frmBtlInv")
End Sub

Private Sub Saving_Log_HO(nStatus As String, isSQl As String)
1        On Error GoTo Saving_Log_HO_Error

2         Call objHO.AddNew_LogHO(Trim(text(0)), nStatus, "INVR", isSQl)

         'On Error GoTo 0
3        Exit Sub

Saving_Log_HO_Error:
'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
varErlError = Erl
varErrNumber = err.Number
'' varErlError,varErrNumber
4         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure Saving_Log_HO of Form frmBtlInv")
End Sub

Private Sub Command1_Click()
    If gUkuranRetur Then
        Call Cetak
        If nDuoInv Then
            Call Cetak2
        End If
        If isInvoiceC Then
            Call Cetak3
        End If
    Else
        Call CetakFormKecil
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm
    text(1) = TglDDMMYYYY(isnow)
End Sub

Private Sub tblBatal_Click()
     
    frmBtlInvoice.Visible = True
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub


Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
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
                If Log.newCetak = 1 And gTipeUser = "SA" Then
                    gBantuan = " and front = '" & gSA & "'"
                Else
                    gBantuan = ""
                End If
                gMainKeyCari = text(Index)
                gSearchOption = "INVBATAL"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
    If Index = 0 Then
        text(0).text = Replace(text(0).text, "-", "")
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim nKey As String
Dim rsSearch As ADODB.Recordset
    Select Case Index
        Case 0
            If text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_INVOICE WITH(NOLOCK) WHERE DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV  like '" & text(Index) & "%'")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "INVBATAL"
                    frmCari.show (vbModal)
                Else
                    nKey = text(Index)
                    Call hapusForm
                    text(Index) = nKey
                    TblBatal.Enabled = True
                    Call panggilINV(Trim(text(Index)))
                End If
            End If
    End Select
End Sub

Private Sub VScroll1_Change()
'    Frame6.Top = -VScroll1.value + 20
    If VScroll1.value > 0 Then
        Frame6.Top = -VScroll1.value + 20
    Else
        Frame6.Top = 0
    End If
End Sub

Private Sub VScroll1_Scroll()
    Frame6.Top = -VScroll1.value + 20
End Sub


Private Sub SimpanACCT_invoice()
          
          
          Dim SsqlNis As String
          Dim isSudahDikurangi As Boolean
1        On Error GoTo SimpanACCT_invoice_Error

2         isSudahDikurangi = False
        
          Dim vSyaratByr As String
          'Dim vJmlHari As Double
          Dim vTglSkr As Date
          Dim vTglTempo As Date
          Dim vTPenjualan As Double
          Dim vTPenjSublet As Double
          Dim vTPenjNonMet As Double
          Dim vTPPN As Double
          
          Dim vTPembSublet As Double
          
          Dim vJSPPN As Double                ' NILAI PPN JASA + SUBLET
          Dim vPaPPN As Double                ' NILAI PPN Part
          Dim vOlPPN As Double                ' NILAI PPN Oli
          Dim vAcPPN As Double                ' NILAI PPN Accs
          
          Dim vTotal As Double
          Dim vTotalNonPPn As Double
          Dim vTMeterai As Double
          Dim vKeterangan As String
          
          Dim vPiutangGL As String
          Dim vPenjualanGL As String
          Dim vPenjSubletGL As String
          Dim vPpnGL As String
          Dim vMeteraiGL As String
          Dim vHPPSublet As String
          Dim vPDPSublet As String
          
          Dim vDelIndex As String
          Dim vKodeLgn As String
          
          Dim NilaiGLNonPart As Double
          
          Dim totalKreditTemp As Double
          Dim totalDebitTemp As Double
          
          Dim adaSelisih As Boolean
          Dim kreditSelisih As Boolean
          
          Dim rsInvoice   As ADODB.Recordset
          Dim rsInvoiceB  As ADODB.Recordset
          Dim rsInvoiceC  As ADODB.Recordset
          
          Dim vLaborA       As Double
          Dim vSubletA       As Double
          
          Dim isFOP       As Boolean
3         isFOP = False
4         Set rsInvoice = objHT.FetchDataBengkel("SELECT OIL_G+OIL_NG as TotOil,PART_G+PART_NG as TotPart,PART_G,PART_NG,OIL_G,OIL_NG,MTRL,SBLT,LABOR,TOTAL,MET FROM t_invoice WHERE FS_LOG_DEL = '' AND NOINV = '" & NulSID(text(0)) & IIf(nDuoInv, "-A", "") & "' ")
5         If nDuoInv Then
6             Set rsInvoiceB = objHT.FetchDataBengkel("SELECT OIL_G+OIL_NG as TotOil,PART_G+PART_NG as TotPart,PART_G,PART_NG,OIL_G,OIL_NG,MTRL,SBLT,LABOR,TOTAL,MET FROM t_invoice WHERE FS_LOG_DEL = '' AND NOINV = '" & NulSID(text(0)) & IIf(nDuoInv, "-B", "") & "' ")
7             If Log.freepart = 1 Then
8                 Set rsInvoiceC = objHT.FetchDataBengkel("SELECT OIL_G+OIL_NG as TotOil,PART_G+PART_NG as TotPart,PART_G,PART_NG,OIL_G,OIL_NG,MTRL,SBLT,LABOR,TOTAL,MET,FS_KD_PLG FROM t_invoice WHERE FS_LOG_DEL = '' AND NOINV = '" & NulSID(text(0)) & "-C" & "' ")
9                 If rsInvoiceC.RecordCount > 0 Then
10                    isFOP = True
11                End If
12            End If
13        End If
14        kreditSelisih = True
15        adaSelisih = False
          
16        vKodeLgn = ""
17        vKeterangan = ""
18        vDelIndex = IIf(opWO(2), "TW", gKode) & Left(McDelIndex(7), 8)
19        vTglSkr = gTanggal
20        If Trim(text(41)) = "" Then
21            vTglTempo = gTanggal
22        Else
23            vTglTempo = DateAdd(LCase(Left(text(41), 1)), Val(Right(text(41), Len(text(41)) - 1)), vTglSkr)
24        End If
          
25        If gJurnalPart Then
              '' JURNAL PART TRANSFER KE ACC
              
              ''vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
              ''vTotal = CDbl(rsInvoice!total)
26            vTotal = Round(CDbl(rsInvoice!PART_G) + CDbl(rsInvoice!PART_NG) + CDbl(rsInvoice!OIL_G) + CDbl(rsInvoice!OIL_NG) + CDbl(rsInvoice!MTRL) + CDbl(rsInvoice!sblt) + CDbl(rsInvoice!labor) + CDbl(rsInvoice!met))
              
              ''vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaidppINv / nilaiTotalInv), 0)
27            vTPenjSublet = Round(CDbl(rsInvoice!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''Round(CDbl(rsInvoice!sblt) * 10 / 11, 0)
              
              ''vTPenjNonMet = Round(Koma2Titik(text(22)) * (nilaidppINv / nilaiTotalInv), 0)
28            vTPenjNonMet = Round(CDbl(rsInvoice!labor) * nilaiDppInv / nilaiTotalInv, 0)
29            If Log.fakturMaterai = 1 Then
30                vTMeterai = Koma2Titik(text(42))
31            Else
32                If nDuoInv Then
33                   Meterai = 0
34                Else
35                   vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))                         ' materai
36                End If
37            End If
            
        
            ''vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaidppINv / nilaiTotalInv), 0)  ' jasa non ppn + meterai
38          vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(rsInvoice!ToTPart) + Koma2Titik(rsInvoice!TotOil)) * nilaiDppInv / nilaiTotalInv), 0)
            
39          vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(rsInvoice!ToTPart) + Koma2Titik(rsInvoice!TotOil)) * nilaiDppInv / nilaiTotalInv), 0)
40          vTPPN = vTotal - vTotalNonPPn
        '' REVISI THANIA 09/05/2014
41          vJSPPN = vTotal - Round(((Koma2Titik(text(24)) + Koma2Titik(text(25)))), 0) - vTPenjSublet - vTPenjNonMet
        
      '        vTPPN = Round((vTotal * 10 / 110), 0)
      '        vTPenjualan = (vTotal - vTPPN) + vTMeterai
        
42        Else
        '' JURNAL PART TDK TRANSFER KE ACC
            ''vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
            ''vTotal = CDbl(rsInvoice!total)
43          vTotal = Round(CDbl(rsInvoice!PART_G) + CDbl(rsInvoice!PART_NG) + CDbl(rsInvoice!OIL_G) + CDbl(rsInvoice!OIL_NG) + CDbl(rsInvoice!MTRL) + CDbl(rsInvoice!sblt) + CDbl(rsInvoice!labor) + CDbl(rsInvoice!met))
44          vTPenjSublet = Round(CDbl(rsInvoice!sblt) * (nilaiDppInv / nilaiTotalInv), 0)
45          vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0) - vTPenjSublet
            
46          If Log.fakturMaterai = 1 Then
47                vTMeterai = Koma2Titik(text(42))
48          Else
49            If nDuoInv Then
50                vTMeterai = 0
51            Else
52                vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))                         ' materai
53            End If
54          End If
           
55          vTPenjualan = vTPenjNonMet + vTPenjSublet + vTMeterai
56          vTPPN = Round(vTotal * (nilaiPPnInv / nilaiTotalInv), 0)
57        End If
          
58        If opWO(2) Then
                ' WARRANTY
59              vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[ " & text(43) & " - " & text(2) & " ]"
60              vKodeLgn = gKdPlgHo
61        ElseIf woStat = "SBI" Then
                ' SBI
62              vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[ " & text(43) & " - " & text(2) & " ]"
63              vKodeLgn = Trim(gKdPlgDealer)
64        Else
65              vKeterangan = text(4) & " " & text(20) & gKodeRetur & "[WO: " & text(43) & " ]"
66              vKodeLgn = text(2)
67        End If

          ' KODE AKUN
          
            Dim isDyna As Boolean
68          isDyna = IIf(InStr(1, UCase(text(14).text), "DYNA") > 0, True, False)
            Dim isCust As String
            Dim isExAflAR As Boolean
            Dim isExAflSI As Boolean
            Dim isExIntSI As Boolean
            Dim nKodeSi As String
        
      ''      nKodeSi = Trim(vKodeLgn)
      ''      isCust = Trim(Text(40))
69            If nDuoInv = True And woStat = "SBI" Then
70                nKodeSi = Nul(text(2)) ''Trim(Text(2))
71                vKodeLgn = Nul(text(2))
72                isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
73                If isExAflSI = True Then
74                    isCust = "AFFPPN"
75                Else
76                    isCust = "EXT"
77                End If
78            Else
79                nKodeSi = vKodeLgn 'Trim(text(2)) 'vKodeLgn
80                isCust = Trim(text(40))
81            End If
              
82          isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(nKodeSi)  'cex dari ARMST apakah AFL / not AFL
83          If isExAflAR = True Then  'AFL
84              isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
85              If isExAflSI = True Then
86                  isCust = "AFFPPN"
87              Else
88                  MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
89                  GoTo vbError
90              End If
91          Else
92              If isCust = "AFFPPN" Then  'cex jika tidak sesuai antara ARMST dan M_Asuransi
93                  MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
94                  GoTo vbError
95              End If
96              If isCust = "INT" Then
97                  isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah INT terdaftar
98                  If isExIntSI = False Then
99                      MsgBox "Data Customer '" & nKodeSi & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
100                     GoTo vbError
101                 End If
102             End If
103         End If
        
        'AKUNBARU
        Dim rsGL As ADODB.Recordset
104     Select Case isCust
            Case "EXT"
105             Set rsGL = objO.Get_AkunGL_NEW("OTHERS", isDyna, True)
106         Case Else  'INT/AFL
107             Set rsGL = objO.Get_AkunGL_NEW(nKodeSi, isDyna, True)
108     End Select
            
      '''  piutang gr
      '''  penj bengkel
      '''  penj sublet
      '''  penj Part
      '''  penj oli
      '''  penj Accs
      '''  PPn Jasa
      '''  PPn Part + oli
      '''  PPn Accs
      '''  Materai
      '''  HPP sublet
      '''  Persediaan sublet
          
109       If Not rsGL.BOF Then
110         vPiutangGL = cekKodeAkun(rsGL.Fields(0).value)
111         vPenjualanGL = cekKodeAkun(rsGL.Fields(1).value)
112         vPenjSubletGL = cekKodeAkun(rsGL.Fields(2).value)
113         vPenjPartGL = cekKodeAkun(rsGL.Fields(3).value)
114         vPenjOlieGL = cekKodeAkun(rsGL.Fields(4).value)
115         vPenjAccsGL = cekKodeAkun(rsGL.Fields(5).value)
116         vPpnGL = cekKodeAkun(rsGL.Fields(6).value)
117         vPpnPartGL = cekKodeAkun(rsGL.Fields(7).value)
118         vPpnAccsGL = cekKodeAkun(rsGL.Fields(8).value)
119         vMeteraiGL = cekKodeAkun(rsGL.Fields(9).value)
120         vHPPSublet = cekKodeAkun(rsGL.Fields(10).value)
121         vPDPSublet = cekKodeAkun(rsGL.Fields(11).value)
      ''                If Trim(vPenjSubletGL) = "" Then
      ''                    vTPenjNonMet = vTPenjNonMet + vTPenjSublet
      ''                    vTPenjSublet = 0
      ''                End If
122       End If

123       If rsGL.State <> adStateClosed Then rsGL.Close



124       If vPiutangGL = "" Or vPenjualanGL = "" Or vPenjSubletGL = "" Or vPenjPartGL = "" Or vPenjOlieGL = "" Or vPenjAccsGL = "" _
       Or vPpnGL = "" Or vPpnPartGL = "" Or vPpnAccsGL = "" Or vMeteraiGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Then
125         pesanIT = pesanIT + "Kode GL belum lengkap !"
126         WarnError = True
127         erTranAll = True
128       End If
          
129       If Trim(text(40)) <> "AFFNONPPN" And vPpnGL = "" Then
130         pesanIT = pesanIT + "Kode GL belum lengkap !"
131         WarnError = True
132         erTranAll = True
133       End If
134       If Trim(vKodeLgn) = "" Then
135         pesanIT = pesanIT + "Kode Pelanggan belum lengkap !"
136         WarnError = True
137         erTranAll = True
138       End If
139       If rsInvoice.EOF Then
140           pesanIT = pesanIT + "No Invoice tidak ditemukan !"
141           WarnError = True
142           erTranAll = True
143       End If
144       If nDuoInv Then
145           If rsInvoiceB.EOF Then
146               pesanIT = pesanIT + "No Invoice B tidak ditemukan !"
147               WarnError = True
148               erTranAll = True
149           End If
150       End If
          
          '-------------------- PROSES PENGECEKAN SELISIH --------------------
151       totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTotal + vTMeterai))
152       NilaiGLNonPart = vTotal + vTMeterai
          
153       totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjNonMet))
154       NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
          
155       If vTPenjSublet <> 0 Then
156         totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjSublet))
157         NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
158       End If
          
159       If vTMeterai <> 0 Then
160         totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTMeterai))
161         NilaiGLNonPart = NilaiGLNonPart - vTMeterai
162       End If

163       If (vTPenjNonMet + vTPenjSublet) <> 0 Then
164         totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik((Round(Koma2Titik(rsInvoice!labor) + Koma2Titik(rsInvoice!sblt), 0)) - (vTPenjNonMet + vTPenjSublet)))
165         NilaiGLNonPart = NilaiGLNonPart - (Round(((vTPenjNonMet + vTPenjSublet) / 10), 0))
166       End If
          
167       If gJurnalPart Then
              Dim SumPartTemp As Double
              Dim SumOlieTemp As Double
              Dim SumAccsTemp As Double
              Dim SumPartFullTemp As Double
              Dim SumOlieFullTemp As Double
              Dim SumAccsFullTemp As Double
              '' NILAI SPAREPART
168           vTPenjPart = 0
              Dim rsSearchPartTemp As ADODB.Recordset
              
169           If rsPart.State <> adStateClosed Then
170               SumPartTemp = 0
171               SumOlieTemp = 0
172               SumAccsTemp = 0
173               SumPartFullTemp = 0
174               SumOlieFullTemp = 0
175               SumAccsFullTemp = 0
176               vPaPPN = 0
177               vOlPPN = 0
178               vAcPPN = 0
179               rsPart.MoveFirst
180               Do Until rsPart.EOF
181                   If Log.freepart = 1 Then
182                       If rsPart!FREEC = False Then
183                           GoTo lanjutloop
184                       End If
185                   End If
186                   Select Case Left(UCase(objO.FetchData_String_Acc("select B.Jenis from Inmst A WITH(NOLOCK) Inner Join Profilacc..SCC B WITH(NOLOCK) On A.SCC_Code = B.Scc_Code Where A.KodeItem = '" & Trim(rsPart("KodeItem")) & "'")), 3)
                          Case "PAR"
187                           SumPartTemp = SumPartTemp + SpyRound((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
188                           SumPartFullTemp = SumPartFullTemp + CDbl(rsPart("HARGA"))
189                       Case "OLI"
190                           SumOlieTemp = SumOlieTemp + SpyRound((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
191                           SumOlieFullTemp = SumOlieFullTemp + CDbl(rsPart("HARGA"))
192                       Case "ACC"
193                           SumAccsTemp = SumAccsTemp + SpyRound((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
194                           SumAccsFullTemp = SumAccsFullTemp + CDbl(rsPart("HARGA"))
195                   End Select
lanjutloop:
196                   rsPart.MoveNext
197               Loop
198               vTPenjPart = SumPartTemp + SumOlieTemp + SumAccsTemp
              
199                 If NilaiGLNonPart <> vTPenjPart Then
200                     If NilaiGLNonPart - vTPenjPart = 1 Then
201                         If SumPartTemp <> 0 Then
202                             SumPartTemp = SumPartTemp + 1
203                         Else
204                             If SumOlieTemp <> 0 Then
205                                 SumOlieTemp = SumOlieTemp + 1
206                             Else
207                                 If SumAccsTemp <> 0 Then
208                                     SumAccsTemp = SumAccsTemp + 1
209                                 End If
210                             End If
211                         End If
212                         vTPenjPart = NilaiGLNonPart
213                     ElseIf NilaiGLNonPart - vTPenjPart = -1 Then
214                         If SumPartTemp <> 0 Then
215                             SumPartTemp = SumPartTemp - 1
216                         Else
217                             If SumOlieTemp <> 0 Then
218                                 SumOlieTemp = SumOlieTemp - 1
219                             Else
220                                 If SumAccsTemp <> 0 Then
221                                     SumAccsTemp = SumAccsTemp - 1
222                                 End If
223                             End If
224                         End If
225                         vTPenjPart = NilaiGLNonPart
226                     End If
227                 End If
                  
228                 vPaPPN = SumPartFullTemp - SumPartTemp
229                 vOlPPN = SumOlieFullTemp - SumOlieTemp
230                 vAcPPN = SumAccsFullTemp - SumAccsTemp
              
231               If SumPartTemp <> 0 Then
232                   totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(SumPartTemp))
233               End If
                  
234               If SumOlieTemp <> 0 Then
235                   totalDebitTemp = totalDebitTemp + SimpanKoma2Titik(SumOlieTemp)
236               End If
                  
237               If (SumPartTemp + SumOlieTemp) <> 0 Then
238                   totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0)))
239               End If
                
240               If SumAccsTemp <> 0 Then
241                   totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(SumAccsTemp))
                      
242                   totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(Round((vAcPPN), 0)))
243               End If
244           End If
245       End If
          
          Dim vTPpnSTemp As String
          Dim vTotalSTemp As Double
          Dim vTBeliSTemp As Double
          Dim nonJurnalTemp As Boolean
          Dim rsSearchTemp As New ADODB.Recordset
          Dim vKeterangan1Temp As String
          
246       nonJurnalTemp = False
                
247       vTPembSublet = 0
248       vTPenjSublet = 0
249       Set rsSearchTemp = objO.FetchData("SELECT PKP,NAMABARANG,NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, NONJURNAL FROM JOBSUBLET a WITH(NOLOCK) INNER JOIN T_SUPPLIER B WITH(NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(text(43)) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
250       If Not rsSearchTemp.EOF Then
251           Do While Not rsSearchTemp.EOF
252               nonJurnalTemp = IIf(Trim(rsSearchTemp("NONJURNAL")) = 1, True, False)
253               If Not nonJurnalTemp Then
254                   vTotalSTemp = Round(CDbl(rsSearchTemp("HPP")))
255                   vTPpnSTemp = Round(vTotalSTemp * (nilaiPPnInv / nilaiTotalInv))
256                   vTBeliSTemp = Round(vTotalSTemp - vTPpnSTemp)
                     ' vKeterangan1Temp = vKeterangan & " SUBLET" '(" & rsSearch("NAMABARANG") & ")"   'No. WO :" & Trim(Text(0))
257                   If Trim(rsSearchTemp("PKP")) = "1" Then
258                       vPpnSGL = ""
259                       vTPpnSTemp = 0
260                       vTotalSTemp = Round(vTotalSTemp)
261                   End If
262                   vTPenjSublet = vTPenjSublet + vTotalSTemp
263               End If
264               rsSearchTemp.MoveNext
265           Loop
266       End If
          
267       If vTPenjSublet <> 0 Then
268         totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjSublet))
269         totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTPenjSublet))
270       End If
          
271       If nDuoInv Then
              ' UNTUK INVOICE KEDUA -B KE DEALER PENJUAL
              ' MEMBUAT TRANSAKSI DUO INV DI AR
              
              ''vTotal = subJob
              ''vTotal = CDbl(rsInvoiceB!labor)
272           vTotal = Round(CDbl(rsInvoiceB!PART_G) + CDbl(rsInvoiceB!PART_NG) + CDbl(rsInvoiceB!OIL_G) + CDbl(rsInvoiceB!OIL_NG) + CDbl(rsInvoiceB!MTRL) + CDbl(rsInvoiceB!sblt) + CDbl(rsInvoiceB!labor) + CDbl(rsInvoiceB!met))
273           vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0)  'dikurangi ppn
274           vTMeterai = Koma2Titik(metInvB) ''Koma2Titik(text(42))               'materai
275           vTPenjualan = vTPenjNonMet + vTMeterai         'non ppn + meterai
              ''2015-10-19 || AKTIF AFFPPN
276           If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
277               If Trim(gNpwp) = Trim(gInvNPWP2) Then
278                   vTotal = vTPenjNonMet
279                   vTPPN = 0     'Round(vTotal * (10 / 110), 0)
280               Else
281                   vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
282               End If
283           Else
284               vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
285           End If
              'vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          'ppn
              'vTPenjualan = (vTotal - vTPPN) + vTMeterai
286           vKeterangan = gInsCust2 & ". AN/ " & vKeterangan
              
              ' MEMBUAT TRANSAKSI DUO INV DI GL
              
287           totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTotal + vTMeterai))
              
288           totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPenjNonMet))
              
289           If vTMeterai <> 0 Then
290               totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTMeterai))
291           End If
              
292           If vTPPN <> 0 Then
293               totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPPN))
294           End If
295       End If
          
296       If totalDebitTemp <> totalKreditTemp Then
297         If (totalDebitTemp - totalKreditTemp) = -1 Or (totalDebitTemp - totalKreditTemp) = 1 Then
298             If (totalDebitTemp - totalKreditTemp) = -1 Then
299                 kreditSelisih = False
300             End If
301             adaSelisih = True
302         End If
303       End If
          
          '-------------------- AKHIR PENGECEKAN SELISIH ---------------------
          
304       If gJurnalPart Then
                '' JURNAL PART TRANSFER KE ACC
305             vTotal = Round(CDbl(rsInvoice!PART_G) + CDbl(rsInvoice!PART_NG) + CDbl(rsInvoice!OIL_G) + CDbl(rsInvoice!OIL_NG) + CDbl(rsInvoice!MTRL) + CDbl(rsInvoice!sblt) + CDbl(rsInvoice!labor) + CDbl(rsInvoice!met)) ''Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
306              vTPenjSublet = Round(CDbl(rsInvoice!sblt) * (nilaiDppInv / nilaiTotalInv), 0)
307             vTPenjNonMet = Round(CDbl(rsInvoice!labor) * nilaiDppInv / nilaiTotalInv, 0) 'Round(vTotal * (nilaidppINv / nilaiTotalInv), 0) - vTPenjSublet
308             vLaborA = Round(CDbl(rsInvoice!labor), 0)
309              vSubletA = Round(CDbl(rsInvoice!sblt), 0)
      '          vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaidppINv / nilaiTotalInv), 0)
      '          vTPenjNonMet = Round(Koma2Titik(text(22)) * (nilaidppINv / nilaiTotalInv), 0)
310             If Log.fakturMaterai = 1 Then
311               vTMeterai = Koma2Titik(text(42))
312             Else
313               If nDuoInv Then
314                   vTMeterai = 0
315               Else
316                   vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42)) ' materai
317               End If
318             End If
                  
319             vTPenjualan = vTPenjNonMet + vTPenjSublet + vTMeterai
320             vTPPN = Round(vTotal * (nilaiPPnInv / nilaiTotalInv), 0)
321             vTotalNonPPn = vTotal - vTPPN
322             vTPenjualan = vTotalNonPPn
      '          vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaidppINv / nilaiTotalInv), 0)  ' jasa non ppn + meterai
      '          vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * nilaidppINv / nilaiTotalInv), 0)
      '          vTPPN = vTotal - vTotalNonPPn
                '' REVISI THANIA 09/05/2014
323             vJSPPN = vTotal - Round(((Koma2Titik(text(24)) + Koma2Titik(text(25)))), 0) - vTPenjSublet - vTPenjNonMet
            
          '        vTPPN = Round((vTotal * 10 / 110), 0)
          '        vTPenjualan = (vTotal - vTPPN) + vTMeterai
        
324       Else
              '' JURNAL PART TDK TRANSFER KE ACC
      '        vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
      '        vTPenjSublet = Round(Koma2Titik(text(27)) * (nilaidppINv / nilaiTotalInv), 0)
      '        vTPenjNonMet = Round(vTotal * (nilaidppINv / nilaiTotalInv), 0) - vTPenjSublet
              
325           vTotal = Round(CDbl(rsInvoice!PART_G) + CDbl(rsInvoice!PART_NG) + CDbl(rsInvoice!OIL_G) + CDbl(rsInvoice!OIL_NG) + CDbl(rsInvoice!MTRL) + CDbl(rsInvoice!sblt) + CDbl(rsInvoice!labor) + CDbl(rsInvoice!met)) ''Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
326           vTPenjSublet = Round(CDbl(rsInvoice!sblt) * (nilaiDppInv / nilaiTotalInv), 0)
327           vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0) - vTPenjSublet
328           vLaborA = Round(CDbl(rsInvoice!labor), 0)
329           vSubletA = Round(CDbl(rsInvoice!sblt), 0)
330           If Log.fakturMaterai = 1 Then
331               vTMeterai = Koma2Titik(text(42))
332           Else
333               If nDuoInv Then
334                   vTMeterai = 0
335               Else
336                   vTMeterai = Koma2Titik(metInvB)  '' Koma2Titik(Text(42))                         ' materai
337               End If
338           End If
        
339         vTPenjualan = vTPenjNonMet + vTPenjSublet + vTMeterai
340         vTPPN = Round(vTotal * (nilaiPPnInv / nilaiTotalInv), 0)
341       End If
          
342       If erTranAll = True Then GoTo vbError ' Exit Sub
          
343         sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                  & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                  & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                  & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & "', 'Q'," _
                  & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                  & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                  & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "'," _
                  & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
                  & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
344         Call objHT.InsertData_ACC(sSQL1)
            
            'PPN
345         If vTPPN <> 0 Then
346             sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                      & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                      & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                      & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & "', 'S'," _
                      & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                      & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                      & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "'," _
                      & " " & SimpanKoma2Titik(Round(vTPPN, 0) * -1) & ", '" & gSA & "'," _
                      & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
347             Call objHT.InsertData_ACC(sSQL1)
348         End If
        
            'Update Armst Saldo berjalan
349         sSQL1 = "Update Armst set SaldoJalan=SaldoJalan + " & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(text(2)) & "'"
350         Call objHT.UpdateData_ACC(sSQL1)
        'End If
            
          'Membuat transaksi di GL
        'If vTotal + vTMeterai <> 0 Then
            'Piutang
351         sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                  & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                  & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                  & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                  & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
352         Call objHT.InsertData_ACC(sSQL1)
353         NilaiGLNonPart = vTotal + vTMeterai
            'Penj Jasa
354         sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                  & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                  & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                  & " " & SimpanKoma2Titik(SpyRound((vTPenjNonMet))) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                  & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
355         If Log.dpplain = 1 Then
356           sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                  & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                  & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                  & " " & SimpanKoma2Titik(SpyRound((vTPenjNonMet * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                  & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
357         End If
            
358         Call objHT.InsertData_ACC(sSQL1)
359         sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                  & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                  & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                  & " " & SimpanKoma2Titik(vTPenjNonMet - SpyRound((vTPenjNonMet * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                  & " '" & vDelIndex & "','SNL : " & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
360         If Log.dpplain = 1 Then Call objHT.InsertData_ACC(sSQL1)
361         NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
            'Penj Sublet
362         If vTPenjSublet <> 0 Then
363             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(SpyRound((vTPenjSublet))) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
364             If Log.dpplain = 1 Then
365               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(SpyRound((vTPenjSublet * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                      
366             End If
367             Call objHT.InsertData_ACC(sSQL1)
368             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(vTPenjSublet - SpyRound((vTPenjSublet * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','SNL : " & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
369             If Log.dpplain = 1 Then Call objHT.InsertData_ACC(sSQL1)
370             NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
371         End If
            'Materai
372         If vTMeterai <> 0 Then
373             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
374             Call objHT.InsertData_ACC(sSQL1)
375             NilaiGLNonPart = NilaiGLNonPart - vTMeterai
376         End If
            'PPn Jasa + Sublet
377          If (vTPenjNonMet + vTPenjSublet) <> 0 Then
      '                sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                      & " ('" & Trim(gKodeFaktur) & Trim(Text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                      & " " & SimpanKoma2Titik(Round(((vTPenjNonMet + vTPenjSublet) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                      
                ' REVISI THANIA 09/05/2014
                 'sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(vJSPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                      
                ' REVISI ACEP 10/06/2014
      '           sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                & " ('" & Trim(gKodeFaktur) & Trim(Text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                & " " & SimpanKoma2Titik((Round(Koma2Titik(Text(22)) + Koma2Titik(Text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                
378             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik((Round(Koma2Titik(vLaborA) + Koma2Titik(vSubletA), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                
379             If adaSelisih Then
380                 If kreditSelisih Then
      '                  sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                & " ('" & Trim(gKodeFaktur) & Trim(Text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                & " " & SimpanKoma2Titik((Round(Koma2Titik(Text(22)) + Koma2Titik(Text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
381                     sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik((Round(Koma2Titik(vLaborA) + Koma2Titik(vSubletA), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"

382                 Else
      '                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                & " ('" & Trim(gKodeFaktur) & Trim(Text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                & " " & SimpanKoma2Titik((Round(Koma2Titik(Text(22)) + Koma2Titik(Text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                      
383                      sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik((Round(Koma2Titik(vLaborA) + Koma2Titik(vSubletA), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"

384                 End If
385                 isSudahDikurangi = True
386             End If
387             Call objHT.InsertData_ACC(sSQL1)
388             NilaiGLNonPart = NilaiGLNonPart - (Round(((vTPenjNonMet + vTPenjSublet) / 10), 0))
389         End If
            
390         If gJurnalPart Then
                '' NILAI SPAREPART
                    Dim SumPart As Double
                    Dim SumOlie As Double
                    Dim SumAccs As Double
                    Dim SumPartFull As Double
                    Dim SumOlieFull As Double
                    Dim SumAccsFull As Double
                    '' NILAI SPAREPART
391                 vTPenjPart = 0
                    Dim rsSearchPart As ADODB.Recordset
392                 If rsPart.State <> adStateClosed Then
393                     SumPart = 0
394                     SumOlie = 0
395                     SumAccs = 0
396                     SumPartFull = 0
397                     SumOlieFull = 0
398                     SumAccsFull = 0
399                     vPaPPN = 0
400                     vOlPPN = 0
401                     vAcPPN = 0
402                     rsPart.MoveFirst
403                     Do Until rsPart.EOF
404                         If Log.freepart = 1 Then
405                           If rsPart!FREEC = False Then
406                               GoTo lanjutLoop2
407                           End If
408                         End If
409                         Select Case Left(UCase(objO.FetchData_String_Acc("select B.Jenis from Inmst A WITH(NOLOCK) Inner Join Profilacc..SCC B WITH(NOLOCK) On A.SCC_Code = B.Scc_Code Where A.KodeItem = '" & Trim(rsPart("KodeItem")) & "'")), 3)
                                Case "PAR"
410                                 SumPart = SumPart + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
411                                 SumPartFull = SumPartFull + CDbl(rsPart("HARGA"))
412                             Case "OLI"
413                                 SumOlie = SumOlie + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
414                                 SumOlieFull = SumOlieFull + CDbl(rsPart("HARGA"))
415                             Case "ACC"
416                                 SumAccs = SumAccs + Round((CDbl(rsPart("HARGA")) * nilaiDppInv / nilaiTotalInv), 0)
417                                 SumAccsFull = SumAccsFull + CDbl(rsPart("HARGA"))
418                         End Select
lanjutLoop2:
419                         rsPart.MoveNext
420                     Loop
421                     vTPenjPart = SumPart + SumOlie + SumAccs
                        
422                     If NilaiGLNonPart <> vTPenjPart Then
423                         If NilaiGLNonPart - vTPenjPart = 1 Then
424                             If SumPart <> 0 Then
425                                 SumPart = SumPart + 1
426                             Else
427                                 If SumOlie <> 0 Then
428                                     SumOlie = SumOlie + 1
429                                 Else
430                                     If SumAccs <> 0 Then
431                                         SumAccs = SumAccs + 1
432                                     End If
433                                 End If
434                             End If
435                             vTPenjPart = NilaiGLNonPart
436                         ElseIf NilaiGLNonPart - vTPenjPart = -1 Then
437                             If SumPart <> 0 Then
438                                 SumPart = SumPart - 1
439                             Else
440                                 If SumOlie <> 0 Then
441                                     SumOlie = SumOlie - 1
442                                 Else
443                                     If SumAccs <> 0 Then
444                                         SumAccs = SumAccs - 1
445                                     End If
446                                 End If
447                             End If
448                             vTPenjPart = NilaiGLNonPart
449                         End If
450                     End If
                
                
451                     vPaPPN = SumPartFull - SumPart
452                     vOlPPN = SumOlieFull - SumOlie
453                     vAcPPN = SumAccsFull - SumAccs
                
      '''
454                 If SumPart <> 0 Then
455                     sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SpyRound((SumPart))) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
456                     If Log.dpplain = 1 Then
                          
457                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SpyRound((SumPart * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
458                     End If
459                           Call objHT.InsertData_ACC(sSQL1)
460                     sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SumPart - SpyRound((SumPart * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','SNL : " & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
461                           If Log.dpplain = 1 Then Call objHT.InsertData_ACC(sSQL1)
462                     sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumPart) & ", " _
                              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumPart) & " " _
                              & " where KodeGl = '" & vPenjPartGL & "'"
463                           Call objHT.UpdateData_ACC(sSQL1)
464                 End If
465                 If SumOlie <> 0 Then
466                     sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjOlieGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SpyRound((SumOlie))) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
467                     If Log.dpplain = 1 Then
                          
468                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjOlieGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SpyRound((SumOlie * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                              
469                     End If
470                           Call objHT.InsertData_ACC(sSQL1)
471                     sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjOlieGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SumOlie - SpyRound((SumOlie * 11) / 12)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','SNL : " & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
472                           If Log.dpplain = 1 Then Call objHT.InsertData_ACC(sSQL1)
473                     sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumOlie) & ", " _
                              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumOlie) & " " _
                              & " where KodeGl = '" & vPenjOlieGL & "'"
474                           Call objHT.UpdateData_ACC(sSQL1)
475                 End If
                    
476                         If (SumPart + SumOlie) <> 0 Then
      '                                sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                                      & " ('" & Trim(gKodeFaktur) & Trim(Text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                                      & " " & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                                      & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(Text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
      '                                      Call objHT.InsertData_ACC(sSQL1)
      '                                sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
      '                                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & " " _
      '                                      & " where KodeGl = '" & vPpnPartGL & "'"
      '                                      Call objHT.InsertData_ACC(sSQL1)

477                             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                      & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                      & " " & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                      & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
478                                 If adaSelisih Then
479                                     If isSudahDikurangi = False Then
480                                         If kreditSelisih Then
481                                             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                                & " " & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) + 1) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
482                                         Else
483                                             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                                & " " & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) - 1) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
484                                         End If
485                                     End If
486                                 End If
                                        
487                             Call objHT.InsertData_ACC(sSQL1)
                                'sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _

                                'sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(((SumPart + SumOlie) / 10)) & ", " _

488                             sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
                                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0)) & " " _
                                      & " where KodeGl = '" & vPpnPartGL & "'"
                                
489                             If adaSelisih Then
490                                 If isSudahDikurangi = False Then
491                                     If kreditSelisih Then
492                                         sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0) + 1) & ", " _
                                            & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) + 1) & " " _
                                            & " where KodeGl = '" & vPpnPartGL & "'"
493                                     Else
494                                         sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0) - 1) & ", " _
                                            & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(Round(vPaPPN + vOlPPN, 0) - 1) & " " _
                                            & " where KodeGl = '" & vPpnPartGL & "'"
495                                     End If
496                                 End If
497                             End If
                                    
498                             Call objHT.InsertData_ACC(sSQL1)
499                         End If
                    
500                 If SumAccs <> 0 Then
501                     sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPenjAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SumAccs) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
502                           Call objHT.InsertData_ACC(sSQL1)
503                     sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumAccs) & ", " _
                              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumAccs) & " " _
                              & " where KodeGl = '" & vPenjAccsGL & "'"
504                           Call objHT.UpdateData_ACC(sSQL1)
                              
      '                        sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                              & " ('" & Trim(gKodeFaktur) & Trim(Text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                              & " " & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                              & " '" & vDelIndex & "','" & vKeterangan & "- PPn Accs" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(Text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
      '                              Call objHT.InsertData_ACC(sSQL1)
      '                        sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", " _
      '                              & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & " " _
      '                              & " where KodeGl = '" & vPpnAccsGL & "'"
      '                              Call objHT.InsertData_ACC(sSQL1)
                              
505                     sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPpnAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(Round(vAcPPN, 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "- PPn Accs" & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
506                           Call objHT.InsertData_ACC(sSQL1)
507                     sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", " _
                              & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(Round(vAcPPN, 0)) & " " _
                              & " where KodeGl = '" & vPpnAccsGL & "'"
508                           Call objHT.InsertData_ACC(sSQL1)

509                 End If
                    
510             End If
511         End If
            
            ' AMBIL HPP SUBLET
            Dim vTPpnS As String
            Dim vTotalS As Double
            Dim vTBeliS As Double
            Dim nonJurnal As Boolean
            Dim rsSearch As New ADODB.Recordset
            Dim vKeterangan1 As String
            
512         nonJurnal = False
513         vTPembSublet = 0
           ' vTPenjSublet = 0
514         Set rsSearch = objO.FetchData("SELECT PKP,NAMABARANG,NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, NONJURNAL FROM JOBSUBLET a WITH(NOLOCK) INNER JOIN T_SUPPLIER B WITH(NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(text(43)) & "' and no_inv not like '%-C' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
515         If Not rsSearch.EOF Then
516             Do While Not rsSearch.EOF
                                        
517                 nonJurnal = IIf(Trim(rsSearch("NONJURNAL")) = 1, True, False)
518                 If Not nonJurnal Then
519                     vTotalS = Round(CDbl(rsSearch("HPP")))
520                     vTPpnS = Round(vTotalS * (nilaiPPnInv / nilaiTotalInv), 0)
521                     vTBeliS = Round(vTotalS - vTPpnS, 0)
522                     vKeterangan1 = vKeterangan & " SUBLET" '(" & rsSearch("NAMABARANG") & ")"   'No. WO :" & Trim(Text(0))
523                     If Trim(rsSearch("PKP")) = "1" Then
524                         vPpnSGL = ""
525                         vTPpnS = 0
526                         vTotalS = Round(vTBeliS, 0)
527                     End If
                        'vTPenjSublet = vTPenjSublet + vTotalS
528                     vTPembSublet = vTPembSublet + vTotalS
529                 End If
530                 rsSearch.MoveNext
531             Loop
532         End If
            
                
            'If Not nonJurnal Then
533             If vTPembSublet <> 0 Then
534                 sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vHPPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
535                 Call objHT.InsertData_ACC(sSQL1)
                    
536                 sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-A", "") & Trim(gKodeRetur) & "', '" & vPDPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
537                 Call objHT.InsertData_ACC(sSQL1)
                    
                    ' HPP SUBLET TERHADAP PDP
538                 sSQL1 = "update glmst set JbulanIni = JbulanIni -" & (vTPembSublet) & ", " _
                          & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                          & " where KodeGl = '" & vHPPSublet & "'"
539                 Call objHT.InsertData_ACC(sSQL1)
                    
540                 sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPembSublet) & ", " _
                          & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                          & " where KodeGl = '" & vPDPSublet & "'"
541                 Call objHT.InsertData_ACC(sSQL1)
542             End If
            'End If
                
                        
543         sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTotal + vTMeterai) & ", " _
                  & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                  & " where KodeGl = '" & vPiutangGL & "'"
544         Call objHT.UpdateData_ACC(sSQL1)
            
545         sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                  & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                  & " where KodeGl = '" & vPenjualanGL & "'"
546         Call objHT.UpdateData_ACC(sSQL1)
            
547         sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjSublet) & ", " _
                  & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjSublet) & " " _
                  & " where KodeGl = '" & vPenjSubletGL & "'"
548         Call objHT.UpdateData_ACC(sSQL1)
            
549         sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTMeterai) & ", " _
                  & " JDebitBi = JDebitBi +" & vTMeterai & " " _
                  & " where KodeGl = '" & vMeteraiGL & "'"
550         Call objHT.UpdateData_ACC(sSQL1)
            
      '            sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPPN) & ", " _
      '                  & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPPN) & " " _
      '                  & " where KodeGl = '" & vPpnGL & "'"
      '            Call objHT.UpdateData_ACC(sSQL1)



551     sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", " _
        & " JDebitBi = JDebitBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & " " _
        & " where KodeGl = '" & vPpnGL & "'"
        
552     If adaSelisih Then
553         If kreditSelisih Then
554             sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & ", " _
                & " JDebitBi = JDebitBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & " " _
                & " where KodeGl = '" & vPpnGL & "'"
555         Else
556             sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & ", " _
                & " JDebitBi = JDebitBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & " " _
                & " where KodeGl = '" & vPpnGL & "'"
557         End If
558     End If
        
559     Call objHT.InsertData_ACC(sSQL1)
                
            
        'End If
        
        
560     If nDuoInv Then
           '' INV B
            ''vTotal = subJob
            ''vTotal = CDbl(rsInvoiceB!labor)
561         vTotal = Round(CDbl(rsInvoiceB!PART_G) + CDbl(rsInvoiceB!PART_NG) + CDbl(rsInvoiceB!OIL_G) + CDbl(rsInvoiceB!OIL_NG) + CDbl(rsInvoiceB!MTRL) + CDbl(rsInvoiceB!sblt) + CDbl(rsInvoiceB!labor)) '+ CDbl(rsInvoiceB!met))
562         vTPenjNonMet = Round(vTotal * (nilaiDppInv / nilaiTotalInv), 0)  'dikurangi ppn
563         vTMeterai = Koma2Titik(metInvB) '' Koma2Titik(Text(42))
564         vTPenjualan = vTPenjNonMet + vTMeterai          'non ppn + meterai
            ''2015-10-19 || AKTIF AFFPPN
565         If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
566             If Trim(gNpwp) = Trim(gInvNPWP2) Then
567                 vTotal = vTPenjNonMet
568                 vTPPN = 0     'Round(vTotal * (nilaiPPnINv / nilaiTotalInv), 0)
569             Else
570                 vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
571             End If
572         Else
573             vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
574         End If
            'vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          ' ppn
            'vTPenjualan = (vTotal - vTPPN) + vTMeterai
            
            
            ' UNTUK RETUR INVOICE KE 2
575         If vPiutangGL = "" Or vPenjualanGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Or vMeteraiGL = "" Then
576             pesanIT = pesanIT + "Kode GL belum lengkap !"
577             WarnError = True
578             erTranAll = True
579         End If
580         If Trim(vTipeCustomer) <> "AFFNONPPN" And vPpnGL = "" Then
581             pesanIT = pesanIT + "Kode GL belum lengkap !"
582             WarnError = True
583             erTranAll = True
584         End If
585         If Trim(vKodeLgn) = "" Then
586             pesanIT = pesanIT + "Kode Pelanggan belum lengkap !"
587             WarnError = True
588             erTranAll = True
589         End If
            
        
590             If nDuoInv = True Then
                    
591                 nKodeSi = Trim(gKdPlgDealer2)
592                 isCust = Trim(text(40))
593                 isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(nKodeSi)  'cex dari ARMST apakah AFL / not AFL
594                 If isExAflAR = True Then  'AFL
595                     isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
596                     If isExAflSI = True Then
597                         isCust = "AFFPPN"
598                     Else
599                         MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
600                         GoTo vbError
601                     End If
602                 Else
603                     If isCust = "AFFPPN" Then  'cex jika tidak sesuai antara ARMST dan M_Asuransi
604                         MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
605                         GoTo vbError
606                     End If
607                     If isCust = "INT" Then
608                         isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah INT terdaftar
609                         If isExIntSI = False Then
610                             MsgBox "Data Customer '" & nKodeSi & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
611                             GoTo vbError
612                         End If
613                     End If
614                 End If
615                 If nDuoInv = True And woStat = "SBI" Then
616                   isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
617                   If isExAflSI = True Then
618                       isCust = "AFFPPN"
619                   Else
620                       isCust = "EXT"
621                   End If
622                 End If
                    Dim rsGL2 As ADODB.Recordset
623                 Select Case isCust
                        Case "EXT"
624                         Set rsGL2 = objO.Get_AkunGL_NEW("OTHERS", isDyna, False)
625                     Case Else  'INT/AFL
626                         Set rsGL2 = objO.Get_AkunGL_NEW(nKodeSi, isDyna, False)
627                 End Select

628                 If Not rsGL2.BOF Then
                        'vPiutangGL = cekKodeAkun(rsGL2.Fields(0).value)
                        'vPenjualanGL = cekKodeAkun(rsGL2.Fields(1).value)
                        'vPpnGL = cekKodeAkun(rsGL2.Fields(6).value)
                        'vMeteraiGL = cekKodeAkun(rsGL2.Fields(9).value)
                        
629                     vPiutangGL = (rsGL2.Fields(0).value)
630                     vPenjualanGL = (rsGL2.Fields(1).value)
631                     vPpnGL = (rsGL2.Fields(6).value)
632                     vMeteraiGL = (rsGL2.Fields(9).value)
633                 End If
634             End If
                

            ' BUAT ARTRN
635             sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                      & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                      & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                      & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & "', 'Q'," _
                      & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                      & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                      & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "'," _
                      & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
                      & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
636             Call objHT.InsertData_ACC(sSQL1)
                
                'PPN
637             If vTPPN <> 0 Then
638                 sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                          & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                          & " JumlahTrn, Kodeuser, DelIndex, Kodesumber,nowo, nopolisi, tglEntry,kodejurnal) values " _
                          & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & "', 'S'," _
                          & " '" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Inv : " & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "'," _
                          & " " & SimpanKoma2Titik(vTPPN * -1) & ", '" & gSA & "'," _
                          & " '" & vDelIndex & "','" & gKode & "','" & Trim(text(43)) & "','" & Trim(text(16)) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
639                 Call objHT.InsertData_ACC(sSQL1)
640             End If
            
                'Update Armst Saldo berjalan
641             sSQL1 = "Update Armst set SaldoJalan=SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(text(2)) & "'"
642             Call objHT.UpdateData_ACC(sSQL1)
            'End If
            
            'Membuat transaksi di GL
643             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
644             Call objHT.InsertData_ACC(sSQL1)
                
645             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(SpyRound((vTPenjNonMet))) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
646             If Log.dpplain = 1 Then
                  
647               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(SpyRound((vTPenjNonMet * 11) / 12)) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                      & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                      
648             End If
649             Call objHT.InsertData_ACC(sSQL1)
650             sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                      & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                      & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                      & " " & SimpanKoma2Titik(vTPenjNonMet - SpyRound((vTPenjNonMet * 11) / 12)) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                      & " '" & vDelIndex & "','SNL : " & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
651             If Log.dpplain = 1 Then Call objHT.InsertData_ACC(sSQL1)
                
652             If vTMeterai <> 0 Then
653                 sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
654                 Call objHT.InsertData_ACC(sSQL1)
655             End If
                
656             If vTPPN <> 0 Then
657                 sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & IIf(nDuoInv, "-B", "") & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
658                 Call objHT.InsertData_ACC(sSQL1)
659             End If
            
660             sSQL1 = "update glmst set JbulanIni = JbulanIni -" & (vTotal + vTMeterai) & ", " _
                      & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                      & " where KodeGl = '" & vPiutangGL & "'"
661             Call objHT.UpdateData_ACC(sSQL1)
                
662             sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                      & " where KodeGl = '" & vPenjualanGL & "'"
663             Call objHT.UpdateData_ACC(sSQL1)
                
664             sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTMeterai) & ", " _
                      & " JDebitBi = JDebitBi +" & vTMeterai & " " _
                      & " where KodeGl = '" & vMeteraiGL & "'"
665             Call objHT.UpdateData_ACC(sSQL1)
                
666             sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPPN) & ", " _
                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPPN) & " " _
                      & " where KodeGl = '" & vPpnGL & "'"
667             Call objHT.UpdateData_ACC(sSQL1)
            'End If
668     End If
        
        
669           If isFOP Then
                  Dim rsGL3 As ADODB.Recordset
                   
670               Set rsGL3 = objO.Get_AkunGL_NEW(NulSID(rsInvoiceC!FS_KD_PLG), isDyna, False)
671               vPiutangGL = ""
672               vPenjSubletGL = ""
673               vPenjPartGL = ""
674               vPenjOlieGL = ""
675               vPpnGL = ""
676               vPpnPartGL = ""
677               vMeteraiGL = ""
678               vHPPSublet = ""
679               vPDPSublet = ""
680               If Not rsGL3.BOF Then
681                   vPiutangGL = cekKodeAkun(rsGL3.Fields(0).value)
682                   vPenjSubletGL = cekKodeAkun(rsGL3.Fields(2).value)
683                   vPenjPartGL = cekKodeAkun(rsGL3.Fields(3).value)
684                   vPenjOlieGL = cekKodeAkun(rsGL3.Fields(4).value)
685                   vPpnGL = cekKodeAkun(rsGL3.Fields(6).value)
686                   vPpnPartGL = cekKodeAkun(rsGL3.Fields(7).value)
687                   vMeteraiGL = cekKodeAkun(rsGL3.Fields(9).value)
688                   vHPPSublet = cekKodeAkun(rsGL3.Fields(10).value)
689                   vPDPSublet = cekKodeAkun(rsGL3.Fields(11).value)
690               End If
                  
691               If vPiutangGL = "" Or vPenjSubletGL = "" Or vPenjPartGL = "" Or vPenjOlieGL = "" _
                      Or vPpnGL = "" Or vPpnPartGL = "" Or vMeteraiGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Then
692                   pesanIT = pesanIT + "Kode GL belum lengkap !"
693                   GoTo vbError
694               End If
                  
695               vKodeLgn = NulSID(rsInvoiceC!FS_KD_PLG)
696               vTPenjPart = 0
697               SumPart = 0
698               SumOlie = 0
699               SumAccs = 0
700               SumPartFull = 0
701               SumOlieFull = 0
702               SumAccsFull = 0
703               vPaPPN = 0
704               vOlPPN = 0
705               vAcPPN = 0
                  Dim tamPartGenuine  As Double
                  Dim tamOliGenuine   As Double
                  Dim subSubletTAM    As Double
                  
706               subSubletTAM = SpyRound(CDbl(rsInvoiceC!sblt), 0)
707               tamPartGenuine = SpyRound(CDbl(rsInvoiceC!PART_G) + CDbl(rsInvoiceC!PART_NG), 0)
708               tamOliGenuine = SpyRound(CDbl(rsInvoiceC!OIL_G) + CDbl(rsInvoiceC!OIL_NG), 0)
709               SumPart = SpyRound(tamPartGenuine * nilaiDppInv / nilaiTotalInv, 0)
710               SumPartFull = tamPartGenuine
711               SumOlie = SpyRound(tamOliGenuine * nilaiDppInv / nilaiTotalInv, 0)
712               SumOlieFull = tamOliGenuine
713               SumAccs = 0
714               vTPenjualan = SpyRound(SumPartFull + SumOlieFull + subSubletTAM, 0)
715               vTPenjPart = SumPart + SumOlie + SumAccs + SpyRound((subSubletTAM * nilaiDppInv / nilaiTotalInv), 0)
716               vJSPPN = subSubletTAM - SpyRound((subSubletTAM * nilaiDppInv / nilaiTotalInv), 0)
717               vTPenjSublet = Round((subSubletTAM * nilaiDppInv / nilaiTotalInv), 0)
718               vPaPPN = SumPartFull - SumPart
719               vOlPPN = SumOlieFull - SumOlie
720               vAcPPN = SumAccsFull - SumAccs
                  
721               If vTPenjualan - (vTPenjPart + vJSPPN + vOlPPN + vPaPPN + vAcPPN) > 0 Then
722                   If vJSPPN > 0 Then
723                       vJSPPN = vJSPPN + (vTPenjPart - (vTPenjPart + vJSPPN + vOlPPN + vPaPPN + vAcPPN))
724                   Else
725                       vPaPPN = vPaPPN + (vTPenjPart - (vTPenjPart + vJSPPN + vOlPPN + vPaPPN + vAcPPN))
726                   End If
727               ElseIf vTPenjualan - (vTPenjPart + vJSPPN + vOlPPN + vPaPPN + vAcPPN) < 0 Then
728                   If vJSPPN > 0 Then
729                       vJSPPN = vJSPPN - (vTPenjPart - (vTPenjPart + vJSPPN + vOlPPN + vPaPPN + vAcPPN))
730                   Else
731                       vPaPPN = vPaPPN - (vTPenjPart - (vTPenjPart + vJSPPN + vOlPPN + vPaPPN + vAcPPN))
732                   End If
733               End If
                  
734               If erTranAll = True Then GoTo vbError ' Exit Sub
          
735               Call objHT.InsertData_ACC(sSQL1)
             
                    
736                sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                          & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) " _
                          & "values " _
                          & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & "', 'Q'," _
                          & " '" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit No. Inv : " & Trim(gKodeFaktur) & Trim(Trim(text(0))) & "-C" & Trim(gKodeRetur) & "'," _
                          & " " & SimpanKoma2Titik(vTPenjPart * -1) & ", '" & gSA & "'," _
                          & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(43))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
737               Call objHT.InsertData_ACC(sSQL1)
                  
                  
                  'PPN
738               If (vPaPPN + vOlPPN + vJSPPN) <> 0 Then
                   
739                   sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                              & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                              & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                              & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & "', 'S'," _
                              & " '" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
                              & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                              & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Inv : " & Trim(Trim(text(0))) & "'," _
                              & " " & SimpanKoma2Titik(SpyRound((vPaPPN + vOlPPN + vJSPPN), 0) * -1) & ", '" & gSA & "'," _
                              & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(43))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
740                   Call objHT.InsertData_ACC(sSQL1)
741               End If
                  
                  
                  
742               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPenjualan) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
743               Call objHT.InsertData_ACC(sSQL1)
                  
                  ''PENJUALAN SUBLET
744               If vTPenjSublet <> 0 Then
745                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(vTPenjSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
746                   Call objHT.InsertData_ACC(sSQL1)
747                   NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
748               End If
                  
                  ''ppn bengkel
749               If vJSPPN <> 0 Then
750                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                  & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                  & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                  & " " & SimpanKoma2Titik(vJSPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                  & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
751                   Call objHT.InsertData_ACC(sSQL1)
752               End If
                  
                  
753               If SumPart <> 0 Then
754                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SumPart) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & " - PART','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
755                   Call objHT.InsertData_ACC(sSQL1)
756                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumPart) & ", " _
                              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumPart) & " " _
                              & " where KodeGl = '" & vPenjPartGL & "'"
757                   Call objHT.InsertData_ACC(sSQL1)
758               End If
759               If SumOlie <> 0 Then
760                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vPenjOlieGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(SumOlie) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & " - OLI','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
761                   Call objHT.InsertData_ACC(sSQL1)
762                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(SumOlie) & ", " _
                              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(SumOlie) & " " _
                              & " where KodeGl = '" & vPenjOlieGL & "'"
763                   Call objHT.InsertData_ACC(sSQL1)
764               End If

765               If (SumPart + SumOlie) <> 0 Then
766                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(43))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
767                   Call objHT.InsertData_ACC(sSQL1)

768                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
                              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0)) & " " _
                              & " where KodeGl = '" & vPpnPartGL & "'"
769                  Call objHT.InsertData_ACC(sSQL1)
770               End If
771               vTPembSublet = 0
                  ' vTPenjSublet = 0
772                Set rsSearch = objO.FetchData("SELECT PKP,NAMABARANG,NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, NONJURNAL FROM JOBSUBLET a WITH(NOLOCK) INNER JOIN T_SUPPLIER B WITH(NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(text(43)) & "' and no_inv like '%-C' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
773                If Not rsSearch.EOF Then
774                    Do While Not rsSearch.EOF
                                               
775                        nonJurnal = IIf(Trim(rsSearch("NONJURNAL")) = 1, True, False)
776                        If Not nonJurnal Then
777                            vTotalS = Round(CDbl(rsSearch("HPP")))
778                            vTPpnS = Round(vTotalS * (nilaiPPnInv / nilaiTotalInv), 0)
779                            vTBeliS = Round(vTotalS - vTPpnS, 0)
780                            vKeterangan1 = vKeterangan & " SUBLET" '(" & rsSearch("NAMABARANG") & ")"   'No. WO :" & Trim(Text(0))
781                            If Trim(rsSearch("PKP")) = "1" Then
782                                vPpnSGL = ""
783                                vTPpnS = 0
784                                vTotalS = Round(vTBeliS, 0)
785                            End If
                               'vTPenjSublet = vTPenjSublet + vTotalS
786                            vTPembSublet = vTPembSublet + vTotalS
787                        End If
788                        rsSearch.MoveNext
789                    Loop
790                End If
791               If vTPembSublet <> 0 Then
792                 sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vHPPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
793                 Call objHT.InsertData_ACC(sSQL1)
                    
794                 sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(0)) & "-C" & Trim(gKodeRetur) & "', '" & vPDPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(43)) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
795                 Call objHT.InsertData_ACC(sSQL1)
                    
                    ' HPP SUBLET TERHADAP PDP
796                 sSQL1 = "update glmst set JbulanIni = JbulanIni -" & (vTPembSublet) & ", " _
                          & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                          & " where KodeGl = '" & vHPPSublet & "'"
797                 Call objHT.InsertData_ACC(sSQL1)
                    
798                 sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPembSublet) & ", " _
                          & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                          & " where KodeGl = '" & vPDPSublet & "'"
799                 Call objHT.InsertData_ACC(sSQL1)
800               End If
801           End If
802       Exit Sub
vbError:
803       erTranAll = True
804       Exit Sub

         'On Error GoTo 0
805      Exit Sub

SimpanACCT_invoice_Error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
806   varErlError = Erl
807   varErrNumber = err.Number
      '' varErlError,varErrNumber
808       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure SimpanACCT_invoice of Form frmBtlInv")
End Sub


Private Sub loadJobDetailMultiple(noWO As String, noRangka As String)
          Dim rsHlp   As ADODB.Recordset
          Dim rsHlp2  As ADODB.Recordset
          Dim rsHlp3  As ADODB.Recordset
          Dim rsHlp4  As ADODB.Recordset
          
          Dim nilaiSubTotal   As Double
          
          
1        On Error GoTo loadJobDetailMultiple_Error

2         sSQL = "select * from [t_invoicejobdetail] with (nolock) where no_wo = '" & NulSID(text(43)) & "' and no_chassis = '" & NulSID(noRangka) & "' and fs_log_del = ''"
3         Set rsJobDetail = objO.FetchData(sSQL)
4         If Not rsJobDetail.EOF Then
5             rsJobDetail.MoveFirst
6             If rsJob.State = adStateOpen Then rsJob.Close
7             If rsJob.State = adStateClosed Then doCreateJob
8             Do While Not rsJobDetail.EOF
9                 With rsJob
10                    If NulSID(rsJobDetail!id_program) <> "" Then
11                        If objBilling.isDuoInvoice(NulSID(rsJobDetail!id_program)) Then
12                            nDuoInv = True
13                            If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
14                                Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
15                            Else
16                                Set rsHlp = objBilling.takeInvoiceName(NulSID(noRangka))
17                            End If
                              ''Set rsHlp = objBilling.takeInvoiceName(rsInvoice("NO_CHASSIS"))
18                            If Not rsHlp.EOF Then
19                                gKdPlgDealer2 = Nul(NulSID(rsHlp("KDPLG")))
20                                gKodePlg2 = Nul(NulSID(rsHlp("KDPLG")))
21                                gInvCust2 = Nul(NulSID(rsHlp("FS_NM_DEALER")))
22                                gInvAlmt2 = Nul(NulSID(rsHlp("FS_ALAMAT")))
23                                gInvKota2 = Nul(NulSID(rsHlp("FS_KOTA")))
24                                gInvNPWP2 = Nul(NulSID(rsHlp("FS_NPWP")))
                                  
25                                If NulSID(rsHlp("INTERNAL")) = 1 Then
26                                    If NulSID(gNpwp) = NulSID(gInvNPWP2) Then
27                                        vTipeCustomer = "AFFNONPPN"
28                                    Else
29                                        vTipeCustomer = "AFFPPN"
30                                    End If
31                                Else
32                                    vTipeCustomer = "EXT"
33                                End If
34                            End If
35                            Call giveDiscount("100", text(21), False)
                              '' 20200107 catatan maks 60
36                            text(30) = Left(NulSID("A/N " & text(4) & "  [" & text(2) & "]"), 60)
37                        End If
38                    End If
39                    namaProgram = ""
40                   .AddNew
41                   !job = NulSID(rsJobDetail!rep)
42                   !Rate = CDbl(NulSID(rsJobDetail!rte))
43                   If NulSID(rsJobDetail!isFP) = 1 Then
44                    !hrg = objO.FetchData_Double("select price from [" & dbReplikasiHO & "]..HO_FIXPRICE where id_FP = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
45                   Else
46                    !hrg = CDbl(rsJobDetail!offn)
47                   End If
                     '!addJob = IIf(NulSID(rsJobDetail!addJob) = 1, True, False)
48                   !disc = CDbl(rsJobDetail!disc)
49                   !DISCRP = IIf(rsJobDetail!isFP, 0, Round((!hrg) - CDbl(rsJobDetail!Price), 0)) 'IIf(rsJobDetail!isFP, 0, Round((!hrg * !Rate) - CDbl(rsJobDetail!Price), 0))
50                   If NulSID(rsJobDetail!isFP) = 1 Then
51                    !STLDISC = !hrg
52                   Else
53                    !STLDISC = Round(CDbl(rsJobDetail!Price), 0)
54                   End If
55                   If NulSID(rsJobDetail!FREEC) = 1 Then
56                       !FREEC = True
57                   Else
58                       !FREEC = False
59                   End If
60                   !idProgram = NulSID(rsJobDetail!id_program)
                     
                     '!isFP = IIf(NulSID(rsJobDetail!isFP) = 1, True, False)
61                   If NulSID(rsJobDetail!isFP) = 1 Then
62                    namaProgram = objO.FetchData_String("select namaFP from [" & dbReplikasiHO & "]..HO_FIXPRICE where id_FP = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
63                   Else
64                    namaProgram = objO.FetchData_String("select PROGRAM from [" & dbReplikasiHO & "]..HO_PROGRAM where id_Program = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
65                   End If
66                   !program = namaProgram
67                   If NulSID(rsJobDetail!isFP) = 1 And !FREEC Then
68                    text(22) = text(22) + !hrg
69                   ElseIf nDuoInv = False Then
70                    text(22) = text(22) + !STLDISC
71                   ElseIf !FREEC Then
72                    text(22) = text(22) + !STLDISC
73                   ElseIf !FREEC = False Then
74                    text(22) = text(22) + 0 '!STLDISC
75                   End If
76                   subJob = subJob + IIf(!FREEC, 0, CDbl(!hrg))
77                   .Update
78                End With
                    ''subJob + (CDbl(nHarga) * CDbl(nilaiRate))
79                totRate = totRate + CDbl(nilaiRate)
80                rsJobDetail.MoveNext
81            Loop
82        End If

83       On Error GoTo 0
84       Exit Sub

loadJobDetailMultiple_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
85        varErlError = Erl
86        varErrNumber = err.Number
87        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure loadJobDetailMultiple of Form frmBtlInv")
End Sub
 
Private Sub simpanJobDetail()
1     On Error GoTo simpanJobDetail_Error
           
              '' simpan repl
2             sSQL = "UPDATE [" & dbReplikasiHO & "]..[jobdetail] SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = getdate() WHERE [no_wo] = '" & NulSID(text(43)) & "' and [no_chassis] = '" & NulSID(text(17)) & "' and [CABANG] = '" & gDb & "' and FS_LOG_DEL = ''"
3             Call objHT.InsertData_Bengkel(sSQL)
              
              ''simpan bengkel
4             sSQL = "UPDATE [t_invoicejobdetail] SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = getdate() WHERE [no_wo] = '" & NulSID(text(43)) & "' and [no_chassis] = '" & NulSID(text(17)) & "' and FS_LOG_DEL = ''"
5             Call objHT.InsertData_Bengkel(sSQL)
                   
6         Exit Sub
simpanJobDetail_Error:
7         varErlError = Erl
8         varErrNumber = err.Number
9         erTranAll = True
10        ErrMsg.judul = ""
11        ErrMsg.message = ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure simpanJobDetail of Form frmBtlInv")
12        ErrMsg.show
End Sub


Private Sub Cetak3()
          Dim rpt As New RptFromRetur1
          Dim gInvUrai(8) As String
          Dim gInvRt(8) As String
          Dim gInvRp(8) As String
          Dim gInvRup(8) As String
          Dim gcat As String
          Dim gDisc As String
          Dim gDiscount(8) As String
          Dim gRateAdd As Double
          Dim gRpAdd As Double
          Dim gRupAdd As Double
          Dim adaAddJob As Boolean
          
1        On Error GoTo Cetak2_Error
         
         '' 20190911 nilai terbilang diambil dari nilai t_invoice
         '' --------------------------------------------------------------
2        subJob = objO.FetchData_Double("SELECT total FROM t_invoice WITH (NOLOCK) WHERE noinv = '" & NulSID(text(20)) & "-C'")
3         If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
4           If Trim(gNpwp) = Trim(gInvNPWP2) Then
               '' tanpa ppn
5              subJob = subJob * nilaiDppInv / nilaiTotalInv
6           End If
7         End If

8         gKodePlg = IIf(gKodePlg2 = "", text(2), gKodePlg2)
9         gInvNo = text(20) & "-C" & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
10        gNoInvoice = text(0) & "-C"
11        gInvTgl = vTglInv 'TglDDMMYYYY(isnow) '
12        gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
13        gInvAlmt = IIf(gInvAlmt2 = "", text(5), gInvAlmt2)
14        gInvKota = IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota) 'vNoTelp
15        gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
16        gInvNoPolisi = text(16)
17        gInvModel = text(14) & " (" & text(13) & ")"
18        gInvNoMesin = text(18)
19        gInvNoChas = text(17)
20        gInvNoOrder = text(43)
21        gInvSA = text(19)
22        gInvStatus = text(40)
23        gInvMech = text(38)
24        gInvPrint = vJam
25        gInvBilling = gSA
          
         
26        For i = 1 To 7
27            gInvRt(i) = ""
28            gInvUrai(i) = ""
29            gInvRup(i) = ""
30        Next
31        gDisc = CDbl(text(21))
32        gRateAdd = 0
33        gRupAdd = 0
34        gRpAdd = 0
35        adaAddJob = False
          
36        If rsJob.State <> adStateClosed Then
37            rsJob.MoveFirst
38            i = 1
39            Do While Not rsJob.EOF
40                If Not rsJob("ADDJOB") Then
41                    gInvUrai(i) = rsJob("JOB")
42                    gInvRt(i) = rsJob("RATE")
43                    gInvRp(i) = Format(CDbl(rsJob("HRG")), "#,#0")
44                    gInvRup(i) = Format(CDbl(rsJob("HRG")) * ((100 - rsJob("DISC")) / 100), "#,#0")
45                    gDiscount(i) = CDbl(rsJob("DISC"))
46                    i = i + 1
47                Else
48                    adaAddJob = True
49                    gRateAdd = gRateAdd + rsJob("RATE")
50                    gRpAdd = gRpAdd + CDbl(rsJob("HRG"))
51                    gRupAdd = gRupAdd + (CDbl(rsJob("HRG")) * (100 - rsJob("DISC")) / 100)
                      'gDiscount(i) = CDbl(rsJob("DISC"))
52                End If
53                rsJob.MoveNext
54            Loop
55        End If

56        gGInvRp = Format(subJob, "#,##0")
57        gInvTRp1 = 0    'Format(Text(22), "#,##0")
58        gInvTRp2 = 0    'Format(Text(27), "#,##0")
59        gInvTRp3 = 0    'Format(Text(24), "#,##0")
60        gInvTRp4 = 0    'Format(Text(25), "#,##0")
61        gInvTRp5 = 0    'Format(Text(28), "#,##0")
          'gInvTRp6 = Format(.Text(6), "#,##0")
62        gInvTRp7 = 0    'Format(Text(42), "#,##0")
63        gGInvTRp = Format(subJob, "#,##0")
64        gInvTerbilang = Angka2Huruf(SpyRound(subJob), 1)
65        gInvSS = text(32)
66        gInvCatatan = text(30)

67        gInvSM = gSecmanager
68        gcat = ""
          
69        rpt.Field67.DataValue = vNoRetur
70        rpt.Field66.DataValue = gNoInvoice & gKodeRetur & "/" & text(40)
71        rpt.Field65.DataValue = gKodePlg
72        rpt.Field39.DataValue = gInvNo
73        rpt.Field40.DataValue = gInvTgl

74        rpt.Field1.DataValue = gInvCust
75        rpt.Field4.DataValue = gInvAlmt
76        rpt.Field28.DataValue = gInvKota
77        rpt.Field7.DataValue = gInvNPWP

78        rpt.Field2.DataValue = gInvNoPolisi
79        rpt.Field5.DataValue = gInvModel
80        rpt.Field12.DataValue = gInvNoMesin
81        rpt.Field14.DataValue = gInvNoChas
          
82        rpt.Field29.DataValue = gInvNoOrder
83        rpt.Field30.DataValue = gInvSA

84        rpt.Field31.DataValue = gInvStatus
85        rpt.Field32.DataValue = gInvMech
86        rpt.Field68.DataValue = gcat

87        rpt.Field33.DataValue = gInvPrint
88        rpt.Field38.DataValue = gInvBilling
          
89        rpt.Field16.DataValue = gInvUrai(1)
90        rpt.Field17.DataValue = gInvUrai(2)
91        rpt.Field18.DataValue = gInvUrai(3)
92        rpt.Field19.DataValue = gInvUrai(4)
93        rpt.Field20.DataValue = gInvUrai(5)
94        rpt.Field21.DataValue = gInvUrai(6)
95        rpt.Field22.DataValue = gInvUrai(7)
          
96        rpt.Field41.DataValue = gInvRt(1)
97        rpt.Field42.DataValue = gInvRt(2)
98        rpt.Field43.DataValue = gInvRt(3)
99        rpt.Field44.DataValue = gInvRt(4)
100       rpt.Field45.DataValue = gInvRt(5)
101       rpt.Field46.DataValue = gInvRt(6)
102       rpt.Field47.DataValue = gInvRt(7)
          
103       rpt.Field48.DataValue = gInvRup(1)
104       rpt.Field49.DataValue = gInvRup(2)
105       rpt.Field50.DataValue = gInvRup(3)
106       rpt.Field51.DataValue = gInvRup(4)
107       rpt.Field52.DataValue = gInvRup(5)
108       rpt.Field53.DataValue = gInvRup(6)
109       rpt.Field54.DataValue = gInvRup(7)
          
110       rpt.Field55.DataValue = gInvTRp1
111       rpt.Field56.DataValue = gInvTRp2
112       rpt.Field57.DataValue = gInvTRp3
113       rpt.Field58.DataValue = gInvTRp4
114       rpt.Field59.DataValue = gInvTRp5
115       rpt.Field60.DataValue = gInvTRp6
116       rpt.Field61.DataValue = gInvTRp7
          
117       rpt.Field23.DataValue = gGInvRp
118       rpt.Field27.DataValue = gGInvTRp
119       rpt.Field25.DataValue = gInvTerbilang
          
120       rpt.Field62.DataValue = gInvSS
121       rpt.Field63.DataValue = gInvCatatan
122       rpt.Field64.DataValue = gInvSM
          '9=A4
          'kwarto

      '    rpt.Printer.PaperSize = 9
123       rpt.Printer.PaperSize = 256
124       rpt.Printer.PaperHeight = 1440 * 11
125       rpt.Printer.PaperWidth = 1440 * 8.5
          
126       rpt.Printer.Orientation = 1
          
127       If gPreviewCetak = True Then
128           rpt.PrintReport False
129           Unload rpt
130       Else
131           rpt.show 'vbModal
132       End If

         'On Error GoTo 0
133      Exit Sub

Cetak2_Error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
134   varErlError = Erl
135   varErrNumber = err.Number
      '' varErlError,varErrNumber
136       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure Cetak2 of Form frmBtlInv")
End Sub
