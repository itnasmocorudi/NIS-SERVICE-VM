VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOrderComplain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Complain User"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13485
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   13485
   Begin VB.Frame FrameKeluhDmsNonDSS 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   1320
      TabIndex        =   88
      Top             =   2500
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   680
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   91
         Top             =   3230
         Width           =   5655
      End
      Begin VB.TextBox txtNextJob 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   90
         Top             =   3230
         Width           =   5655
      End
      Begin VB.TextBox txtFI 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   680
         Width           =   5655
      End
      Begin VB.Line Line18 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3840
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
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
         Left            =   4320
         TabIndex        =   98
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         X1              =   7560
         X2              =   11400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         X1              =   7560
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11655
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
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
         Left            =   4680
         TabIndex        =   97
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   11640
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   375
         Width           =   3615
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Keluhan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   2900
         Width           =   3615
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Final Inspection dan Test Drive : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   94
         Top             =   375
         Width           =   5655
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Next Job : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   93
         Top             =   2900
         Width           =   3615
      End
   End
   Begin VB.CommandButton tblSimpanComplain 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Simpan Complain"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   9240
      Width           =   1890
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   9240
      Width           =   1890
   End
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
      Height          =   300
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13800
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame framKeluh 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1320
      TabIndex        =   74
      Top             =   2500
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox txtKeluhan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtPermintaan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   480
         Width           =   5655
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   11640
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
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
         Left            =   4440
         TabIndex        =   82
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   12885
      Begin VB.CommandButton cmdKeluh 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1680
         Width           =   2130
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   4
         Top             =   480
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
         Index           =   4
         Left            =   930
         TabIndex        =   5
         Top             =   750
         Width           =   3735
         _ExtentX        =   6588
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
         Index           =   5
         Left            =   930
         TabIndex        =   6
         Top             =   1035
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
         TabIndex        =   7
         Top             =   1350
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
         TabIndex        =   8
         Top             =   1680
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
         Left            =   5550
         TabIndex        =   9
         Top             =   450
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
         Left            =   5550
         TabIndex        =   10
         Top             =   750
         Width           =   2295
         _ExtentX        =   4048
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
         Left            =   5550
         TabIndex        =   11
         Top             =   1350
         Width           =   3135
         _ExtentX        =   5530
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
         Left            =   5550
         TabIndex        =   12
         Top             =   1050
         Width           =   3135
         _ExtentX        =   5530
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
         Left            =   7890
         TabIndex        =   13
         Top             =   750
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   14
         Top             =   480
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
         Index           =   16
         Left            =   9510
         TabIndex        =   15
         Top             =   450
         Width           =   1170
         _ExtentX        =   2064
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
         Left            =   9510
         TabIndex        =   16
         Top             =   735
         Width           =   3210
         _ExtentX        =   5662
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
         Left            =   7470
         TabIndex        =   17
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   3480
         TabIndex        =   18
         Top             =   1680
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
         Left            =   6090
         TabIndex        =   19
         Top             =   1680
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
         Left            =   9510
         TabIndex        =   20
         Top             =   1035
         Width           =   3210
         _ExtentX        =   5662
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
         Left            =   3840
         TabIndex        =   21
         Top             =   480
         Width           =   825
         _ExtentX        =   1455
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   40
         Left            =   2505
         TabIndex        =   22
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
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
         Index           =   34
         Left            =   9510
         TabIndex        =   23
         Top             =   1350
         Width           =   3210
         _ExtentX        =   5662
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
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   8640
         X2              =   12720
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   8640
         X2              =   12720
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label74 
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
         Left            =   4320
         TabIndex        =   81
         Top             =   45
         Width           =   3675
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   240
         Y2              =   240
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
         Left            =   8820
         TabIndex        =   38
         Top             =   1080
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
         Left            =   5505
         TabIndex        =   37
         Top             =   1725
         Width           =   435
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
         Left            =   6720
         TabIndex        =   36
         Top             =   480
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
         Left            =   8850
         TabIndex        =   35
         Top             =   780
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
         Left            =   8715
         TabIndex        =   34
         Top             =   495
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
         TabIndex        =   33
         Top             =   480
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
         Left            =   4920
         TabIndex        =   32
         Top             =   1365
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
         Left            =   4980
         TabIndex        =   31
         Top             =   765
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
         Left            =   4755
         TabIndex        =   30
         Top             =   495
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
         TabIndex        =   29
         Top             =   1695
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
         TabIndex        =   28
         Top             =   1395
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
         TabIndex        =   27
         Top             =   1095
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
         TabIndex        =   26
         Top             =   795
         Width           =   405
      End
      Begin VB.Label Label16 
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
         Left            =   8955
         TabIndex        =   25
         Top             =   1395
         Width           =   345
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
         Left            =   3060
         TabIndex        =   24
         Top             =   1725
         Width           =   315
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   12855
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3375
      LargeChange     =   800
      Left            =   12960
      Max             =   2600
      SmallChange     =   600
      TabIndex        =   67
      Top             =   3600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   480
      ScaleHeight     =   3315
      ScaleWidth      =   12795
      TabIndex        =   61
      Top             =   3600
      Width           =   12855
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8055
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   12375
         Begin VB.CommandButton cmdCetak 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cetak"
            Height          =   390
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   8520
            Width           =   1890
         End
         Begin TrueOleDBGrid70.TDBGrid JobGrid 
            Height          =   2295
            Left            =   0
            TabIndex        =   63
            Top             =   240
            Width           =   12255
            _ExtentX        =   21616
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
            Columns(4).NumberFormat=   "Standard"
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
            Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(40)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(44)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(48)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(51)=   "Column(10).Order=11"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.locked=-1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.locked=-1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=32,.parent=13,.locked=0"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   2055
            Left            =   0
            TabIndex        =   65
            Top             =   2640
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   3625
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
            Columns(3).NumberFormat=   "Standard"
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
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
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
            AllowDelete     =   -1  'True
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
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
         Begin TrueOleDBGrid70.TDBGrid subletGrid 
            Height          =   1815
            Left            =   0
            TabIndex        =   66
            Top             =   4800
            Width           =   12255
            _ExtentX        =   21616
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
            Columns(7)._VlistStyle=   0
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
            AllowDelete     =   -1  'True
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   480
      TabIndex        =   39
      Top             =   2400
      Width           =   12885
      Begin VB.ComboBox cboPelanggan 
         Height          =   315
         ItemData        =   "frmOrderComplain.frx":0000
         Left            =   8880
         List            =   "frmOrderComplain.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   800
         Width           =   735
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Regular"
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
         Index           =   0
         Left            =   840
         TabIndex        =   42
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Program"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warranty"
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
         Index           =   2
         Left            =   2880
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   20
         Left            =   4920
         TabIndex        =   43
         Top             =   800
         Width           =   1170
         _ExtentX        =   2064
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
         Index           =   19
         Left            =   6795
         TabIndex        =   44
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
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
         Left            =   120
         TabIndex        =   45
         Top             =   525
         Width           =   1410
         _ExtentX        =   2487
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   36
         Left            =   120
         TabIndex        =   46
         Top             =   810
         Width           =   3855
         _ExtentX        =   6800
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
         Index           =   37
         Left            =   4920
         TabIndex        =   47
         Top             =   170
         Width           =   1170
         _ExtentX        =   2064
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
         Left            =   6795
         TabIndex        =   48
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
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
         Index           =   31
         Left            =   6795
         TabIndex        =   49
         Top             =   795
         Width           =   1170
         _ExtentX        =   2064
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
         Index           =   43
         Left            =   4920
         TabIndex        =   50
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
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
         Index           =   28
         Left            =   8880
         TabIndex        =   77
         Top             =   165
         Width           =   3090
         _ExtentX        =   5450
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
         Index           =   32
         Left            =   8880
         TabIndex        =   78
         Top             =   480
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   270
         Index           =   44
         Left            =   9600
         TabIndex        =   84
         Top             =   825
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   476
         _Version        =   393216
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
      Begin VB.Label Label28 
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
         Left            =   8160
         TabIndex        =   85
         Top             =   860
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Phone"
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
         Left            =   8025
         TabIndex        =   80
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pembawa"
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
         Left            =   8085
         TabIndex        =   79
         Top             =   210
         Width           =   690
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
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
         Left            =   4200
         TabIndex        =   59
         Top             =   845
         Width           =   600
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
         Left            =   6300
         TabIndex        =   58
         Top             =   210
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
         TabIndex        =   57
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
         Left            =   4560
         TabIndex        =   56
         Top             =   215
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
         Left            =   6315
         TabIndex        =   55
         Top             =   525
         Width           =   375
      End
      Begin VB.Label lblPDS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   12120
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book"
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
         Left            =   6345
         TabIndex        =   53
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label19 
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
         Left            =   4200
         TabIndex        =   52
         Top             =   525
         Width           =   570
      End
      Begin VB.Label LblReturn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   12120
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2600
      TabIndex        =   68
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
      Mask            =   ">CCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   69
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
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
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin VB.TextBox txtComplain 
      Height          =   1815
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   99
      Top             =   7320
      Width           =   12015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COMPLAIN USER :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   100
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label JualNRM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   7800
      TabIndex        =   87
      Top             =   120
      Width           =   5490
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETURN JOB"
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
      Left            =   5280
      TabIndex        =   72
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   71
      Top             =   120
      Width           =   570
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
      Left            =   3015
      TabIndex        =   70
      Top             =   135
      Width           =   570
   End
End
Attribute VB_Name = "frmOrderComplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsSublet As New ADODB.Recordset
Dim objBilling As New BClsBilling
Dim newValid As Boolean
Dim gSpart As Boolean
Dim adaSublet As Boolean
Dim simpan As Boolean
Dim noInv As Boolean
Dim totRate As Double
Dim nWKT As String
Dim support As New clsFormDealer
Dim totalNilaiPart As Double

Dim isBaru As Boolean

Private Sub doCreatePart()
    If rsPart.State Then rsPart.Close
    'Set rsPart = Nothing
    With rsPart
        'If .State Then .Close
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "NAMAITEM", adVarChar, 200
        .Fields.Append "QTYOrder", adInteger
        .Fields.Append "QTY", adInteger
        .Fields.Append "QTYBO", adInteger
        .Fields.Append "QTYCANCEL", adInteger
        .Fields.Append "HARGA", adDouble
        .Fields.Append "OLI", adBoolean
        .Fields.Append "EXTDISC", adDouble
        .Open
    End With
End Sub

Private Sub doHeadPart()
    With partGrid
        .Columns(0).Width = 2000
        .Columns(0).Caption = "KODE PART"
        .Columns(0).DataField = "KODEITEM"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 2500
        .Columns(1).Caption = "NAMA PART"
        .Columns(1).DataField = "NAMAITEM"
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 1000
        .Columns(2).Caption = "QTY Order"
        .Columns(2).DataField = "QTYOrder"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 1000
        .Columns(3).Caption = "QTY Supply"
        .Columns(3).DataField = "QTY"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 1000
        .Columns(4).Caption = "QTY BO"
        .Columns(4).DataField = "QTYBO"
        .Splits(0).Columns(3).Locked = True
        On Error GoTo Column_Ini_Untuk_Grid_FillRate
        
        
        .Columns(5).Visible = True
        .Columns(5).Width = 1000
        .Columns(5).Caption = "HARGA"
        .Columns(5).DataField = "Harga"
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Visible = False
        .Columns(6).Caption = "OLI"
        .Columns(6).DataField = "OLI"
        .Splits(0).Columns(6).Locked = True
        
        .Columns(7).Width = 1000
        .Columns(7).Caption = "QTY CANCEL"
        .Columns(7).DataField = "QTYCANCEL"
        .Splits(0).Columns(7).Locked = True
Column_Ini_Untuk_Grid_FillRate:
    End With
End Sub

Private Sub doCreateJob()
    If rsJob.State Then rsJob.Close
    'Set rsJob = Nothing
    With rsJob
        'If .State Then .Close
        .Fields.Append "JOB", adVarChar, 50
        .Fields.Append "RATE", adDouble
        .Fields.Append "ACTRT", adDouble
        .Fields.Append "GRP", adVarChar, 6
        .Fields.Append "MECH", adVarChar, 10
        .Fields.Append "ADDJOB", adBoolean
        .Fields.Append "HRG", adDouble
        .Fields.Append "DISC", adDouble
        .Fields.Append "DISCRP", adDouble
        .Fields.Append "STLDISC", adDouble
        .Fields.Append "VALID", adVarChar, 10
        .Fields.Append "FREEC", adBoolean           ' TANDA JIKA DIBEBANKAN KE CUSTOMER
        .Open
    End With
End Sub

Private Sub doHeadJob()
    With JobGrid
        .Columns(0).Width = 4500
        .Columns(0).Caption = "JOB ORDER"
        .Columns(0).DataField = "JOB"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 800
        .Columns(1).Caption = "RATE"
        .Columns(1).DataField = "RATE"
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 0
        .Columns(2).Caption = "ACTRATE"
        .Columns(2).DataField = "ACTRT"
        .Columns(2).Visible = False
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 0
        .Columns(3).Caption = "GROUP"
        .Columns(3).DataField = "GRP"
        .Columns(3).Visible = False
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 0
        .Columns(4).Caption = "MEKANIK"
        .Columns(4).DataField = "MECH"
        .Columns(4).Visible = False
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 0
        .Columns(5).Caption = "Additional Job"
        .Columns(5).DataField = "ADDJOB"
        .Columns(5).Visible = False
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Width = 1200
        .Columns(6).Caption = "HARGA"
        .Columns(6).DataField = "HRG"
        .Splits(0).Columns(6).Locked = True
               
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
End Sub

Private Sub doHeadSublet()
    
    With subletGrid
        .Columns(0).Width = 1500
        .Columns(0).Caption = "NO NOTA"
        .Columns(0).DataField = "NONOTA"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 1800
        .Columns(1).Caption = "SUPPLIER"
        .Columns(1).DataField = "SUPPLIER"
        .Columns(1).Visible = True
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 3000
        .Columns(2).Caption = "NAMA BARANG"
        .Columns(2).DataField = "NAMABARANG"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 1500
        .Columns(3).Caption = "KODE SATUAN"
        .Columns(3).DataField = "KODESATUAN"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 800
        .Columns(4).Caption = "JUMLAH"
        .Columns(4).DataField = "QTY"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 800
        .Columns(5).Caption = "VALID"
        .Columns(5).DataField = "VALID"
        .Columns(5).Visible = False
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Width = 1500
        .Columns(6).Caption = "BELI"
        .Columns(6).DataField = "HARGA"
        .Columns(6).NumberFormat = "#,#0"
        .Splits(0).Columns(6).Locked = True
        
        .Columns(7).Width = 1500
        .Columns(7).Caption = "JUAL"
        .Columns(7).DataField = "PRICE"
        .Columns(7).NumberFormat = "#,#0"
        .Splits(0).Columns(7).Locked = True
    End With
End Sub

Private Function returnJob(labelReturn As Object) As Boolean
    returnJob = False
    If labelReturn.Visible Then
        returnJob = True
    End If
End Function

Private Function typeJob(statusJob As String) As String
    If statusJob = "W" Then
        typeJob = "WARRANTY"
    ElseIf statusJob = "P" Then
        typeJob = "PROGRAM"
    Else
        typeJob = "REGULAR"
    End If
End Function

Private Function giveDiscount(discountValue As String, discountText As Object, statusText As Boolean)
   ' discountText = discountValue
    'discountText.Enabled = statusText
    JobGrid.Splits(0).Columns(7).Locked = Not statusText
    JobGrid.Splits(0).Columns(8).Locked = Not statusText
End Function

Private Sub panggilWO(noWO As String)
      Dim rsHlp As ADODB.Recordset
      Dim rsHlp2 As ADODB.Recordset
      Dim rsHlp3 As ADODB.Recordset
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim fOth As Boolean
      'Dim vHitung As Double
      Dim nHarga As Double
          
1        On Error GoTo panggilWO_Error

2         gKdPlgDealer = ""
3         gKodePlg = ""
4         gInvCust = ""
5         gInvAlmt = ""
6         gInvKota = ""
7         gInvNPWP = ""
8         nDuoInv = False
          
9         Set rsTemp = objBilling.searchOrderDetail(noWO)
10        If Not rsTemp.EOF Then
              
              ' TIPE PEMBAYARAN
11            Set rsHlp = objBilling.searchPaymentType(noWO)
12            If Not rsHlp.EOF Then
13                tipeByr = Trim(rsHlp("PEMBAYARAN"))
14                gVerifikator = Trim(rsHlp("VERIFIKATOR"))
15                cboPelanggan = Trim(rsHlp("KARAKTER"))
16                Text(44) = Trim(rsHlp("DTLKARAKTER"))
17            End If
18            If rsHlp.State <> adStateClosed Then rsHlp.Close
              
19            Set rsHlp = objBilling.searchDtlOrder(rsTemp("NO_WO"), rsTemp("NO_CHASSIS"))
20            If Not rsHlp.EOF Then
21                nHeavy = IIf(rsHlp("HEAVYUSE") <> 0, True, False)
22            End If
23            If rsHlp.State <> adStateClosed Then rsHlp.Close
              
              ' TAMPIL DETAIL CUSTOMER
24            Text(2) = rsTemp("FS_KD_PLG")
25            Text(3) = rsTemp("FS_KD_WIL")
26            Text(4) = rsTemp("CUSTOMER")
27            Text(5) = Left(rsTemp("CUST_ADRES"), 60)
28            Text(6) = rsTemp("KOTA")
29            Text(7) = rsTemp("TELPON")
30            Text(8) = rsTemp("GSM")
31            Text(9) = rsTemp("CDMA")
32            Text(10) = rsTemp("FS_TH_MOBIL")
33            Text(11).Mask = ""
34            Text(11) = Nul(IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsTemp("DOTAM"), "/", "-"))))
35            Text(12) = rsTemp("MDL")
36            Text(13) = rsTemp("MODEL")
37            Text(14) = rsTemp("FMDL")
38            Text(15) = rsTemp("COLOR")
39            Text(16) = rsTemp("NO_POLISI")
40            Text(17) = rsTemp("NO_CHASSIS")
41            Text(18) = rsTemp("NO_ENGINE")
42            Text(19) = rsTemp("FS_LOG_ADD")
43            If rsTemp("NOINV") <> "" Then
44                Text(20) = rsTemp("NOINV")
45            End If
46            Text(31) = rsTemp("NOFAK")
              'text(33) = rsTemp("FS_RC")
47            Text(37) = rsTemp("KM")
48            Text(38) = rsTemp("MECH")
49            Text(40) = rsTemp("STATUS")
50            Text(43) = TglDDMMYYYY(rsTemp("DATE"))
51            nTHS = IIf(rsTemp("THS") = 1, True, False)
52            LblReturn.Visible = IIf(rsTemp("FN_RETURN") = "1", True, False)
53            lblPDS.Visible = IIf(rsTemp("PDS") = "1", True, False)
54            Text(28) = rsTemp("ANCUST")
55            Text(32) = rsTemp("TELPANCUST")
              
              ' CHECK DEFAULT DISCOUNT
              'text(30) = ""
56            Text(35).Visible = False
57            Text(36).Visible = False
58            vTipeCustomer = ""
59            myChassis = Trim(Text(17))
60            If returnJob(LblReturn) Then
61                If typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
62                    opWO(1).value = True
                       
63                    Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
64                    If Not rsHlp.EOF Then
65                        Text(35) = rsHlp("ID_PROGRAM")
66                        Text(36) = rsHlp("PROGRAM")
67                        Text(35).Visible = True
68                        Text(36).Visible = True
69                    End If
70                    If rsHlp.State <> adStateClosed Then rsHlp.Close
71                    If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
72                        nDuoInv = True
                          'RTJ WO BERKEMBANG
                           
73                        Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
74                        If Not rsHlp.EOF Then
75                            gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
76                            gKodePlg2 = Trim(rsHlp("KDPLG"))
77                            gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
78                            gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
79                            gInvKota2 = Trim(rsHlp("FS_KOTA"))
80                            gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                              
81                            If Trim(rsHlp("INTERNAL")) = 1 Then
82                                If Trim(gNpwp) = Trim(gInvNPWP2) Then
83                                    vTipeCustomer = "AFFNONPPN"
84                                Else
85                                    vTipeCustomer = "AFFPPN"
86                                End If
87                            Else
88                                vTipeCustomer = "EXT"
89                            End If
90                        End If
91                        If rsHlp.State <> adStateClosed Then rsHlp.Close
92                        Call giveDiscount("100", Text(21), False)
93                        Text(30) = "A/N " & Text(4) & "  [" & Text(2) & "]"
94                    End If
95                Else
                      ' HANYA RTJ
96                    Call giveDiscount("100", Text(21), False)
97                    vTipeCustomer = "INT"
98                End If
99            Else
100               If typeJob(Trim(rsTemp("TPSVC"))) = "WARRANTY" Then
101                   opWO(2).value = True
                       
102                   Set rsHlp = objBilling.searchWarrantyDetail(rsTemp("KODETIPE"))
103                   If Not rsHlp.EOF Then
104                       Text(35) = rsHlp("ID_WARRANTY")
105                       Text(36) = rsHlp("WARRANTY")
106                       Text(35).Visible = True
107                       Text(36).Visible = True
108                   End If
109                   If rsHlp.State <> adStateClosed Then rsHlp.Close
110                   Call giveDiscount("0", Text(21), True)
111                   Text(30) = "A/N " & Text(4) & "  [" & Text(2) & "]"
112                   Set rsHlp = objO.FetchData("SELECT * FROM T_PROFIL WITH (NOLOCK)")
113                   If Not rsHlp.EOF Then
114                       gKdPlgDealer = Trim(rsHlp("FN_HOKDPLG"))
115                       gKodePlg = Trim(rsHlp("FN_HOKDPLG"))
116                       gInvCust = Trim(rsHlp("FN_HO"))
117                       gInvAlmt = Trim(rsHlp("FN_HOADRES"))
118                       gInvKota = Trim(rsHlp("FN_HOKOTA"))
119                       gInvNPWP = Trim(rsHlp("FN_HONPWP"))
120                       vTipeCustomer = "AFFPPN"
121                   End If
122                   If rsHlp.State <> adStateClosed Then rsHlp.Close
123               ElseIf typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
124                   opWO(1).value = True
                       
125                   Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
126                   If Not rsHlp.EOF Then
127                       Text(35) = rsHlp("ID_PROGRAM")
128                       Text(36) = rsHlp("PROGRAM")
129                       Text(35).Visible = True
130                       Text(36).Visible = True
131                   End If
132                   If rsHlp.State <> adStateClosed Then rsHlp.Close
133                   If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
134                       nDuoInv = True
                          'SBI WO BERKEMBANG
135                       If rsTemp("FS_RC") = "R001" Then
                               
136                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
137                           If Not rsHlp.EOF Then
138                               gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
139                               gKodePlg2 = Trim(rsHlp("KDPLG"))
140                               gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
141                               gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
142                               gInvKota2 = Trim(rsHlp("FS_KOTA"))
143                               gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                                  
144                               If Trim(rsHlp("INTERNAL")) = 1 Then
145                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
146                                       vTipeCustomer = "AFFNONPPN"
147                                   Else
148                                       vTipeCustomer = "AFFPPN"
149                                   End If
150                               Else
151                                   vTipeCustomer = "EXT"
152                               End If
153                           End If
154                           If rsHlp.State <> adStateClosed Then rsHlp.Close
155                           Call giveDiscount("100", Text(21), False)
156                           Text(30) = "A/N " & Text(4) & "  [" & Text(2) & "]"
157                       Else
                               
158                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
159                           If Not rsHlp.EOF Then
160                               gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
161                               gKodePlg2 = Trim(rsHlp("KDPLG"))
162                               gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
163                               gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
164                               gInvKota2 = Trim(rsHlp("FS_KOTA"))
165                               gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                                  
166                               If Trim(rsHlp("INTERNAL")) = 1 Then
167                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
168                                       vTipeCustomer = "AFFNONPPN"
169                                   Else
170                                       vTipeCustomer = "AFFPPN"
171                                   End If
172                               Else
173                                   vTipeCustomer = "EXT"
174                               End If
175                           End If
176                           If rsHlp.State <> adStateClosed Then rsHlp.Close
177                           Call giveDiscount("100", Text(21), False)
                           '   text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
178                       End If
179                   Else
180                       Call giveDiscount(objBilling.takeDiscountProgram(rsTemp("KODETIPE")), Text(21), False)
181                   End If
182               ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
183                   If lblPDS.Visible Then
184                       Call giveDiscount("0", Text(21), False)
185                       Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
186                       If Not rsHlp.EOF Then
187                           gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
188                           gKodePlg2 = Trim(rsHlp("KDPLG"))
189                           gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
190                           gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
191                           gInvKota2 = Trim(rsHlp("FS_KOTA"))
192                           gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                              
193                           If Trim(rsHlp("INTERNAL")) = 1 Then
194                               If Trim(gNpwp) = Trim(gInvNPWP2) Then
195                                   vTipeCustomer = "AFFNONPPN"
196                               Else
197                                   vTipeCustomer = "AFFPPN"
198                               End If
199                           Else
200                               vTipeCustomer = "EXT"
201                           End If
202                       End If
203                       If rsHlp.State <> adStateClosed Then rsHlp.Close
204                   ElseIf rsTemp("FS_RC") = "R001" Then
                          ' HANYA SBI
205                       woStat = "SBI"
                           
206                       Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
207                       If Not rsHlp.EOF Then
208                           gKdPlgDealer = Trim(rsHlp("KDPLG"))
209                           gKodePlg = Trim(rsHlp("KDPLG"))
210                           gInvCust = Trim(rsHlp("FS_NM_DEALER"))
211                           gInvAlmt = Trim(rsHlp("FS_ALAMAT"))
212                           gInvKota = Trim(rsHlp("FS_KOTA"))
213                           gInvNPWP = Trim(rsHlp("FS_NPWP"))
214                           If Trim(rsHlp("INTERNAL")) = 1 Then
215                               If Trim(gNpwp) = Trim(gInvNPWP) Then
216                                   Text(40) = "AFFNONPPN"
217                               Else
218                                   Text(40) = "AFFPPN"
219                               End If
220                           Else
221                               Text(40) = "EXT"
222                           End If
223                       End If
224                       If rsHlp.State <> adStateClosed Then rsHlp.Close
225                       Call giveDiscount("0", Text(21), False)
                         ' text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
226                   ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
                           
227                       Set rsHlp = objBilling.searchMOUDetail(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS"))
228                       If Not rsHlp.EOF Then
229                           Text(34) = Trim(rsHlp("NO_MOU"))
230                       End If
231                       If rsHlp.State <> adStateClosed Then rsHlp.Close
                         ' Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")), text(21), False)
232                   Else
                          'Call giveDiscount("0", text(21), True)
233                   End If
234               End If
235           End If
              
              ' TAMPIL JOB ORDER
              Dim nilaiRate As Double
236           totRate = 0
             ' text(22) = 0
237           subJob = 0
238           If rsJob.State Then rsJob.Close
239           myChassis = Trim(Text(17))
240           For i = 1 To 7
241               hlp = "REP" & i
242               hlp2 = "RTE" & i
243               hlp3 = "KLPKERJA" & i
244               hlp4 = "ACTRATE" & i
245               hlp5 = "MEKANIK" & i
                  
246               If rsTemp(hlp) <> "" Then
247                   fOth = False
248                   Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "' AND MDL = '" & rsTemp("MDL") & "'")
249                   If rsHlp.EOF Then
250                       fOth = True
251                   End If
252                   If rsHlp.State <> adStateClosed Then rsHlp.Close
                      
253                   If returnJob(LblReturn) Then
254                       If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
255                           If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
256                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
257                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
258                           End If
259                       Else
260                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
261                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
262                       End If
263                   Else
264                       If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
265                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
266                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
267                       ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
268                           If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
269                               If rsTemp("FS_RC") = "R001" Then
270                                   If Trim(rsTemp(hlp3)) = "SBI" Then
                                           
271                                       Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
272                                       If Not rsHlp2.EOF Then
273                                           nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
274                                           If nStatusDealer = 1 Then
275                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
276                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
277                                           Else
278                                               If nStatusDealer = 2 Then
279                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
280                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
281                                               ElseIf nStatusDealer = 3 Then
282                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
283                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
284                                               End If
285                                           End If
286                                       End If
287                                       If rsHlp2.State <> adStateClosed Then rsHlp2.Close
288                                   Else
289                                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
290                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
291                                   End If
292                               Else
                                       
293                                   Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
294                                   If Not rsHlp2.EOF Then
                                          'Dim nStatusDealer As Integer
295                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
296                                       If nStatusDealer = 1 Then
297                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0  '
298                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
299                                       Else
300                                           If nStatusDealer = 2 Then
301                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
302                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
303                                           ElseIf nStatusDealer = 3 Then
304                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
305                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
306                                           End If
307                                       End If
                                          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
308                                   End If
309                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
310                               End If
311                           Else
312                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
313                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
314                           End If
315                       ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
316                           If lblPDS.Visible Then
317                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
318                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
319                           ElseIf rsTemp("FS_RC") = "R001" Then
                                   
320                               Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
321                               If Not rsHlp2.EOF Then
322                                   nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
323                                   If nStatusDealer = 1 Then
324                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
325                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
326                                   Else
327                                       If nStatusDealer = 2 Then
328                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
329                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
330                                       ElseIf nStatusDealer = 3 Then
331                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
332                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
333                                       End If
334                                   End If
                                      'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
335                               End If
336                               If rsHlp2.State <> adStateClosed Then rsHlp2.Close
337                           ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
338                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
339                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
340                           Else
341                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
342                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
343                           End If
344                       End If
345                   End If
346                   nilaiRate = CDbl(rsTemp(hlp2))
                   '   text(22) = text(22) + (CDbl(nHarga) * nilaiRate) * ((100 - text(21)) / 100)
                      
347                   With rsJob
348                   If .State = adStateClosed Then doCreateJob
349                      .AddNew
350                      !job = Trim(rsTemp(hlp))
351                      !Rate = nilaiRate 'CDbl(rsTemp(hlp2))
352                      !ACTRT = CDbl(rsTemp(hlp4))
353                      !grp = Trim(rsTemp(hlp3))
354                      !Mech = Trim(rsTemp(hlp5))
355                      !hrg = CDbl(nHarga) * nilaiRate
356                      !addJob = False
357                      !disc = 0 'CDbl(text(21))
358                      !DISCRP = 0 'CDbl(!hrg) * (CDbl(text(21)) / 100)
359                      !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
360                      !valid = gVerifikator
361                      !FREEC = False
362                      .Update
363                   End With
364                   subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
365                   totRate = totRate + CDbl(nilaiRate)
366               End If
367           Next
              
              ' add JOB
368           Set rsHlp2 = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & noWO & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
369           If Not rsHlp2.EOF Then
370               Do While Not rsHlp2.EOF
      '            For i = 1 To 10
      '                hlp = "REP" & i
      '                hlp2 = "RTE" & i
      '                hlp3 = "KLPKERJA" & i
      '                hlp4 = "ACTRATE" & i
      '                hlp5 = "MEKANIK" & i
                      
371                   If rsHlp2("REP") <> "" Then
372                       If returnJob(LblReturn) Then
373                           If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
374                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
375                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
376                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
377                               End If
378                           Else
379                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
380                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
381                           End If
382                       Else
383                           If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
384                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
385                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                  'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
386                           ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
387                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
388                                   If rsTemp("FS_RC") = "R001" Then
389                                       If rsHlp2("KLPKERJA") = "SBI" Then
                                               
390                                           Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
391                                           If Not rsHlp3.EOF Then
392                                               nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
393                                               If nStatusDealer = 1 Then
394                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
395                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
396                                               Else
397                                                   If nStatusDealer = 2 Then
398                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
399                                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
400                                                   ElseIf nStatusDealer = 3 Then
401                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
402                                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
403                                                   End If
404                                               End If
                                                  'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
405                                           End If
406                                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
407                                       Else
408                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
409                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
410                                       End If
411                                   Else
                                           
412                                       Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
413                                       If Not rsHlp3.EOF Then
414                                           nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
415                                           If nStatusDealer = 1 Then
416                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0 '
417                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
418                                           Else
419                                               If nStatusDealer = 2 Then
420                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
421                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
422                                               ElseIf nStatusDealer = 3 Then
423                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
424                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
425                                               End If
426                                           End If
                                              'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
427                                       End If
428                                       If rsHlp3.State <> adStateClosed Then rsHlp3.Close
429                                   End If
430                               Else
431                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
432                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
433                               End If
434                           ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
435                               If lblPDS.Visible Then
436                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
437                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
438                               ElseIf rsTemp("FS_RC") = "R001" Then
                                       
439                                   Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
440                                   If Not rsHlp3.EOF Then
441                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
442                                       If nStatusDealer = 1 Then
443                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
444                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
445                                       Else
446                                           If nStatusDealer = 2 Then
447                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
448                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
449                                           ElseIf nStatusDealer = 3 Then
450                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
451                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
452                                           End If
453                                       End If
                                          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
454                                   End If
455                                   If rsHlp3.State <> adStateClosed Then rsHlp3.Close
456                               ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
457                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
458                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
459                               Else
460                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
461                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
462                               End If
463                           End If
464                       End If
                          
                          ' ADJUSTMENT RATE
465                       nilaiRate = CDbl(rsHlp2("RTE"))
      '                    If Trim(rsTemp(hlp3)) <> "SBE" And Trim(rsTemp(hlp3)) <> "SBI" Then
      '                        Me.Tag = ""
      '                        gVerifikator = ""
      '                        If MsgBox("Pekerjaan : " & Trim(rsHlp2("REP")) & vbCr & "Adjustment Rate? ", vbYesNo, "PENYESUAIAN RATE") = vbYes Then
      '                            frmValidasi.Tag = "ADJ"
      '                            frmValidasi.Show vbModal
      '                        End If
      '                        If Me.Tag = "valid" Then
      '                            'nilaiRate = 0
      '                            nHarga = 0
      '                        End If
      '                    End If
                          
                        '  text(22) = text(22) + (CDbl(nHarga) * CDbl(nilaiRate)) * ((100 - text(21)) / 100)
                              
466                       With rsJob
467                       If .State = adStateClosed Then doCreateJob
468                           .AddNew
469                           !job = Trim(rsHlp2("REP"))
470                           !Rate = nilaiRate 'CDbl(rsHlp2("RTE"))
471                           !ACTRT = CDbl(rsHlp2("ACTRTE"))
472                           !grp = Trim(rsHlp2("KLPKERJA"))
473                           !Mech = Trim(rsHlp2("MEKANIK"))
474                           !hrg = CDbl(nHarga) * CDbl(nilaiRate)
475                           !addJob = True
476                           !disc = 0 'CDbl(text(21))
477                           !DISCRP = CDbl(!hrg) * 0 '(CDbl(text(21)) / 100)
478                           !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
479                           !valid = gVerifikator
480                           !FREEC = False
481                           .Update
482                           adaAddJob = True
483                       End With
484                       subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
485                       totRate = totRate + CDbl(nilaiRate)
486                   End If
      '            Next
487                   rsHlp2.MoveNext
488               Loop
489           End If
              
      '        subJob = CDbl(Text(22))
490           Set JobGrid.DataSource = rsJob
491           Call doHeadJob
              
              ' TAMPIL PART ORDER
492           noKFaktur = ""
493           gSpart = True
              'text(24) = 0
              'ptext(25) = 0
494           subPartGenuine = 0
495           subPartNGenuine = 0
496           subOliGenuine = 0
497           subOliNGenuine = 0
              Dim HargaDisc As Double
498           If rsPart.State Then rsPart.Close
              
              ' /// AMBIL PART YG BAWA SENDIRI //
              ' AND B.TIPE = 'O'
499           Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, TIPE, GENUINE, BAWA FROM jobpart a WITH (NOLOCK) LEFT JOIN T_oli b WITH (NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli WHERE no_WO = '" & noWO & "' " _
                                      & "AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y' ") 'AND BAWA <> 'Y'
500           If Not rsHlp.EOF Then
501               rsHlp.MoveFirst
502               Do While Not rsHlp.EOF
503                   With rsPart
504                   If .State = adStateClosed Then doCreatePart
505                       gSpart = True
                          
506                       If rsHlp("BAWA") = "Y" Then
507                           .AddNew
508                           !KodeItem = "BAWA SENDIRI"
509                           !NAMAITEM = rsHlp("PARTBAWA")
510                           !qty = rsHlp("QTY")
511                           !QtyOrder = rsHlp("QTY")
512                           !QtyBO = 0
513                           !QtyCancel = 0
514                           !Harga = 0
515                           !oli = IIf(rsHlp("TIPE") <> "P", True, False)
516                           .Update
517                       Else
518                           If Not gFPart Then
                                  '///// TIDAK CONNECT ACCOUNTING SPAREPART ////
519                               .AddNew
520                               !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
521                               !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
522                               !oli = IIf(rsHlp("TIPE") <> "P", True, False)
523                               !qty = rsHlp("QTY")
                                      'Set rsHlp2 = objO.FetchDataACC("SELECT hargajualA FROM inmst WHERE kodeItem = '" & rsHlp("KODE_ITEM") & "'")
524                                   Set rsHlp2 = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & rsHlp("KODE_ITEM") & "'")
525                                   If Not rsHlp2.EOF Then
526                                       HargaDisc = CDbl(rsHlp2("HARGAJUALA"))
527                                       !Harga = CDbl(HargaDisc)
528                                       If rsHlp("OLI") <> "" Then
529                                           If rsHlp("TIPE") = "P" Then
530                                               If rsHlp2("TGenuine") = False Then
                                                      ' PART GENUINE
531                                                   subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
532                                               ElseIf rsHlp2("TGenuine") = True Then
                                                      ' PART NON GENUINE
533                                                   subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
534                                               End If
535                                               Text(24) = Text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
536                                           ElseIf rsHlp("TIPE") = "O" Then
537                                               If rsHlp2("TGenuine") = False Then
                                                      ' OLI GENUINE
538                                                   subOliNGenuine = subOliNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
539                                               ElseIf rsHlp2("TGenuine") = True Then
                                                      ' OLI NON GENUINE
540                                                   subOliGenuine = subOliGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
541                                               End If
542                                               Text(25) = Text(25) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
543                                           End If
544                                       Else
                                              ' PART GENUINE
545                                           subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
546                                           'text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
547                                       End If
548                                   End If
549                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                                  
550                               .Update
551                           End If
552                       End If
553                   End With
554                   rsHlp.MoveNext
555               Loop
556           End If
557           If rsHlp.State <> adStateClosed Then rsHlp.Close
              
558           If gFPart Then
                  '/// KONEKSI KE SPAREPART
559               If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
560                   With rsPart
561                       If noInv Then
                              Dim strSql As String
562                           If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
563                               strSql = "select (case When KodeItem = '' Then namaBYSA Else KodeItem end) as KodeItem, " & _
                                       "Keterangan as Namabarang, SatuanNs as Satuan, Qty as QtyOrder, (isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0)) AS Jml, " & _
                                       "(Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) as QtyBO, HargaJual as HargaJualA, NOSO as NoRack, " & _
                                       "0 as DisItem, 0 as ExtDis, Case BO When 'CANCEL' THEN (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) ELSE 0 END AS QTYCANCEL from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(Text(0)) & "'"
564                           Else
565                               strSql = "Select Z.*, 0 as QtyOrder, 0 as QtyBO FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                      "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                      "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
                                      "INMST.NORACK,0 as DisItem, 0 as ExtDis, 0 as QtyCANCEL FROM INMST " & _
                                      "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                      "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                      "Where Z.Norack = '" & Trim(Text(0)) & "'"
566                           End If
567                           Set rsHlp3 = objO.FetchDataACC(strSql)

568                           If Not rsHlp3.EOF Then
569                               Do While Not rsHlp3.EOF
570                                   If .State = adStateClosed Then doCreatePart
571                                   .AddNew
572                                   !KodeItem = Trim(rsHlp3("KODEITEM"))
573                                   !NAMAITEM = Trim(rsHlp3("NAMABARANG"))
574                                   !QtyOrder = CDbl(rsHlp3("QtyOrder"))
575                                   !qty = CDbl(rsHlp3("JML"))
576                                    If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
577                                       !QtyBO = IIf(UCase(isBO) = "CANCEL", 0, objO.FetchData_DoubleACC("select Qty FROM BOFILE WITH (NOLOCK) Where TutupOrderFlag = 0 and (NoSO = '" & Trim(Text(0)) & "' Or NOSO =(Select Distinct NoSPK From SOORD WITH (NOLOCK) Where NOSO ='" & Trim(Text(0)) & "' and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "')) and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "'"))
578                                    Else
579                                       !QtyBO = 0
580                                    End If
                                      
581                                   !QtyCancel = CDbl(rsHlp3("QtyCancel"))
582                                   !Harga = CDbl(rsHlp3("HARGAJUALA"))
      '                                Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
      '                                If Not rsHlp.EOF Then
      '                                    If rsHlp("TIPE") = "O" Then
      '                                        !oli = True
      '                                        Text(25) = Text(25) + (CDbl(!Harga) * CDbl(!qty))
      '                                    ElseIf rsHlp("TIPE") = "P" Then
      '                                        !oli = False
      '                                        Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                    End If
      '                                Else
      '                                    !oli = False
      '                                    Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                End If
583                                       Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp3("KODEITEM") & "'")
584                                       If Not rsHlp.EOF Then
585                                           If Left(rsHlp("SCC"), 3) = "OLI" Then
586                                               !oli = True
                                                  'text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
587                                           Else 'If Left(rsHlp("SCC"), 3) = "PAR" Then
588                                               !oli = False
                                                  'text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
      '                                            !oli = False
      '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
589                                           End If
590                                       Else
591                                           !oli = False
592                                           Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
593                                       End If
594                                   .Update
595                                   rsHlp3.MoveNext
596                               Loop
597                           Else 'If rsHlp3.EOF And newValid Then
                                  Dim rsInvPart As New ADODB.Recordset
                                  Dim vHitung As Double
598                               Set rsInvPart = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                                   & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(Text(0)) & "' and KodeLgn = '" & Text(2) & "' ")
599                               If Not rsInvPart.EOF Then
600                                   Do While Not rsInvPart.EOF
601                                       If .State = adStateClosed Then doCreatePart
602                                       .AddNew
603                                       !KodeItem = Trim(rsInvPart("KODEITEM"))
604                                       !NAMAITEM = Trim(rsInvPart("KETERANGAN"))
605                                       vHitung = (Koma2Titik(rsInvPart("Qty")) * (Koma2Titik(rsInvPart("HargaJual")) - (Koma2Titik(rsInvPart("ItemDisc")) / Abs(Koma2Titik(rsInvPart("Qty"))))))
606                                       vHitung = vHitung - (Koma2Titik(rsInvPart("ExtDisPersen")) * vHitung)
607                                       If Trim(UCase(rsInvPart("TypePpn"))) = "E" Then
608                                           vHitung = Round(vHitung + vHitung * (profilACC.nilaiPPn / 100), 2)
609                                       End If
610                                       If gJurnalPart Then
611                                           vHitung = cekIncludePPn(vHitung)
612                                       End If
613                                       !qty = Val(rsInvPart("QTY"))
614                                       If !qty = 0 Then
615                                           !Harga = 0
616                                       Else
617                                           !Harga = CDbl(vHitung / !qty)
618                                       End If
619                                       !QtyOrder = Val(rsInvPart("QTY"))
620                                       !QtyBO = 0
621                                       !QtyCancel = 0
                                           
                                          'Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsInvPart("KODEITEM"), 1) & "'")
622                                       Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsInvPart("KODEITEM") & "'")
623                                       If Not rsHlp.EOF Then
624                                           If Left(rsHlp("SCC"), 3) = "OLI" Then
625                                               !oli = True
626                                               'text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
627                                           Else 'If Left(rsHlp("SCC"), 3) = "PAR" Then
628                                               !oli = False
629                                               'text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
      '                                            !oli = False
      '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
630                                           End If
631                                       Else
632                                           !oli = False
633                                           'text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
634                                       End If
      '                                    If Not rsHlp.EOF Then
      '                                        If rsHlp("TIPE") = "O" Then
      '                                            !oli = True
      '                                            Text(25) = Text(25) + (CDbl(!Harga) * CDbl(!qty))
      '                                        ElseIf rsHlp("TIPE") = "P" Then
      '                                            !oli = False
      '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                        End If
      '                                    Else
      '                                        !oli = False
      '                                        Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                    End If
635                                       .Update
636                                       rsInvPart.MoveNext
637                                   Loop
638                               Else
639                                   gSpart = False
      '                                If TblValid.Enabled Then
      '                                    temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
      '                                    If temp = vbYes Then
      '                                        gSpart = False
      '                                    End If
      '                                End If
640                               End If
641                               If rsInvPart.State <> adStateClosed Then rsInvPart.Close
642                           End If
643                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
644                       Else 'Not newValid Then
                               
645                           Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NOINV LIKE '" & Trim(Text(20)) & "%'")
646                           If Not rsHlp2.EOF Then
647                               Do While Not rsHlp2.EOF
648                                   If .State = adStateClosed Then doCreatePart
649                                   'With rsPart
                                          
650                                       .AddNew
651                                       !KodeItem = Trim(rsHlp2("KODEITEM"))
652                                       !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
653                                       !qty = CDbl(rsHlp2("QTY"))
654                                       !Harga = CDbl(rsHlp2("HARGA")) '* ((100 - CDbl(rsHlp2("FN_EXT_DISKON"))) / 100)
655                                       !oli = Trim(rsHlp2("OLI")) 'IIf(Trim(rsHlp2("OLI")) = 1, True, False)
656                                       If !oli Then
657                                          ' text(25) = text(25) + CDbl(!Harga)
658                                       Else
659                                           'text(24) = text(24) + CDbl(!Harga)
660                                       End If
661                                       !EXTDISC = CDbl(rsHlp2("FN_EXT_DISKON"))
662                                       !QtyOrder = CDbl(rsHlp2("QTY"))
663                                       !QtyBO = 0
664                                       !QtyCancel = 0
665                                       .Update
666                                   'End With
667                                   rsHlp2.MoveNext
668                               Loop
669                           Else
670                               gSpart = False
671                           End If
672                           If rsHlp2.State <> adStateClosed Then rsHlp2.Close
673                       End If
674                   End With
675               End If
              
676               If Not gSpart Then
677                       With rsPart
678                           Set rsHlp3 = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                             & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(Text(0)) & "' and KodeLgn = '" & Text(2) & "' ")
                              
679                           If Not rsHlp3.EOF Then
680                               Do While Not rsHlp3.EOF
681                                   If .State = adStateClosed Then doCreatePart
682                                   .AddNew
683                                   !KodeItem = Trim(rsHlp3("KODEITEM"))
684                                   !NAMAITEM = Trim(rsHlp3("KETERANGAN"))
685                                   vHitung = (Koma2Titik(rsHlp3("Qty")) * (Koma2Titik(rsHlp3("HargaJual")) - (Koma2Titik(rsHlp3("ItemDisc")) / Abs(Koma2Titik(rsHlp3("Qty"))))))
686                                   vHitung = vHitung - (Koma2Titik(rsHlp3("ExtDisPersen")) * vHitung)
687                                   If Trim(UCase(rsHlp3("TypePpn"))) = "E" Then
688                                       vHitung = Round(vHitung + vHitung * (profilACC.nilaiPPn / 100), 2)
689                                   End If
690                                   If gJurnalPart Then
691                                       vHitung = cekIncludePPn(vHitung)
692                                   End If
693                                   !qty = CDbl(rsHlp3("Qty"))
694                                   !Harga = CDbl(vHitung / !qty)
695                                   Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
696                                   If Not rsHlp.EOF Then
697                                       If rsHlp("TIPE") = "O" Then
698                                           !oli = True
699                                           'text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
700                                       ElseIf rsHlp("TIPE") = "P" Then
701                                           !oli = False
702                                           'text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
703                                       End If
704                                   End If
705                                   !QtyOrder = CDbl(rsHlp3("Qty"))
706                                   !QtyBO = 0
707                                   !QtyCancel = 0
708                                   .Update
709                                   rsHlp3.MoveNext
710                               Loop
711                               gSpart = True
712                           Else
      '                            If TblValid.Enabled Then
      '                                temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
      '                                If temp = vbYes Then
      '                                    gSpart = False
      '                                Else
      '                                    gSpart = True
      '                                End If
      '                            End If
713                           End If
714                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
715                       End With
716                   Else
717                       gSpart = True
718                   End If
719               End If
720           End If
              
721           subPart = 0 ' CDbl(text(24))
722           subOli = 0 'CDbl(text(25))
723           If rsPart.State <> adStateClosed Then
724               If Not rsPart.BOF Then
725                   rsPart.MoveFirst
726                   rsPart.Sort = "OLI"
727                   Set partGrid.DataSource = rsPart
728                   Call doHeadPart
729               End If
730           End If
              
              '---------
              
              Dim rsWashing As ADODB.Recordset
      '        Set rsWashing = objO.FetchData("select washing from jobdtl where no_wo = '" & noWO & "'")
              
      '        If rsWashing.RecordCount > 0 Then
      '            If rsWashing("washing") = True Then
      '                'pengecekan sudah proses cuci
      '                Set rsWashing = objO.FetchData("SELECT * FROM TEMPPROGRESSCUCI WHERE NOORDER = '" & noWO & "'")
      '                If rsWashing.RecordCount > 0 Then
      '                    btnBatalWashing.Enabled = False
      '                Else
      '                    btnBatalWashing.Enabled = True
      '                End If
      '            Else
      '                btnBatalWashing.Enabled = False
      '            End If
      '        Else
      '            btnBatalWashing.Enabled = False
      '        End If
          
              ' TAMPIL SUBLET ORDER
731           adaSublet = False
              'cmdSublet.Visible = False
             ' text(27) = 0
732           If rsSublet.State Then rsSublet.Close
733           Set rsSublet = objO.FetchData("SELECT NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, valid, HARGA, PRICE FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noWO & "' AND no_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
734           If Not rsSublet.EOF Then
735               Do While Not rsSublet.EOF
736                   If rsSublet("VALID") = 0 Then
737                       adaSublet = True
738                       MsgBox "Sublet belum di validasi"
739                       If rsSublet.State Then rsSublet.Close
740                       TblValid.Enabled = False
741                       Exit Sub
742                   End If
                      'text(27) = text(27) + (IIf(Trim(text(20)) = "", CDbl(rsSublet("HARGA")), CDbl(rsSublet("PRICE"))) * CDbl(rsSublet("QTY")))
743                   rsSublet.MoveNext
744               Loop
745               Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
                       & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "') and no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "' " _
                       & "ORDER BY nonota, kodeitem")
                 ' subSublet = CDbl(text(27))
746               adaSublet = True
                  'cmdSublet.Visible = True
747               Set subletGrid.DataSource = rsSublet
748               Call doHeadSublet
749           Else
750               If newValid Then
      '                If TblValid.Enabled Then
      '                    If MsgBox("Ada Sublet ?", vbYesNo, "Konfirmasi Sublet") = vbNo Then
      '                        adaSublet = True
      '                    End If
      '                End If
751               End If
752           End If
              
              'text(29) = Round(CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)), 0) '+ CDbl(text(28))
              ' TAMPIL TERBILANG
      '''''''''''''''''        txtTotalRp = Angka2Huruf(text(29), 1)
753               If OnTopSV = 0 Then
754                   Set rsHlp = objO.FetchDataACC("select KodeLgn, SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(Text(2)) & "'")
755               Else
                  'UNTUK FEBUARI - TOP PELANGGAN
756                   Set rsHlp = objO.FetchDataACC("select KodeLgn, TOPGRP as SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(Text(2)) & "'")
757               End If
758               If Not rsHlp.BOF Then
759                   Text(41) = rsHlp("SyaratBayar")
760               End If
761               If rsHlp.State <> adStateClosed Then rsHlp.Close
              
762           If Not gSpart Or Not adaSublet Then
                '  TblValid.Enabled = False
763           End If
764           If Log.dmsNonDss = 1 Then
                  'lakukan pengecekan apakah harus menginputkan DP part
                  'total > plafon = dp >= prosentase * total
                  '2014-12-19 RUDI KRISTANTO
      '
      '            If Trim(Text(64)) <> "" Or Trim(Text(78)) <> "" Then
      '                totalNilaiPart = getTotalPart(IIf(Trim(Text(64)) <> "", Trim(Text(64)), Trim(Text(78))), False)
      '            Else
      '                totalNilaiPart = CDbl(Trim(Text(57))) + CDbl(Trim(Text(59)))
      '            End If
765               totalNilaiPart = 0
               '   totalNilaiPart = CDbl(Trim(text(24))) + CDbl(Trim(text(25)))
                              
766               If totalNilaiPart > Log.plafonDP Then
                      'lakukan pengisian DP
767                   MsgBox "Nilai part membutuhkan DP minimal Rp." & Format(((Log.persentaseDP / 100) * totalNilaiPart), "0,#"), vbInformation, "Informasi DP Part"
                      'boleh = False
768               End If
769           End If
              

         
770      Exit Sub

panggilWO_Error:

771       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilWO of Form frmOrderComplain")
              
End Sub

Private Sub hapusForm()
    Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & Text(0) & "'")
    For i = 0 To 44
        Select Case i
            Case 43
                Text(i).Mask = "" '"99-99-9999"
            Case 10 ', 21, 22, 23, 24, 25, 26, 27, 29, 39
                Text(i) = ""
            Case 42, 30, 33, 21, 22, 23, 24, 25, 26, 27, 29, 39

            Case Else
                Text(i) = ""
        End Select
    Next
'    For i = 0 To 44
'        Select Case i
'            Case 1, 11, 43
'                text(i).Mask = "" '"99-99-9999"
'            Case 21, 22, 23, 24, 25, 26, 27, 29, 39
'
'            Case 10
'                text(i) = 0
'            Case 42, 30, 33
'            Case Else
'                text(i) = ""
'        End Select
'    Next
    
    VScroll1.value = 0
    opWO(0).value = True
    
    If rsJob.State Then rsJob.Close
    Set rsJob = Nothing
    Set JobGrid.DataSource = rsJob
    'JobGrid.ReBind
    
    If rsPart.State Then rsPart.Close
    Set rsPart = Nothing
    Set partGrid.DataSource = rsPart
    'partGrid.ReBind
    
    Set subletGrid.DataSource = Nothing
    adaAddJob = False
    
    ' ISI KARAKTER PELANGGAN
    Set rsTemp = objO.FetchData("SELECT * FROM T_Karakter WITH (NOLOCK) ORDER BY KODE")
    If Not rsTemp.EOF Then
        cboPelanggan.Clear
        Do While Not rsTemp.EOF
            cboPelanggan.AddItem (rsTemp("KODE"))
            rsTemp.MoveNext
        Loop
    End If
    cboPelanggan.ListIndex = 0
    
    vTipeCustomer = ""
    woStat = ""
    nDuoInv = False
    'chPS.value = Unchecked
   ' chNea.value = Unchecked
   isBaru = True
   txtComplain.Text = ""
   lblStatus.Caption = ""
   JualNRM.Caption = ""

End Sub

Private Sub btnBatalWashing_Click()
    If Trim(Text(0)) <> "" Then
        Set rsTmp = objO.FetchData("SELECT FLAG,JAMMULAI,NOSTALL,JAMSELESAI FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & Trim(Text(0)) & "|%'")
        If Not rsTmp.EOF Then
            ' TUTUP TAMBAH PEKERJAAN
            If UCase(rsTmp("FLAG")) = "G" Then
                framePembatalanWashing.Visible = True
            End If
        End If
    End If
End Sub

Private Sub btnSimpan_Click()
    Call objO.UpdateData("Update Jobdta Set Catat = '" & Trim(txt.Text) & "' Where NO_WO = '" & Trim(Text(0).Text) & "'")
    MsgBox "Catatan telah disimpan !", vbInformation + vbOKOnly
End Sub

Private Sub chNea_Click()
    If chNea.value = 1 Then
        txtInfoNea.Visible = True
        txtInfoNea.Text = ""
    Else
        txtInfoNea.Visible = False
    End If
End Sub

Private Sub cmdKeluh_Click()
      Dim rsSearch As New ADODB.Recordset
      Dim rscr7 As ADODB.Recordset
1         If Log.dmsNonDss = 1 Then
2             txtPermintaan = ""
3             txtKeluhan = ""
4             txtFI = ""
5             txtNextJob = ""
              
6             If FrameKeluhDmsNonDSS.Visible Then
7                 FrameKeluhDmsNonDSS.Visible = False
8             Else
9                 Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(Text(0)) & "'")
10                If Not rsSearch.EOF Then
11                    Text3.Text = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
12                    Text2.Text = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
13                End If
                  ''20180515REV01'INSPECTION
      '            Set rsSearch = objO.FetchData("SELECT * FROM INSPECTION A WITH (NOLOCK) WHERE NOWO = '" & Trim(text(0)) & "'")
14                Set rsSearch = objO.FetchData("SELECT top 1 * FROM INSPECTION A WITH (NOLOCK) WHERE NOWO = '" & Trim(Text(0)) & "' order by tglentry desc")
15                If Not rsSearch.EOF Then
16                    txtFI = Trim(rsSearch("HASILFI")) & vbCrLf & Trim(rsSearch("HasilTestDrive"))
17                    txtNextJob = Trim(rsSearch("SNJ1")) & vbCrLf & Trim(rsSearch("SNJ2")) & vbCrLf & Trim(rsSearch("SNJ3")) & vbCrLf & Trim(rsSearch("SNJ4")) & vbCrLf & Trim(rsSearch("SNJ5")) & vbCrLf & Trim(rsSearch("SNJ6")) & vbCrLf & Trim(rsSearch("SNJ7"))
18                End If
19                FrameKeluhDmsNonDSS.Visible = True
20            End If
              
              
21        Else
          
22            txtPermintaan = ""
23            txtKeluhan = ""
24            If framKeluh.Visible Then
25                framKeluh.Visible = False
26            Else
27                Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(Text(0)) & "'")
28                If Not rsSearch.EOF Then
29                    txtPermintaan = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
30                    txtKeluhan = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
31                End If
32                framKeluh.Visible = True
33            End If
              
           
34            Set rscr7 = objO.FetchData("Select Catat FROM JOBDTA WITH (NOLOCK) Where NO_WO = '" & Trim(Text(0)) & "'")
35            If rscr7.RecordCount > 0 Then
36                txt.Text = rscr7.Fields(0).value
37            Else
38                txt.Text = ""
39            End If
              
40        End If
End Sub

Private Sub Command1_Click()
Call cmdKeluh_Click
End Sub

Private Sub Command2_Click()
    If Trim(txtAlasanBtlWashing.Text) <> "" Then
        If Len(txtAlasanBtlWashing.Text) < 5 Then
            MsgBox "Alasan Pembatalan Washing minimal 5 karakter!", vbCritical, "Informasi Pembatalan Washing"
        Else
            On Error GoTo vbError
            erTranAll = False
            objHT.OpenTransactionBengkel
            
            Call objHT.InsertData_Bengkel("INSERT INTO WASHING VALUES ('" & Trim(Text(0)) & "',GETDATE(),'" & gSA & "','" & txtAlasanBtlWashing.Text & "')")
            Call objHT.InsertData_Bengkel("UPDATE ANTRICUCI SET FLAG = 'CANCEL' WHERE NOORDER = '" & Trim(Text(0)) & "'")
            Call objHT.InsertData_Bengkel("UPDATE JOBDTL SET WASHING = 0 WHERE NO_wo = '" & Trim(Text(0)) & "'")
            Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%" & Trim(Text(0)) & "%'")
            
            If erTranAll = False Then
            objHT.CloseTransactionBengkel
            End If
            MsgBox "Proses Pembatalan Washing berhasil dilakukan", vbInformation, "Informasi Pembatalan Washing"
            txtAlasanBtlWashing.Text = ""
            framePembatalanWashing.Visible = False
            btnBatalWashing.Enabled = False
            Exit Sub
vbError:
            objHT.CloseTransactionBengkel_err
            MsgBox "Proses Pembatalan Washing gagal dilakukan", vbCritical
            Exit Sub
        End If
    Else
        MsgBox "Alasan Pembatalan Washing harus diisi!", vbCritical, "Informasi Pembatalan Washing"
    End If
End Sub

Private Sub Form_Activate()
    If orderAktif Or subletAktif Or subRetAktif Then
        MsgBox "Tutup Menu sebelumnya terlebih Dahulu !!", vbCritical
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    validAktif = True
    Text(1) = TglDDMMYYYY(isnow)
    simpan = True
    Frame6.Top = 0
   isBaru = True
End Sub

Private Sub Form_Resize()
    'asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    validAktif = False
End Sub

Private Sub tblBatal_Click()
    'If gSpart And adaSublet Then
    Set rsTemp = objO.FetchData("SELECT DISTINCT NO_WO FROM VALIDBILLING WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND DELLOG = ''")
    If rsTemp.EOF Then
        Call objO.UpdateData("UPDATE VALIDORDER SET USERDEL = '" & Trim(gSA) & "', USERDEL_TGL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(17)) & "' AND USERDEL = ''")
        Call objO.DeleteData("DELETE FROM NEA_PS Where noWO = '" & Trim(Text(0)) & "'")
        chNea.value = 0
        chPS.value = 0
        txtInfoNea.Text = ""
        lblStatus.Caption = "BATAL VALIDASI"
        TblBatal.Enabled = False
        TblValid.Enabled = True
    Else
        MsgBox "WO ini sudah diberi Diskon Khusus oleh HO, sehingga tidak bisa dibatal Validasi" & vbCr & "Hubungi HO untuk membatalkan diskon khusus. ", vbCritical + vbOKOnly, "Batal Diskon Khusus"
    End If
    'End If
End Sub

Private Sub Label41_Click()

End Sub



Private Sub cetakForm()
          Dim rpt As Object
          Dim NamaSAOrder As String
          Dim rsRepo As New ADODB.Recordset
          Dim rsHlp As New ADODB.Recordset
          Dim bodyClear As Boolean
          
1        On Error GoTo cetakForm_Error

2         Select Case gPrintCetak
              Case "2"
3                 Set rpt = New RptOrder4            ' KGW + PEMUDA ' 2190
4             Case "1"
5                 Set rpt = New RptOrder5            ' MLATI + ALL  ' 2180
6             Case "3"
7                 Set rpt = New RptOrder6            ' SLAMET  ' LX 300+II
8             Case "4"
9                 Set rpt = New RptOrder7            ' BLANKO UTK KERTAS KOSONG
10                rpt.txtDealer.SetText support.kopOrder(0)
11        End Select
              
12        If MsgBox("Mencetak Order ?", vbYesNo, "Konfirmasi") = vbYes Then
13            gNoOrder = Text(0) & "/" & Trim(Text(41)) 'CboStatus.text
14            gKetSA = Text(19)
15            NamaSAOrder = objO.FetchData_String("Select FS_NM_PETUGAS from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS ='" & Trim(gKetSA) & "'")
16            gCust = Text(4)
17            gAlmt = Text(5)
18            gKota = Text(6)
19            gTelpon = Text(7)
20            gTahun = Text(10)
21            gNoPolisi = Text(16)
22            gModel = Text(14)
23            gWarna = Text(15)
24            gTglMasuk = Text(43)
25            gMDLOrder = Text(12)
26            gFRM = Text(17)
27            gEG = Text(18)
              
               
28            Set rsHlp = objO.FetchData("SELECT CATAT, DATE, TIME, B.FS_NM_DEALER, B.DDATE, REMARK, REQ1, REQ2, REQ3, REQ4, REQ5, REP1, REP2, REP3, REP4, REP5, REP6, REP7 FROM JOBDTA A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS WHERE NO_WO = '" & Text(0) & "' AND A.NO_CHASSIS = '" & Text(17) & "'")
29            If Not rsHlp.EOF Then
30                gCatatan = Trim(rsHlp("CATAT"))
31                gJamMasuk = Trim(rsHlp("TIME"))
32                gDealer = Trim(rsHlp("FS_NM_DEALER"))
33                gDD = Trim(rsHlp("DDATE"))
34                gPKhusus = Trim(rsHlp("REMARK"))
                  
35                nWKT = hitungWaktu(CDbl(totRate), rsHlp("TIME"), rsHlp("DATE"))
36                gPSelesai = Left(nWKT, InStr(nWKT, "|") - 1)
                  
37                gKeluhan(1) = Trim(rsHlp("REQ1"))
38                gKeluhan(2) = Trim(rsHlp("REQ2"))
39                gKeluhan(3) = Trim(rsHlp("REQ3"))
40                gKeluhan(4) = Trim(rsHlp("REQ4"))
41                gKeluhan(5) = Trim(rsHlp("REQ5"))
42                gOrder(1) = Trim(rsHlp("REP1"))
43                gOrder(2) = Trim(rsHlp("REP2"))
44                gOrder(3) = Trim(rsHlp("REP3"))
45                gOrder(4) = Trim(rsHlp("REP4"))
46                gOrder(5) = Trim(rsHlp("REP5"))
47                gOrder(6) = Trim(rsHlp("REP6"))
48                gOrder(7) = Trim(rsHlp("REP7"))
49            End If
               
50            Set rsHlp = objO.FetchData("SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Text(0) & "' AND NO_CHASSIS = '" & Text(17) & "'")
51            If Not rsHlp.EOF Then
52                i = 1
53                Do While Not rsHlp.EOF
54                    If i < 10 Then
55                        gSublet(i) = rsHlp("NAMABARANG")
56                        i = i + 1
57                    End If
58                    rsHlp.MoveNext
59                Loop
60            End If
61            If rsHlp.State <> adStateClosed Then rsHlp.Close
              
62            gKM = Format(Text(37), "#,#0")
63            gSAdvisor = NamaSAOrder
64            gKodePlg = Text(2)
              
              Dim sHlp As Object  'New ADODB.Recordset
              Dim sHlp2 As Object
              Dim sHlp3 As Object
              Dim sHlp4 As Object
              Dim rRs As ADODB.Recordset
              Dim crxSubReportHistory As CRAXDRT.Report
              Dim crxSubReportBooking As CRAXDRT.Report
              Dim crxSubReportBooking2 As CRAXDRT.Report
              Dim crxSubReportBookingFU As CRAXDRT.Report
65            If gUkuranOrder = False Then
                  
66                Set rRs = objO.FetchData("SELECT A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.NO_POLISI, A.NO_ENGINE, A.NO_CHASSIS, A.MDL, A.MODEL, A.FMDL, A.COLOR, " _
                                        & "D.FS_KD_DEALER, D.FS_NM_DEALER, D.DDATE, A.DOTAM, A.CODE, A.NPWP, A.PPN, A.DISC, A.DISCT, A.CRDTCD, A.FRONT, A.[GROUP], A.MECH, " _
                                        & "A.KM, A.DATE, A.TIME, A.TIME_M, A.TIME_S, A.TIME_A, A.TIME_AT, A.TIME_SA, A.PART, A.MTRL, A.OIL, A.MET, A.SBLT, A.LABOR, A.BODY, A.TOTAL, " _
                                        & "A.JOB, A.REQ1, A.REQ2, A.REQ3, A.REQ4, A.REQ5, A.JOBCD, A.CR, A.REP1, A.REP2, A.REP3, A.REP4, A.REP5, A.REP6, A.REP7, A.REP8, A.SLIP, " _
                                        & "A.RTE1, A.RTE2, A.RTE3, A.RTE4, A.RTE5, A.RTE6, A.RTE7, A.RTE8, A.INVDATE, A.NO_WO, A.TPSVC, A.NOINV, A.STATUS, A.PAID, A.NOFAK, " _
                                        & "A.CATAT, A.DELIV, A.SPK, A.RPR_COND, A.KD_K, A.MEKANIK1, A.MEKANIK2, A.MEKANIK3, A.MEKANIK4, A.MEKANIK5, " _
                                        & "A.MEKANIK6, A.MEKANIK7, A.MEKANIK8, A.GRP1, A.GRP2, A.GRP3, A.GRP4, A.GRP5, A.GRP6, A.GRP7, A.GRP8, " _
                                        & "A.OFF1, A.OFF2, A.OFF3, A.OFF4, A.OFF5, A.OFF6, A.OFF7, A.OFF8, " _
                                        & "A.KLPKERJA1, A.KLPKERJA2, A.KLPKERJA3, A.KLPKERJA4, A.KLPKERJA5, A.KLPKERJA6, A.KLPKERJA7, A.KLPKERJA8, A.ALASAN, A.FS_STATUS_INV, " _
                                        & "A.FS_TH_MOBIL, A.FN_RETURN, A.FS_LOG_ADD, A.FS_LOG_DEL, A.FD_TGL_DEL, A.FS_RC, A.FS_NM_HOBI, A.FS_KD_PLG, A.FS_KD_WIL, " _
                                        & "A.KODETIPE, A.PDS, A.THS, A.SBE, B.COND1, B.COND2, B.COND3, B.COND4, B.COND5, B.COND6, B.WAK1, B.WAK2, B.WAK3, B.WAK4, B.WAK5, " _
                                        & "B.WAE1, B.WAE2, B.WAE3, B.WAE4, B.WAE5, B.BBM, B.HEAVYUSE, B.GANTIPART, B.TUNGGUSVC, B.PEMBAYARAN, B.TGLFU, B.JAMFU, B.TELPFU, " _
                                        & "B.KETFU, B.VERIFIKATOR, B.DATESERAH, B.TIMESERAH, B.CMPL1, B.CMPL2, B.CMPL3, B.CMPL4, B.CMPL5, B.DATANG, B.DTLDATANG, " _
                                        & "B.KARAKTER, B.DTLKARAKTER, D.KODEPOS, D.REMARK, D.KEYNO, D.[USER], D.TGL_REP, D.FS_KM_AKHIR, D.USEJOBDSC, D.RCDATE, D.CDK, " _
                                        & "D.TGL_D , D.CDD, D.USEFS_NM_HOBI, D.NOTE, E.FS_KD_PETUGAS, E.FS_NM_PETUGAS, E.FS_KD_PASS, E.FS_KD_GROUP, E.tipe, E.DIVISI, A.ANCUST, D.GSM, D.CDMA " _
                                        & "FROM JOBDTA AS A WITH (NOLOCK) INNER JOIN JOBDTL AS B WITH (NOLOCK) ON A.NO_WO = B.NO_WO LEFT OUTER JOIN CUSTDAT AS D WITH (NOLOCK) ON A.NO_CHASSIS = D.NO_CHASSIS INNER JOIN " _
                                        & "T_PETUGAS AS E WITH (NOLOCK) ON A.FRONT = E.FS_KD_PETUGAS WHERE A.NO_WO = '" & Trim(Text(0)) & "'")
67                rpt.database.SetDataSource rRs
                             
68                gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
                  
69                sSQL = "SELECT * FROM historyUnitFunc('" & Trim(gFRM) & "')"
70                Set sHlp = objO.FetchData(sSQL)
71                Set crxSubReportHistory = rpt.Subreport2.OpenSubreport
72                crxSubReportHistory.database.SetDataSource sHlp
                  
73                sSQL = "select DATE_BOOK, TIME_D, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(Text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
74                Set sHlp2 = objO.FetchData(sSQL)
75                Set crxSubReportBooking = rpt.Subreport1.OpenSubreport
76                crxSubReportBooking.database.SetDataSource sHlp2
                  
77                sSQL = "select DATE, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(Text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
78                Set sHlp4 = objO.FetchData(sSQL)
79                Set crxSubReportBooking2 = rpt.Subreport4.OpenSubreport
80                crxSubReportBooking2.database.SetDataSource sHlp4
                  
81                sSQL = "select WAKTU FROM M_FUBOOK WITH (NOLOCK) where NOFAK = '" & Trim(Text(31)) & "' "
82                Set sHlp3 = objO.FetchData(sSQL)
83                Set crxSubReportBookingFU = rpt.Subreport3.OpenSubreport
84                crxSubReportBookingFU.database.SetDataSource sHlp3
                  
85                Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(Text(0)) & "|%'")
86                If Not rsHlp.EOF Then
87                    rpt.txtStall.SetText Trim(rsHlp("NOSTALL"))
88                End If
                  
89                rpt.mechPlan.SetText ""
                  
                  Dim n As Integer
90                bodyClear = True
91                Set rsHlp = objO.FetchData("SELECT * FROM JOBKOORD WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "'")
92                If Not rsHlp.EOF Then
93                    n = 1
94                    Do While Not rsHlp.EOF
95                        Call pasangTanda(rpt, n, CDbl(rsHlp("KOORDY")), CDbl(rsHlp("KOORDX")))
96                        n = n + 1
97                        rsHlp.MoveNext
98                    Loop
99                    bodyClear = False
100               End If
101               rsHlp.Close
                  
102               If bodyClear Then
103                   rpt.Text93.TextColor = &HFF&
104                   rpt.Text94.TextColor = &HFF&
105                   rpt.Text95.TextColor = &HFF&
106                   rpt.Text96.TextColor = &HFF&
107                   rpt.Text97.TextColor = &HFF&
108               End If
                  
109               rpt.Text2.TextColor = &HFF&
110               rpt.Text3.TextColor = &HFF&
111               rpt.Text4.TextColor = &HFF&
112               rpt.Text5.TextColor = &HFF&
                  
                   
113               Set rsHlp = objO.FetchData("SELECT BBM FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "'")
114               If Not rsHlp.EOF Then
115                   Select Case rsHlp("BBM")
                          Case 0
                              ' FULL
116                           rpt.Text90.TextColor = &H0&
117                       Case 1
                              ' 3/4
118                           rpt.Text62.TextColor = &H0&
119                       Case 2
                              ' 1/2
120                           rpt.Text64.TextColor = &H0&
121                       Case 3
                              ' 1/4
122                           rpt.Text88.TextColor = &H0&
123                       Case 4
                              ' EMPTY
124                           rpt.Text89.TextColor = &H0&
125                   End Select
126               End If
                  
127               With FrmctkOrder.CRViewer1
128                   .ReportSource = rpt
129                   .ViewReport
130                   .Zoom 100
131                   .Visible = True
132               End With
133               FrmctkOrder.cmdTutup.Visible = True
134               Screen.MousePointer = vbDefault
135               FrmctkOrder.WindowState = vbMaximized
136               FrmctkOrder.show vbModal
137           End If
138       End If

139      On Error GoTo 0
140      Exit Sub

cetakForm_Error:

141       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakForm of Form frmOrderComplain")
End Sub

Private Sub tblCetakWO_Click()

End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblValid_Click()
      Dim rsSearch As New ADODB.Recordset

1         On Error GoTo vbError
2         Set rsSearch = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND USERDEL = ''")
3         If Not rsSearch.EOF Then
4             MsgBox "Sudah di validasi sebelumnya !!", vbInformation, "Validasi Order"
5             lblStatus.Caption = "DATA SUDAH DIVALIDASI"
6             TblValid.Enabled = False
7             TblBatal.Enabled = True
              
8         Else
9             If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(43)), "DD/MM/YYYY") Then
                  Dim rsSO As ADODB.Recordset
10                Set rsSO = objO.FetchDataACC("select KodeItem  from SOORD WITH (NOLOCK) " & _
                           "Where (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0)+ IsNull(QtyKeluar,0))) <> 0 and NOSO = '" & Trim(Text(0)) & "' and isnull(bo,'') <> 'CANCEL'")
11                If rsSO.RecordCount > 0 Then
12                    MsgBox "Terdapat " & rsSO.RecordCount & " Line Item Part yang belum tersuplay sesuai order !" & Chr(13) & "Validasi tidak dapat diproses !", vbInformation, "Validasi Order"
13                    Exit Sub
14                End If
15            Else
                  
16            End If
17            Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderComplain", "tblValid", 0)
18                erTranAll = False
19                onTR = True
20                Call objHT.OpenTransactionBengkel
                  
21            If gSpart And adaSublet And Trim(Text(17)) <> "" Then
22                Call objHT.InsertData_BengkelOnly("INSERT INTO VALIDORDER VALUES ('" & Trim(gSA) & "',getdate(),'" & IIf(rsJob.State = adStateClosed, "TANPA PEKERJAAN| ", "") + IIf(rsPart.State = adStateClosed, "TANPA PART| ", "") + IIf(rsSublet.State = adStateClosed, "TANPA SUBLET| ", "") & "','" & Trim(Text(0)) & "','" & Trim(Text(17)) & "','','')")
23                Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTL SET KARAKTER = '" & Trim(cboPelanggan) & "', DTLKARAKTER = '" & Trim(Text(44)) & "' WHERE NO_WO = '" & Trim(Text(0)) & "'")
24                If chNea.value = 1 Or chPS.value = 1 Then Call objHT.InsertData_BengkelOnly("Insert Into NEA_PS values ('" & Trim(Text(0)) & "'," & chNea.value & "," & chPS.value & ",'" & Trim(txtInfoNea.Text) & "','" & Trim(gSA) & "',getdate(),'','01/01/1900','')")
25                lblStatus.Caption = "SUDAH VALIDASI"
26                TblValid.Enabled = False
27                TblBatal.Enabled = True
28            Else
29                lblStatus.Caption = "DATA BELUM LENGKAP"
30            End If
31                If erTranAll = False Then
32                    Call objHT.CloseTransactionBengkel
33                    onTR = False
34                    Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderComplain", "tblValid", 1)
35                    Exit Sub
36                Else
37                    Call objHT.CloseTransactionBengkel_err
38                    onTR = False
39                    Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderComplain", "tblValid", 2)
40                    MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
41                    Exit Sub
42                End If
43        End If
44        Exit Sub

vbError:
45        Call objHT.CloseTransactionBengkel_err
46        onTR = False
47        Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderComplain", "tblValid", 2)
48        MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmBooking"), vbInformation, "Gagal Simpan"
49        Exit Sub
 End Sub

Private Sub tblSimpanComplain_Click()
1        On Error GoTo tblSimpanComplain_Click_Error

2         If txtComplain.Text <> "" Then
3             If Len(Trim(txtComplain.Text)) > 5 Then
4                 If isBaru Then
5                     Call objO.InsertData("INSERT INTO COMPLAIN VALUES ('" & Trim(Text(0)) & "','" & Trim(txtComplain.Text) & "',getdate(),'" & gSA & "')")
6                 Else
7                     Call objO.InsertData("UPDATE COMPLAIN SET KOMPLAIN = '" & Trim(txtComplain.Text) & "' WHERE NO_WO = '" & Trim(Text(0)) & "'")
8                 End If
9                 MsgBox "Data berhasil disimpan", vbInformation
10                Call hapusForm
11                Exit Sub
12            Else
13                MsgBox "Complain user minimal 5 karakter", vbInformation
14            End If
15        Else
16            MsgBox "Complain user harus diisi!", vbCritical
17        End If

18       On Error GoTo 0
19       Exit Sub

tblSimpanComplain_Click_Error:

20        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpanComplain_Click of Form frmOrderComplain")
End Sub

Private Sub Text_Change(Index As Integer)
1        On Error GoTo Text_Change_Error

2         If Index = 17 Then
3             If Trim(Text(17)) <> "" Then
4                 myChassis = Trim(Text(17))
5                 Call CexChassis_Discon(Trim(Text(13)))
6                 If isJualNRM = True Then
7                     Me.JualNRM.Caption = nJualNRM
8                 Else
9                     Me.JualNRM.Caption = ""
10                End If
11            End If
12        End If

13       On Error GoTo 0
14       Exit Sub

Text_Change_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Text_Change of Form frmOrderComplain")
End Sub

Private Sub Text_GotFocus(Index As Integer)
          Dim DiTab As Boolean
1        On Error GoTo text_GotFocus_Error

2         DiTab = False
3         If Text(Index).BackColor = &HC0E0FF Then
4             DiTab = True
5         Else
6             If gMainKeyCari <> "" Then
7                 Text(Index).Mask = ""
8                 Text(Index) = gMainKeyCari
9                 gMainKeyCari = ""
10                DiTab = True
11            End If
12        End If

13        Select Case Index
              Case 31
14                VScroll1.value = 2500
15            Case 32
16                VScroll1.value = 2700
17            Case 28
18                VScroll1.value = 2000
19        End Select
              
20        If DiTab Then
21            Sendkeys ("{TAB}")
22        Else
23            Sendkeys ("{HOME}+{END}")
24        End If

25       On Error GoTo 0
26       Exit Sub

text_GotFocus_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_GotFocus of Form frmOrderComplain")
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
1        On Error GoTo Text_KeyDown_Error

2         Call CekTombol(KeyCode)
3         Select Case KeyCode
              Case vbKeyF12
4                 Select Case Index
                      Case 0
5                         gMainKeyCari = Text(Index)
6                         gSearchOption = "INV2"
7                         gBantuan = "AND DELIV <> 'BTL'"
8                         frmCari.show (vbModal)
9                     Case 42
10                        gMainKeyCari = Text(Index)
11                        gSearchOption = "METERAI"
12                        frmCari.show (vbModal)
13                End Select
14            Case vbKeyEscape
15                If simpan = False Then
16                    If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
17                        Call hapusForm
18                        Text(0).SelLength = Len(Text(0))
19                        Text(0).SetFocus
20                    End If
21                End If
22        End Select

23       On Error GoTo 0
24       Exit Sub

Text_KeyDown_Error:

25        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Text_KeyDown of Form frmOrderComplain")
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
1        On Error GoTo text_KeyPress_Error

2         KeyAscii = UpCase(KeyAscii)
3         If KeyAscii <> 13 Then
4             Select Case Index
                  Case 28, 42
5                     KeyAscii = CekNumerik(KeyAscii)
6                     simpan = False
7             End Select
8         End If

9        On Error GoTo 0
10       Exit Sub

text_KeyPress_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_KeyPress of Form frmOrderComplain")
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsHlp As New ADODB.Recordset

1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 0
3                 If Text(Index) <> "" Then
4                     boleh = True
5                     noInv = True
6                     If simpan = False Then
7                         If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
8                             boleh = False
9                         End If
10                    End If
11                    If boleh Then
12                        n = Trim(Text(0))
13                        Call hapusForm
14                        Text(0) = n
15                        If rsTemp.State Then rsTemp.Close
16                        Set rsTemp = objO.FetchData("SELECT * FROM jobDta A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.no_WO = '" & Text(Index) & "' ")
17                        If Not rsTemp.EOF Then
                              'TblValid.Enabled = False
18                            newValid = False
19                            lblStatus.Visible = True
20                            lblStatus.Caption = ""
21                            Set rsHlp = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND USERDEL = ''")
22                            If Not rsHlp.EOF Then
23                                lblStatus.Caption = "SUDAH DIVALIDASI"
                                  'TblValid.Enabled = False
                                  'TblBatal.Enabled = True
24                                adaSublet = True
                                  
                                  Dim rsNPS As ADODB.Recordset
25                                Set rsNPS = objO.FetchData("Select * from NEA_PS WITH (NOLOCK) Where NOWO = '" & Text(0).Text & "'")
26                                If rsNPS.RecordCount > 0 Then
                                      'chNea.value = rsNPS.Fields(1).value
                                      'chPS.value = rsNPS.Fields(2).value
27                                    'txtInfoNea.Text = Trim(rsNPS.Fields(3).value)
28                                End If
                                  
29                                If rsTemp("NOINV") <> "" And rsTemp("RPR_COND") <> "K" Then
30                                    MsgBox ("Telah dibuat invoice, Tanggal " & TglDDMMMMYYYY(Trim(rsTemp("INVDATE"))) & " No. Invoice = " & Trim(rsTemp("NOINV")))
31                                    lblStatus.Caption = "Sudah Invoice"
                                      'TblBatal.Enabled = False
32                                    noInv = False
                                      'Exit Sub
33                                ElseIf rsTemp("NOINV") <> "" And rsTemp("RPR_COND") = "K" Then
34                                    MsgBox ("Telah diclosing !")
35                                    lblStatus.Caption = "Closing"
                                      'TblBatal.Enabled = False
36                                    noInv = False
                                      'Exit Sub
37                                End If
38                            Else
                                  'TblBatal.Enabled = False
39                                If (rsTemp("DELIV") = "PRC" Or rsTemp("DELIV") = "OKE") And rsTemp("RPR_COND") = "S" And rsTemp("NOINV") = "" Then
40                                    newValid = True
                                     ' text(21).SelLength = Len(text(21))
                                      'TblValid.Enabled = True
                                      'TblValid.SetFocus
41                                    MsgBox ("Unit telah selesai dikerjakan !")
42                                    lblStatus.Caption = "Selesai dikerjakan"
43                                    GoTo prosesStall
44                                ElseIf rsTemp("DELIV") = "BTL" Then
45                                    lblStatus.Caption = "WO dibatalkan"
46                                ElseIf rsTemp("RPR_COND") = "K" Then
47                                    lblStatus.Caption = "WO Sudah Closing"
48                                    noInv = False
49                                Else
prosesStall:
50                                    Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(rsTemp("NO_POLISI")) & " |" & Trim(rsTemp("NO_WO")) & "%'")
51                                    If Not rsHlp.EOF Then
52                                        Select Case Trim(rsHlp("FLAG"))
                                              Case ""
53                                                MsgBox "Unit belum dibuat WO masih booking waiting !", vbInformation, "Status Order"
54                                                lblStatus.Caption = "Booking"
55                                            Case "N"
56                                                MsgBox "Unit belum dikerjakan !", vbInformation, "Status Order"
57                                                lblStatus.Caption = "Menunggu Foreman di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
58                                            Case "S"
59                                                MsgBox "Unit sedang dikerjakan !", vbInformation, "Status Order"
60                                                lblStatus.Caption = "Proses di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
61                                            Case "P"
62                                                MsgBox "Unit sedang dihentikan sementara !", vbInformation, "Status Order"
63                                                lblStatus.Caption = "Pause di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
64                                            Case "H"
65                                                MsgBox "Unit menunggu proses Final Inspection", vbInformation, "Status Order"
66                                                lblStatus.Caption = "Tunggu Inspection di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
67                                            Case "Q"
68                                                MsgBox "Unit dalam proses Final Inspection", vbInformation, "Status Order"
69                                                lblStatus.Caption = "Proses Inspection di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
70                                            Case "G"
71                                                MsgBox "Unit menunggu proses Washing", vbInformation, "Status Order"
72                                                lblStatus.Caption = "Tunggu Washing di " & getStallWashing(Trim(Text(0)), True, True) '& ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
73                                            Case "D"
74                                                MsgBox "Unit menunggu proses Drying", vbInformation, "Status Order"
75                                                lblStatus.Caption = "Tunggu Drying di " & getStallWashing(Trim(Text(0)), True, False) ' & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
76                                            Case "A"
77                                                MsgBox "Unit sedang proses Washing", vbInformation, "Status Order"
78                                                lblStatus.Caption = "Washing"
79                                            Case "R"
80                                                MsgBox "Unit sedang proses Drying", vbInformation, "Status Order"
81                                                lblStatus.Caption = "Drying"
82                                            Case "F"
83                                                MsgBox ("Unit telah selesai dikerjakan !")
84                                                lblStatus.Caption = "Selesai dikerjakan"
85                                        End Select
86                                    End If
87                                End If
88                            End If
89                            Call panggilWO(Text(Index))
90                            Set rsTemp = objO.FetchData("SELECT * FROM COMPLAIN WHERE NO_WO = '" & Trim(Text(Index)) & "'")
91                            If rsTemp.RecordCount > 0 Then
92                                isBaru = False
93                                rsTemp.MoveFirst
94                                txtComplain.Text = rsTemp("komplain")
95                            End If
96                        End If
97                        If rsTemp.State <> adStateClosed Then rsTemp.Close
98                        simpan = False
99                    End If
100               End If
101       End Select
102       txtComplain.SetFocus

103      On Error GoTo 0
104      Exit Sub

text_Validate_Error:

105       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmOrderComplain")
End Sub

Public Function getStallWashing(noOrder As String, isTunggu As Boolean, isCuci As Boolean) As String
          Dim rsTemp As ADODB.Recordset
1        On Error GoTo getStallWashing_Error

2         If isTunggu = True Then
3             Set rsTemp = objO.FetchData("SELECT * FROM ANTRICUCI WHERE NOORDER = '" & noOrder & "' AND TIPESTALL = '" & IIf(isCuci, "Cuci", "Pengering") & "'")
4         Else
5             Set rsTemp = objO.FetchData("SELECT * FROM TEMPPROGRESSCUCI WHERE NOORDER = '" & noOrder & "' AND FINISH = ''")
6         End If
7         If rsTemp.RecordCount > 0 Then
8             getStallWashing = rsTemp("NoStall")
              
9         Else
10            getStallWashing = ""
11        End If

12       On Error GoTo 0
13       Exit Function

getStallWashing_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getStallWashing of Form frmOrderComplain")
End Function

Private Sub txtComplain_KeyPress(KeyAscii As Integer)
KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub txtKeluhan_GotFocus()
    Dim DiTab As Boolean
    DiTab = False
    If txtKeluhan.BackColor = &HC0E0FF Then
        DiTab = True
    End If

    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub txtPermintaan_GotFocus()
    Dim DiTab As Boolean
    DiTab = False
    If txtPermintaan.BackColor = &HC0E0FF Then
        DiTab = True
    End If

    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub VScroll1_Change()
    If VScroll1.value > 0 Then
        Frame6.Top = -VScroll1.value + 20
    Else
        Frame6.Top = 0
    End If
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.value > 0 Then
        Frame6.Top = -VScroll1.value + 20
    Else
        Frame6.Top = 0
    End If
End Sub

Private Sub pasangTanda(rpt As Report, Urut As Integer, nilX As Integer, nilY As Integer)
      Dim nTop As Integer
      Dim nLeft As Integer
          
1        On Error GoTo pasangTanda_Error

2         nTop = 6750  '6750 'repo3kosong    '7420 'repo2
3         nLeft = 4300 '4300 'repo3kosong    '4440 'repo2
          
4         Select Case Urut
              Case 1
5                 rpt.imgTnd1.TextColor = &H0&
6                 rpt.imgTnd1.Top = nTop + nilX - (195 / 2)
7                 rpt.imgTnd1.Left = nLeft + nilY - (210 / 2)
8             Case 2
9                 rpt.imgTnd2.TextColor = &H0&
10                rpt.imgTnd2.Top = nTop + nilX - (195 / 2)
11                rpt.imgTnd2.Left = nLeft + nilY - (210 / 2)
12            Case 3
13                rpt.imgTnd3.TextColor = &H0&
14                rpt.imgTnd3.Top = nTop + nilX - (195 / 2)
15                rpt.imgTnd3.Left = nLeft + nilY - (210 / 2)
16            Case 4
17                rpt.imgTnd4.TextColor = &H0&
18                rpt.imgTnd4.Top = nTop + nilX - (195 / 2)
19                rpt.imgTnd4.Left = nLeft + nilY - (210 / 2)
20            Case 5
21                rpt.imgTnd5.TextColor = &H0&
22                rpt.imgTnd5.Top = nTop + nilX - (195 / 2)
23                rpt.imgTnd5.Left = nLeft + nilY - (210 / 2)
24            Case 6
25                rpt.imgTnd6.TextColor = &H0&
26                rpt.imgTnd6.Top = nTop + nilX - (195 / 2)
27                rpt.imgTnd6.Left = nLeft + nilY - (210 / 2)
28            Case 7
29                rpt.imgTnd7.TextColor = &H0&
30                rpt.imgTnd7.Top = nTop + nilX - (195 / 2)
31                rpt.imgTnd7.Left = nLeft + nilY - (210 / 2)
32            Case 8
33                rpt.imgTnd8.TextColor = &H0&
34                rpt.imgTnd8.Top = nTop + nilX - (195 / 2)
35                rpt.imgTnd8.Left = nLeft + nilY - (210 / 2)
36            Case 9
37                rpt.imgTnd9.TextColor = &H0&
38                rpt.imgTnd9.Top = nTop + nilX - (195 / 2)
39                rpt.imgTnd9.Left = nLeft + nilY - (210 / 2)
40            Case 10
41                rpt.imgTnd10.TextColor = &H0&
42                rpt.imgTnd10.Top = nTop + nilX - (195 / 2)
43                rpt.imgTnd10.Left = nLeft + nilY - (210 / 2)
44            Case 11
45                rpt.imgTnd11.TextColor = &H0&
46                rpt.imgTnd11.Top = nTop + nilX - (195 / 2)
47                rpt.imgTnd11.Left = nLeft + nilY - (210 / 2)
48            Case 12
49                rpt.imgTnd12.TextColor = &H0&
50                rpt.imgTnd12.Top = nTop + nilX - (195 / 2)
51                rpt.imgTnd12.Left = nLeft + nilY - (210 / 2)
52            Case 13
53                rpt.imgTnd13.TextColor = &H0&
54                rpt.imgTnd13.Top = nTop + nilX - (195 / 2)
55                rpt.imgTnd13.Left = nLeft + nilY - (210 / 2)
56            Case 14
57                rpt.imgTnd14.TextColor = &H0&
58                rpt.imgTnd14.Top = nTop + nilX - (195 / 2)
59                rpt.imgTnd14.Left = nLeft + nilY - (210 / 2)
60            Case 15
61                rpt.imgTnd15.TextColor = &H0&
62                rpt.imgTnd15.Top = nTop + nilX - (195 / 2)
63                rpt.imgTnd15.Left = nLeft + nilY - (210 / 2)
64            Case 16
65                rpt.imgTnd16.TextColor = &H0&
66                rpt.imgTnd16.Top = nTop + nilX - (195 / 2)
67                rpt.imgTnd16.Left = nLeft + nilY - (210 / 2)
68            Case 17
69                rpt.imgTnd17.TextColor = &H0&
70                rpt.imgTnd17.Top = nTop + nilX - (195 / 2)
71                rpt.imgTnd17.Left = nLeft + nilY - (210 / 2)
72            Case 18
73                rpt.imgTnd18.TextColor = &H0&
74                rpt.imgTnd18.Top = nTop + nilX - (195 / 2)
75                rpt.imgTnd18.Left = nLeft + nilY - (210 / 2)
76            Case 19
77                rpt.imgTnd19.TextColor = &H0&
78                rpt.imgTnd19.Top = nTop + nilX - (195 / 2)
79                rpt.imgTnd19.Left = nLeft + nilY - (210 / 2)
80            Case 20
81                rpt.imgTnd20.TextColor = &H0&
82                rpt.imgTnd20.Top = nTop + nilX - (195 / 2)
83                rpt.imgTnd20.Left = nLeft + nilY - (210 / 2)
84            Case 21
85                rpt.imgTnd21.TextColor = &H0&
86                rpt.imgTnd21.Top = nTop + nilX - (195 / 2)
87                rpt.imgTnd21.Left = nLeft + nilY - (210 / 2)
88            Case 22
89                rpt.imgTnd22.TextColor = &H0&
90                rpt.imgTnd22.Top = nTop + nilX - (195 / 2)
91                rpt.imgTnd22.Left = nLeft + nilY - (210 / 2)
92            Case 23
93                rpt.imgTnd23.TextColor = &H0&
94                rpt.imgTnd23.Top = nTop + nilX - (195 / 2)
95                rpt.imgTnd23.Left = nLeft + nilY - (210 / 2)
96            Case 24
97                rpt.imgTnd24.TextColor = &H0&
98                rpt.imgTnd24.Top = nTop + nilX - (195 / 2)
99                rpt.imgTnd24.Left = nLeft + nilY - (210 / 2)
100           Case 25
101               rpt.imgTnd25.TextColor = &H0&
102               rpt.imgTnd25.Top = nTop + nilX - (195 / 2)
103               rpt.imgTnd25.Left = nLeft + nilY - (210 / 2)
104           Case 26
105               rpt.imgTnd26.TextColor = &H0&
106               rpt.imgTnd26.Top = nTop + nilX - (195 / 2)
107               rpt.imgTnd26.Left = nLeft + nilY - (210 / 2)
108           Case 27
109               rpt.imgTnd27.TextColor = &H0&
110               rpt.imgTnd27.Top = nTop + nilX - (195 / 2)
111               rpt.imgTnd27.Left = nLeft + nilY - (210 / 2)
112           Case 28
113               rpt.imgTnd28.TextColor = &H0&
114               rpt.imgTnd28.Top = nTop + nilX - (195 / 2)
115               rpt.imgTnd28.Left = nLeft + nilY - (210 / 2)
116           Case 29
117               rpt.imgTnd29.TextColor = &H0&
118               rpt.imgTnd29.Top = nTop + nilX - (195 / 2)
119               rpt.imgTnd29.Left = nLeft + nilY - (210 / 2)
120           Case 30
121               rpt.imgTnd30.TextColor = &H0&
122               rpt.imgTnd30.Top = nTop + nilX - (195 / 2)
123               rpt.imgTnd30.Left = nLeft + nilY - (210 / 2)
124           Case 31
125               rpt.imgTnd31.TextColor = &H0&
126               rpt.imgTnd31.Top = nTop + nilX - (195 / 2)
127               rpt.imgTnd31.Left = nLeft + nilY - (210 / 2)
128           Case 32
129               rpt.imgTnd32.TextColor = &H0&
130               rpt.imgTnd32.Top = nTop + nilX - (195 / 2)
131               rpt.imgTnd32.Left = nLeft + nilY - (210 / 2)
132           Case 33
133               rpt.imgTnd33.TextColor = &H0&
134               rpt.imgTnd33.Top = nTop + nilX - (195 / 2)
135               rpt.imgTnd33.Left = nLeft + nilY - (210 / 2)
136           Case 34
137               rpt.imgTnd34.TextColor = &H0&
138               rpt.imgTnd34.Top = nTop + nilX - (195 / 2)
139               rpt.imgTnd34.Left = nLeft + nilY - (210 / 2)
140           Case 35
141               rpt.imgTnd35.TextColor = &H0&
142               rpt.imgTnd35.Top = nTop + nilX - (195 / 2)
143               rpt.imgTnd35.Left = nLeft + nilY - (210 / 2)
144           Case 36
145               rpt.imgTnd36.TextColor = &H0&
146               rpt.imgTnd36.Top = nTop + nilX - (195 / 2)
147               rpt.imgTnd36.Left = nLeft + nilY - (210 / 2)
148           Case 37
149               rpt.imgTnd37.TextColor = &H0&
150               rpt.imgTnd37.Top = nTop + nilX - (195 / 2)
151               rpt.imgTnd37.Left = nLeft + nilY - (210 / 2)
152           Case 38
153               rpt.imgTnd38.TextColor = &H0&
154               rpt.imgTnd38.Top = nTop + nilX - (195 / 2)
155               rpt.imgTnd38.Left = nLeft + nilY - (210 / 2)
156           Case 39
157               rpt.imgTnd39.TextColor = &H0&
158               rpt.imgTnd39.Top = nTop + nilX - (195 / 2)
159               rpt.imgTnd39.Left = nLeft + nilY - (210 / 2)
160           Case 40
161               rpt.imgTnd40.TextColor = &H0&
162               rpt.imgTnd40.Top = nTop + nilX - (195 / 2)
163               rpt.imgTnd40.Left = nLeft + nilY - (210 / 2)
164           Case 41
165               rpt.imgTnd41.TextColor = &H0&
166               rpt.imgTnd41.Top = nTop + nilX - (195 / 2)
167               rpt.imgTnd41.Left = nLeft + nilY - (210 / 2)
168           Case 42
169               rpt.imgTnd42.TextColor = &H0&
170               rpt.imgTnd42.Top = nTop + nilX - (195 / 2)
171               rpt.imgTnd42.Left = nLeft + nilY - (210 / 2)
172           Case 43
173               rpt.imgTnd43.TextColor = &H0&
174               rpt.imgTnd43.Top = nTop + nilX - (195 / 2)
175               rpt.imgTnd43.Left = nLeft + nilY - (210 / 2)
176           Case 44
177               rpt.imgTnd44.TextColor = &H0&
178               rpt.imgTnd44.Top = nTop + nilX - (195 / 2)
179               rpt.imgTnd44.Left = nLeft + nilY - (210 / 2)
180           Case 45
181               rpt.imgTnd45.TextColor = &H0&
182               rpt.imgTnd45.Top = nTop + nilX - (195 / 2)
183               rpt.imgTnd45.Left = nLeft + nilY - (210 / 2)
184           Case 46
185               rpt.imgTnd46.TextColor = &H0&
186               rpt.imgTnd46.Top = nTop + nilX - (195 / 2)
187               rpt.imgTnd46.Left = nLeft + nilY - (210 / 2)
188           Case 47
189               rpt.imgTnd47.TextColor = &H0&
190               rpt.imgTnd47.Top = nTop + nilX - (195 / 2)
191               rpt.imgTnd47.Left = nLeft + nilY - (210 / 2)
192           Case 48
193               rpt.imgTnd48.TextColor = &H0&
194               rpt.imgTnd48.Top = nTop + nilX - (195 / 2)
195               rpt.imgTnd48.Left = nLeft + nilY - (210 / 2)
196           Case 49
197               rpt.imgTnd49.TextColor = &H0&
198               rpt.imgTnd49.Top = nTop + nilX - (195 / 2)
199               rpt.imgTnd49.Left = nLeft + nilY - (210 / 2)
200           Case 50
201               rpt.imgTnd50.TextColor = &H0&
202               rpt.imgTnd50.Top = nTop + nilX - (195 / 2)
203               rpt.imgTnd50.Left = nLeft + nilY - (210 / 2)
204           Case 51
205               rpt.imgTnd51.TextColor = &H0&
206               rpt.imgTnd51.Top = nTop + nilX - (195 / 2)
207               rpt.imgTnd51.Left = nLeft + nilY - (210 / 2)
208           Case 52
209               rpt.imgTnd52.TextColor = &H0&
210               rpt.imgTnd52.Top = nTop + nilX - (195 / 2)
211               rpt.imgTnd52.Left = nLeft + nilY - (210 / 2)
212           Case 53
213               rpt.imgTnd53.TextColor = &H0&
214               rpt.imgTnd53.Top = nTop + nilX - (195 / 2)
215               rpt.imgTnd53.Left = nLeft + nilY - (210 / 2)
216       End Select

217      On Error GoTo 0
218      Exit Sub

pasangTanda_Error:

219       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure pasangTanda of Form frmOrderComplain")
End Sub



