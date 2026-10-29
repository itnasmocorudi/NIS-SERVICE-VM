VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOrderValidasi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Validasi Order"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   13485
   Begin VB.CommandButton cmdBatalRequest 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Batal Re&q Disc"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   9255
      Width           =   1890
   End
   Begin VB.Frame framePembatalanWashing 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pembatalan Washing"
      Height          =   2295
      Left            =   7200
      TabIndex        =   137
      Top             =   2160
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Simpan"
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   1800
         Width           =   1725
      End
      Begin VB.TextBox txtAlasanBtlWashing 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   138
         Top             =   240
         Width           =   3375
      End
   End
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
      Height          =   5895
      Left            =   1440
      TabIndex        =   123
      Top             =   2400
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
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
         Left            =   10200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOrderValidasi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   5400
         Width           =   1245
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   128
         Top             =   680
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   127
         Top             =   3230
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   249
         TabIndex        =   126
         Top             =   5520
         Width           =   9210
      End
      Begin VB.TextBox txtNextJob 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   125
         Top             =   3230
         Width           =   5655
      End
      Begin VB.TextBox txtFI 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   124
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
         TabIndex        =   135
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
         TabIndex        =   134
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
      Begin VB.Label Label42 
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
         Left            =   120
         TabIndex        =   133
         Top             =   5520
         Width           =   585
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
         Top             =   375
         Width           =   5655
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CR 7 : "
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
         TabIndex        =   129
         Top             =   2900
         Width           =   3615
      End
   End
   Begin VB.CommandButton tblCetakWO 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak Ulang WO"
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
      TabIndex        =   117
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
      TabIndex        =   65
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
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13800
      TabIndex        =   7
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
      Height          =   3495
      Left            =   1680
      TabIndex        =   105
      Top             =   2400
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton btnSimpan 
         Appearance      =   0  'Flat
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
         Left            =   10200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOrderValidasi.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   2790
         Width           =   1245
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   249
         TabIndex        =   118
         Top             =   2840
         Width           =   9210
      End
      Begin VB.TextBox txtKeluhan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   107
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtPermintaan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   106
         Top             =   480
         Width           =   5655
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   11640
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label31 
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
         Left            =   120
         TabIndex        =   119
         Top             =   2880
         Width           =   585
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
         TabIndex        =   113
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
      TabIndex        =   8
      Top             =   480
      Width           =   12885
      Begin VB.CommandButton btnBatalWashing 
         BackColor       =   &H000000FF&
         Caption         =   "Batal Washing"
         Enabled         =   0   'False
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   1680
         Width           =   1890
      End
      Begin VB.CommandButton cmdKeluh 
         BackColor       =   &H0000FFFF&
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
         TabIndex        =   104
         Top             =   1680
         Width           =   2130
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   9
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   10
         Top             =   750
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   11
         Top             =   1035
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   12
         Top             =   1350
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   435
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   15
         Top             =   750
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   16
         Top             =   1350
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   17
         Top             =   1050
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   18
         Top             =   750
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   19
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
         TabIndex        =   20
         Top             =   450
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   21
         Top             =   735
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   22
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
         Top             =   1035
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   26
         Top             =   480
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   27
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   28
         Top             =   1350
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   112
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
      Height          =   5655
      LargeChange     =   800
      Left            =   12960
      Max             =   2600
      SmallChange     =   600
      TabIndex        =   96
      Top             =   3600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   480
      ScaleHeight     =   5595
      ScaleWidth      =   12795
      TabIndex        =   66
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
         Left            =   120
         TabIndex        =   67
         Top             =   -600
         Width           =   12375
         Begin VB.Frame Frame3 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   735
            Left            =   1440
            TabIndex        =   121
            Top             =   7350
            Width           =   6855
            Begin VB.TextBox txtInfoNea 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2520
               MaxLength       =   49
               TabIndex        =   3
               ToolTipText     =   ".:: Keterangan NEA ::."
               Top             =   360
               Visible         =   0   'False
               Width           =   4290
            End
            Begin VB.CheckBox chPS 
               BackColor       =   &H0080FF80&
               Caption         =   "Profit Sharing"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3600
               TabIndex        =   2
               Top             =   60
               Width           =   1935
            End
            Begin VB.CheckBox chNea 
               BackColor       =   &H0080FF80&
               Caption         =   "N.E.A."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2520
               TabIndex        =   1
               Top             =   60
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Additional Order Tipe : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   122
               Top             =   60
               Width           =   2265
            End
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
            Height          =   390
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   7560
            Width           =   1890
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Left            =   10320
            TabIndex        =   85
            Top             =   4800
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   26
               Left            =   915
               TabIndex        =   86
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               TabIndex        =   87
               Top             =   480
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
               TabIndex        =   89
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
               TabIndex        =   88
               Top             =   120
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10320
            TabIndex        =   78
            Top             =   2640
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   23
               Left            =   915
               TabIndex        =   79
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               Left            =   480
               TabIndex        =   80
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
               TabIndex        =   81
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
               TabIndex        =   84
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
               TabIndex        =   83
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
               TabIndex        =   82
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.CommandButton TblValid 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Validasi"
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
            Left            =   10320
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   7560
            Width           =   1890
         End
         Begin VB.CommandButton cmdCetak 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cetak"
            Height          =   390
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   8520
            Width           =   1890
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
            Left            =   1440
            TabIndex        =   76
            Top             =   6960
            Width           =   10815
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10320
            TabIndex        =   68
            Top             =   240
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   21
               Left            =   915
               TabIndex        =   69
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               TabIndex        =   70
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
               TabIndex        =   71
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
               TabIndex        =   74
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
               TabIndex        =   73
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
               TabIndex        =   72
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin TrueOleDBGrid70.TDBGrid JobGrid 
            Height          =   2295
            Left            =   0
            TabIndex        =   75
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
            TabIndex        =   90
            Top             =   2640
            Width           =   10335
            _ExtentX        =   18230
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
            TabIndex        =   91
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
         Begin MSMask.MaskEdBox text 
            Height          =   405
            Index           =   29
            Left            =   10365
            TabIndex        =   92
            Top             =   6510
            Width           =   1815
            _ExtentX        =   3201
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
            Index           =   30
            Left            =   1440
            TabIndex        =   101
            Top             =   6680
            Width           =   7515
            _ExtentX        =   13256
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   102
            Top             =   6700
            Width           =   555
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
            TabIndex        =   95
            Top             =   6240
            Width           =   2040
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            Height          =   195
            Left            =   9360
            TabIndex        =   94
            Top             =   6600
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Terbilang (Rupiah)"
            Height          =   195
            Left            =   60
            TabIndex        =   93
            Top             =   7080
            Width           =   1305
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   480
      TabIndex        =   44
      Top             =   2400
      Width           =   12885
      Begin VB.ComboBox cboPelanggan 
         Height          =   315
         ItemData        =   "frmOrderValidasi.frx":0A64
         Left            =   8880
         List            =   "frmOrderValidasi.frx":0A66
         Style           =   2  'Dropdown List
         TabIndex        =   114
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   20
         Left            =   4920
         TabIndex        =   48
         Top             =   800
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   49
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   50
         Top             =   525
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   51
         Top             =   810
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   52
         Top             =   170
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   53
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   54
         Top             =   795
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   55
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   108
         Top             =   165
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   109
         Top             =   480
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         TabIndex        =   115
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
      Begin VB.Label JualNRM 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   12360
         TabIndex        =   141
         Top             =   600
         Width           =   60
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
         TabIndex        =   116
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
         Top             =   525
         Width           =   375
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
         Height          =   285
         Left            =   12120
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   510
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
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   525
         Width           =   570
      End
      Begin VB.Label LblReturn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTJ"
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
         Height          =   285
         Left            =   12120
         TabIndex        =   56
         Top             =   360
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2600
      TabIndex        =   97
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
      TabIndex        =   98
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
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmOrderValidasi.frx":0A68
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
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
      TabIndex        =   103
      Top             =   82
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
      TabIndex        =   100
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
      TabIndex        =   99
      Top             =   135
      Width           =   570
   End
End
Attribute VB_Name = "frmOrderValidasi"
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
Dim noinv As Boolean
Dim totRate As Double
Dim nWKT As String
Dim support As New clsFormDealer
Dim totalNilaiPart As Double
''2017-08-01 || Rev20170801A - penambahan request disc menampilkan data wo walaupun belum validasi tapi button disable
Dim isBolehBatalReqDisc As Boolean

Private Sub doCreatePart()
1        On Error GoTo doCreatePart_Error
2         If rsPart.State Then rsPart.Close
3         'Set rsPart = Nothing
4         With rsPart
             ' If .State Then .Close
5             .Fields.Append "KODEITEM", adVarChar, 50
6             .Fields.Append "NAMAITEM", adVarChar, 200
7             .Fields.Append "QTYOrder", adInteger
8             .Fields.Append "QTY", adInteger
9             .Fields.Append "QTYBO", adInteger
10            .Fields.Append "QTYCANCEL", adInteger
11            .Fields.Append "HARGA", adDouble
12            .Fields.Append "OLI", adBoolean
13            .Fields.Append "EXTDISC", adDouble
14            .Open
15        End With

16       On Error GoTo 0
17       Exit Sub

doCreatePart_Error:

18        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreatePart of Form frmOrderValidasi")
End Sub

Private Sub doHeadPart()
1        On Error GoTo doHeadPart_Error

2         With partGrid
3             .Columns(0).Width = 2000
4             .Columns(0).Caption = "KODE PART"
5             .Columns(0).DataField = "KODEITEM"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 2500
8             .Columns(1).Caption = "NAMA PART"
9             .Columns(1).DataField = "NAMAITEM"
10            .Splits(0).Columns(1).Locked = True
              
11            .Columns(2).Width = 1000
12            .Columns(2).Caption = "QTY Order"
13            .Columns(2).DataField = "QTYOrder"
14            .Splits(0).Columns(2).Locked = True
              
15            .Columns(3).Width = 1000
16            .Columns(3).Caption = "QTY Supply"
17            .Columns(3).DataField = "QTY"
18            .Splits(0).Columns(3).Locked = True
              
19            .Columns(4).Width = 1000
20            .Columns(4).Caption = "QTY BO"
21            .Columns(4).DataField = "QTYBO"
22            .Splits(0).Columns(3).Locked = True
23            On Error GoTo Column_Ini_Untuk_Grid_FillRate
              
              
24            .Columns(5).Visible = True
25            .Columns(5).Width = 1000
26            .Columns(5).Caption = "HARGA"
27            .Columns(5).DataField = "Harga"
28            .Splits(0).Columns(5).Locked = True
              
29            .Columns(6).Visible = False
30            .Columns(6).Caption = "OLI"
31            .Columns(6).DataField = "OLI"
32            .Splits(0).Columns(6).Locked = True
              
33            .Columns(7).Width = 1000
34            .Columns(7).Caption = "QTY CANCEL"
35            .Columns(7).DataField = "QTYCANCEL"
36            .Splits(0).Columns(7).Locked = True
Column_Ini_Untuk_Grid_FillRate:
37        End With

38       On Error GoTo 0
39       Exit Sub

doHeadPart_Error:

40        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadPart of Form frmOrderValidasi")
End Sub

Private Sub doCreateJob()
1        On Error GoTo doCreateJob_Error
2         If rsJob.State Then rsJob.Close
3        ' Set rsJob = Nothing
4         With rsJob
              'If .State Then .Close
5             .Fields.Append "JOB", adVarChar, 50
6             .Fields.Append "RATE", adDouble
7             .Fields.Append "ACTRT", adDouble
8             .Fields.Append "GRP", adVarChar, 6
9             .Fields.Append "MECH", adVarChar, 10
10            .Fields.Append "ADDJOB", adBoolean
11            .Fields.Append "HRG", adDouble
12            .Fields.Append "DISC", adDouble
13            .Fields.Append "DISCRP", adDouble
14            .Fields.Append "STLDISC", adDouble
15            .Fields.Append "VALID", adVarChar, 10
16            .Fields.Append "FREEC", adBoolean           ' TANDA JIKA DIBEBANKAN KE CUSTOMER
17            .Open
18        End With

19       On Error GoTo 0
20       Exit Sub

doCreateJob_Error:

21        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreateJob of Form frmOrderValidasi")
End Sub

Private Sub doHeadJob()
1        On Error GoTo doHeadJob_Error

2         With JobGrid
3             .Columns(0).Width = 4500
4             .Columns(0).Caption = "JOB ORDER"
5             .Columns(0).DataField = "JOB"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 800
8             .Columns(1).Caption = "RATE"
9             .Columns(1).DataField = "RATE"
10            .Splits(0).Columns(1).Locked = True
              
11            .Columns(2).Width = 0
12            .Columns(2).Caption = "ACTRATE"
13            .Columns(2).DataField = "ACTRT"
14            .Columns(2).Visible = False
15            .Splits(0).Columns(2).Locked = True
              
16            .Columns(3).Width = 0
17            .Columns(3).Caption = "GROUP"
18            .Columns(3).DataField = "GRP"
19            .Columns(3).Visible = False
20            .Splits(0).Columns(3).Locked = True
              
21            .Columns(4).Width = 0
22            .Columns(4).Caption = "MEKANIK"
23            .Columns(4).DataField = "MECH"
24            .Columns(4).Visible = False
25            .Splits(0).Columns(4).Locked = True
              
26            .Columns(5).Width = 0
27            .Columns(5).Caption = "Additional Job"
28            .Columns(5).DataField = "ADDJOB"
29            .Columns(5).Visible = False
30            .Splits(0).Columns(5).Locked = True
              
31            .Columns(6).Width = 1200
32            .Columns(6).Caption = "HARGA"
33            .Columns(6).DataField = "HRG"
34            .Splits(0).Columns(6).Locked = True
                     
35            .Columns(7).Width = 600
36            .Columns(7).Caption = "DISC"
37            .Columns(7).DataField = "DISC"
              
38            .Columns(8).Width = 1200
39            .Columns(8).Caption = "DISC RUPIAH"
40            .Columns(8).DataField = "DISCRP"
              
41            .Columns(9).Width = 1200
42            .Columns(9).Caption = "SUB TOTAL"
43            .Columns(9).DataField = "STLDISC"
44            .Splits(0).Columns(9).Locked = True
              
45            .Columns(10).Width = 800
46            .Columns(10).Caption = "Customer"
47            .Columns(10).DataField = "FREEC"
48            .Columns(10).Alignment = dbgCenter
49            .Splits(0).Columns(10).Locked = False
50            .Columns(10).Visible = False
51            If nDuoInv Then
52                .Columns(10).Visible = True
53            End If
54        End With

55       On Error GoTo 0
56       Exit Sub

doHeadJob_Error:

57        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadJob of Form frmOrderValidasi")
End Sub

Private Sub doHeadSublet()
          
1        On Error GoTo doHeadSublet_Error

2         With subletGrid
3             .Columns(0).Width = 1500
4             .Columns(0).Caption = "NO NOTA"
5             .Columns(0).DataField = "NONOTA"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 1800
8             .Columns(1).Caption = "SUPPLIER"
9             .Columns(1).DataField = "SUPPLIER"
10            .Columns(1).Visible = True
11            .Splits(0).Columns(1).Locked = True
              
12            .Columns(2).Width = 3000
13            .Columns(2).Caption = "NAMA BARANG"
14            .Columns(2).DataField = "NAMABARANG"
15            .Splits(0).Columns(2).Locked = True
              
16            .Columns(3).Width = 1500
17            .Columns(3).Caption = "KODE SATUAN"
18            .Columns(3).DataField = "KODESATUAN"
19            .Splits(0).Columns(3).Locked = True
              
20            .Columns(4).Width = 800
21            .Columns(4).Caption = "JUMLAH"
22            .Columns(4).DataField = "QTY"
23            .Splits(0).Columns(4).Locked = True
              
24            .Columns(5).Width = 800
25            .Columns(5).Caption = "VALID"
26            .Columns(5).DataField = "VALID"
27            .Columns(5).Visible = False
28            .Splits(0).Columns(5).Locked = True
              
29            .Columns(6).Width = 1500
30            .Columns(6).Caption = "BELI"
31            .Columns(6).DataField = "HARGA"
32            .Columns(6).NumberFormat = "#,#0"
33            .Splits(0).Columns(6).Locked = True
              
34            .Columns(7).Width = 1500
35            .Columns(7).Caption = "JUAL"
36            .Columns(7).DataField = "PRICE"
37            .Columns(7).NumberFormat = "#,#0"
38            .Splits(0).Columns(7).Locked = True
39        End With

40       On Error GoTo 0
41       Exit Sub

doHeadSublet_Error:

42        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadSublet of Form frmOrderValidasi")
End Sub

Private Function returnJob(labelReturn As Object) As Boolean
1        On Error GoTo returnJob_Error

2         returnJob = False
3         If labelReturn.Visible Then
4             returnJob = True
5         End If

6        On Error GoTo 0
7        Exit Function

returnJob_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure returnJob of Form frmOrderValidasi")
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

9        On Error GoTo 0
10       Exit Function

typeJob_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure typeJob of Form frmOrderValidasi")
End Function

Private Function giveDiscount(discountValue As String, discountText As Object, statusText As Boolean)
1        On Error GoTo giveDiscount_Error

2         discountText = discountValue
3         discountText.Enabled = statusText
4         JobGrid.Splits(0).Columns(7).Locked = Not statusText
5         JobGrid.Splits(0).Columns(8).Locked = Not statusText

6        On Error GoTo 0
7        Exit Function

giveDiscount_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure giveDiscount of Form frmOrderValidasi")
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
              
11            Text1.text = Trim(rsTemp!catat)
              ' TIPE PEMBAYARAN
12            Set rsHlp = objBilling.searchPaymentType(noWO)
13            If Not rsHlp.EOF Then
14                tipeByr = Trim(rsHlp("PEMBAYARAN"))
15                gVerifikator = Trim(rsHlp("VERIFIKATOR"))
16                cboPelanggan = Trim(rsHlp("KARAKTER"))
17                text(44) = Trim(rsHlp("DTLKARAKTER"))
18            End If
19            If rsHlp.State <> adStateClosed Then rsHlp.Close
              
20            Set rsHlp = objBilling.searchDtlOrder(rsTemp("NO_WO"), rsTemp("NO_CHASSIS"))
21            If Not rsHlp.EOF Then
22                nHeavy = IIf(rsHlp("HEAVYUSE") <> 0, True, False)
23            End If
24            If rsHlp.State <> adStateClosed Then rsHlp.Close
              
              ' TAMPIL DETAIL CUSTOMER
25            text(2) = rsTemp("FS_KD_PLG")
26            text(3) = rsTemp("FS_KD_WIL")
27            text(4) = rsTemp("CUSTOMER")
28            text(5) = Left(rsTemp("CUST_ADRES"), 60)
29            text(6) = rsTemp("KOTA")
30            text(7) = rsTemp("TELPON")
31            text(8) = rsTemp("GSM")
32            text(9) = rsTemp("CDMA")
33            text(10) = rsTemp("FS_TH_MOBIL")
34            text(11).Mask = ""
35            text(11) = Nul(IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsTemp("DOTAM"), "/", "-"))))
36            text(12) = rsTemp("MDL")
37            text(13) = rsTemp("MODEL")
38            text(14) = rsTemp("FMDL")
39            text(15) = rsTemp("COLOR")
40            text(16) = rsTemp("NO_POLISI")
41            text(17) = rsTemp("NO_CHASSIS")
42            text(18) = rsTemp("NO_ENGINE")
43            text(19) = rsTemp("FS_LOG_ADD")
44            If rsTemp("NOINV") <> "" Then
45                text(20) = rsTemp("NOINV")
46            End If
47            text(31) = rsTemp("NOFAK")
              'text(33) = rsTemp("FS_RC")
48            text(37) = rsTemp("KM")
49            text(38) = rsTemp("MECH")
50            text(40) = rsTemp("STATUS")
51            text(43) = TglDDMMYYYY(rsTemp("DATE"))
52            nTHS = IIf(rsTemp("THS") = 1, True, False)
53            LblReturn.Visible = IIf(rsTemp("FN_RETURN") = "1", True, False)
54            lblPDS.Visible = IIf(rsTemp("PDS") = "1", True, False)
55            text(28) = rsTemp("ANCUST")
56            text(32) = rsTemp("TELPANCUST")
              
              ' CHECK DEFAULT DISCOUNT
57            text(30) = ""
58            text(35).Visible = False
59            text(36).Visible = False
60            vTipeCustomer = ""
61            myChassis = Trim(text(17))
62            If returnJob(LblReturn) Then
63                If typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
64                    opWO(1).value = True
                       
65                    Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
66                    If Not rsHlp.EOF Then
67                        text(35) = rsHlp("ID_PROGRAM")
68                        text(36) = rsHlp("PROGRAM")
69                        text(35).Visible = True
70                        text(36).Visible = True
71                    End If
72                    If rsHlp.State <> adStateClosed Then rsHlp.Close
73                    If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
74                        nDuoInv = True
                          'RTJ WO BERKEMBANG
                           
75                        Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
76                        If Not rsHlp.EOF Then
77                            gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
78                            gKodePlg2 = Trim(rsHlp("KDPLG"))
79                            gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
80                            gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
81                            gInvKota2 = Trim(rsHlp("FS_KOTA"))
82                            gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                              
83                            If Trim(rsHlp("INTERNAL")) = 1 Then
84                                If Trim(gNpwp) = Trim(gInvNPWP2) Then
85                                    vTipeCustomer = "AFFNONPPN"
86                                Else
87                                    vTipeCustomer = "AFFPPN"
88                                End If
89                            Else
90                                vTipeCustomer = "EXT"
91                            End If
92                        End If
93                        If rsHlp.State <> adStateClosed Then rsHlp.Close
94                        Call giveDiscount("100", text(21), False)
95                        text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
96                    End If
97                Else
                      ' HANYA RTJ
98                    Call giveDiscount("100", text(21), False)
99                    vTipeCustomer = "INT"
100               End If
101           Else
102               If typeJob(Trim(rsTemp("TPSVC"))) = "WARRANTY" Then
103                   opWO(2).value = True
                       
104                   Set rsHlp = objBilling.searchWarrantyDetail(rsTemp("KODETIPE"))
105                   If Not rsHlp.EOF Then
106                       text(35) = rsHlp("ID_WARRANTY")
107                       text(36) = rsHlp("WARRANTY")
108                       text(35).Visible = True
109                       text(36).Visible = True
110                   End If
111                   If rsHlp.State <> adStateClosed Then rsHlp.Close
112                   Call giveDiscount("0", text(21), True)
113                   text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
114                   Set rsHlp = objO.FetchData("SELECT * FROM T_PROFIL WITH (NOLOCK)")
115                   If Not rsHlp.EOF Then
116                       gKdPlgDealer = Trim(rsHlp("FN_HOKDPLG"))
117                       gKodePlg = Trim(rsHlp("FN_HOKDPLG"))
118                       gInvCust = Trim(rsHlp("FN_HO"))
119                       gInvAlmt = Trim(rsHlp("FN_HOADRES"))
120                       gInvKota = Trim(rsHlp("FN_HOKOTA"))
121                       gInvNPWP = Trim(rsHlp("FN_HONPWP"))
122                       vTipeCustomer = "AFFPPN"
123                   End If
124                   If rsHlp.State <> adStateClosed Then rsHlp.Close
125               ElseIf typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
126                   opWO(1).value = True
                       
127                   Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
128                   If Not rsHlp.EOF Then
129                       text(35) = rsHlp("ID_PROGRAM")
130                       text(36) = rsHlp("PROGRAM")
131                       text(35).Visible = True
132                       text(36).Visible = True
133                   End If
134                   If rsHlp.State <> adStateClosed Then rsHlp.Close
135                   If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
136                       nDuoInv = True
                          'SBI WO BERKEMBANG
137                       If rsTemp("FS_RC") = "R001" Then
                               
138                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
139                           If Not rsHlp.EOF Then
140                               gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
141                               gKodePlg2 = Trim(rsHlp("KDPLG"))
142                               gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
143                               gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
144                               gInvKota2 = Trim(rsHlp("FS_KOTA"))
145                               gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                                  
146                               If Trim(rsHlp("INTERNAL")) = 1 Then
147                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
148                                       vTipeCustomer = "AFFNONPPN"
149                                   Else
150                                       vTipeCustomer = "AFFPPN"
151                                   End If
152                               Else
153                                   vTipeCustomer = "EXT"
154                               End If
155                           End If
156                           If rsHlp.State <> adStateClosed Then rsHlp.Close
157                           Call giveDiscount("100", text(21), False)
158                           text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
159                       Else
                               
160                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
161                           If Not rsHlp.EOF Then
162                               gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
163                               gKodePlg2 = Trim(rsHlp("KDPLG"))
164                               gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
165                               gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
166                               gInvKota2 = Trim(rsHlp("FS_KOTA"))
167                               gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                                  
168                               If Trim(rsHlp("INTERNAL")) = 1 Then
169                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
170                                       vTipeCustomer = "AFFNONPPN"
171                                   Else
172                                       vTipeCustomer = "AFFPPN"
173                                   End If
174                               Else
175                                   vTipeCustomer = "EXT"
176                               End If
177                           End If
178                           If rsHlp.State <> adStateClosed Then rsHlp.Close
179                           Call giveDiscount("100", text(21), False)
180                           text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
181                       End If
182                   Else
183                       Call giveDiscount(objBilling.takeDiscountProgram(rsTemp("KODETIPE")), text(21), False)
184                   End If
185               ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
186                   If lblPDS.Visible Then
187                       Call giveDiscount("0", text(21), False)
188                       Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
189                       If Not rsHlp.EOF Then
190                           gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
191                           gKodePlg2 = Trim(rsHlp("KDPLG"))
192                           gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
193                           gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
194                           gInvKota2 = Trim(rsHlp("FS_KOTA"))
195                           gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                              
196                           If Trim(rsHlp("INTERNAL")) = 1 Then
197                               If Trim(gNpwp) = Trim(gInvNPWP2) Then
198                                   vTipeCustomer = "AFFNONPPN"
199                               Else
200                                   vTipeCustomer = "AFFPPN"
201                               End If
202                           Else
203                               vTipeCustomer = "EXT"
204                           End If
205                       End If
206                       If rsHlp.State <> adStateClosed Then rsHlp.Close
207                   ElseIf rsTemp("FS_RC") = "R001" Then
                          ' HANYA SBI
208                       woStat = "SBI"
                           
209                       Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
210                       If Not rsHlp.EOF Then
211                           gKdPlgDealer = Trim(rsHlp("KDPLG"))
212                           gKodePlg = Trim(rsHlp("KDPLG"))
213                           gInvCust = Trim(rsHlp("FS_NM_DEALER"))
214                           gInvAlmt = Trim(rsHlp("FS_ALAMAT"))
215                           gInvKota = Trim(rsHlp("FS_KOTA"))
216                           gInvNPWP = Trim(rsHlp("FS_NPWP"))
217                           If Trim(rsHlp("INTERNAL")) = 1 Then
218                               If Trim(gNpwp) = Trim(gInvNPWP) Then
219                                   text(40) = "AFFNONPPN"
220                               Else
221                                   text(40) = "AFFPPN"
222                               End If
223                           Else
224                               text(40) = "EXT"
225                           End If
226                       End If
227                       If rsHlp.State <> adStateClosed Then rsHlp.Close
228                       Call giveDiscount("0", text(21), False)
229                       text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
230                   ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
                           
231                       Set rsHlp = objBilling.searchMOUDetail(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS"))
232                       If Not rsHlp.EOF Then
233                           text(34) = Trim(rsHlp("NO_MOU"))
234                       End If
235                       If rsHlp.State <> adStateClosed Then rsHlp.Close
236                       Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")), text(21), False)
237                   Else
238                       Call giveDiscount("0", text(21), True)
239                   End If
240               End If
241           End If
              
              ' TAMPIL JOB ORDER
              Dim nilaiRate As Double
242           totRate = 0
243           text(22) = 0
244           subJob = 0
245           If rsJob.State Then rsJob.Close
246           myChassis = Trim(text(17))
              Dim nilaiDiscJasa As Integer
              
247           For i = 1 To 7
248               hlp = "REP" & i
249               hlp2 = "RTE" & i
250               hlp3 = "KLPKERJA" & i
251               hlp4 = "ACTRATE" & i
252               hlp5 = "MEKANIK" & i
253               nilaiDiscJasa = 0
254               If rsTemp(hlp) <> "" Then
255                   fOth = False
256                   Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "' AND MDL = '" & rsTemp("MDL") & "'")
257                   If rsHlp.EOF Then
258                       fOth = True
259                   End If
260                   If rsHlp.State <> adStateClosed Then rsHlp.Close
                      
261                   If returnJob(LblReturn) Then
262                       If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
263                           If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
264                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
265                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
266                           End If
267                       Else
268                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
269                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
270                       End If
271                   Else
272                       If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
273                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
274                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
275                       ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
276                           If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
277                               If rsTemp("FS_RC") = "R001" Then
278                                   If Trim(rsTemp(hlp3)) = "SBI" Then
                                           
279                                       Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
280                                       If Not rsHlp2.EOF Then
281                                           nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
282                                           If nStatusDealer = 1 Then
283                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
284                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
285                                           Else
286                                               If nStatusDealer = 2 Then
287                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
288                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
289                                               ElseIf nStatusDealer = 3 Then
290                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
291                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
292                                               End If
293                                           End If
294                                       End If
295                                       If rsHlp2.State <> adStateClosed Then rsHlp2.Close
296                                   Else
297                                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
298                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
299                                   End If
300                               Else
                                       
301                                   Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
302                                   If Not rsHlp2.EOF Then
                                          'Dim nStatusDealer As Integer
303                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
304                                       If nStatusDealer = 1 Then
305                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0  '
306                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
307                                       Else
308                                           If nStatusDealer = 2 Then
309                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
310                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
311                                           ElseIf nStatusDealer = 3 Then
312                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
313                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
314                                           End If
315                                       End If
                                          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
316                                   End If
317                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
318                               End If
319                           Else
320                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
321                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
322                           End If
323                       ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
324                           If lblPDS.Visible Then
325                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
326                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
327                           ElseIf rsTemp("FS_RC") = "R001" Then
                                   
328                               Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
329                               If Not rsHlp2.EOF Then
330                                   nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
331                                   If nStatusDealer = 1 Then
332                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
333                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
334                                   Else
335                                       If nStatusDealer = 2 Then
336                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
337                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
338                                       ElseIf nStatusDealer = 3 Then
339                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
340                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
341                                       End If
342                                   End If
                                      'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
343                               End If
344                               If rsHlp2.State <> adStateClosed Then rsHlp2.Close
345                           ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
346                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
347                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
348                           Else
349                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
350                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
351                           End If
352                       End If
353                   End If
354                   nilaiRate = CDbl(rsTemp(hlp2))
355                   nilaiDiscJasa = CInt(text(21))
356                   If Log.discKacab = 1 Then
357                       Set rsHlp2 = objO.FetchData("SELECT discJasaValid as jasa FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(text(Index)) & "' ")
358                       If Not rsHlp2.EOF And nDuoInv = False Then
359                           nilaiDiscJasa = CInt(rsHlp2!Jasa)
360                       End If
361                   End If
362                   text(22) = text(22) + (CDbl(nHarga) * nilaiRate) * ((100 - nilaiDiscJasa) / 100)
                      
363                   With rsJob
364                   If .State = adStateClosed Then doCreateJob
365                      .AddNew
366                      !job = Trim(rsTemp(hlp))
367                      !Rate = nilaiRate 'CDbl(rsTemp(hlp2))
368                      !ACTRT = CDbl(rsTemp(hlp4))
369                      !grp = Trim(rsTemp(hlp3))
370                      !Mech = Trim(rsTemp(hlp5))
371                      !hrg = CDbl(nHarga) * nilaiRate
372                      !addJob = False
373                      !disc = CDbl(nilaiDiscJasa)
374                      !DISCRP = CDbl(!hrg) * (CDbl(nilaiDiscJasa) / 100)
375                      !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
376                      !valid = gVerifikator
377                      !FREEC = False
378                      .Update
379                   End With
380                   subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
381                   totRate = totRate + CDbl(nilaiRate)
382               End If
383           Next
              
              ' add JOB
384           Set rsHlp2 = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & noWO & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
385           If Not rsHlp2.EOF Then
386               Do While Not rsHlp2.EOF
      '            For i = 1 To 10
      '                hlp = "REP" & i
      '                hlp2 = "RTE" & i
      '                hlp3 = "KLPKERJA" & i
      '                hlp4 = "ACTRATE" & i
      '                hlp5 = "MEKANIK" & i
                      
387                   If rsHlp2("REP") <> "" Then
388                       If returnJob(LblReturn) Then
389                           If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
390                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
391                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
392                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
393                               End If
394                           Else
395                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
396                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
397                           End If
398                       Else
399                           If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
400                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
401                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                  'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
402                           ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
403                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
404                                   If rsTemp("FS_RC") = "R001" Then
405                                       If rsHlp2("KLPKERJA") = "SBI" Then
                                               
406                                           Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
407                                           If Not rsHlp3.EOF Then
408                                               nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
409                                               If nStatusDealer = 1 Then
410                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
411                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
412                                               Else
413                                                   If nStatusDealer = 2 Then
414                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
415                                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
416                                                   ElseIf nStatusDealer = 3 Then
417                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
418                                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
419                                                   End If
420                                               End If
                                                  'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
421                                           End If
422                                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
423                                       Else
424                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
425                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
426                                       End If
427                                   Else
                                           
428                                       Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
429                                       If Not rsHlp3.EOF Then
430                                           nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
431                                           If nStatusDealer = 1 Then
432                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0 '
433                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
434                                           Else
435                                               If nStatusDealer = 2 Then
436                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
437                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
438                                               ElseIf nStatusDealer = 3 Then
439                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
440                                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
441                                               End If
442                                           End If
                                              'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
443                                       End If
444                                       If rsHlp3.State <> adStateClosed Then rsHlp3.Close
445                                   End If
446                               Else
447                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
448                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
449                               End If
450                           ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
451                               If lblPDS.Visible Then
452                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
453                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
454                               ElseIf rsTemp("FS_RC") = "R001" Then
                                       
455                                   Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
456                                   If Not rsHlp3.EOF Then
457                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
458                                       If nStatusDealer = 1 Then
459                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
460                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
461                                       Else
462                                           If nStatusDealer = 2 Then
463                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
464                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
465                                           ElseIf nStatusDealer = 3 Then
466                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
467                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
468                                           End If
469                                       End If
                                          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
470                                   End If
471                                   If rsHlp3.State <> adStateClosed Then rsHlp3.Close
472                               ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
473                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
474                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
475                               Else
476                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
477                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
478                               End If
479                           End If
480                       End If
                          
                          ' ADJUSTMENT RATE
481                       nilaiRate = CDbl(rsHlp2("RTE"))
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
                          
482                       text(22) = text(22) + (CDbl(nHarga) * CDbl(nilaiRate)) * ((100 - text(21)) / 100)
                              
483                       With rsJob
484                       If .State = adStateClosed Then doCreateJob
485                           .AddNew
486                           !job = Trim(rsHlp2("REP"))
487                           !Rate = nilaiRate 'CDbl(rsHlp2("RTE"))
488                           !ACTRT = CDbl(rsHlp2("ACTRTE"))
489                           !grp = Trim(rsHlp2("KLPKERJA"))
490                           !Mech = Trim(rsHlp2("MEKANIK"))
491                           !hrg = CDbl(nHarga) * CDbl(nilaiRate)
492                           !addJob = True
493                           !disc = CDbl(text(21))
494                           !DISCRP = CDbl(!hrg) * (CDbl(text(21)) / 100)
495                           !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
496                           !valid = gVerifikator
497                           !FREEC = False
498                           .Update
499                           adaAddJob = True
500                       End With
501                       subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
502                       totRate = totRate + CDbl(nilaiRate)
503                   End If
      '            Next
504                   rsHlp2.MoveNext
505               Loop
506           End If
              
      '        subJob = CDbl(Text(22))
507           Set JobGrid.DataSource = rsJob
508           Call doHeadJob
              
              ' TAMPIL PART ORDER
509           noKFaktur = ""
510           gSpart = True
511           text(24) = 0
512           text(25) = 0
513           subPartGenuine = 0
514           subPartNGenuine = 0
515           subOliGenuine = 0
516           subOliNGenuine = 0
              Dim HargaDisc As Double
517           If rsPart.State Then rsPart.Close
              
              ' /// AMBIL PART YG BAWA SENDIRI //
              ' AND B.TIPE = 'O'
518           Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, TIPE, GENUINE, BAWA FROM jobpart a WITH (NOLOCK) LEFT JOIN T_oli b WITH (NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli WHERE no_WO = '" & noWO & "' " _
                                      & "AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y' ") 'AND BAWA <> 'Y'
519           If Not rsHlp.EOF Then
520               rsHlp.MoveFirst
521               Do While Not rsHlp.EOF
522                   With rsPart
523                   If .State = adStateClosed Then doCreatePart
524                       gSpart = True
                          
525                       If rsHlp("BAWA") = "Y" Then
526                           .AddNew
527                           !KodeItem = "BAWA SENDIRI"
528                           !NAMAITEM = rsHlp("PARTBAWA")
529                           !qty = rsHlp("QTY")
530                           !QtyOrder = rsHlp("QTY")
531                           !QtyBO = 0
532                           !QtyCancel = 0
533                           !Harga = 0
534                           !oli = IIf(rsHlp("TIPE") <> "P", True, False)
535                           .Update
536                       Else
537                           If Not gFPart Then
                                  '///// TIDAK CONNECT ACCOUNTING SPAREPART ////
538                               .AddNew
539                               !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
540                               !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
541                               !oli = IIf(rsHlp("TIPE") <> "P", True, False)
542                               !qty = rsHlp("QTY")
                                      'Set rsHlp2 = objO.FetchDataACC("SELECT hargajualA FROM inmst WHERE kodeItem = '" & rsHlp("KODE_ITEM") & "'")
543                                   Set rsHlp2 = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & rsHlp("KODE_ITEM") & "'")
544                                   If Not rsHlp2.EOF Then
545                                       HargaDisc = CDbl(rsHlp2("HARGAJUALA"))
546                                       !Harga = CDbl(HargaDisc)
547                                       If rsHlp("OLI") <> "" Then
548                                           If rsHlp("TIPE") = "P" Then
549                                               If rsHlp2("TGenuine") = False Then
                                                      ' PART GENUINE
550                                                   subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
551                                               ElseIf rsHlp2("TGenuine") = True Then
                                                      ' PART NON GENUINE
552                                                   subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
553                                               End If
554                                               text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
555                                           ElseIf rsHlp("TIPE") = "O" Then
556                                               If rsHlp2("TGenuine") = False Then
                                                      ' OLI GENUINE
557                                                   subOliNGenuine = subOliNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
558                                               ElseIf rsHlp2("TGenuine") = True Then
                                                      ' OLI NON GENUINE
559                                                   subOliGenuine = subOliGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
560                                               End If
561                                               text(25) = text(25) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
562                                           End If
563                                       Else
                                              ' PART GENUINE
564                                           subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
565                                           text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
566                                       End If
567                                   End If
568                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                                  
569                               .Update
570                           End If
571                       End If
572                   End With
573                   rsHlp.MoveNext
574               Loop
575           End If
576           If rsHlp.State <> adStateClosed Then rsHlp.Close
              
577           If gFPart Then
                  '/// KONEKSI KE SPAREPART
578               If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
579                   With rsPart
580                       If noinv Then
                              Dim strSql As String
581                           If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
582                               strSql = "select (case When KodeItem = '' Then namaBYSA Else KodeItem end) as KodeItem, " & _
                                       "Keterangan as Namabarang, SatuanNs as Satuan, Qty as QtyOrder, (isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0)) AS Jml, " & _
                                       "(Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) as QtyBO, HargaJual as HargaJualA, NOSO as NoRack, " & _
                                       "0 as DisItem, 0 as ExtDis, Case BO When 'CANCEL' THEN (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) ELSE 0 END AS QTYCANCEL from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(text(0)) & "'"
583                           Else
584                               strSql = "Select Z.*, 0 as QtyOrder, 0 as QtyBO FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                      "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                      "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
                                      "INMST.NORACK,0 as DisItem, 0 as ExtDis, 0 as QtyCANCEL FROM INMST " & _
                                      "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                      "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                      "Where Z.Norack = '" & Trim(text(0)) & "'"
585                           End If
586                           Set rsHlp3 = objO.FetchDataACC(strSql)

587                           If Not rsHlp3.EOF Then
588                               Do While Not rsHlp3.EOF
589                                   If .State = adStateClosed Then doCreatePart
590                                   .AddNew
591                                   !KodeItem = Trim(rsHlp3("KODEITEM"))
592                                   !NAMAITEM = Trim(rsHlp3("NAMABARANG"))
593                                   !QtyOrder = CDbl(rsHlp3("QtyOrder"))
594                                   !qty = CDbl(rsHlp3("JML"))
595                                    If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
596                                       !QtyBO = IIf(UCase(isBO) = "CANCEL", 0, objO.FetchData_DoubleACC("select Qty FROM BOFILE WITH (NOLOCK) Where TutupOrderFlag = 0 and (NoSO = '" & Trim(text(0)) & "' Or NOSO =(Select Distinct NoSPK From SOORD WITH (NOLOCK) Where NOSO ='" & Trim(text(0)) & "' and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "')) and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "'"))
597                                    Else
598                                       !QtyBO = 0
599                                    End If
                                      
600                                   !QtyCancel = CDbl(rsHlp3("QtyCancel"))
601                                   !Harga = CDbl(rsHlp3("HARGAJUALA"))
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
602                                       Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp3("KODEITEM") & "'")
603                                       If Not rsHlp.EOF Then
604                                           If Left(rsHlp("SCC"), 3) = "OLI" Then
605                                               !oli = True
606                                               text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
607                                           Else 'If Left(rsHlp("SCC"), 3) = "PAR" Then
608                                               !oli = False
609                                               text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
      '                                            !oli = False
      '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
610                                           End If
611                                       Else
612                                           !oli = False
613                                           text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
614                                       End If
615                                   .Update
616                                   rsHlp3.MoveNext
617                               Loop
618                           Else 'If rsHlp3.EOF And newValid Then
                                  Dim rsInvPart As New ADODB.Recordset
                                  Dim vHitung As Double
619                               Set rsInvPart = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                                   & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' ")
620                               If Not rsInvPart.EOF Then
621                                   Do While Not rsInvPart.EOF
622                                       If .State = adStateClosed Then doCreatePart
623                                       .AddNew
624                                       !KodeItem = Trim(rsInvPart("KODEITEM"))
625                                       !NAMAITEM = Trim(rsInvPart("KETERANGAN"))
626                                       vHitung = (Koma2Titik(rsInvPart("Qty")) * (Koma2Titik(rsInvPart("HargaJual")) - (Koma2Titik(rsInvPart("ItemDisc")) / Abs(Koma2Titik(rsInvPart("Qty"))))))
627                                       vHitung = vHitung - (Koma2Titik(rsInvPart("ExtDisPersen")) * vHitung)
628                                       If Trim(UCase(rsInvPart("TypePpn"))) = "E" Then
629                                           vHitung = Round(vHitung + vHitung * 0.1, 2)
630                                       End If
631                                       If gJurnalPart Then
632                                           vHitung = cekIncludePPn(vHitung)
633                                       End If
634                                       !qty = Val(rsInvPart("QTY"))
635                                       If !qty = 0 Then
636                                           !Harga = 0
637                                       Else
638                                           !Harga = CDbl(vHitung / !qty)
639                                       End If
640                                       !QtyOrder = Val(rsInvPart("QTY"))
641                                       !QtyBO = 0
642                                       !QtyCancel = 0
                                           
                                          'Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsInvPart("KODEITEM"), 1) & "'")
643                                       Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsInvPart("KODEITEM") & "'")
644                                       If Not rsHlp.EOF Then
645                                           If Left(rsHlp("SCC"), 3) = "OLI" Then
646                                               !oli = True
647                                               text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
648                                           Else 'If Left(rsHlp("SCC"), 3) = "PAR" Then
649                                               !oli = False
650                                               text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
      '                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
      '                                            !oli = False
      '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
651                                           End If
652                                       Else
653                                           !oli = False
654                                           text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
655                                       End If
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
656                                       .Update
657                                       rsInvPart.MoveNext
658                                   Loop
659                               Else
660                                   gSpart = False
      '                                If TblValid.Enabled Then
      '                                    temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
      '                                    If temp = vbYes Then
      '                                        gSpart = False
      '                                    End If
      '                                End If
661                               End If
662                               If rsInvPart.State <> adStateClosed Then rsInvPart.Close
663                           End If
664                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
665                       Else 'Not newValid Then
                               
666                           Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND NOINV LIKE '" & Trim(text(20)) & "%'")
667                           If Not rsHlp2.EOF Then
668                               Do While Not rsHlp2.EOF
669                                   If .State = adStateClosed Then doCreatePart
670                                   .AddNew
671                                   !KodeItem = Trim(rsHlp2("KODEITEM"))
672                                   !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
673                                   !qty = CDbl(rsHlp2("QTY"))
674                                   !Harga = CDbl(rsHlp2("HARGA")) '* ((100 - CDbl(rsHlp2("FN_EXT_DISKON"))) / 100)
675                                   !oli = Trim(rsHlp2("OLI")) 'IIf(Trim(rsHlp2("OLI")) = 1, True, False)
676                                   If !oli Then
677                                       text(25) = text(25) + CDbl(!Harga)
678                                   Else
679                                       text(24) = text(24) + CDbl(!Harga)
680                                   End If
681                                   !EXTDISC = CDbl(rsHlp2("FN_EXT_DISKON"))
682                                   !QtyOrder = CDbl(rsHlp2("QTY"))
683                                   !QtyBO = 0
684                                   !QtyCancel = 0
685                                   .Update
686                                   rsHlp2.MoveNext
687                               Loop
688                           Else
689                               gSpart = False
690                           End If
691                           If rsHlp2.State <> adStateClosed Then rsHlp2.Close
692                       End If
693                   End With
694               End If
              
695               If Not gSpart Then
696                       With rsPart
697                           Set rsHlp3 = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                             & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' ")
                              
698                           If Not rsHlp3.EOF Then
699                               Do While Not rsHlp3.EOF
700                                   If .State = adStateClosed Then doCreatePart
701                                   .AddNew
702                                   !KodeItem = Trim(rsHlp3("KODEITEM"))
703                                   !NAMAITEM = Trim(rsHlp3("KETERANGAN"))
704                                   vHitung = (Koma2Titik(rsHlp3("Qty")) * (Koma2Titik(rsHlp3("HargaJual")) - (Koma2Titik(rsHlp3("ItemDisc")) / Abs(Koma2Titik(rsHlp3("Qty"))))))
705                                   vHitung = vHitung - (Koma2Titik(rsHlp3("ExtDisPersen")) * vHitung)
706                                   If Trim(UCase(rsHlp3("TypePpn"))) = "E" Then
707                                       vHitung = Round(vHitung + vHitung * 0.1, 2)
708                                   End If
709                                   If gJurnalPart Then
710                                       vHitung = cekIncludePPn(vHitung)
711                                   End If
712                                   !qty = CDbl(rsHlp3("Qty"))
713                                   !Harga = CDbl(vHitung / !qty)
714                                   Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
715                                   If Not rsHlp.EOF Then
716                                       If rsHlp("TIPE") = "O" Then
717                                           !oli = True
718                                           text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
719                                       ElseIf rsHlp("TIPE") = "P" Then
720                                           !oli = False
721                                           text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
722                                       End If
723                                   End If
724                                   !QtyOrder = CDbl(rsHlp3("Qty"))
725                                   !QtyBO = 0
726                                   !QtyCancel = 0
727                                   .Update
728                                   rsHlp3.MoveNext
729                               Loop
730                               gSpart = True
731                           Else
732                               If TblValid.Enabled Then
733                                   temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
734                                   If temp = vbYes Then
735                                       gSpart = False
736                                   Else
737                                       gSpart = True
738                                   End If
739                               End If
740                           End If
741                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
742                       End With
743                   Else
744                       gSpart = True
745                   End If
746               End If
747           End If
              
748           subPart = CDbl(text(24))
749           subOli = CDbl(text(25))
750           If rsPart.State <> adStateClosed Then
751               If Not rsPart.BOF Then
752                   rsPart.MoveFirst
753                   rsPart.Sort = "OLI"
754                   Set partGrid.DataSource = rsPart
755                   Call doHeadPart
756               End If
757           End If
              
              '---------
              
              Dim rsWashing As ADODB.Recordset
758           Set rsWashing = objO.FetchData("select washing from jobdtl where no_wo = '" & noWO & "'")
              
759           If rsWashing.RecordCount > 0 Then
760               If rsWashing("washing") = True Then
                      'pengecekan sudah proses cuci
761                   Set rsWashing = objO.FetchData("SELECT * FROM TEMPPROGRESSCUCI WHERE NOORDER = '" & noWO & "'")
762                   If rsWashing.RecordCount > 0 Then
763                       btnBatalWashing.Enabled = False
764                   Else
765                       btnBatalWashing.Enabled = True
766                   End If
767               Else
768                   btnBatalWashing.Enabled = False
769               End If
770           Else
771               btnBatalWashing.Enabled = False
772           End If
          
              ' TAMPIL SUBLET ORDER
773           adaSublet = False
              'cmdSublet.Visible = False
774           text(27) = 0
775           If rsSublet.State Then rsSublet.Close
776           Set rsSublet = objO.FetchData("SELECT NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, valid, HARGA, PRICE FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noWO & "' AND no_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
777           If Not rsSublet.EOF Then
778               Do While Not rsSublet.EOF
779                   If rsSublet("VALID") = 0 Then
780                       adaSublet = True
781                       MsgBox "Sublet belum di validasi"
782                       If rsSublet.State Then rsSublet.Close
783                       TblValid.Enabled = False
784                       Exit Sub
785                   End If
786                   text(27) = text(27) + (IIf(Trim(text(20)) = "", CDbl(rsSublet("HARGA")), CDbl(rsSublet("PRICE"))) * CDbl(rsSublet("QTY")))
787                   rsSublet.MoveNext
788               Loop
                  ''2016-07-25 || remark sintaks insert tempjobsublet waktu panggil order di  menu validasi order.
      '            Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
      '                 & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "') and no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "' " _
      '                 & "ORDER BY nonota, kodeitem")
789               subSublet = CDbl(text(27))
790               adaSublet = True
                  'cmdSublet.Visible = True
791               Set subletGrid.DataSource = rsSublet
792               Call doHeadSublet
793           Else
794               If newValid Then
795                   If TblValid.Enabled Then
796                       If MsgBox("Ada Sublet ?", vbYesNo, "Konfirmasi Sublet") = vbNo Then
797                           adaSublet = True
798                       End If
799                   End If
800               End If
801           End If
              
802           text(29) = Round(CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)), 0) '+ CDbl(text(28))
              ' TAMPIL TERBILANG
803           txtTotalRp = Angka2Huruf(text(29), 1)
804               If OnTopSV = 0 Then
805                   Set rsHlp = objO.FetchDataACC("select KodeLgn, SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(text(2)) & "'")
806               Else
                  'UNTUK FEBUARI - TOP PELANGGAN
807                   Set rsHlp = objO.FetchDataACC("select KodeLgn, TOPGRP as SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(text(2)) & "'")
808               End If
809               If Not rsHlp.BOF Then
810                   text(41) = rsHlp("SyaratBayar")
811               End If
812               If rsHlp.State <> adStateClosed Then rsHlp.Close
              
813           If Not gSpart Or Not adaSublet Then
814               TblValid.Enabled = False
815           End If
816           If Log.dmsNonDss = 1 Then
                  'lakukan pengecekan apakah harus menginputkan DP part
                  'total > plafon = dp >= prosentase * total
                  '2014-12-19 RUDI KRISTANTO
      '
      '            If Trim(Text(64)) <> "" Or Trim(Text(78)) <> "" Then
      '                totalNilaiPart = getTotalPart(IIf(Trim(Text(64)) <> "", Trim(Text(64)), Trim(Text(78))), False)
      '            Else
      '                totalNilaiPart = CDbl(Trim(Text(57))) + CDbl(Trim(Text(59)))
      '            End If
817               totalNilaiPart = 0
818               totalNilaiPart = CDbl(Trim(text(24))) + CDbl(Trim(text(25)))
                              
819               If totalNilaiPart > Log.plafonDP Then
                      'lakukan pengisian DP
820                   MsgBox "Nilai part membutuhkan DP minimal Rp." & Format(((Log.persentaseDP / 100) * totalNilaiPart), "0,#"), vbInformation, "Informasi DP Part"
                      'boleh = False
821               End If
822           End If
              
              ''2017-08-01 || Rev20170801A
823           If Log.discKacab = 1 Then
824               If cmdBatalRequest.Enabled = True Then
                      ''tombol valid dibuat disable saat request disc belum disetujui
825                   TblValid.Enabled = False
826               End If
827           End If
              

828      On Error GoTo 0
829      Exit Sub

panggilWO_Error:

830       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilWO of Form frmOrderValidasi")
              
End Sub

Private Sub hapusForm()
1        On Error GoTo hapusForm_Error

2         'Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & text(0) & "'")
          
3         For i = 0 To 44
4             Select Case i
                  Case 1, 11, 43
5                     text(i).Mask = "99-99-9999"
6                 Case 10, 21, 22, 23, 24, 25, 26, 27, 29, 39
7                     text(i) = 0
8                 Case 42, 30, 33
9                 Case Else
10                    text(i) = ""
11            End Select
12        Next
          
13        VScroll1.value = 0
14        opWO(0).value = True
          
15        If rsJob.State Then rsJob.Close
16        Set rsJob = Nothing
17        Set JobGrid.DataSource = rsJob
          'JobGrid.ReBind
          
18        If rsPart.State Then rsPart.Close
19        Set rsPart = Nothing
20        Set partGrid.DataSource = rsPart
          'partGrid.ReBind
          
21        Set subletGrid.DataSource = Nothing
22        adaAddJob = False
          
          ' ISI KARAKTER PELANGGAN
23        Set rsTemp = objO.FetchData("SELECT * FROM T_Karakter WITH (NOLOCK) ORDER BY KODE")
24        If Not rsTemp.EOF Then
25            cboPelanggan.Clear
26            Do While Not rsTemp.EOF
27                cboPelanggan.AddItem (rsTemp("KODE"))
28                rsTemp.MoveNext
29            Loop
30        End If
31        cboPelanggan.ListIndex = 0
          
32        vTipeCustomer = ""
33        woStat = ""
34        nDuoInv = False
35        chPS.value = Unchecked
36        chNea.value = Unchecked

37       On Error GoTo 0
38       Exit Sub

hapusForm_Error:

39        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmOrderValidasi")
End Sub

Private Sub btnBatalWashing_Click()
1        On Error GoTo btnBatalWashing_Click_Error
         
2         If Trim(text(0)) <> "" Then
3              If Log.lockWashing = 1 Then
4                 If MsgBox("Apakah Anda yakin unit tidak dicuci?" & vbCrLf & "Jika Yes maka pilihan anda akan direport ke HO!", vbYesNo, "") = vbNo Then
5                    Exit Sub
6                 End If
7              End If
               
8             Set rsTmp = objO.FetchData("SELECT FLAG,JAMMULAI,NOSTALL,JAMSELESAI FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & Trim(text(0)) & "|%'")
9             If Not rsTmp.EOF Then
                  ' TUTUP TAMBAH PEKERJAAN
10                If UCase(rsTmp("FLAG")) = "G" Then
11                    framePembatalanWashing.Visible = True
12                End If
13            End If
14        End If

15       On Error GoTo 0
16       Exit Sub

btnBatalWashing_Click_Error:

17        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure btnBatalWashing_Click of Form frmOrderValidasi")
End Sub

Private Sub btnSimpan_Click()
1        On Error GoTo btnSimpan_Click_Error

2         Call objO.UpdateData("Update Jobdta Set Catat = '" & Trim(Txt.text) & "' Where NO_WO = '" & Trim(text(0).text) & "'")
3         MsgBox "Catatan telah disimpan !", vbInformation + vbOKOnly

4        On Error GoTo 0
5        Exit Sub

btnSimpan_Click_Error:

6         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure btnSimpan_Click of Form frmOrderValidasi")
End Sub

Private Sub chNea_Click()
    If chNea.value = 1 Then
        txtInfoNea.Visible = True
        txtInfoNea.text = ""
    Else
        txtInfoNea.Visible = False
    End If
End Sub

Private Sub chNea_GotFocus()
    VScroll1.value = 2500
End Sub

Private Sub chPS_GotFocus()
    VScroll1.value = 2500
End Sub

Private Sub cmdKeluh_Click()
      Dim rsSearch As New ADODB.Recordset
      Dim rscr7 As ADODB.Recordset
1        On Error GoTo cmdKeluh_Click_Error

2         If Log.dmsNonDss = 1 Then
3             txtPermintaan = ""
4             txtKeluhan = ""
5             txtFI = ""
6             txtNextJob = ""
              
7             If FrameKeluhDmsNonDSS.Visible Then
8                 FrameKeluhDmsNonDSS.Visible = False
9             Else
10                Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(text(0)) & "'")
11                If Not rsSearch.EOF Then
12                    Text3.text = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
13                    Text2.text = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
14                End If
15                Set rsSearch = objO.FetchData("SELECT * FROM INSPECTION A WITH (NOLOCK) WHERE NOWO = '" & Trim(text(0)) & "'")
16                If Not rsSearch.EOF Then
17                    txtFI = Trim(rsSearch("HASILFI")) & vbCrLf & Trim(rsSearch("HasilTestDrive"))
18                    txtNextJob = Trim(rsSearch("SNJ1")) & vbCrLf & Trim(rsSearch("SNJ2")) & vbCrLf & Trim(rsSearch("SNJ3")) & vbCrLf & Trim(rsSearch("SNJ4")) & vbCrLf & Trim(rsSearch("SNJ5")) & vbCrLf & Trim(rsSearch("SNJ6")) & vbCrLf & Trim(rsSearch("SNJ7"))
19                End If
20                FrameKeluhDmsNonDSS.Visible = True
21            End If
              
              
22            Set rscr7 = objO.FetchData("Select Catat FROM JOBDTA WITH (NOLOCK) Where NO_WO = '" & Trim(text(0)) & "'")
23            If rscr7.RecordCount > 0 Then
24                Text1.text = rscr7.Fields(0).value
25            Else
26                Text1.text = ""
27            End If
28        Else
          
29            txtPermintaan = ""
30            txtKeluhan = ""
31            If framKeluh.Visible Then
32                framKeluh.Visible = False
33            Else
34                Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(text(0)) & "'")
35                If Not rsSearch.EOF Then
36                    txtPermintaan = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
37                    txtKeluhan = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
38                End If
39                framKeluh.Visible = True
40            End If
              
           
41            Set rscr7 = objO.FetchData("Select Catat FROM JOBDTA WITH (NOLOCK) Where NO_WO = '" & Trim(text(0)) & "'")
42            If rscr7.RecordCount > 0 Then
43                Txt.text = rscr7.Fields(0).value
44            Else
45                Txt.text = ""
46            End If
              
47        End If

48       On Error GoTo 0
49       Exit Sub

cmdKeluh_Click_Error:

50        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdKeluh_Click of Form frmOrderValidasi")
End Sub

Private Sub Command1_Click()
1        On Error GoTo Command1_Click_Error

      'Call cmdKeluh_Click
2         Call objO.UpdateData("Update Jobdta Set Catat = '" & Trim(Text1.text) & "' Where NO_WO = '" & Trim(text(0).text) & "'")
3         MsgBox "Catatan telah disimpan !", vbInformation + vbOKOnly

4        On Error GoTo 0
5        Exit Sub

Command1_Click_Error:

6         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Command1_Click of Form frmOrderValidasi")
End Sub

Private Sub Command2_Click()
    If Trim(txtAlasanBtlWashing.text) <> "" Then
        If Len(txtAlasanBtlWashing.text) < 5 Then
            MsgBox "Alasan Pembatalan Washing minimal 5 karakter!", vbCritical, "Informasi Pembatalan Washing"
        Else
            On Error GoTo vbError
            erTranAll = False
            objHT.OpenTransactionBengkel
            
            Call objHT.InsertData_Bengkel("INSERT INTO WASHING VALUES ('" & Trim(text(0)) & "',GETDATE(),'" & gSA & "','" & txtAlasanBtlWashing.text & "')")
            Call objHT.InsertData_Bengkel("UPDATE ANTRICUCI SET FLAG = 'CANCEL' WHERE NOORDER = '" & Trim(text(0)) & "'")
            Call objHT.InsertData_Bengkel("UPDATE JOBDTL SET WASHING = 0 WHERE NO_wo = '" & Trim(text(0)) & "'")
            Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%" & Trim(text(0)) & "%'")
            
            If erTranAll = False Then
            objHT.CloseTransactionBengkel
            End If
            MsgBox "Proses Pembatalan Washing berhasil dilakukan", vbInformation, "Informasi Pembatalan Washing"
            txtAlasanBtlWashing.text = ""
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
1        On Error GoTo Form_Activate_Error

2         If orderAktif Or subletAktif Or subRetAktif Then
3             MsgBox "Tutup Menu sebelumnya terlebih Dahulu !!", vbCritical
4             Unload Me
5         End If

6        On Error GoTo 0
7        Exit Sub

Form_Activate_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Activate of Form frmOrderValidasi")
End Sub

Private Sub Form_Load()
1        On Error GoTo Form_Load_Error

2         validAktif = True
3         text(1) = TglDDMMYYYY(isnow)
4         simpan = True
5         Frame6.Top = 0
6         If Log.dmsNonDss = 1 Then
7             btnBatalWashing.Visible = True
8         Else
9             btnBatalWashing.Visible = False
10        End If
          
          ''2017-08-01 || REVISI penambahan tombol batal request discount
          ''---------------------------------------------------------------------
11        If Log.discKacab = 1 Then
12            cmdBatalRequest.Visible = True
13            cmdBatalRequest.Enabled = False
14        Else
15            cmdBatalRequest.Visible = False
16        End If
          ''---------------------------------------------------------------------


17       On Error GoTo 0
18       Exit Sub

Form_Load_Error:

19        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Load of Form frmOrderValidasi")
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    validAktif = False
End Sub

Private Sub tblBatal_Click()
          'If gSpart And adaSublet Then
1        On Error GoTo tblBatal_Click_Error

2         Set rsTemp = objO.FetchData("SELECT DISTINCT NO_WO FROM VALIDBILLING WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND DELLOG = ''")
3         If rsTemp.EOF Then
          
4             Call objO.InsertData("UPDATE CHASSINGFORM SET TANGGALVALIDASI = NULL WHERE NOWO = '" & Trim(text(0)) & "'")
              
5             Call objO.UpdateData("UPDATE VALIDORDER SET USERDEL = '" & Trim(gSA) & "', USERDEL_TGL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND USERDEL = ''")
6             Call objO.DeleteData("DELETE FROM NEA_PS Where noWO = '" & Trim(text(0)) & "'")
7             chNea.value = 0
8             chPS.value = 0
9             txtInfoNea.text = ""
10            lblStatus.Caption = "BATAL VALIDASI"
11            tblBatal.Enabled = False
12            TblValid.Enabled = True
              
13            If Log.discKacab = 1 Then
                  ''Rev20170801A
14                isBolehBatalReqDisc = True
15                Set rsTemp = objO.FetchData("SELECT userValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(text(Index)) & "' ")
16                If Not rsTemp.EOF Then
                      ''2017-08-01 || ENABLE / DISABLE BUTTON BATAL
17                    If rsTemp!uservalid = "" Then
18                        cmdBatalRequest.Enabled = True
19                        MsgBox "Request belum divalidasi." & vbCrLf & "Proses validasi order dibatalkan.", vbInformation
                          'Exit Sub
20                    Else
21                        cmdBatalRequest.Enabled = False
22                    End If
23                End If
24            End If
              
25        Else
26            MsgBox "WO ini sudah diberi Diskon Khusus oleh HO, sehingga tidak bisa dibatal Validasi" & vbCr & "Hubungi HO untuk membatalkan diskon khusus. ", vbCritical + vbOKOnly, "Batal Diskon Khusus"
27        End If
          'End If

28       On Error GoTo 0
29       Exit Sub

tblBatal_Click_Error:

30        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblBatal_Click of Form frmOrderValidasi")
End Sub

Private Sub tblCetakWO_Click()
1        On Error GoTo tblCetakWO_Click_Error

2         If Trim(text(0)) <> "" Then
3             Call cetakForm
4         End If

5        On Error GoTo 0
6        Exit Sub

tblCetakWO_Click_Error:

7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblCetakWO_Click of Form frmOrderValidasi")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cetakForm
' DateTime  : 30/04/2015 13:58
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cetakForm()
        Dim rpt As Object
        Dim NamaSAOrder As String
        Dim rsRepo As New ADODB.Recordset
        Dim rsHlp As New ADODB.Recordset
        Dim bodyClear As Boolean
        
       On Error GoTo cetakForm_Error
    
        Select Case gPrintCetak
      Case "2"
          Set rpt = New RptOrder4            ' KGW + PEMUDA ' 2190
      Case "1"
          Set rpt = New RptOrder5            ' MLATI + ALL  ' 2180
      Case "3"
          Set rpt = New RptOrder6            ' SLAMET  ' LX 300+II
      Case "4"
          Set rpt = New RptOrder7            ' BLANKO UTK KERTAS KOSONG
          rpt.txtDealer.SetText support.kopOrder(0)
        End Select
      
        If MsgBox("Mencetak Order ?", vbYesNo, "Konfirmasi") = vbYes Then
      gNoOrder = text(0) & "/" & Trim(text(41)) 'CboStatus.text
      gKetSA = text(19)
      NamaSAOrder = objO.FetchData_String("Select FS_NM_PETUGAS from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS ='" & Trim(gKetSA) & "'")
      gCust = text(4)
      gAlmt = text(5)
      gKota = text(6)
      gTelpon = text(7)
      gTahun = text(10)
      gNoPolisi = text(16)
      gModel = text(14)
      gWarna = text(15)
      gTglMasuk = text(43)
      gMDLOrder = text(12)
      gFRM = text(17)
      gEG = text(18)
      
       
      Set rsHlp = objO.FetchData("SELECT CATAT, DATE, TIME, B.FS_NM_DEALER, B.DDATE, REMARK, REQ1, REQ2, REQ3, REQ4, REQ5, REP1, REP2, REP3, REP4, REP5, REP6, REP7 FROM JOBDTA A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS WHERE NO_WO = '" & text(0) & "' AND A.NO_CHASSIS = '" & text(17) & "'")
      If Not rsHlp.EOF Then
          gCatatan = Trim(rsHlp("CATAT"))
          gJamMasuk = Trim(rsHlp("TIME"))
          gDealer = Trim(rsHlp("FS_NM_DEALER"))
          gDD = Trim(rsHlp("DDATE"))
          gPKhusus = Trim(rsHlp("REMARK"))
          
          nWKT = hitungWaktu(CDbl(totRate), rsHlp("TIME"), rsHlp("DATE"))
          gPSelesai = Left(nWKT, InStr(nWKT, "|") - 1)
          
          gKeluhan(1) = Trim(rsHlp("REQ1"))
          gKeluhan(2) = Trim(rsHlp("REQ2"))
          gKeluhan(3) = Trim(rsHlp("REQ3"))
          gKeluhan(4) = Trim(rsHlp("REQ4"))
          gKeluhan(5) = Trim(rsHlp("REQ5"))
          gOrder(1) = Trim(rsHlp("REP1"))
          gOrder(2) = Trim(rsHlp("REP2"))
          gOrder(3) = Trim(rsHlp("REP3"))
          gOrder(4) = Trim(rsHlp("REP4"))
          gOrder(5) = Trim(rsHlp("REP5"))
          gOrder(6) = Trim(rsHlp("REP6"))
          gOrder(7) = Trim(rsHlp("REP7"))
      End If
       
      Set rsHlp = objO.FetchData("SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & text(0) & "' AND NO_CHASSIS = '" & text(17) & "'")
      If Not rsHlp.EOF Then
          i = 1
          Do While Not rsHlp.EOF
              If i < 10 Then
                  gSublet(i) = rsHlp("NAMABARANG")
                  i = i + 1
              End If
              rsHlp.MoveNext
          Loop
      End If
      If rsHlp.State <> adStateClosed Then rsHlp.Close
      
      gKM = Format(text(37), "#,#0")
      gSAdvisor = NamaSAOrder
      gKodePlg = text(2)
      
      Dim sHlp As Object  'New ADODB.Recordset
      Dim sHlp2 As Object
      Dim sHlp3 As Object
      Dim sHlp4 As Object
      Dim rRs As ADODB.Recordset
      Dim crxSubReportHistory As CRAXDRT.Report
      Dim crxSubReportBooking As CRAXDRT.Report
      Dim crxSubReportBooking2 As CRAXDRT.Report
      Dim crxSubReportBookingFU As CRAXDRT.Report
      If gUkuranOrder = False Then
          
          Set rRs = objO.FetchData("SELECT A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.NO_POLISI, A.NO_ENGINE, A.NO_CHASSIS, A.MDL, A.MODEL, A.FMDL, A.COLOR, " _
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
                                & "T_PETUGAS AS E WITH (NOLOCK) ON A.FRONT = E.FS_KD_PETUGAS WHERE A.NO_WO = '" & Trim(text(0)) & "'")
          rpt.database.SetDataSource rRs
                     
          gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
          
          sSQL = "SELECT * FROM historyUnitFunc('" & Trim(gFRM) & "')"
          Set sHlp = objO.FetchData(sSQL)
          Set crxSubReportHistory = rpt.Subreport2.OpenSubreport
          crxSubReportHistory.database.SetDataSource sHlp
          
          sSQL = "select DATE_BOOK, TIME_D, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
          Set sHlp2 = objO.FetchData(sSQL)
          Set crxSubReportBooking = rpt.Subreport1.OpenSubreport
          crxSubReportBooking.database.SetDataSource sHlp2
          
          sSQL = "select DATE, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
          Set sHlp4 = objO.FetchData(sSQL)
          Set crxSubReportBooking2 = rpt.Subreport4.OpenSubreport
          crxSubReportBooking2.database.SetDataSource sHlp4
          
          sSQL = "select WAKTU FROM M_FUBOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(31)) & "' "
          Set sHlp3 = objO.FetchData(sSQL)
          Set crxSubReportBookingFU = rpt.Subreport3.OpenSubreport
          crxSubReportBookingFU.database.SetDataSource sHlp3
          
          Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(text(0)) & "|%'")
          If Not rsHlp.EOF Then
              rpt.txtStall.SetText Trim(rsHlp("NOSTALL"))
          End If
          
          rpt.mechPlan.SetText ""
          
          Dim n As Integer
          bodyClear = True
          Set rsHlp = objO.FetchData("SELECT * FROM JOBKOORD WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "'")
          If Not rsHlp.EOF Then
              n = 1
              Do While Not rsHlp.EOF
                  Call pasangTanda(rpt, n, CDbl(rsHlp("KOORDY")), CDbl(rsHlp("KOORDX")))
                  n = n + 1
                  rsHlp.MoveNext
              Loop
              bodyClear = False
          End If
          rsHlp.Close
          
          If bodyClear Then
              rpt.Text93.TextColor = &HFF&
              rpt.Text94.TextColor = &HFF&
              rpt.Text95.TextColor = &HFF&
              rpt.Text96.TextColor = &HFF&
              rpt.Text97.TextColor = &HFF&
          End If
          
          rpt.Text2.TextColor = &HFF&
          rpt.Text3.TextColor = &HFF&
          rpt.Text4.TextColor = &HFF&
          rpt.Text5.TextColor = &HFF&
          
    
        ' ===================================================================================================================================
        ' 23 Feb 2017 | THANIA | tambahan info bill di cetak ulang WO
        ' ===================================================================================================================================
                                
            Set rsHlp = objO.FetchData("SELECT ISNULL(INFOBILL,'') INFO FROM T_PROFIL WITH(NOLOCK)")
            If Not rsHlp.EOF Then
               rpt.txtInfo.SetText Trim(rsHlp!info)
            End If
        ' ===================================================================================================================================
            
          Set rsHlp = objO.FetchData("SELECT BBM FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "'")
          If Not rsHlp.EOF Then
                Select Case rsHlp("BBM")
                    Case 0
                        ' FULL
                        rpt.Text90.TextColor = &H0&
                    Case 1
                        ' 3/4
                        rpt.Text62.TextColor = &H0&
                    Case 2
                        ' 1/2
                        rpt.Text64.TextColor = &H0&
                    Case 3
                        ' 1/4
                        rpt.Text88.TextColor = &H0&
                    Case 4
                        ' EMPTY
                        rpt.Text89.TextColor = &H0&
                End Select
            End If
            
            With FrmctkOrder.CRViewer1
                .ReportSource = rpt
                .ViewReport
                .Zoom 100
                .Visible = True
            End With
            FrmctkOrder.cmdTutup.Visible = True
            Screen.MousePointer = vbDefault
            FrmctkOrder.WindowState = vbMaximized
            FrmctkOrder.Show vbModal
        End If
    End If

   On Error GoTo 0
   Exit Sub

cetakForm_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakForm of Form frmOrderValidasi")
End Sub

Private Sub TblKeluar_Click()
1        On Error GoTo TblKeluar_Click_Error
2        Call hapusForm
         
3         Unload Me

4        On Error GoTo 0
5        Exit Sub

TblKeluar_Click_Error:

6         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TblKeluar_Click of Form frmOrderValidasi")
End Sub

Private Sub TblValid_Click()
Dim rsSearch As New ADODB.Recordset

    On Error GoTo vbError
    Set rsSearch = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND USERDEL = ''")
    If Not rsSearch.EOF Then
        MsgBox "Sudah di validasi sebelumnya !!", vbInformation, "Validasi Order"
        lblStatus.Caption = "DATA SUDAH DIVALIDASI"
        TblValid.Enabled = False
        tblBatal.Enabled = True
        
    Else
        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(43)), "DD/MM/YYYY") Then
            Dim rsSO As ADODB.Recordset
            Set rsSO = objO.FetchDataACC("select KodeItem  from SOORD WITH (NOLOCK) " & _
                     "Where (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0)+ IsNull(QtyKeluar,0))) <> 0 and NOSO = '" & Trim(text(0)) & "' and isnull(bo,'') <> 'CANCEL'")
            If rsSO.RecordCount > 0 Then
                MsgBox "Terdapat " & rsSO.RecordCount & " Line Item Part yang belum tersuplay sesuai order !" & Chr(13) & "Validasi tidak dapat diproses !", vbInformation, "Validasi Order"
                Exit Sub
            End If
        Else
            
        End If
        Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 0)
            erTranAll = False
            onTR = True
            Call objHT.OpenTransactionBengkel
            
        If gSpart And adaSublet And Trim(text(17)) <> "" Then
            Call objHT.InsertData_BengkelOnly("INSERT INTO VALIDORDER VALUES ('" & Trim(gSA) & "',getdate(),'" & IIf(rsJob.State = adStateClosed, "TANPA PEKERJAAN| ", "") + IIf(rsPart.State = adStateClosed, "TANPA PART| ", "") + IIf(rsSublet.State = adStateClosed, "TANPA SUBLET| ", "") & "','" & Trim(text(0)) & "','" & Trim(text(17)) & "','','')")
            Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTL SET KARAKTER = '" & Trim(cboPelanggan) & "', DTLKARAKTER = '" & Trim(text(44)) & "' WHERE NO_WO = '" & Trim(text(0)) & "'")
            Call objHT.UpdateData_BengkelOnly("UPDATE CHASSINGFORM SET TANGGALVALIDASI = GETDATE() WHERE NOWO = '" & Trim(text(0)) & "'")
            If chNea.value = 1 Or chPS.value = 1 Then Call objHT.InsertData_BengkelOnly("Insert Into NEA_PS values ('" & Trim(text(0)) & "'," & chNea.value & "," & chPS.value & ",'" & Trim(txtInfoNea.text) & "','" & Trim(gSA) & "',getdate(),'','01/01/1900','')")
            lblStatus.Caption = "SUDAH VALIDASI"
            TblValid.Enabled = False
            tblBatal.Enabled = True
        Else
            lblStatus.Caption = "DATA BELUM LENGKAP"
        End If
            If erTranAll = False Then
                Call objHT.CloseTransactionBengkel
                onTR = False
                Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 1)
                Exit Sub
            Else
                Call objHT.CloseTransactionBengkel_err
                onTR = False
                Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 2)
                MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
                Exit Sub
            End If
    End If
    Exit Sub

vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 2)
    MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TblValid_Click of Form frmOrderValidasi"), vbInformation, "Gagal Simpan"
    Exit Sub
 End Sub

Private Sub Text_Change(Index As Integer)
1        On Error GoTo Text_Change_Error

2         If Index = 17 Then
3             If Trim(text(17)) <> "" Then
4                 myChassis = Trim(text(17))
5                 Call CexChassis_Discon(Trim(text(13)))
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

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Text_Change of Form frmOrderValidasi")
End Sub

Private Sub Text_GotFocus(Index As Integer)
          Dim DiTab As Boolean
1        On Error GoTo text_GotFocus_Error

2         DiTab = False
3         If text(Index).BackColor = &HC0E0FF Then
4             DiTab = True
5         Else
6             If gMainKeyCari <> "" Then
7                 text(Index).Mask = ""
8                 text(Index) = gMainKeyCari
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
21            SendKeys ("{TAB}")
22        Else
23            SendKeys ("{HOME}+{END}")
24        End If

25       On Error GoTo 0
26       Exit Sub

text_GotFocus_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_GotFocus of Form frmOrderValidasi")
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
1        On Error GoTo Text_KeyDown_Error

2         Call CekTombol(KeyCode)
3         Select Case KeyCode
              Case vbKeyF12
4                 Select Case Index
                      Case 0
5                         gMainKeyCari = text(Index)
6                         gSearchOption = "INV2"
7                         gBantuan = "AND DELIV <> 'BTL'"
8                         frmCari.Show (vbModal)
9                     Case 42
10                        gMainKeyCari = text(Index)
11                        gSearchOption = "METERAI"
12                        frmCari.Show (vbModal)
13                End Select
14            Case vbKeyEscape
15                If simpan = False Then
16                    If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
17                        Call hapusForm
18                        text(0).SelLength = Len(text(0))
19                        text(0).SetFocus
20                    End If
21                End If
22        End Select

23       On Error GoTo 0
24       Exit Sub

Text_KeyDown_Error:

25        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Text_KeyDown of Form frmOrderValidasi")
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

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_KeyPress of Form frmOrderValidasi")
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsHlp As New ADODB.Recordset

   On Error GoTo text_Validate_Error

    Select Case Index
        Case 0
            If text(Index) <> "" Then
                boleh = True
                noinv = True
                If simpan = False Then
                    If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                        boleh = False
                    End If
                End If
                If boleh Then
                    n = Trim(text(0))
                    Call hapusForm
                    text(0) = n
                    If rsTemp.State Then rsTemp.Close
                    Set rsTemp = objO.FetchData("SELECT * FROM jobDta A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.no_WO = '" & text(Index) & "' ")
                    If Not rsTemp.EOF Then
                        TblValid.Enabled = False
                        newValid = False
                        lblStatus.Visible = True
                        lblStatus.Caption = ""
                        Set rsHlp = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND USERDEL = ''")
                        If Not rsHlp.EOF Then
                            lblStatus.Caption = "SUDAH DIVALIDASI"
                            TblValid.Enabled = False
                            tblBatal.Enabled = True
                            adaSublet = True
                            
                            Dim rsNPS As ADODB.Recordset
                            Set rsNPS = objO.FetchData("Select * from NEA_PS WITH (NOLOCK) Where NOWO = '" & text(0).text & "'")
                            If rsNPS.RecordCount > 0 Then
                                chNea.value = rsNPS.Fields(1).value
                                chPS.value = rsNPS.Fields(2).value
                                txtInfoNea.text = Trim(rsNPS.Fields(3).value)
                            End If
                            
                            If rsTemp("NOINV") <> "" And rsTemp("RPR_COND") <> "K" Then
                                MsgBox ("Telah dibuat invoice, Tanggal " & TglDDMMMMYYYY(Trim(rsTemp("INVDATE"))) & " No. Invoice = " & Trim(rsTemp("NOINV")))
                                lblStatus.Caption = "Sudah Invoice"
                                tblBatal.Enabled = False
                                noinv = False
                                'Exit Sub
                            ElseIf rsTemp("NOINV") <> "" And rsTemp("RPR_COND") = "K" Then
                                MsgBox ("Telah diclosing !")
                                lblStatus.Caption = "Closing"
                                tblBatal.Enabled = False
                                noinv = False
                                'Exit Sub
                            End If
                        Else
                            tblBatal.Enabled = False
                            If (rsTemp("DELIV") = "PRC" Or rsTemp("DELIV") = "OKE") And rsTemp("RPR_COND") = "S" And rsTemp("NOINV") = "" Then
                                newValid = True
                                text(21).SelLength = Len(text(21))
                                TblValid.Enabled = True
                                TblValid.SetFocus
                                MsgBox ("Unit telah selesai dikerjakan !")
                                lblStatus.Caption = "Selesai dikerjakan"
                                GoTo prosesStall
                            ElseIf rsTemp("DELIV") = "BTL" Then
                                lblStatus.Caption = "WO dibatalkan"
                            ElseIf rsTemp("RPR_COND") = "K" Then
                                lblStatus.Caption = "WO Sudah Closing"
                                noinv = False
                            Else
prosesStall:
                                Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(rsTemp("NO_POLISI")) & " |" & Trim(rsTemp("NO_WO")) & "%'")
                                If Not rsHlp.EOF Then
                                    Select Case Trim(rsHlp("FLAG"))
                                        Case ""
                                            MsgBox "Unit belum dibuat WO masih booking waiting !", vbInformation, "Status Order"
                                            lblStatus.Caption = "Booking"
                                        Case "N"
                                            MsgBox "Unit belum dikerjakan !", vbInformation, "Status Order"
                                            lblStatus.Caption = "Menunggu Foreman di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "S"
                                            MsgBox "Unit sedang dikerjakan !", vbInformation, "Status Order"
                                            lblStatus.Caption = "Proses di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "P"
                                            MsgBox "Unit sedang dihentikan sementara !", vbInformation, "Status Order"
                                            lblStatus.Caption = "Pause di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "H"
                                            MsgBox "Unit menunggu proses Final Inspection", vbInformation, "Status Order"
                                            lblStatus.Caption = "Tunggu Inspection di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "Q"
                                            MsgBox "Unit dalam proses Final Inspection", vbInformation, "Status Order"
                                            lblStatus.Caption = "Proses Inspection di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "G"
                                            MsgBox "Unit menunggu proses Washing", vbInformation, "Status Order"
                                            lblStatus.Caption = "Tunggu Washing di " & getStallWashing(Trim(text(0)), True, True) '& ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "D"
                                            MsgBox "Unit menunggu proses Drying", vbInformation, "Status Order"
                                            lblStatus.Caption = "Tunggu Drying di " & getStallWashing(Trim(text(0)), True, False) ' & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "A"
                                            MsgBox "Unit sedang proses Washing", vbInformation, "Status Order"
                                            lblStatus.Caption = "Washing"
                                        Case "R"
                                            MsgBox "Unit sedang proses Drying", vbInformation, "Status Order"
                                            lblStatus.Caption = "Drying"
                                        Case "F"
                                            MsgBox ("Unit telah selesai dikerjakan !")
                                            lblStatus.Caption = "Selesai dikerjakan"
                                    End Select
                                End If
                            End If
                        End If
                        ''------------------------------------------------------------
                        ''2017-06-07 CEK JIKA ADA REQUEST DISC HARUS DIVALIDASI KACAB
                        ''------------------------------------------------------------
                        
                        If Log.discKacab = 1 Then
                            ''Rev20170801A
                            isBolehBatalReqDisc = True
                            Set rsTemp = objO.FetchData("SELECT userValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(text(Index)) & "' ")
                            If Not rsTemp.EOF Then
                                ''2017-08-01 || ENABLE / DISABLE BUTTON BATAL
                                If rsTemp!uservalid = "" Then
                                    cmdBatalRequest.Enabled = True
                                    MsgBox "Request belum divalidasi." & vbCrLf & "Proses validasi order dibatalkan.", vbInformation
                                    'Exit Sub
                                Else
                                    cmdBatalRequest.Enabled = False
                                End If
                            End If
                        End If
                        
                        Call panggilWO(text(Index))
                    End If
                    If rsTemp.State <> adStateClosed Then rsTemp.Close
                    simpan = False
                End If
            End If
    End Select

   On Error GoTo 0
   Exit Sub

text_Validate_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmOrderValidasi")
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

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getStallWashing of Form frmOrderValidasi")
End Function
 
Private Sub txtKeluhan_GotFocus()
    Dim DiTab As Boolean
    DiTab = False
    If txtKeluhan.BackColor = &HC0E0FF Then
        DiTab = True
    End If

    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub txtPermintaan_GotFocus()
    Dim DiTab As Boolean
    DiTab = False
    If txtPermintaan.BackColor = &HC0E0FF Then
        DiTab = True
    End If

    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
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

219       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure pasangTanda of Form frmOrderValidasi")
End Sub



