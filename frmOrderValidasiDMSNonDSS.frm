VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOrderValidasiDMSNonDSS 
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
   Begin VB.Frame framePembatalanWashing 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pembatalan Washing"
      Height          =   2295
      Left            =   5400
      TabIndex        =   131
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtAlasanBtlWashing 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   133
         Top             =   240
         Width           =   3375
      End
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
         TabIndex        =   132
         Top             =   1800
         Width           =   1725
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
      Height          =   5895
      Left            =   1440
      TabIndex        =   105
      Top             =   2640
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox txtFI 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   127
         Top             =   680
         Width           =   5655
      End
      Begin VB.TextBox txtNextJob 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   126
         Top             =   3230
         Width           =   5655
      End
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
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   5400
         Width           =   1245
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   249
         TabIndex        =   118
         Top             =   5520
         Width           =   9210
      End
      Begin VB.TextBox txtKeluhan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   107
         Top             =   3230
         Width           =   5655
      End
      Begin VB.TextBox txtPermintaan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   106
         Top             =   680
         Width           =   5655
      End
      Begin VB.Label Label36 
         BackColor       =   &H80000009&
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
         TabIndex        =   129
         Top             =   2900
         Width           =   3615
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000009&
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
         TabIndex        =   128
         Top             =   375
         Width           =   5655
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000009&
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
         TabIndex        =   125
         Top             =   2900
         Width           =   3615
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000009&
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
         TabIndex        =   124
         Top             =   375
         Width           =   3615
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   11640
         Y1              =   2880
         Y2              =   2880
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
         Top             =   5520
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
         Left            =   4680
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
         BackColor       =   &H00DCFCFB&
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   1680
         Width           =   1890
      End
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1680
         Width           =   2730
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
         Left            =   0
         TabIndex        =   67
         Top             =   -240
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
         Left            =   8880
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
         TabIndex        =   59
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
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   615
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
      TabIndex        =   123
      Top             =   120
      Width           =   5490
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
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
Attribute VB_Name = "frmOrderValidasiDMSNonDSS"
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

Private Sub doCreatePart()
    With rsPart
        If .State Then .Close
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
    With rsJob
        If .State Then .Close
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
    discountText = discountValue
    discountText.Enabled = statusText
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
Dim rsTmp As ADODB.Recordset
    
    gKdPlgDealer = ""
    gKodePlg = ""
    gInvCust = ""
    gInvAlmt = ""
    gInvKota = ""
    gInvNPWP = ""
    nDuoInv = False
    btnBatalWashing.Enabled = False
    Set rsTemp = objBilling.searchOrderDetail(noWO)
    If Not rsTemp.EOF Then
        
        ' TIPE PEMBAYARAN
        Set rsHlp = objBilling.searchPaymentType(noWO)
        If Not rsHlp.EOF Then
            tipeByr = Trim(rsHlp("PEMBAYARAN"))
            gVerifikator = Trim(rsHlp("VERIFIKATOR"))
            cboPelanggan = Trim(rsHlp("KARAKTER"))
            text(44) = Trim(rsHlp("DTLKARAKTER"))
        End If
        If rsHlp.State <> adStateClosed Then rsHlp.Close
        
        Set rsHlp = objBilling.searchDtlOrder(rsTemp("NO_WO"), rsTemp("NO_CHASSIS"))
        If Not rsHlp.EOF Then
            nHeavy = IIf(rsHlp("HEAVYUSE") <> 0, True, False)
        End If
        If rsHlp.State <> adStateClosed Then rsHlp.Close
        
        ' TAMPIL DETAIL CUSTOMER
        text(2) = rsTemp("FS_KD_PLG")
        text(3) = rsTemp("FS_KD_WIL")
        text(4) = rsTemp("CUSTOMER")
        text(5) = Left(rsTemp("CUST_ADRES"), 60)
        text(6) = rsTemp("KOTA")
        text(7) = rsTemp("TELPON")
        text(8) = rsTemp("GSM")
        text(9) = rsTemp("CDMA")
        text(10) = rsTemp("FS_TH_MOBIL")
        text(11).Mask = ""
        text(11) = Nul(IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsTemp("DOTAM"), "/", "-"))))
        text(12) = rsTemp("MDL")
        text(13) = rsTemp("MODEL")
        text(14) = rsTemp("FMDL")
        text(15) = rsTemp("COLOR")
        text(16) = rsTemp("NO_POLISI")
        text(17) = rsTemp("NO_CHASSIS")
        text(18) = rsTemp("NO_ENGINE")
        text(19) = rsTemp("FS_LOG_ADD")
        If rsTemp("NOINV") <> "" Then
            text(20) = rsTemp("NOINV")
        End If
        text(31) = rsTemp("NOFAK")
        'text(33) = rsTemp("FS_RC")
        text(37) = rsTemp("KM")
        text(38) = rsTemp("MECH")
        text(40) = rsTemp("STATUS")
        text(43) = TglDDMMYYYY(rsTemp("DATE"))
        nTHS = IIf(rsTemp("THS") = 1, True, False)
        LblReturn.Visible = IIf(rsTemp("FN_RETURN") = "1", True, False)
        lblPDS.Visible = IIf(rsTemp("PDS") = "1", True, False)
        text(28) = rsTemp("ANCUST")
        text(32) = rsTemp("TELPANCUST")
        
        ' CHECK DEFAULT DISCOUNT
        text(30) = ""
        text(35).Visible = False
        text(36).Visible = False
        vTipeCustomer = ""
        myChassis = Trim(text(17))
        If returnJob(LblReturn) Then
            If typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
                opWO(1).value = True
                 
                Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
                If Not rsHlp.EOF Then
                    text(35) = rsHlp("ID_PROGRAM")
                    text(36) = rsHlp("PROGRAM")
                    text(35).Visible = True
                    text(36).Visible = True
                End If
                If rsHlp.State <> adStateClosed Then rsHlp.Close
                If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
                    nDuoInv = True
                    'RTJ WO BERKEMBANG
                     
                    Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
                    If Not rsHlp.EOF Then
                        gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
                        gKodePlg2 = Trim(rsHlp("KDPLG"))
                        gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
                        gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
                        gInvKota2 = Trim(rsHlp("FS_KOTA"))
                        gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                        
                        If Trim(rsHlp("INTERNAL")) = 1 Then
                            If Trim(gNpwp) = Trim(gInvNPWP2) Then
                                vTipeCustomer = "AFFNONPPN"
                            Else
                                vTipeCustomer = "AFFPPN"
                            End If
                        Else
                            vTipeCustomer = "EXT"
                        End If
                    End If
                    If rsHlp.State <> adStateClosed Then rsHlp.Close
                    Call giveDiscount("100", text(21), False)
                    text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
                End If
            Else
                ' HANYA RTJ
                Call giveDiscount("100", text(21), False)
                vTipeCustomer = "INT"
            End If
        Else
            If typeJob(Trim(rsTemp("TPSVC"))) = "WARRANTY" Then
                opWO(2).value = True
                 
                Set rsHlp = objBilling.searchWarrantyDetail(rsTemp("KODETIPE"))
                If Not rsHlp.EOF Then
                    text(35) = rsHlp("ID_WARRANTY")
                    text(36) = rsHlp("WARRANTY")
                    text(35).Visible = True
                    text(36).Visible = True
                End If
                If rsHlp.State <> adStateClosed Then rsHlp.Close
                Call giveDiscount("0", text(21), True)
                text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
                Set rsHlp = objO.FetchData("SELECT * FROM T_PROFIL WITH (NOLOCK)")
                If Not rsHlp.EOF Then
                    gKdPlgDealer = Trim(rsHlp("FN_HOKDPLG"))
                    gKodePlg = Trim(rsHlp("FN_HOKDPLG"))
                    gInvCust = Trim(rsHlp("FN_HO"))
                    gInvAlmt = Trim(rsHlp("FN_HOADRES"))
                    gInvKota = Trim(rsHlp("FN_HOKOTA"))
                    gInvNPWP = Trim(rsHlp("FN_HONPWP"))
                    vTipeCustomer = "AFFPPN"
                End If
                If rsHlp.State <> adStateClosed Then rsHlp.Close
            ElseIf typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
                opWO(1).value = True
                 
                Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
                If Not rsHlp.EOF Then
                    text(35) = rsHlp("ID_PROGRAM")
                    text(36) = rsHlp("PROGRAM")
                    text(35).Visible = True
                    text(36).Visible = True
                End If
                If rsHlp.State <> adStateClosed Then rsHlp.Close
                If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
                    nDuoInv = True
                    'SBI WO BERKEMBANG
                    If rsTemp("FS_RC") = "R001" Then
                         
                        Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                        If Not rsHlp.EOF Then
                            gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
                            gKodePlg2 = Trim(rsHlp("KDPLG"))
                            gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
                            gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
                            gInvKota2 = Trim(rsHlp("FS_KOTA"))
                            gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                            
                            If Trim(rsHlp("INTERNAL")) = 1 Then
                                If Trim(gNpwp) = Trim(gInvNPWP2) Then
                                    vTipeCustomer = "AFFNONPPN"
                                Else
                                    vTipeCustomer = "AFFPPN"
                                End If
                            Else
                                vTipeCustomer = "EXT"
                            End If
                        End If
                        If rsHlp.State <> adStateClosed Then rsHlp.Close
                        Call giveDiscount("100", text(21), False)
                        text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
                    Else
                         
                        Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                        If Not rsHlp.EOF Then
                            gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
                            gKodePlg2 = Trim(rsHlp("KDPLG"))
                            gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
                            gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
                            gInvKota2 = Trim(rsHlp("FS_KOTA"))
                            gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                            
                            If Trim(rsHlp("INTERNAL")) = 1 Then
                                If Trim(gNpwp) = Trim(gInvNPWP2) Then
                                    vTipeCustomer = "AFFNONPPN"
                                Else
                                    vTipeCustomer = "AFFPPN"
                                End If
                            Else
                                vTipeCustomer = "EXT"
                            End If
                        End If
                        If rsHlp.State <> adStateClosed Then rsHlp.Close
                        Call giveDiscount("100", text(21), False)
                        text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
                    End If
                Else
                    Call giveDiscount(objBilling.takeDiscountProgram(rsTemp("KODETIPE")), text(21), False)
                End If
            ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
                If lblPDS.Visible Then
                    Call giveDiscount("0", text(21), False)
                    Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
                    If Not rsHlp.EOF Then
                        gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
                        gKodePlg2 = Trim(rsHlp("KDPLG"))
                        gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
                        gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
                        gInvKota2 = Trim(rsHlp("FS_KOTA"))
                        gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                        
                        If Trim(rsHlp("INTERNAL")) = 1 Then
                            If Trim(gNpwp) = Trim(gInvNPWP2) Then
                                vTipeCustomer = "AFFNONPPN"
                            Else
                                vTipeCustomer = "AFFPPN"
                            End If
                        Else
                            vTipeCustomer = "EXT"
                        End If
                    End If
                    If rsHlp.State <> adStateClosed Then rsHlp.Close
                ElseIf rsTemp("FS_RC") = "R001" Then
                    ' HANYA SBI
                    woStat = "SBI"
                     
                    Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                    If Not rsHlp.EOF Then
                        gKdPlgDealer = Trim(rsHlp("KDPLG"))
                        gKodePlg = Trim(rsHlp("KDPLG"))
                        gInvCust = Trim(rsHlp("FS_NM_DEALER"))
                        gInvAlmt = Trim(rsHlp("FS_ALAMAT"))
                        gInvKota = Trim(rsHlp("FS_KOTA"))
                        gInvNPWP = Trim(rsHlp("FS_NPWP"))
                        If Trim(rsHlp("INTERNAL")) = 1 Then
                            If Trim(gNpwp) = Trim(gInvNPWP) Then
                                text(40) = "AFFNONPPN"
                            Else
                                text(40) = "AFFPPN"
                            End If
                        Else
                            text(40) = "EXT"
                        End If
                    End If
                    If rsHlp.State <> adStateClosed Then rsHlp.Close
                    Call giveDiscount("0", text(21), False)
                    text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
                ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
                     
                    Set rsHlp = objBilling.searchMOUDetail(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS"))
                    If Not rsHlp.EOF Then
                        text(34) = Trim(rsHlp("NO_MOU"))
                    End If
                    If rsHlp.State <> adStateClosed Then rsHlp.Close
                    Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")), text(21), False)
                Else
                    Call giveDiscount("0", text(21), True)
                End If
            End If
        End If
        
        ' TAMPIL JOB ORDER
        Dim nilaiRate As Double
        totRate = 0
        text(22) = 0
        subJob = 0
        If rsJob.State Then rsJob.Close
        myChassis = Trim(text(17))
        For i = 1 To 7
            hlp = "REP" & i
            hlp2 = "RTE" & i
            hlp3 = "KLPKERJA" & i
            hlp4 = "ACTRATE" & i
            hlp5 = "MEKANIK" & i
            
            If rsTemp(hlp) <> "" Then
                fOth = False
                Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "' AND MDL = '" & rsTemp("MDL") & "'")
                If rsHlp.EOF Then
                    fOth = True
                End If
                If rsHlp.State <> adStateClosed Then rsHlp.Close
                
                If returnJob(LblReturn) Then
                    If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
                        If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        End If
                    Else
                        nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                    End If
                Else
                    If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
                        nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                    ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
                        If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
                            If rsTemp("FS_RC") = "R001" Then
                                If Trim(rsTemp(hlp3)) = "SBI" Then
                                     
                                    Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                                    If Not rsHlp2.EOF Then
                                        nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
                                        If nStatusDealer = 1 Then
                                            nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
                                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                        Else
                                            If nStatusDealer = 2 Then
                                                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
                                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                            ElseIf nStatusDealer = 3 Then
                                                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
                                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                            End If
                                        End If
                                    End If
                                    If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                                Else
                                    nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                End If
                            Else
                                 
                                Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                                If Not rsHlp2.EOF Then
                                    'Dim nStatusDealer As Integer
                                    nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
                                    If nStatusDealer = 1 Then
                                        nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0  '
                                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                    Else
                                        If nStatusDealer = 2 Then
                                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                        ElseIf nStatusDealer = 3 Then
                                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
                                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                        End If
                                    End If
                                    'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
                                End If
                                If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                            End If
                        Else
                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        End If
                    ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
                        If lblPDS.Visible Then
                            nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        ElseIf rsTemp("FS_RC") = "R001" Then
                             
                            Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                            If Not rsHlp2.EOF Then
                                nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
                                If nStatusDealer = 1 Then
                                    nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                Else
                                    If nStatusDealer = 2 Then
                                        nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
                                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                    ElseIf nStatusDealer = 3 Then
                                        nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
                                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                    End If
                                End If
                                'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
                            End If
                            If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                        ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        Else
                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        End If
                    End If
                End If
                nilaiRate = CDbl(rsTemp(hlp2))
                text(22) = text(22) + (CDbl(nHarga) * nilaiRate) * ((100 - text(21)) / 100)
                
                With rsJob
                If .State = adStateClosed Then doCreateJob
                   .AddNew
                   !Job = Trim(rsTemp(hlp))
                   !Rate = nilaiRate 'CDbl(rsTemp(hlp2))
                   !ACTRT = CDbl(rsTemp(hlp4))
                   !grp = Trim(rsTemp(hlp3))
                   !MECH = Trim(rsTemp(hlp5))
                   !hrg = CDbl(nHarga) * nilaiRate
                   !addJob = False
                   !disc = CDbl(text(21))
                   !DISCRP = CDbl(!hrg) * (CDbl(text(21)) / 100)
                   !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
                   !valid = gVerifikator
                   !FREEC = False
                   .Update
                End With
                subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
                totRate = totRate + CDbl(nilaiRate)
            End If
        Next
        
        ' add JOB
        Set rsHlp2 = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & noWO & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
        If Not rsHlp2.EOF Then
            Do While Not rsHlp2.EOF
'            For i = 1 To 10
'                hlp = "REP" & i
'                hlp2 = "RTE" & i
'                hlp3 = "KLPKERJA" & i
'                hlp4 = "ACTRATE" & i
'                hlp5 = "MEKANIK" & i
                
                If rsHlp2("REP") <> "" Then
                    If returnJob(LblReturn) Then
                        If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
                            If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            End If
                        Else
                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        End If
                    Else
                        If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
                        ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
                            If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
                                If rsTemp("FS_RC") = "R001" Then
                                    If rsHlp2("KLPKERJA") = "SBI" Then
                                         
                                        Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                                        If Not rsHlp3.EOF Then
                                            nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
                                            If nStatusDealer = 1 Then
                                                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
                                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                            Else
                                                If nStatusDealer = 2 Then
                                                    nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
                                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                                ElseIf nStatusDealer = 3 Then
                                                    nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
                                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                                End If
                                            End If
                                            'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
                                        End If
                                        If rsHlp3.State <> adStateClosed Then rsHlp3.Close
                                    Else
                                        nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                    End If
                                Else
                                     
                                    Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                                    If Not rsHlp3.EOF Then
                                        nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
                                        If nStatusDealer = 1 Then
                                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0 '
                                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                        Else
                                            If nStatusDealer = 2 Then
                                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                            ElseIf nStatusDealer = 3 Then
                                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
                                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                            End If
                                        End If
                                        'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
                                    End If
                                    If rsHlp3.State <> adStateClosed Then rsHlp3.Close
                                End If
                            Else
                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            End If
                        ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
                            If lblPDS.Visible Then
                                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            ElseIf rsTemp("FS_RC") = "R001" Then
                                 
                                Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
                                If Not rsHlp3.EOF Then
                                    nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
                                    If nStatusDealer = 1 Then
                                        nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
                                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                    Else
                                        If nStatusDealer = 2 Then
                                            nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
                                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                        ElseIf nStatusDealer = 3 Then
                                            nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
                                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                                        End If
                                    End If
                                    'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
                                End If
                                If rsHlp3.State <> adStateClosed Then rsHlp3.Close
                            ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            Else
                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            End If
                        End If
                    End If
                    
                    ' ADJUSTMENT RATE
                    nilaiRate = CDbl(rsHlp2("RTE"))
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
                    
                    text(22) = text(22) + (CDbl(nHarga) * CDbl(nilaiRate)) * ((100 - text(21)) / 100)
                        
                    With rsJob
                    If .State = adStateClosed Then doCreateJob
                        .AddNew
                        !Job = Trim(rsHlp2("REP"))
                        !Rate = nilaiRate 'CDbl(rsHlp2("RTE"))
                        !ACTRT = CDbl(rsHlp2("ACTRTE"))
                        !grp = Trim(rsHlp2("KLPKERJA"))
                        !MECH = Trim(rsHlp2("MEKANIK"))
                        !hrg = CDbl(nHarga) * CDbl(nilaiRate)
                        !addJob = True
                        !disc = CDbl(text(21))
                        !DISCRP = CDbl(!hrg) * (CDbl(text(21)) / 100)
                        !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
                        !valid = gVerifikator
                        !FREEC = False
                        .Update
                        adaAddJob = True
                    End With
                    subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
                    totRate = totRate + CDbl(nilaiRate)
                End If
'            Next
                rsHlp2.MoveNext
            Loop
        End If
        
'        subJob = CDbl(Text(22))
        Set JobGrid.DataSource = rsJob
        Call doHeadJob
        
        ' TAMPIL PART ORDER
        noKFaktur = ""
        gSpart = True
        text(24) = 0
        text(25) = 0
        subPartGenuine = 0
        subPartNGenuine = 0
        subOliGenuine = 0
        subOliNGenuine = 0
        Dim HargaDisc As Double
        If rsPart.State Then rsPart.Close
        
        ' /// AMBIL PART YG BAWA SENDIRI //
        ' AND B.TIPE = 'O'
        Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, TIPE, GENUINE, BAWA FROM jobpart a WITH (NOLOCK) LEFT JOIN T_oli b WITH (NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli WHERE no_WO = '" & noWO & "' " _
                                & "AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y' ") 'AND BAWA <> 'Y'
        If Not rsHlp.EOF Then
            rsHlp.MoveFirst
            Do While Not rsHlp.EOF
                With rsPart
                If .State = adStateClosed Then doCreatePart
                    gSpart = True
                    
                    If rsHlp("BAWA") = "Y" Then
                        .AddNew
                        !KodeItem = "BAWA SENDIRI"
                        !NAMAITEM = rsHlp("PARTBAWA")
                        !qty = rsHlp("QTY")
                        !QtyOrder = rsHlp("QTY")
                        !QtyBO = 0
                        !QtyCancel = 0
                        !Harga = 0
                        !oli = IIf(rsHlp("TIPE") <> "P", True, False)
                        .Update
                    Else
                        If Not gFPart Then
                            '///// TIDAK CONNECT ACCOUNTING SPAREPART ////
                            .AddNew
                            !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
                            !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
                            !oli = IIf(rsHlp("TIPE") <> "P", True, False)
                            !qty = rsHlp("QTY")
                                'Set rsHlp2 = objO.FetchDataACC("SELECT hargajualA FROM inmst WHERE kodeItem = '" & rsHlp("KODE_ITEM") & "'")
                                Set rsHlp2 = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & rsHlp("KODE_ITEM") & "'")
                                If Not rsHlp2.EOF Then
                                    HargaDisc = CDbl(rsHlp2("HARGAJUALA"))
                                    !Harga = CDbl(HargaDisc)
                                    If rsHlp("OLI") <> "" Then
                                        If rsHlp("TIPE") = "P" Then
                                            If rsHlp2("TGenuine") = False Then
                                                ' PART GENUINE
                                                subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                            ElseIf rsHlp2("TGenuine") = True Then
                                                ' PART NON GENUINE
                                                subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                            End If
                                            text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                        ElseIf rsHlp("TIPE") = "O" Then
                                            If rsHlp2("TGenuine") = False Then
                                                ' OLI GENUINE
                                                subOliNGenuine = subOliNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                            ElseIf rsHlp2("TGenuine") = True Then
                                                ' OLI NON GENUINE
                                                subOliGenuine = subOliGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                            End If
                                            text(25) = text(25) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                        End If
                                    Else
                                        ' PART GENUINE
                                        subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                        text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                    End If
                                End If
                                If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                            
                            .Update
                        End If
                    End If
                End With
                rsHlp.MoveNext
            Loop
        End If
        If rsHlp.State <> adStateClosed Then rsHlp.Close
        
        If gFPart Then
            '/// KONEKSI KE SPAREPART
            If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
                With rsPart
                    If noinv Then
                        Dim strSQL As String
                        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
                            strSQL = "select (case When KodeItem = '' Then namaBYSA Else KodeItem end) as KodeItem, " & _
                                 "Keterangan as Namabarang, SatuanNs as Satuan, Qty as QtyOrder, (isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0)) AS Jml, " & _
                                 "(Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) as QtyBO, HargaJual as HargaJualA, NOSO as NoRack, " & _
                                 "0 as DisItem, 0 as ExtDis, Case BO When 'CANCEL' THEN (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) ELSE 0 END AS QTYCANCEL from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(text(0)) & "'"
                        Else
                            strSQL = "Select Z.*, 0 as QtyOrder, 0 as QtyBO FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
                                "INMST.NORACK,0 as DisItem, 0 as ExtDis, 0 as QtyCANCEL FROM INMST " & _
                                "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                "Where Z.Norack = '" & Trim(text(0)) & "'"
                        End If
                        Set rsHlp3 = objO.FetchDataACC(strSQL)

                        If Not rsHlp3.EOF Then
                            Do While Not rsHlp3.EOF
                                If .State = adStateClosed Then doCreatePart
                                .AddNew
                                !KodeItem = Trim(rsHlp3("KODEITEM"))
                                !NAMAITEM = Trim(rsHlp3("NAMABARANG"))
                                !QtyOrder = CDbl(rsHlp3("QtyOrder"))
                                !qty = CDbl(rsHlp3("JML"))
                                 If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
                                    !QtyBO = IIf(UCase(isBO) = "CANCEL", 0, objO.FetchData_DoubleACC("select Qty FROM BOFILE WITH (NOLOCK) Where TutupOrderFlag = 0 and (NoSO = '" & Trim(text(0)) & "' Or NOSO =(Select Distinct NoSPK From SOORD WITH (NOLOCK) Where NOSO ='" & Trim(text(0)) & "' and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "')) and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "'"))
                                 Else
                                    !QtyBO = 0
                                 End If
                                
                                !QtyCancel = CDbl(rsHlp3("QtyCancel"))
                                !Harga = CDbl(rsHlp3("HARGAJUALA"))
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
                                    Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp3("KODEITEM") & "'")
                                    If Not rsHlp.EOF Then
                                        If Left(rsHlp("SCC"), 3) = "OLI" Then
                                            !oli = True
                                            text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
                                        ElseIf Left(rsHlp("SCC"), 3) = "PAR" Then
                                            !oli = False
                                            text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
                                            !oli = False
                                            text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
                                        End If
                                    Else
                                        !oli = False
                                        text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
                                    End If
                                .Update
                                rsHlp3.MoveNext
                            Loop
                        Else 'If rsHlp3.EOF And newValid Then
                            Dim rsInvPart As New ADODB.Recordset
                            Dim vHitung As Double
                            Set rsInvPart = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                             & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' ")
                            If Not rsInvPart.EOF Then
                                Do While Not rsInvPart.EOF
                                    If .State = adStateClosed Then doCreatePart
                                    .AddNew
                                    !KodeItem = Trim(rsInvPart("KODEITEM"))
                                    !NAMAITEM = Trim(rsInvPart("KETERANGAN"))
                                    vHitung = (Koma2Titik(rsInvPart("Qty")) * (Koma2Titik(rsInvPart("HargaJual")) - (Koma2Titik(rsInvPart("ItemDisc")) / Abs(Koma2Titik(rsInvPart("Qty"))))))
                                    vHitung = vHitung - (Koma2Titik(rsInvPart("ExtDisPersen")) * vHitung)
                                    If Trim(UCase(rsInvPart("TypePpn"))) = "E" Then
                                        vHitung = Round(vHitung + vHitung * 0.1, 2)
                                    End If
                                    If gJurnalPart Then
                                        vHitung = cekIncludePPn(vHitung)
                                    End If
                                    !qty = Val(rsInvPart("QTY"))
                                    If !qty = 0 Then
                                        !Harga = 0
                                    Else
                                        !Harga = CDbl(vHitung / !qty)
                                    End If
                                    !QtyOrder = Val(rsInvPart("QTY"))
                                    !QtyBO = 0
                                    !QtyCancel = 0
                                     
                                    'Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsInvPart("KODEITEM"), 1) & "'")
                                    Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsInvPart("KODEITEM") & "'")
                                    If Not rsHlp.EOF Then
                                        If Left(rsHlp("SCC"), 3) = "OLI" Then
                                            !oli = True
                                            text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
                                        ElseIf Left(rsHlp("SCC"), 3) = "PAR" Then
                                            !oli = False
                                            text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
                                            !oli = False
                                            text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
                                        End If
                                    Else
                                        !oli = False
                                        text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
                                    End If
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
                                    .Update
                                    rsInvPart.MoveNext
                                Loop
                            Else
                                gSpart = False
'                                If TblValid.Enabled Then
'                                    temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
'                                    If temp = vbYes Then
'                                        gSpart = False
'                                    End If
'                                End If
                            End If
                            If rsInvPart.State <> adStateClosed Then rsInvPart.Close
                        End If
                        If rsHlp3.State <> adStateClosed Then rsHlp3.Close
                    Else 'Not newValid Then
                         
                        Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND NOINV LIKE '" & Trim(text(20)) & "%'")
                        If Not rsHlp2.EOF Then
                            Do While Not rsHlp2.EOF
                                If .State = adStateClosed Then doCreatePart
                                .AddNew
                                !KodeItem = Trim(rsHlp2("KODEITEM"))
                                !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
                                !qty = CDbl(rsHlp2("QTY"))
                                !Harga = CDbl(rsHlp2("HARGA")) '* ((100 - CDbl(rsHlp2("FN_EXT_DISKON"))) / 100)
                                !oli = Trim(rsHlp2("OLI")) 'IIf(Trim(rsHlp2("OLI")) = 1, True, False)
                                If !oli Then
                                    text(25) = text(25) + CDbl(!Harga)
                                Else
                                    text(24) = text(24) + CDbl(!Harga)
                                End If
                                !EXTDISC = CDbl(rsHlp2("FN_EXT_DISKON"))
                                !QtyOrder = CDbl(rsHlp2("QTY"))
                                !QtyBO = 0
                                !QtyCancel = 0
                                .Update
                                rsHlp2.MoveNext
                            Loop
                        Else
                            gSpart = False
                        End If
                        If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                    End If
                End With
            End If
        
            If Not gSpart Then
                    With rsPart
                        Set rsHlp3 = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                       & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' ")
                        
                        If Not rsHlp3.EOF Then
                            Do While Not rsHlp3.EOF
                                If .State = adStateClosed Then doCreatePart
                                .AddNew
                                !KodeItem = Trim(rsHlp3("KODEITEM"))
                                !NAMAITEM = Trim(rsHlp3("KETERANGAN"))
                                vHitung = (Koma2Titik(rsHlp3("Qty")) * (Koma2Titik(rsHlp3("HargaJual")) - (Koma2Titik(rsHlp3("ItemDisc")) / Abs(Koma2Titik(rsHlp3("Qty"))))))
                                vHitung = vHitung - (Koma2Titik(rsHlp3("ExtDisPersen")) * vHitung)
                                If Trim(UCase(rsHlp3("TypePpn"))) = "E" Then
                                    vHitung = Round(vHitung + vHitung * 0.1, 2)
                                End If
                                If gJurnalPart Then
                                    vHitung = cekIncludePPn(vHitung)
                                End If
                                !qty = CDbl(rsHlp3("Qty"))
                                !Harga = CDbl(vHitung / !qty)
                                Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
                                If Not rsHlp.EOF Then
                                    If rsHlp("TIPE") = "O" Then
                                        !oli = True
                                        text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
                                    ElseIf rsHlp("TIPE") = "P" Then
                                        !oli = False
                                        text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
                                    End If
                                End If
                                !QtyOrder = CDbl(rsHlp3("Qty"))
                                !QtyBO = 0
                                !QtyCancel = 0
                                .Update
                                rsHlp3.MoveNext
                            Loop
                            gSpart = True
                        Else
                            If TblValid.Enabled Then
                                temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
                                If temp = vbYes Then
                                    gSpart = False
                                Else
                                    gSpart = True
                                End If
                            End If
                        End If
                        If rsHlp3.State <> adStateClosed Then rsHlp3.Close
                    End With
                Else
                    gSpart = True
                End If
            End If
        End If
        
        subPart = CDbl(text(24))
        subOli = CDbl(text(25))
        If rsPart.State <> adStateClosed Then
            If Not rsPart.BOF Then
                rsPart.MoveFirst
                rsPart.Sort = "OLI"
                Set partGrid.DataSource = rsPart
                Call doHeadPart
            End If
        End If
        
        ' TAMPIL SUBLET ORDER
        adaSublet = False
        'cmdSublet.Visible = False
        text(27) = 0
        If rsSublet.State Then rsSublet.Close
        Set rsSublet = objO.FetchData("SELECT NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, valid, HARGA, PRICE FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noWO & "' AND no_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
        If Not rsSublet.EOF Then
            Do While Not rsSublet.EOF
                If rsSublet("VALID") = 0 Then
                    adaSublet = True
                    MsgBox "Sublet belum di validasi"
                    If rsSublet.State Then rsSublet.Close
                    TblValid.Enabled = False
                    Exit Sub
                End If
                text(27) = text(27) + (IIf(Trim(text(20)) = "", CDbl(rsSublet("HARGA")), CDbl(rsSublet("PRICE"))) * CDbl(rsSublet("QTY")))
                rsSublet.MoveNext
            Loop
            Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
                 & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "') and no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "' " _
                 & "ORDER BY nonota, kodeitem")
            subSublet = CDbl(text(27))
            adaSublet = True
            'cmdSublet.Visible = True
            Set subletGrid.DataSource = rsSublet
            Call doHeadSublet
        Else
            If newValid Then
                If TblValid.Enabled Then
                    If MsgBox("Ada Sublet ?", vbYesNo, "Konfirmasi Sublet") = vbNo Then
                        adaSublet = True
                    End If
                End If
            End If
        End If
        
        text(29) = Round(CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)), 0) '+ CDbl(text(28))
        ' TAMPIL TERBILANG
        txtTotalRp = Angka2Huruf(text(29), 1)
            If OnTopSV = 0 Then
                Set rsHlp = objO.FetchDataACC("select KodeLgn, SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(text(2)) & "'")
            Else
            'UNTUK FEBUARI - TOP PELANGGAN
                Set rsHlp = objO.FetchDataACC("select KodeLgn, TOPGRP as SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(text(2)) & "'")
            End If
            If Not rsHlp.BOF Then
                text(41) = rsHlp("SyaratBayar")
            End If
            If rsHlp.State <> adStateClosed Then rsHlp.Close
        
        If Not gSpart Or Not adaSublet Then
            TblValid.Enabled = False
        End If
        
    ' PEMBACAAN STALL
   
    
    Set rsTmp = objO.FetchData("SELECT FLAG,JAMMULAI,NOSTALL,JAMSELESAI FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & nWO & "|%'")
    If Not rsTmp.EOF Then
    
        If UCase(rsTmp("FLAG")) = "G" Then
            MsgBox "Unit sedang menunggu proses cuci", vbInformation, "Informasi Unit"
            btnBatalWashing.Enabled = True
        ElseIf UCase(rsTmp("FLAG")) = "A" Then
            MsgBox "Unit sedang proses cuci", vbInformation, "Informasi Unit"
        ElseIf UCase(rsTmp("FLAG")) = "F" Then
            MsgBox "Unit telah selesai proses cuci", vbInformation, "Informasi Unit"
        End If
        
        
    End If
    
End Sub

Private Sub hapusForm()
    Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & text(0) & "'")
    
    For i = 0 To 44
        Select Case i
            Case 1, 11, 43
                text(i).Mask = "99-99-9999"
            Case 10, 21, 22, 23, 24, 25, 26, 27, 29, 39
                text(i) = 0
            Case 42, 30, 33
            Case Else
                text(i) = ""
        End Select
    Next
    
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
    chPS.value = Unchecked
    chNea.value = Unchecked
End Sub

Private Sub btnBatalWashing_Click()
    If Trim(text(0)) <> "" Then
        Set rsTmp = objO.FetchData("SELECT FLAG,JAMMULAI,NOSTALL,JAMSELESAI FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & Trim(text(0)) & "|%'")
        If Not rsTmp.EOF Then
            ' TUTUP TAMBAH PEKERJAAN
            If UCase(rsTmp("FLAG")) = "G" Then
                framePembatalanWashing.Visible = True
            End If
        End If
    End If
End Sub

Private Sub btnSimpan_Click()
    Call objO.UpdateData("Update Jobdta Set Catat = '" & Trim(Txt.text) & "' Where NO_WO = '" & Trim(text(0).text) & "'")
    MsgBox "Catatan telah disimpan !", vbInformation + vbOKOnly
End Sub

Private Sub chNea_Click()
    If chNea.value = 1 Then
        txtInfoNea.Visible = True
        txtInfoNea.text = ""
    Else
        txtInfoNea.Visible = False
    End If
End Sub

Private Sub cmdKeluh_Click()
Dim rsSearch As New ADODB.Recordset

    txtPermintaan = ""
    txtKeluhan = ""
    txtFI = ""
    txtNextJob = ""
    
    If framKeluh.Visible Then
        framKeluh.Visible = False
    Else
        Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(text(0)) & "'")
        If Not rsSearch.EOF Then
            txtPermintaan = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
            txtKeluhan = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
        End If
        Set rsSearch = objO.FetchData("SELECT * FROM INSPECTION A WITH (NOLOCK) WHERE NOWO = '" & Trim(text(0)) & "'")
        If Not rsSearch.EOF Then
            txtFI = Trim(rsSearch("HASILFI")) & vbCrLf & Trim(rsSearch("HasilTestDrive"))
            txtNextJob = Trim(rsSearch("SNJ1")) & vbCrLf & Trim(rsSearch("SNJ2")) & vbCrLf & Trim(rsSearch("SNJ3")) & vbCrLf & Trim(rsSearch("SNJ4")) & vbCrLf & Trim(rsSearch("SNJ5")) & vbCrLf & Trim(rsSearch("SNJ6")) & vbCrLf & Trim(rsSearch("SNJ7"))
        End If
        framKeluh.Visible = True
    End If
    
    Dim rscr7 As ADODB.Recordset
    Set rscr7 = objO.FetchData("Select Catat FROM JOBDTA WITH (NOLOCK) Where NO_WO = '" & Trim(text(0)) & "'")
    If rscr7.RecordCount > 0 Then
        Txt.text = rscr7.Fields(0).value
    Else
        Txt.text = ""
    End If

End Sub

Private Sub Command2_Click()
    If Trim(txtAlasanBtlWashing.text) <> "" Then
        If Len(txtAlasanBtlWashing.text) < 5 Then
            MsgBox "Alasan Pembatalan Washing minimal 5 karakter!", vbCritical, "Informasi Pembatalan Washing"
        Else
            objO.InsertData ("INSERT INTO WASHING VALUES ('" & Trim(text(64)) & "',GETDATE(),'" & gUser & "','" & txtAlasanBtlWashing.text & "')")
            MsgBox "Proses Pembatalan Washing berhasil dilakukan", vbInformation, "Informasi Pembatalan Washing"
            txtAlasanBtlWashing.text = ""
            framePembatalanWashing.Visible = False
            btnBatalWashing.Enabled = False
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
    text(1) = TglDDMMYYYY(isnow)
    simpan = True
    Frame6.Top = 0
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    validAktif = False
End Sub

Private Sub TblBatal_Click()
    'If gSpart And adaSublet Then
    Set rsTemp = objO.FetchData("SELECT DISTINCT NO_WO FROM VALIDBILLING WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND DELLOG = ''")
    If rsTemp.EOF Then
        Call objO.UpdateData("UPDATE VALIDORDER SET USERDEL = '" & Trim(gSA) & "', USERDEL_TGL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND USERDEL = ''")
        Call objO.DeleteData("DELETE FROM NEA_PS Where noWO = '" & Trim(text(0)) & "'")
        chNea.value = 0
        chPS.value = 0
        txtInfoNea.text = ""
        lblStatus.Caption = "BATAL VALIDASI"
        TblBatal.Enabled = False
        TblValid.Enabled = True
    Else
        MsgBox "WO ini sudah diberi Diskon Khusus oleh HO, sehingga tidak bisa dibatal Validasi" & vbCr & "Hubungi HO untuk membatalkan diskon khusus. ", vbCritical + vbOKOnly, "Batal Diskon Khusus"
    End If
    'End If
End Sub

Private Sub tblCetakWO_Click()
    If Trim(text(0)) <> "" Then
        Call cetakForm
    End If
End Sub

Private Sub cetakForm()
    Dim rpt As Object
    Dim NamaSAOrder As String
    Dim rsRepo As New ADODB.Recordset
    Dim rsHlp As New ADODB.Recordset
    Dim bodyClear As Boolean
    
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
            rpt.Database.SetDataSource rRs
                       
            gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
            
            sSQL = "SELECT * FROM historyUnitFunc('" & Trim(gFRM) & "')"
            Set sHlp = objO.FetchData(sSQL)
            Set crxSubReportHistory = rpt.Subreport2.OpenSubreport
            crxSubReportHistory.Database.SetDataSource sHlp
            
            sSQL = "select DATE_BOOK, TIME_D, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
            Set sHlp2 = objO.FetchData(sSQL)
            Set crxSubReportBooking = rpt.Subreport1.OpenSubreport
            crxSubReportBooking.Database.SetDataSource sHlp2
            
            sSQL = "select DATE, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
            Set sHlp4 = objO.FetchData(sSQL)
            Set crxSubReportBooking2 = rpt.Subreport4.OpenSubreport
            crxSubReportBooking2.Database.SetDataSource sHlp4
            
            sSQL = "select WAKTU FROM M_FUBOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(31)) & "' "
            Set sHlp3 = objO.FetchData(sSQL)
            Set crxSubReportBookingFU = rpt.Subreport3.OpenSubreport
            crxSubReportBookingFU.Database.SetDataSource sHlp3
            
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
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblValid_Click()
Dim rsSearch As New ADODB.Recordset

    On Error GoTo vbError
    Set rsSearch = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND USERDEL = ''")
    If Not rsSearch.EOF Then
        MsgBox "Sudah di validasi sebelumnya !!", vbInformation, "Validasi Order"
        lblStatus.Caption = "DATA SUDAH DIVALIDASI"
        TblValid.Enabled = False
        TblBatal.Enabled = True
        
    Else
        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(43)), "DD/MM/YYYY") Then
            Dim rsSO As ADODB.Recordset
            Set rsSO = objO.FetchDataACC("select KodeItem  from SOORD WITH (NOLOCK) " & _
                     "Where (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0)+ IsNull(QtyKeluar,0))) <> 0 and NOSO = '" & Trim(text(0)) & "' and BO <> 'CANCEL'")
            If rsSO.RecordCount > 0 Then
                MsgBox "Terdapat " & rsSO.RecordCount & " Line Item Part yang belum tersuplay sesuai order !" & Chr(13) & "Validasi tidak dapat diproses !", vbInformation, "Validasi Order"
                Exit Sub
            End If
        Else
            
        End If
        Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasiDMSNonDSS", "tblValid", 0)
            erTranAll = False
            onTR = True
            Call objHT.OpenTransactionBengkel
            
        If gSpart And adaSublet And Trim(text(17)) <> "" Then
            Call objHT.InsertData_BengkelOnly("INSERT INTO VALIDORDER VALUES ('" & Trim(gSA) & "',getdate(),'" & IIf(rsJob.State = adStateClosed, "TANPA PEKERJAAN| ", "") + IIf(rsPart.State = adStateClosed, "TANPA PART| ", "") + IIf(rsSublet.State = adStateClosed, "TANPA SUBLET| ", "") & "','" & Trim(text(0)) & "','" & Trim(text(17)) & "','','')")
            Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTL SET KARAKTER = '" & Trim(cboPelanggan) & "', DTLKARAKTER = '" & Trim(text(44)) & "' WHERE NO_WO = '" & Trim(text(0)) & "'")
            If chNea.value = 1 Or chPS.value = 1 Then Call objHT.InsertData_BengkelOnly("Insert Into NEA_PS values ('" & Trim(text(0)) & "'," & chNea.value & "," & chPS.value & ",'" & Trim(txtInfoNea.text) & "','" & Trim(gSA) & "',getdate(),'','01/01/1900','')")
            lblStatus.Caption = "SUDAH VALIDASI"
            TblValid.Enabled = False
            TblBatal.Enabled = True
        Else
            lblStatus.Caption = "DATA BELUM LENGKAP"
        End If
            If erTranAll = False Then
                Call objHT.CloseTransactionBengkel
                onTR = False
                Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasiDMSNonDSS", "tblValid", 1)
                Exit Sub
            Else
                Call objHT.CloseTransactionBengkel_err
                onTR = False
                Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasiDMSNonDSS", "tblValid", 2)
                MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
                Exit Sub
            End If
    End If
    Exit Sub

vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasiDMSNonDSS", "tblValid", 2)
    MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
    Exit Sub
 End Sub

Private Sub Text_Change(Index As Integer)
    If Index = 17 Then
        If Trim(text(17)) <> "" Then
            myChassis = Trim(text(17))
            Call CexChassis_Discon(Trim(text(13)))
            If isJualNRM = True Then
                Me.JualNRM.Caption = nJualNRM
            Else
                Me.JualNRM.Caption = ""
            End If
        End If
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If

    Select Case Index
        Case 31
            VScroll1.value = 2500
        Case 32
            VScroll1.value = 2700
        Case 28
            VScroll1.value = 2000
    End Select
        
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    Select Case KeyCode
        Case vbKeyF12
            Select Case Index
                Case 0
                    gMainKeyCari = text(Index)
                    gSearchOption = "INV2"
                    gBantuan = "AND DELIV <> 'BTL'"
                    frmCari.Show (vbModal)
                Case 42
                    gMainKeyCari = text(Index)
                    gSearchOption = "METERAI"
                    frmCari.Show (vbModal)
            End Select
        Case vbKeyEscape
            If simpan = False Then
                If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
                    Call hapusForm
                    text(0).SelLength = Len(text(0))
                    text(0).SetFocus
                End If
            End If
    End Select
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If KeyAscii <> 13 Then
        Select Case Index
            Case 28, 42
                KeyAscii = CekNumerik(KeyAscii)
                simpan = False
        End Select
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim rsHlp As New ADODB.Recordset

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
                            TblBatal.Enabled = True
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
                                TblBatal.Enabled = False
                                noinv = False
                                'Exit Sub
                            ElseIf rsTemp("NOINV") <> "" And rsTemp("RPR_COND") = "K" Then
                                MsgBox ("Telah diclosing !")
                                lblStatus.Caption = "Closing"
                                TblBatal.Enabled = False
                                noinv = False
                                'Exit Sub
                            End If
                        Else
                            TblBatal.Enabled = False
                            If (rsTemp("DELIV") = "PRC" Or rsTemp("DELIV") = "OKE") And rsTemp("RPR_COND") = "S" And rsTemp("NOINV") = "" Then
                                newValid = True
                                text(21).SelLength = Len(text(21))
                                TblValid.Enabled = True
                                TblValid.SetFocus
                                MsgBox ("Unit telah selesai dikerjakan !")
                                lblStatus.Caption = "Selesai dikerjakan"
                            ElseIf rsTemp("DELIV") = "BTL" Then
                                lblStatus.Caption = "WO dibatalkan"
                            ElseIf rsTemp("RPR_COND") = "K" Then
                                lblStatus.Caption = "WO Sudah Closing"
                                noinv = False
                            Else
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
                                            lblStatus.Caption = "Tunggu Washing di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
                                        Case "A"
                                            MsgBox "Unit sedang proses Washing", vbInformation, "Status Order"
                                            lblStatus.Caption = "Washing"
                                        Case "F"
                                            MsgBox ("Unit telah selesai dikerjakan !")
                                            lblStatus.Caption = "Selesai dikerjakan"
                                    End Select
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
End Sub

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
    
    nTop = 6750  '6750 'repo3kosong    '7420 'repo2
    nLeft = 4300 '4300 'repo3kosong    '4440 'repo2
    
    Select Case Urut
        Case 1
            rpt.imgTnd1.TextColor = &H0&
            rpt.imgTnd1.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd1.Left = nLeft + nilY - (210 / 2)
        Case 2
            rpt.imgTnd2.TextColor = &H0&
            rpt.imgTnd2.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd2.Left = nLeft + nilY - (210 / 2)
        Case 3
            rpt.imgTnd3.TextColor = &H0&
            rpt.imgTnd3.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd3.Left = nLeft + nilY - (210 / 2)
        Case 4
            rpt.imgTnd4.TextColor = &H0&
            rpt.imgTnd4.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd4.Left = nLeft + nilY - (210 / 2)
        Case 5
            rpt.imgTnd5.TextColor = &H0&
            rpt.imgTnd5.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd5.Left = nLeft + nilY - (210 / 2)
        Case 6
            rpt.imgTnd6.TextColor = &H0&
            rpt.imgTnd6.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd6.Left = nLeft + nilY - (210 / 2)
        Case 7
            rpt.imgTnd7.TextColor = &H0&
            rpt.imgTnd7.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd7.Left = nLeft + nilY - (210 / 2)
        Case 8
            rpt.imgTnd8.TextColor = &H0&
            rpt.imgTnd8.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd8.Left = nLeft + nilY - (210 / 2)
        Case 9
            rpt.imgTnd9.TextColor = &H0&
            rpt.imgTnd9.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd9.Left = nLeft + nilY - (210 / 2)
        Case 10
            rpt.imgTnd10.TextColor = &H0&
            rpt.imgTnd10.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd10.Left = nLeft + nilY - (210 / 2)
        Case 11
            rpt.imgTnd11.TextColor = &H0&
            rpt.imgTnd11.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd11.Left = nLeft + nilY - (210 / 2)
        Case 12
            rpt.imgTnd12.TextColor = &H0&
            rpt.imgTnd12.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd12.Left = nLeft + nilY - (210 / 2)
        Case 13
            rpt.imgTnd13.TextColor = &H0&
            rpt.imgTnd13.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd13.Left = nLeft + nilY - (210 / 2)
        Case 14
            rpt.imgTnd14.TextColor = &H0&
            rpt.imgTnd14.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd14.Left = nLeft + nilY - (210 / 2)
        Case 15
            rpt.imgTnd15.TextColor = &H0&
            rpt.imgTnd15.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd15.Left = nLeft + nilY - (210 / 2)
        Case 16
            rpt.imgTnd16.TextColor = &H0&
            rpt.imgTnd16.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd16.Left = nLeft + nilY - (210 / 2)
        Case 17
            rpt.imgTnd17.TextColor = &H0&
            rpt.imgTnd17.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd17.Left = nLeft + nilY - (210 / 2)
        Case 18
            rpt.imgTnd18.TextColor = &H0&
            rpt.imgTnd18.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd18.Left = nLeft + nilY - (210 / 2)
        Case 19
            rpt.imgTnd19.TextColor = &H0&
            rpt.imgTnd19.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd19.Left = nLeft + nilY - (210 / 2)
        Case 20
            rpt.imgTnd20.TextColor = &H0&
            rpt.imgTnd20.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd20.Left = nLeft + nilY - (210 / 2)
        Case 21
            rpt.imgTnd21.TextColor = &H0&
            rpt.imgTnd21.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd21.Left = nLeft + nilY - (210 / 2)
        Case 22
            rpt.imgTnd22.TextColor = &H0&
            rpt.imgTnd22.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd22.Left = nLeft + nilY - (210 / 2)
        Case 23
            rpt.imgTnd23.TextColor = &H0&
            rpt.imgTnd23.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd23.Left = nLeft + nilY - (210 / 2)
        Case 24
            rpt.imgTnd24.TextColor = &H0&
            rpt.imgTnd24.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd24.Left = nLeft + nilY - (210 / 2)
        Case 25
            rpt.imgTnd25.TextColor = &H0&
            rpt.imgTnd25.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd25.Left = nLeft + nilY - (210 / 2)
        Case 26
            rpt.imgTnd26.TextColor = &H0&
            rpt.imgTnd26.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd26.Left = nLeft + nilY - (210 / 2)
        Case 27
            rpt.imgTnd27.TextColor = &H0&
            rpt.imgTnd27.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd27.Left = nLeft + nilY - (210 / 2)
        Case 28
            rpt.imgTnd28.TextColor = &H0&
            rpt.imgTnd28.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd28.Left = nLeft + nilY - (210 / 2)
        Case 29
            rpt.imgTnd29.TextColor = &H0&
            rpt.imgTnd29.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd29.Left = nLeft + nilY - (210 / 2)
        Case 30
            rpt.imgTnd30.TextColor = &H0&
            rpt.imgTnd30.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd30.Left = nLeft + nilY - (210 / 2)
        Case 31
            rpt.imgTnd31.TextColor = &H0&
            rpt.imgTnd31.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd31.Left = nLeft + nilY - (210 / 2)
        Case 32
            rpt.imgTnd32.TextColor = &H0&
            rpt.imgTnd32.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd32.Left = nLeft + nilY - (210 / 2)
        Case 33
            rpt.imgTnd33.TextColor = &H0&
            rpt.imgTnd33.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd33.Left = nLeft + nilY - (210 / 2)
        Case 34
            rpt.imgTnd34.TextColor = &H0&
            rpt.imgTnd34.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd34.Left = nLeft + nilY - (210 / 2)
        Case 35
            rpt.imgTnd35.TextColor = &H0&
            rpt.imgTnd35.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd35.Left = nLeft + nilY - (210 / 2)
        Case 36
            rpt.imgTnd36.TextColor = &H0&
            rpt.imgTnd36.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd36.Left = nLeft + nilY - (210 / 2)
        Case 37
            rpt.imgTnd37.TextColor = &H0&
            rpt.imgTnd37.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd37.Left = nLeft + nilY - (210 / 2)
        Case 38
            rpt.imgTnd38.TextColor = &H0&
            rpt.imgTnd38.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd38.Left = nLeft + nilY - (210 / 2)
        Case 39
            rpt.imgTnd39.TextColor = &H0&
            rpt.imgTnd39.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd39.Left = nLeft + nilY - (210 / 2)
        Case 40
            rpt.imgTnd40.TextColor = &H0&
            rpt.imgTnd40.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd40.Left = nLeft + nilY - (210 / 2)
        Case 41
            rpt.imgTnd41.TextColor = &H0&
            rpt.imgTnd41.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd41.Left = nLeft + nilY - (210 / 2)
        Case 42
            rpt.imgTnd42.TextColor = &H0&
            rpt.imgTnd42.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd42.Left = nLeft + nilY - (210 / 2)
        Case 43
            rpt.imgTnd43.TextColor = &H0&
            rpt.imgTnd43.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd43.Left = nLeft + nilY - (210 / 2)
        Case 44
            rpt.imgTnd44.TextColor = &H0&
            rpt.imgTnd44.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd44.Left = nLeft + nilY - (210 / 2)
        Case 45
            rpt.imgTnd45.TextColor = &H0&
            rpt.imgTnd45.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd45.Left = nLeft + nilY - (210 / 2)
        Case 46
            rpt.imgTnd46.TextColor = &H0&
            rpt.imgTnd46.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd46.Left = nLeft + nilY - (210 / 2)
        Case 47
            rpt.imgTnd47.TextColor = &H0&
            rpt.imgTnd47.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd47.Left = nLeft + nilY - (210 / 2)
        Case 48
            rpt.imgTnd48.TextColor = &H0&
            rpt.imgTnd48.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd48.Left = nLeft + nilY - (210 / 2)
        Case 49
            rpt.imgTnd49.TextColor = &H0&
            rpt.imgTnd49.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd49.Left = nLeft + nilY - (210 / 2)
        Case 50
            rpt.imgTnd50.TextColor = &H0&
            rpt.imgTnd50.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd50.Left = nLeft + nilY - (210 / 2)
        Case 51
            rpt.imgTnd51.TextColor = &H0&
            rpt.imgTnd51.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd51.Left = nLeft + nilY - (210 / 2)
        Case 52
            rpt.imgTnd52.TextColor = &H0&
            rpt.imgTnd52.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd52.Left = nLeft + nilY - (210 / 2)
        Case 53
            rpt.imgTnd53.TextColor = &H0&
            rpt.imgTnd53.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd53.Left = nLeft + nilY - (210 / 2)
    End Select
End Sub



