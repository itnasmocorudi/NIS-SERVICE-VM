VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoInvTot 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Revenue Total"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   5715
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1350
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   5715
      Width           =   1440
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2835
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   5715
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetak 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   5715
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5160
      TabIndex        =   88
      Top             =   5745
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   4110
      TabIndex        =   93
      Top             =   5715
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   476
      BTYPE           =   8
      TX              =   "Printer :"
      ENAB            =   0   'False
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoInvTot.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   750
      ScaleHeight     =   1395
      ScaleWidth      =   5235
      TabIndex        =   67
      Top             =   6345
      Visible         =   0   'False
      Width           =   5295
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1290
         TabIndex        =   68
         Top             =   0
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2160
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
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
         Index           =   2
         Left            =   1290
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   5160
         TabIndex        =   71
         Top             =   360
         Visible         =   0   'False
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
         Index           =   3
         Left            =   1290
         TabIndex        =   72
         Top             =   720
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   2160
         TabIndex        =   73
         Top             =   720
         Visible         =   0   'False
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
         Left            =   2280
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
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
         Index           =   4
         Left            =   2280
         TabIndex        =   75
         Top             =   720
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mechanic"
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
         Left            =   540
         TabIndex        =   78
         Top             =   765
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   480
         TabIndex        =   77
         Top             =   405
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Job Order"
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
         Left            =   0
         TabIndex        =   76
         Top             =   45
         Visible         =   0   'False
         Width           =   1200
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   720
      TabIndex        =   20
      Top             =   5400
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
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame frmeAturKredit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "                    ... :: Pengaturan Invoice posisi CASH / KREDIT :: ...                 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   8520
      TabIndex        =   48
      Top             =   5520
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton cmdTo 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&KREDIT"
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
         Index           =   1
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2760
         Width           =   1890
      End
      Begin VB.CommandButton cmdTo 
         BackColor       =   &H00DCFCFB&
         Caption         =   "CAS&H"
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
         Index           =   0
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2760
         Width           =   1890
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   1200
         TabIndex        =   49
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   18
         Left            =   1200
         TabIndex        =   51
         Top             =   840
         Width           =   4455
         _ExtentX        =   7858
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
         Index           =   19
         Left            =   1200
         TabIndex        =   54
         Top             =   1155
         Width           =   4455
         _ExtentX        =   7858
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
         Index           =   20
         Left            =   1200
         TabIndex        =   55
         Top             =   1485
         Width           =   4455
         _ExtentX        =   7858
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
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   0
         Left            =   7680
         Picture         =   "frmRepoInvTot.frx":001C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   735
         TabIndex        =   57
         Top             =   1530
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   570
         TabIndex        =   56
         Top             =   1195
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         TabIndex        =   52
         Top             =   885
         Width           =   690
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Invoice"
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
         TabIndex        =   50
         Top             =   525
         Width           =   825
      End
   End
   Begin VB.CommandButton tblRefresh 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Refresh"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4800
      Width           =   1890
   End
   Begin VB.CommandButton TblSimpan 
      BackColor       =   &H00DCFCFB&
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4800
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   480
      TabIndex        =   21
      Top             =   480
      Width           =   12015
      Begin VB.CheckBox opSource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TMS"
         Height          =   375
         Index           =   2
         Left            =   9360
         TabIndex        =   97
         Top             =   3240
         Width           =   975
      End
      Begin VB.CheckBox opSource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Small SP"
         Height          =   375
         Index           =   1
         Left            =   8280
         TabIndex        =   95
         Top             =   3240
         Width           =   975
      End
      Begin VB.CheckBox opSource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reguler"
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   94
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         TabIndex        =   84
         Top             =   1260
         Width           =   3855
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Beberapa Model"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   86
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1 Model"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin TDBDate6Ctl.TDBDate DtDOAwal 
         Height          =   330
         Left            =   7200
         TabIndex        =   15
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoInvTot.frx":3937
         Caption         =   "frmRepoInvTot.frx":3A63
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoInvTot.frx":3ACF
         Keys            =   "frmRepoInvTot.frx":3AED
         Spin            =   "frmRepoInvTot.frx":3B4B
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7200
         TabIndex        =   60
         Top             =   3000
         Width           =   2895
         Begin VB.OptionButton opLap 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan Shift"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton opLap 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CheckBox chTHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THS"
         Height          =   255
         Left            =   7200
         TabIndex        =   59
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton cmdCash 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Atur Cash / Kredit"
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.OptionButton opTipe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Harian Cetak"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   5
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CheckBox chInvDealer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hanya Invoice B"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton opTipe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Global"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin VB.OptionButton opTipe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Harian"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Top             =   3720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   110559233
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1740
         TabIndex        =   2
         Top             =   840
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   2610
         TabIndex        =   22
         Top             =   840
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
         Index           =   7
         Left            =   1740
         TabIndex        =   3
         Top             =   1560
         Width           =   4005
         _ExtentX        =   7064
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   5760
         TabIndex        =   23
         Top             =   1560
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
         Index           =   6
         Left            =   2750
         TabIndex        =   24
         Top             =   840
         Width           =   3150
         _ExtentX        =   5556
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
         Index           =   8
         Left            =   1740
         TabIndex        =   25
         Top             =   1920
         Width           =   4155
         _ExtentX        =   7329
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
         Index           =   9
         Left            =   1740
         TabIndex        =   7
         Top             =   2280
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   2610
         TabIndex        =   26
         Top             =   2280
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
         Index           =   10
         Left            =   2760
         TabIndex        =   27
         Top             =   2280
         Width           =   3150
         _ExtentX        =   5556
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
         Index           =   11
         Left            =   1740
         TabIndex        =   8
         Top             =   3000
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   3030
         TabIndex        =   28
         Top             =   3000
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
         Index           =   12
         Left            =   7170
         TabIndex        =   9
         Top             =   840
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   270
         Left            =   9015
         TabIndex        =   29
         Top             =   840
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
         Index           =   13
         Left            =   7170
         TabIndex        =   10
         Top             =   1200
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   9015
         TabIndex        =   30
         Top             =   1200
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   111149057
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   7170
         TabIndex        =   11
         Top             =   1560
         Width           =   1755
         _ExtentX        =   3096
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
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   270
         Left            =   9015
         TabIndex        =   40
         Top             =   1560
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
         Index           =   15
         Left            =   7170
         TabIndex        =   12
         Top             =   1920
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   9015
         TabIndex        =   43
         Top             =   1920
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
         Index           =   16
         Left            =   7170
         TabIndex        =   45
         Top             =   2280
         Width           =   4275
         _ExtentX        =   7541
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
         Index           =   21
         Left            =   10800
         TabIndex        =   17
         Top             =   3870
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   270
         Left            =   11400
         TabIndex        =   61
         Top             =   3870
         Visible         =   0   'False
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
         Index           =   22
         Left            =   7200
         TabIndex        =   62
         Top             =   3600
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   23
         Left            =   8280
         TabIndex        =   63
         Top             =   3600
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin TDBDate6Ctl.TDBDate DtDOAkhir 
         Height          =   330
         Left            =   9240
         TabIndex        =   16
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoInvTot.frx":3B73
         Caption         =   "frmRepoInvTot.frx":3C9F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoInvTot.frx":3D0B
         Keys            =   "frmRepoInvTot.frx":3D29
         Spin            =   "frmRepoInvTot.frx":3D87
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   4365
         TabIndex        =   80
         Top             =   3000
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   270
         Left            =   5760
         TabIndex        =   81
         Top             =   3000
         Visible         =   0   'False
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
         Index           =   25
         Left            =   1740
         TabIndex        =   83
         Top             =   3360
         Visible         =   0   'False
         Width           =   4155
         _ExtentX        =   7329
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
         Index           =   26
         Left            =   2280
         TabIndex        =   87
         Top             =   1920
         Visible         =   0   'False
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outlet"
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
         Left            =   6600
         TabIndex        =   96
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Program"
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
         Left            =   3240
         TabIndex        =   82
         Top             =   3045
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "THS"
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
         TabIndex        =   79
         Top             =   2640
         Width           =   285
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl WO"
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
         Left            =   6540
         TabIndex        =   66
         Top             =   3915
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8760
         TabIndex        =   65
         Top             =   3840
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Report"
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
         Left            =   6225
         TabIndex        =   64
         Top             =   3030
         Width           =   840
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
         Left            =   6405
         TabIndex        =   44
         Top             =   1950
         Width           =   660
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
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
         Left            =   1230
         TabIndex        =   42
         Top             =   3810
         Width           =   300
      End
      Begin VB.Label Label45 
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
         Left            =   6600
         TabIndex        =   41
         Top             =   1620
         Width           =   465
      End
      Begin VB.Label Label3 
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
         Left            =   1335
         TabIndex        =   38
         Top             =   885
         Width           =   195
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaraan"
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
         TabIndex        =   37
         Top             =   1245
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Penjual"
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
         Left            =   495
         TabIndex        =   36
         Top             =   2325
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Service"
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
         Left            =   660
         TabIndex        =   35
         Top             =   3045
         Width           =   870
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Order"
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
         TabIndex        =   34
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proses"
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
         Left            =   6600
         TabIndex        =   33
         Top             =   1245
         Width           =   480
      End
      Begin VB.Label Label10 
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
         Left            =   870
         TabIndex        =   32
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3240
         TabIndex        =   31
         Top             =   240
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   12000
         Y1              =   600
         Y2              =   600
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8175
      Left            =   0
      Picture         =   "frmRepoInvTot.frx":3DAF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoInvTot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim Report          As Object
Dim Report2         As Object
Dim Report3         As Object
Dim rsTemp          As ADODB.Recordset
Dim vNamaTabel      As String
Dim objBilling      As New BClsBilling
Dim jmlUnit         As Integer
Dim jmlInv          As Integer
Dim whereMultiple   As String
Dim isAdaData       As Boolean


Private Sub hapusForm()
    For i = 0 To 23
        Text(i) = ""
    Next
    Text(26) = ""
    Tgl1 = isnow
    Tgl2 = isnow
    chTHS.value = Unchecked
End Sub

Private Sub BuatTabelRepoInv(vNamaTabel As String)
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    
    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
         & "    [SA][char](5) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](15) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NO_POLISI][varchar](12) NOT NULL DEFAULT '', " _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [JOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [JRTE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [RTJ][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [THS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [GRP][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TWC][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [LUB][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PDI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [WAS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [BOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [MATERIAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PART][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [OLI][decimal](18,2) NOT NULL DEFAULT 0, "
         
    sSQL = sSQL + "[SUBLET][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SUBTOT][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [METERAI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TOTAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
         & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "    [LABOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
         & "    [HPP][decimal](18,2) NOT NULL DEFAULT 0," _
         & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000') ON [PRIMARY]"
    Call objO.UpdateData(sSQL)
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    reateTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepoInv(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo(Ket As String, vNamaTabel As String, txt As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim rsHlp As ADODB.Recordset
      Dim rsAdd As ADODB.Recordset
      Dim nNilai As Double
      Dim nDuoInv As Boolean
      Dim subletOnly As Boolean
      Dim noInv As String
      Dim nmCust As String
      Dim nStatus As String
      Dim ppnInv As Double
1         If Log.RPLHistori = 0 Then
2             Set rsTemp = objO.FetchData("SELECT CASE DELIV WHEN 'BTL' THEN CASE WHEN FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' THEN 0 ELSE 1 END ELSE 1 END AS HIT_UNIT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                          & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, FD_RETUR, " _
                                          & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP,A.PPN " _
                                          & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                          & "WHERE INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' " & Ket & " ORDER BY NOINV")
3         Else
              ''20201211 - revisi data double ditambahkan distinct
4             Set rsTemp = objO.FetchData("SELECT DISTINCT CASE DELIV WHEN 'BTL' THEN CASE WHEN FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' THEN 0 ELSE 1 END ELSE 1 END AS HIT_UNIT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                          & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, FD_RETUR, " _
                                          & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP,A.PPN " _
                                          & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C WITH(NOLOCK) ON C.FS_LOG_DEL = '' AND A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                          & "WHERE INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' " & Ket & " ORDER BY NOINV")
5         End If
                                   '--:: 23 Agustus 2010 ; Menampilkan nilai retur ; DELIV <> 'BTL' AND ::--
6         If Not rsTemp.EOF Then
7             Call BuatTabelRepoInv(vNamaTabel)
8             jmlUnit = 0
9             jmlInv = 0
10            Do While Not rsTemp.EOF
11                jmlUnit = jmlUnit + IIf(Right(rsTemp("NOINV"), 2) <> "-B", 1, 0) 'rsTemp("HIT_UNIT")
12                jmlInv = jmlInv + 1
13                subletOnly = True
14                ppnInv = CDbl(rsTemp("ppn"))
15                For i = 1 To 8
16                    hlp = "REP" & i
17                    If rsTemp(hlp) <> "" Then
18                        subletOnly = False
19                        hlp2 = "KLPKERJA" & i
20                        hlp3 = "RTE" & i    '"ACTRATE" & i
21                        hlp4 = "MEKANIK" & i
22                        hlp5 = "OFF" & i
23                        hlp6 = "DISC" & i
24                        hlp7 = "FREEC" & i
25                        If createTemp(vNamaTabel) Then
26                            nDuoInv = False
27                            If rsTemp("NDUOINV") = 1 Then
28                                nDuoInv = True
29                            End If
                              '#########################################################'
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                              '#########################################################'
30                            nStatus = rsTemp("STATUS")
31                            nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
32                            noInv = Trim(rsTemp("NOINV"))
33                            nmCust = Trim(rsTemp("CUSTOMER"))
34                            If nStatus = "AFFNONPPN" Then
35                                nNilai = nNilai * (100 / (100 + ppnInv))  ''(100 / 110)
36                            End If
                              
37                            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
38                            If Not rsHlp.EOF Then
      '                            Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
      '                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
      '                                         & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE" And Trim(rsTemp("FS_RC")) <> "" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP" And Trim(rsTemp("FS_RC")) = "" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
      '                                         & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
      '                                         & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
      '                                         & "BOR = '" & setKoma(cekDuo((CDbl(rsTemp("BODY"))), noinv)) & "', MATERIAL = '" & setKoma(cekDuo((CDbl(rsTemp("MTRL"))), noinv)) & "', PART = '" & setKoma(cekDuo((CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noinv)) & "', " _
      '                                         & "OLI = '" & setKoma(cekDuo((CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noinv)) & "', SUBLET = '" & setKoma(cekDuo((CDbl(rsTemp("SBLT"))), noinv)) & "', SUBTOT = '" & setKoma(cekDuo((CDbl(rsTemp("SUBTOT"))), noinv)) & "', " _
      '                                         & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv And Right(rsTemp("NOINV"), 1) = "B", "", Trim(rsTemp("NOFAKTUR"))) & "', HPP = '" & setKoma(cekDuo((CDbl(rsTemp("HPP"))), noinv)) & "' " _
      '                                         & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                                  
                                  '---------------------------------------
                                  '2015-08-03 | hilangkan fs_rc pada grp
                                  '---------------------------------------
                                  
39                                Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE" And Trim(rsTemp("FS_RC")) <> "" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "BOR = '" & setKoma(cekDuo((CDbl(rsTemp("BODY"))), noInv)) & "', MATERIAL = '" & setKoma(cekDuo((CDbl(rsTemp("MTRL"))), noInv)) & "', PART = '" & setKoma(cekDuo((CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv)) & "', " _
                                               & "OLI = '" & setKoma(cekDuo((CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv)) & "', SUBLET = '" & setKoma(cekDuo((CDbl(rsTemp("SBLT"))), noInv)) & "', SUBTOT = '" & setKoma(cekDuo((CDbl(rsTemp("SUBTOT"))), noInv)) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv And Right(rsTemp("NOINV"), 1) = "B", "", Trim(rsTemp("NOFAKTUR"))) & "', HPP = '" & setKoma(cekDuo((CDbl(rsTemp("HPP"))), noInv)) & "' " _
                                               & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
                                               
40                            Else
      '                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noinv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
      '                                         & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
      '                                         & "'" & IIf(rsTemp(hlp2) = "SBE" And rsTemp("FS_RC") <> "" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("FS_RC") = "" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', " _
      '                                         & "'" & IIf(rsTemp(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
      '                                         & "'" & setKoma(cekDuo((CDbl(rsTemp("BODY"))), noinv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("MTRL"))), noinv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noinv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noinv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("SBLT"))), noinv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("SUBTOT"))), noinv)) & "', " _
      '                                         & " '" & setKoma(cekDuo((CDbl(rsTemp("MET"))), noinv, 1)) & "', '" & setKoma(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & IIf(nDuoInv And Right(rsTemp("NOINV"), 1) = "B", "", Trim(rsTemp("NOFAKTUR"))) & "', '" & setKoma(cekDuo((CDbl(rsTemp("HPP"))), noinv)) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
                                  
                                  '---------------------------------------
                                  '2015-08-03 | hilangkan fs_rc pada grp
                                  '---------------------------------------
                                  
41                                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsTemp(hlp2) = "SBE" And rsTemp("FS_RC") <> "" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsTemp(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & setKoma(cekDuo((CDbl(rsTemp("BODY"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("MTRL"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("SBLT"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("SUBTOT"))), noInv)) & "', " _
                                               & " '" & setKoma(cekDuo((CDbl(rsTemp("MET"))), noInv, 1)) & "', '" & setKoma(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & IIf(nDuoInv And Right(rsTemp("NOINV"), 1) = "B", "", Trim(rsTemp("NOFAKTUR"))) & "', '" & setKoma(cekDuo((CDbl(rsTemp("HPP"))), noInv)) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
                              
42                            End If
43                        End If
44                    End If
45                Next
46                If subletOnly Then
47                    nDuoInv = False
48                    If rsTemp("NDUOINV") = 1 Then
49                        nDuoInv = True
50                    End If
                      '#########################################################'
                      '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                      '#########################################################'
51                    noInv = Trim(rsTemp("NOINV"))
52                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & noInv & "'")
53                    If Not rsHlp.EOF Then
54                        Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                       & "JOR = '0', TIPEINV = '" & Trim(rsTemp("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', BOR = '" & setKoma(CDbl(rsTemp("BODY"))) & "', MATERIAL = '" & setKoma(CDbl(rsTemp("MTRL"))) & "', " _
                                       & "PART = '" & setKoma(CDbl(rsTemp("PART_G") + rsTemp("PART_NG"))) & "', OLI = '" & setKoma(CDbl(rsTemp("OIL_G") + rsTemp("OIL_NG"))) & "', SUBLET = '" & setKoma(CDbl(rsTemp("SBLT"))) & "', SUBTOT = '" & setKoma(CDbl(rsTemp("SUBTOT"))) & "', " _
                                       & "LABOR = '" & setKoma(CDbl(rsTemp("LABOR"))) & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & setKoma(CDbl(rsTemp("HPP"))) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
55                    Else
56                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                       & "'0','0', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                       & "'" & setKoma(CDbl(rsTemp("BODY"))) & "', '" & setKoma(CDbl(rsTemp("MTRL"))) & "', '" & setKoma(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))) & "', '" & setKoma(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))) & "', '" & setKoma(CDbl(rsTemp("SBLT"))) & "', '" & setKoma(CDbl(rsTemp("SUBTOT"))) & "', " _
                                       & " '" & setKoma(CDbl(rsTemp("MET"))) & "', '" & setKoma(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(rsTemp("STATUS")) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & setKoma(CDbl(rsTemp("HPP"))) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
57                    End If
58                End If
                  
                  ' DARI ADD JOB
59                Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NO_INV = '" & rsTemp("NOINV") & "' AND FS_LOG_DEL = '" & Trim(rsTemp("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
60                If Not rsAdd.EOF Then
61                    Do While Not rsAdd.EOF
62                        hlp2 = "KLPKERJA"
63                        hlp3 = "RTE"
64                        hlp4 = "MEKANIK"
65                        hlp5 = "OFFN"
66                        hlp6 = "FREEC"
                          
67                        nDuoInv = False
68                        If rsTemp("NDUOINV") = 1 Then
69                            nDuoInv = True
70                        End If
                          '#########################################################'
                          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                          '#########################################################'
71                        nStatus = rsTemp("STATUS")
72                        nNilai = CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsTemp("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
73                        noInv = Trim(rsTemp("NOINV"))
74                        nmCust = Trim(rsTemp("CUSTOMER"))
75                        If Trim(nStatus) = "AFFNONPPN" Then
76                            nNilai = nNilai * (100 / (100 + ppnInv))  ''110)
77                        End If
                          
78                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & (noInv) & "'")
79                        If Not rsHlp.EOF Then
80                            Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                           & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai), 0), 0) & "', " _
                                           & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "BOR = '" & setKoma(cekDuo((CDbl(rsTemp("BODY"))), noInv)) & "', MATERIAL = '" & setKoma(cekDuo((CDbl(rsTemp("MTRL"))), noInv)) & "', PART = '" & setKoma(cekDuo((CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv)) & "', " _
                                           & "OLI = '" & setKoma(cekDuo((CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv)) & "', SUBLET = '" & setKoma(cekDuo((CDbl(rsTemp("SBLT"))), noInv)) & "', SUBTOT = '" & setKoma(CDbl(rsTemp("SUBTOT"))) & "', HPP = '" & setKoma(cekDuo((CDbl(rsTemp("HPP"))), noInv)) & "', " _
                                           & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & setKoma(CDbl(rsTemp("LABOR"))) & "', NOFAKTUR = '" & IIf(nDuoInv And Right(rsTemp("NOINV"), 1) = "B", "", Trim(rsTemp("NOFAKTUR"))) & "' " _
                                           & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
81                        Else
82                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                           & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "SBE" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & setKoma(cekDuo((CDbl(rsTemp("BODY"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("MTRL"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv)) & "', '" & setKoma(cekDuo((CDbl(rsTemp("SBLT"))), noInv)) & "', '" & CDbl(setKoma(rsTemp("SUBTOT"))) & "', " _
                                           & " '" & setKoma(cekDuo((CDbl(rsTemp("MET"))), noInv, 1)) & "', '" & setKoma(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & IIf(nDuoInv And Right(rsTemp("NOINV"), 1) = "B", "", Trim(rsTemp("NOFAKTUR"))) & "','" & setKoma(cekDuo((CDbl(rsTemp("HPP"))), noInv)) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
83                        End If
84                        rsAdd.MoveNext
85                    Loop
86                End If
87                rsTemp.MoveNext
88            Loop
89            Call IsiTempRepoRetur(Ket, vNamaTabel, txt)
90            Call Cetak(vNamaTabel, txt)
91            Call DisableExport
92            Call setLokasiTombol
93        Else
94            MsgBox "Tidak ada data !!"
95            Exit Sub
96        End If
End Sub

Private Sub IsiTempRepoRetur(Ket As String, vNamaTabel As String, txt As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim rsHlp As ADODB.Recordset
      Dim rsAdd As ADODB.Recordset
      Dim nNilai As Double
      Dim nDuoInv As Boolean
      Dim subletOnly As Boolean
      Dim noInv As String
      Dim nmCust As String
      Dim nStatus As String
      Dim rsTemp2 As ADODB.Recordset
          
1         If Log.RPLHistori = 0 Then
2             Set rsTemp2 = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, FD_RETUR, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE DELIV = 'BTL' AND FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' " & Ket & " ORDER BY NOINV")
3         Else
4             Set rsTemp2 = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, FD_RETUR, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C WITH(NOLOCK) ON C.FS_LOG_DEL = '' AND A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE DELIV = 'BTL' AND FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' " & Ket & " ORDER BY NOINV")
5         End If
6         If Not rsTemp2.EOF Then
7             If createTemp(vNamaTabel) Then
8                 Do While Not rsTemp2.EOF
9                     subletOnly = True
10                    jmlUnit = jmlUnit - IIf(Right(rsTemp2("NOINV"), 2) <> "-B", 1, 0)
11                    jmlInv = jmlInv - 1
12                    For i = 1 To 8
13                        hlp = "REP" & i
14                        If rsTemp2(hlp) <> "" Then
15                            subletOnly = False
16                            hlp2 = "KLPKERJA" & i
17                            hlp3 = "RTE" & i    '"ACTRATE" & i
18                            hlp4 = "MEKANIK" & i
19                            hlp5 = "OFF" & i
20                            hlp6 = "DISC" & i
21                            hlp7 = "FREEC" & i
22                            If createTemp(vNamaTabel) Then
23                                nDuoInv = False
24                                If rsTemp2("NDUOINV") = 1 Then
25                                    nDuoInv = True
26                                End If
                                  '#########################################################'
                                  '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                                  '#########################################################'
27                                nStatus = rsTemp2("STATUS")
28                                nNilai = -1 * CDbl(rsTemp2(hlp5)) * (1 - (CDbl(rsTemp2(hlp6)) / 100))
29                                noInv = Trim(rsTemp2("NOINV")) & "R"
30                                nmCust = Trim(rsTemp2("CUSTOMER"))
31                                If nStatus = "AFFNONPPN" Then
32                                    nNilai = nNilai * (100 / 110)
33                                End If
                                  
34                                Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp2("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp2("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
35                                If Not rsHlp.EOF Then
36                                    Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp2("FRONT") & "', NO_POLISI = '" & rsTemp2("NO_POLISI") & "', " _
                                                   & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp2(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp2(hlp2) = "SBI" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                                   & "SBE = SBE + '" & IIf(rsTemp2(hlp2) = "SBE" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp2(hlp2) = "GRP" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                                   & "TWC = TWC + '" & IIf(rsTemp2("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp2(hlp2) = "LUB" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                                   & "PDI = PDI + '" & IIf(rsTemp2(hlp2) = "PDS" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp2(hlp2) = "WAS" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                                   & "BOR = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("BODY")), noInv)) & "', MATERIAL = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("MTRL")), noInv)) & "', PART = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("PART_G")) + CDbl(rsTemp2("PART_NG")), noInv)) & "', " _
                                                   & "OLI = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("OIL_G")) + CDbl(rsTemp2("OIL_NG")), noInv)) & "', SUBLET = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("SBLT")), noInv)) & "', SUBTOT = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("SUBTOT")), noInv)) & "', " _
                                                   & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & -1 * Replace(CDbl(rsTemp2("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv And Right(rsTemp2("NOINV"), 1) = "B", "", Trim(rsTemp2("NOFAKTUR"))) & "', HPP = '" & -1 * cekDuo(CDbl(rsTemp2("HPP")), noInv) & "' " _
                                                   & "WHERE NO_CHASSIS = '" & rsTemp2("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp2("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
37                                Else
38                                    Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp2("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp2("NO_WO") & "', '" & rsTemp2("NO_POLISI") & "','" & rsTemp2("NO_CHASSIS") & "', " _
                                                   & "'0','" & Replace(CDbl(rsTemp2(hlp3)), ",", ".") & "', '" & IIf(rsTemp2("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp2("THS") = 0, 0, 1) & "', '" & IIf(rsTemp2(hlp2) = "SBI" And rsTemp2("TPSVC") <> "W", RepRp(Round(CDbl(nNilai), 0)), 0) & "', " _
                                                   & "'" & IIf(rsTemp2(hlp2) = "SBE" And rsTemp2("TPSVC") <> "W", RepRp(Round(CDbl(nNilai))), 0) & "','" & IIf(rsTemp2(hlp2) = "GRP" And rsTemp2("TPSVC") <> "W", RepRp(Round(CDbl(nNilai))), 0) & "', '" & IIf(rsTemp2("TPSVC") = "W", RepRp(Round(CDbl(nNilai))), 0) & "', " _
                                                   & "'" & IIf(rsTemp2(hlp2) = "LUB" And rsTemp2("TPSVC") <> "W", RepRp(Round(CDbl(nNilai))), 0) & "', '" & IIf(rsTemp2(hlp2) = "PDS" And rsTemp2("TPSVC") <> "W", RepRp(Round(CDbl(nNilai))), 0) & "', '" & IIf(rsTemp2(hlp2) = "WAS" And rsTemp2("TPSVC") <> "W", RepRp(Round(CDbl(nNilai))), 0) & "', " _
                                                   & "'" & RepRp(-1 * cekDuo(CDbl(rsTemp2("BODY")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("MTRL")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("PART_G")) + CDbl(rsTemp2("PART_NG")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("OIL_G")) + CDbl(rsTemp2("OIL_NG")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("SBLT")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("SUBTOT")), noInv)) & "', " _
                                                   & " '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("MET")), noInv, 1)) & "', '" & RepRp(-1 * CDbl(rsTemp2("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & RepRp(-1 * Replace(CDbl(rsTemp2("LABOR")), ",", ".")) & "', '" & IIf(nDuoInv And Right(rsTemp2("NOINV"), 1) = "B", "", Trim(rsTemp2("NOFAKTUR"))) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("HPP")), noInv)) & "','" & Format(rsTemp2("FD_RETUR"), "MM/DD/YYYY") & "')")
39                                End If
40                            End If
41                        End If
42                    Next
43                    If subletOnly Then
44                        nDuoInv = False
45                        If rsTemp2("NDUOINV") = 1 Then
46                            nDuoInv = True
47                        End If
                          '#########################################################'
                          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                          '#########################################################'
48                        noInv = Trim(rsTemp2("NOINV")) & "R"
                          
49                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp2("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp2("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
50                        If Not rsHlp.EOF Then
51                            Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp2("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp2("NO_POLISI") & "', " _
                                           & "JOR = '0', TIPEINV = '" & Trim(rsTemp2("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp2("CUSTOMER")) & "', BOR = '" & RepRp(-1 * CDbl(rsTemp2("BODY"))) & "', MATERIAL = '" & RepRp(-1 * CDbl(rsTemp2("MTRL"))) & "', " _
                                           & "PART = '" & RepRp(-1 * CDbl(rsTemp2("PART_G") + rsTemp2("PART_NG"))) & "', OLI = '" & RepRp(-1 * CDbl(rsTemp2("OIL_G") + rsTemp2("OIL_NG"))) & "', SUBLET = '" & RepRp(-1 * CDbl(rsTemp2("SBLT"))) & "', SUBTOT = '" & RepRp(-1 * CDbl(rsTemp2("SUBTOT"))) & "', " _
                                           & "LABOR = '" & RepRp(-1 * CDbl(rsTemp2("LABOR"))) & "', NOFAKTUR = '" & Trim(rsTemp2("NOFAKTUR")) & "', HPP = '" & RepRp(-1 * CDbl(rsTemp2("HPP"))) & "' WHERE NO_CHASSIS = '" & rsTemp2("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp2("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
52                        Else
53                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp2("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp2("NO_WO") & "', '" & rsTemp2("NO_POLISI") & "','" & rsTemp2("NO_CHASSIS") & "', " _
                                           & "'0','0', '" & IIf(rsTemp2("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp2("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                           & "'" & RepRp(-1 * CDbl(rsTemp2("BODY"))) & "', '" & RepRp(-1 * CDbl(rsTemp2("MTRL"))) & "', '" & RepRp(-1 * (CDbl(rsTemp2("PART_G")) + CDbl(rsTemp2("PART_NG")))) & "', '" & RepRp(-1 * (CDbl(rsTemp2("OIL_G")) + CDbl(rsTemp2("OIL_NG")))) & "', '" & RepRp(-1 * CDbl(rsTemp2("SBLT"))) & "', '" & RepRp(-1 * CDbl(rsTemp2("SUBTOT"))) & "', " _
                                           & " '" & RepRp(-1 * CDbl(rsTemp2("MET"))) & "', '" & RepRp(-1 * CDbl(rsTemp2("TOTALL"))) & "', '" & Trim(rsTemp2("STATUS")) & "', '" & Trim(rsTemp2("CUSTOMER")) & "','" & RepRp(-1 * CDbl(rsTemp2("LABOR"))) & "', '" & Trim(rsTemp2("NOFAKTUR")) & "','" & RepRp(-1 * CDbl(rsTemp2("HPP"))) & "','" & Format(rsTemp2("FD_RETUR"), "MM/DD/YYYY") & "')")
54                        End If
55                    End If
                      
                      ' DARI ADD JOB
56                    Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsTemp2("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp2("NO_WO") & "' AND NO_INV = '" & rsTemp2("NOINV") & "' AND FS_LOG_DEL = '" & Trim(rsTemp2("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsTemp2("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
57                    If Not rsAdd.EOF Then
58                        Do While Not rsAdd.EOF
59                            hlp2 = "KLPKERJA"
60                            hlp3 = "RTE"
61                            hlp4 = "MEKANIK"
62                            hlp5 = "OFFN"
63                            hlp6 = "FREEC"
                              
64                            nDuoInv = False
65                            If rsTemp2("NDUOINV") = 1 Then
66                                nDuoInv = True
67                            End If
                              '#########################################################'
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                              '#########################################################'
68                            nStatus = rsTemp2("STATUS")
69                            nNilai = -1 * CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsTemp2("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
70                            noInv = Trim(rsTemp2("NOINV")) & "R"
71                            nmCust = Trim(rsTemp2("CUSTOMER"))
                              
72                            If Trim(nStatus) = "AFFNONPPN" Then
73                                nNilai = (nNilai * (100 / 110))
74                            End If
                              
75                            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp2("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp2("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
76                            If Not rsHlp.EOF Then
77                                Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp2("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp2("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI" And rsTemp2("TPSVC") <> "W", RepRp(CDbl(nNilai)), 0) & "', " _
                                               & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "TWC = TWC + '" & IIf(rsTemp2("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "WAS" And rsTemp2("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "BOR = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("BODY")), noInv)) & "', MATERIAL = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("MTRL")), noInv)) & "', PART = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("PART_G")) + CDbl(rsTemp2("PART_NG")), noInv)) & "', " _
                                               & "OLI = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("OIL_G")) + CDbl(rsTemp2("OIL_NG")), noInv)) & "', SUBLET = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("SBLT")), noInv)) & "', SUBTOT = '" & RepRp(-1 * CDbl(rsTemp2("SUBTOT"))) & "', HPP = '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("HPP")), noInv)) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & RepRp(CDbl(rsTemp2("LABOR"))) & "', NOFAKTUR = '" & IIf(nDuoInv And Right(rsTemp2("NOINV"), 1) = "B", "", Trim(rsTemp2("NOFAKTUR"))) & "' " _
                                               & "WHERE NO_CHASSIS = '" & rsTemp2("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp2("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
78                            Else
79                                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp2("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp2("NO_WO") & "', '" & rsTemp2("NO_POLISI") & "','" & rsTemp2("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp2("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp2("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI" And rsTemp2("TPSVC") <> "W", RepRp(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsAdd(hlp2) = "SBE" And rsTemp2("TPSVC") <> "W", RepRp(CDbl(nNilai)), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsTemp2("TPSVC") <> "W", RepRp(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp2("TPSVC") = "W", RepRp(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsAdd(hlp2) = "LUB" And rsTemp2("TPSVC") <> "W", RepRp(CDbl(nNilai)), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS" And rsTemp2("TPSVC") <> "W", RepRp(CDbl(nNilai)), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS" And rsTemp2("TPSVC") <> "W", RepRp(CDbl(nNilai)), 0) & "', " _
                                               & "'" & RepRp(-1 * cekDuo(CDbl(rsTemp2("BODY")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("MTRL")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("PART_G")) + CDbl(rsTemp2("PART_NG")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("OIL_G")) + CDbl(rsTemp2("OIL_NG")), noInv)) & "', '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("SBLT")), noInv)) & "', '" & RepRp(-1 * CDbl(rsTemp2("SUBTOT"))) & "', " _
                                               & " '" & RepRp(-1 * cekDuo(CDbl(rsTemp2("MET")), noInv, 1)) & "', '" & RepRp(-1 * CDbl(rsTemp2("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & RepRp(-1 * CDbl(rsTemp2("LABOR"))) & "', '" & IIf(nDuoInv And Right(rsTemp2("NOINV"), 1) = "B", "", Trim(rsTemp2("NOFAKTUR"))) & "','" & RepRp(-1 * cekDuo(CDbl(rsTemp2("HPP")), noInv)) & "','" & Format(rsTemp2("FD_RETUR"), "MM/DD/YYYY") & "')")
80                            End If
81                            rsAdd.MoveNext
82                        Loop
83                    End If
84                    rsTemp2.MoveNext
85                Loop
86            End If
87        End If
End Sub

Private Function statusDealer(codeStatus As Integer)
    Select Case codeStatus
        Case 1
            statusDealer = "AFFNONPPN"
        Case 2
            statusDealer = "AFFPPN"
        Case 3
            statusDealer = "EXT"
    End Select
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
    If Right(Trim(nINV), 2) = "-B" Then
        cekDuo = 0
        If stat = 1 Then
            cekDuo = CDbl(hrg)
        End If
    Else
        cekDuo = CDbl(hrg)
        If stat = 2 Then
            cekDuo = 0
        End If
    End If
End Function

Private Sub Cetak(vNamaTabel As String, txt As String)
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptRepoInv
    Set Report2 = New RptRepoInvHarian
    Set Report3 = New RptRepoInvPrint
    If txt <> "" Then
        txt = " WHERE " & txt
    End If
    
    If opTipe(1).value Then
        sSQL = "SELECT (CASE RIGHT(A.NOINV,2) WHEN '-B' THEN 1 ELSE (CASE WHEN CASHADD = '' THEN (CASE WHEN CONVERT(VarChar(15), B.TGLKASIR, 1) <> B.INVDATE THEN 1 ELSE FS_KREDIT END) ELSE FS_KREDIT END) END) AS KREDIT, SA,A.NOINV,A.NO_WO,A.NO_POLISI,A.NO_CHASSIS,JOR,JRTE,RTJ,A.THS,SBI,SBE,GRP,TWC,LUB,PDI,WAS, " & _
               "BOR,MATERIAL,PART,OLI,SUBLET,SUBTOT,METERAI,A.TOTAL,TIPEINV,A.CUSTOMER,A.LABOR,A.NOFAKTUR,HPP,A.INVDATE,isnull(B.PPN,0) ppn,isnull(B.FN_PAJAK,0) fn_pajak " & _
               "FROM [TEMPDB]..[" & vNamaTabel & "] A LEFT JOIN T_INVOICE B ON A.NO_WO = B.NO_WO AND REPLACE(A.NOINV,'R','') = B.NOINV " & txt & " ORDER BY A.NOINV"     '
               
        Set oRs = objO.FetchData(sSQL)
        Report.database.SetDataSource oRs
        Report.txtPerusahaan.SetText gPerush
        Report.TxtAlamat.SetText gAlamat
        Report.txtWO.SetText jmlUnit
        Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
        
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 100
            .Visible = True
        End With
    ElseIf opTipe(2).value Then
        sSQL = "SELECT (CASE RIGHT(A.NOINV,2) WHEN '-B' THEN 1 ELSE (CASE WHEN CASHADD = '' THEN (CASE WHEN CONVERT(VarChar(15), B.TGLKASIR, 1) <> B.INVDATE THEN 1 ELSE FS_KREDIT END) ELSE FS_KREDIT END) END) AS KREDIT, SA,A.NOINV,A.NO_WO,A.NO_POLISI,A.NO_CHASSIS,JOR,JRTE,RTJ,A.THS,SBI,SBE,GRP,TWC,LUB,PDI,WAS, " & _
               "BOR,MATERIAL,PART,OLI,SUBLET,SUBTOT,METERAI,A.TOTAL,TIPEINV,A.CUSTOMER,A.LABOR,A.NOFAKTUR,HPP,A.INVDATE,isnull(B.PPN,0) ppn,isnull(B.FN_PAJAK,0) fn_pajak " & _
               "FROM [TEMPDB]..[" & vNamaTabel & "] A LEFT JOIN T_INVOICE B ON A.NO_WO = B.NO_WO AND REPLACE(A.NOINV,'R','') = B.NOINV " & txt & " ORDER BY A.NOINV"     '
               
        Set oRs = objO.FetchData(sSQL)
        Report3.database.SetDataSource oRs
        Report3.txtPerusahaan.SetText gPerush
        Report3.TxtAlamat.SetText gAlamat
        Report3.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
        Report3.txtWO.SetText jmlUnit
        Report3.txtInv.SetText jmlInv
        
        With CRViewer1
            .ReportSource = Report3
            .ViewReport
            .Zoom 100
            .Visible = True
        End With
    ElseIf opTipe(0).value Then
        sSQL = "SELECT SA,NOINV,NO_WO,NO_POLISI,NO_CHASSIS,JOR,JRTE,RTJ,THS,SBI,SBE,GRP,TWC,LUB,PDI,WAS,BOR,MATERIAL,PART,OLI,SUBLET,SUBTOT,METERAI,TOTAL,TIPEINV,CUSTOMER,LABOR,NOFAKTUR,HPP,INVDATE,isnull((select top 1 ppn from t_invoice where no_wo = a.no_wo and noinv = REPLACE(A.NOINV,'R','')),0) ppn,isnull((select top 1 fn_pajak from t_invoice where no_wo = a.no_wo and noinv = REPLACE(A.NOINV,'R','')),0) fn_pajak FROM [TEMPDB]..[" & vNamaTabel & "] A " & txt & " ORDER BY NOINV"     '
        Set oRs = objO.FetchData(sSQL)
        Report2.database.SetDataSource oRs
        Report2.txtPerusahaan.SetText gPerush
        Report2.TxtAlamat.SetText gAlamat
        Report2.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
        Report2.txtWO.SetText jmlUnit
        Report2.txtInv.SetText jmlInv
        
        With CRViewer1
            .ReportSource = Report2
            .ViewReport
            .Zoom 100
            .Visible = True
        End With
    End If
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnBatal_Click(Index As Integer)
    frmeAturKredit.Visible = False
End Sub

Private Sub cmdCash_Click()
    For i = 17 To 20
        Text(i) = ""
    Next
    frmeAturKredit.Visible = True
End Sub

Private Sub cmdSettingPrinter_Click()
    If opTipe(1).value Then
        Report.PrinterSetup Me.hWnd
    ElseIf opTipe(0).value Then
        Report2.PrinterSetup Me.hWnd
    ElseIf opTipe(2).value Then
        Report3.PrinterSetup Me.hWnd
    End If
    CRViewer1.Refresh
End Sub

Private Sub cmdTo_Click(Index As Integer)
Dim rsSearch As ADODB.Recordset

    If Trim(Text(17)) <> "" Then
        Select Case Index
            Case 0
                '#########################################################'
                '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                '#########################################################'
                Set rsSearch = objO.FetchData("SELECT * FROM T_INVOICE WITH(NOLOCK) WHERE NOINV LIKE '" & Trim(Text(17)) & "%' AND RPR_COND = 'K'")
                If Not rsSearch.EOF Then
                    'CASH
                    Call objO.UpdateData("UPDATE T_INVOICE SET FS_KREDIT = 0, CASHADD = '" & Trim(gSA) & "', TGLCASH = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "' WHERE NOINV LIKE '" & Trim(Text(17)) & "%'")
                Else
                    MsgBox "Lakukan Registrasi terlebih dahulu !"
                End If
            Case 1
                '#########################################################'
                '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                '#########################################################'
                Call objO.UpdateData("UPDATE T_INVOICE SET FS_KREDIT = 1, CASHADD = '" & Trim(gSA) & "', TGLCASH = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "' WHERE NOINV LIKE '" & Trim(Text(17)) & "%'")
        End Select
        frmeAturKredit.Visible = False
    Else
        MsgBox "No Invoice belum ada !!."
    End If
End Sub

Private Sub cmdTutup_Click()
    sSQL = "DELETE [TEMPDB]..[" & vNamaTabel & "]"
    Call objO.DeleteData(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    Call hapusForm
    vNamaTabel = "InvTot" & DriveSerialNumber("c:")
    If gCASH Then
        cmdCash.Visible = True
    End If
End Sub

Private Sub opLap_Click(Index As Integer)
    Text(21) = ""
    Text(22) = ""
    Text(23) = ""
    DtDOAwal = ""
    DtDOAkhir = ""
    Label16.Visible = False
    Label17.Visible = False
    DtDOAwal.Visible = False
    DtDOAkhir.Visible = False
    Text(21).Visible = False
    MaskEdBox10.Visible = False
    Select Case Index
        Case 1
            Text(21).Visible = True
            MaskEdBox10.Visible = True
            Label16.Visible = True
            Label17.Visible = True
            DtDOAwal.Visible = True
            DtDOAkhir.Visible = True
            Text(21).SelLength = Len(Text(21))
            Text(21).SetFocus
    End Select
End Sub

Private Sub opVehicle_Click(Index As Integer)
    If opVehicle(0) Then
        Text(7).Visible = True
        Text(8).Visible = True
        MaskEdBox4.Visible = True
        Text(26).Visible = False
    Else
        Text(7).Visible = False
        Text(8).Visible = False
        MaskEdBox4.Visible = False
        Text(26).Visible = True
        Text(26).Left = Text(7).Left
        Text(26).SelLength = Len(Text(26))
        Text(26).SetFocus
    End If
End Sub

Private Sub TblKeluar_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 345
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdCetak.Top = CRViewer1.Top + 345
    cmdCetak.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub

Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdCetak.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdCetak_Click()
    For Each pr In Printers
        If UCase(pr.DeviceName) = UCase(cboPrinter) Then
            GoTo Cetak
        End If
    Next pr
    MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
    cboPrinter.SetFocus
    Exit Sub
    
Cetak:
    Dim Psize As Integer
    Dim POr As Integer
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation
    
    Set Printer = Printers(cboPrinter.ListIndex)
    If opTipe(1).value Then
        Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        Report.PaperOrientation = POr
        Report.PaperSize = Psize
        Report.PrintOut
    ElseIf opTipe(2).value Then
        'report3
        Report3.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        Report3.PaperOrientation = POr
        Report3.PaperSize = Psize
        Report3.PrintOut
    ElseIf opTipe(0).value Then
        'Report2
        Report2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        Report2.PaperOrientation = POr
        Report2.PaperSize = Psize
        Report2.PrintOut
    End If
    
    
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    
    Dim f As New FormExplorer
    f.show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    If Log.PdfOnly = 1 Then
         
        NamaFile = MyPath & f.FileName
        
        If UCase(Right(Trim(f.FileName), 5)) = ".XLSX" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 5)
        If UCase(Right(Trim(f.FileName), 4)) = ".XLS" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 4)
        
         If opTipe(1).value Then
            Report.ExportOptions.DestinationType = crEDTDiskFile
            Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
            Report.ExportOptions.RTFExportAllPages = True
            Report.ExportOptions.FormatType = crEFTExactRichText
            Report.Export False
        ElseIf opTipe(2).value Then
            'report3
            Report3.ExportOptions.DestinationType = crEDTDiskFile
            Report3.ExportOptions.DiskFileName = NamaFile & ".Doc"
            Report3.ExportOptions.RTFExportAllPages = True
            Report3.ExportOptions.FormatType = crEFTExactRichText
            Report3.Export False
        ElseIf opTipe(0).value Then
            'Report2
            Report2.ExportOptions.DestinationType = crEDTDiskFile
            Report2.ExportOptions.DiskFileName = NamaFile & ".Doc"
            Report2.ExportOptions.RTFExportAllPages = True
            Report2.ExportOptions.FormatType = crEFTExactRichText
            Report2.Export False
        End If
        
        
        MsgBox NamaFile & ".Doc" & "  ... Data sudah di Export !", vbInformation
        Exit Sub
        With msg
            .PassingMsg = NamaFile & ".Doc" & "  ... Data sudah di Export !"
            .passingTipe = 2
            .show vbModal
        End With
        Exit Sub
    ElseIf onExpo = 0 Then
        Report.Export
    End If
    Exit Sub
Salah:
    If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
    MsgBox "Err.Line : " & Erl & "Err.Numb : " & err.Number & "Err.Desc. : " & err.Description & vbCr & "Procedure : " & Me.Name & " export"
Keluar:
End Sub

Private Sub tblSimpan_Click()
      Dim boleh As Boolean
      Dim txtTemp As String
      Dim txt As String
1         isAdaData = True
2         boleh = False
3         txtTemp = ""
4         txt = ""
5         whereMultiple = ""
6         If Text(0) <> "" Then
7             boleh = True
8             If txt <> "" Then
9                 txt = txt & "AND "
10            End If
11            txt = txt & "KLPKERJA = '" & Text(0) & "' "
12        End If
13        If Text(2) <> "" Then
14            boleh = True
15            If txt <> "" Then
16                txt = txt & "AND "
17            End If
18            txt = txt & "REPS = '" & Text(2) & "' "
19        End If
20        If Text(3) <> "" Then
21            boleh = True
22            If txt <> "" Then
23                txt = txt & "AND "
24            End If
25            txt = txt & "MECHS = '" & Text(3) & "' "
26        End If
27        If Text(5) <> "" Then
28            txtTemp = txtTemp & "AND FRONT = '" & Text(5) & "' "
29        End If
30        If Text(7) <> "" Or Text(26) <> "" Then
31            If opVehicle(0) Then
32                txtTemp = txtTemp & "AND A.MDL = '" & Trim(Text(7)) & "' "
33            Else
34                txtTemp = txtTemp & "AND A.FMDL LIKE '%" & Trim(Text(26)) & "%' "
35            End If
36        End If
37        If Text(9) <> "" Then
38            txtTemp = txtTemp & "AND B.FS_KD_DEALER = '" & Text(9) & "' "
39        End If
40        If Text(11) <> "" Then
41            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
42                txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "WARRANTY", "W", "") & "' "
43            Else
44                txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "REGULAR", "", IIf(Text(11) = "WARRANTY", "W", "P")) & "' "
45            End If
46            If Trim(Text(11)) = "PROGRAM" Then
47                If Trim(Text(24)) <> "" Then
48                    If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
49                        txtTemp = txtTemp & "AND (KODETIPE = '" & Trim(Text(24)) & "' OR A.no_wo in (select no_wo from T_invoiceJobDetail where FS_LOG_DEL = '' and no_wo = A.no_wo and id_program = '" & Trim(Text(24)) & "' )) "
50                    Else
51                        txtTemp = txtTemp & "AND KODETIPE = '" & Trim(Text(24)) & "' "
52                    End If
53                End If
54            End If
55        End If
          
56        If Text(12) <> "" Then
57            If Text(12) = "BOOKING" Then
58                txtTemp = txtTemp & "AND NOFAK <> ''"
59            ElseIf Text(12) = "NON BOOKING" Then
60                txtTemp = txtTemp & "AND NOFAK = ''"
61            End If
62        End If
63        If Text(13) <> "" Then
64            If Text(13) = "PROSES" Then
65                txtTemp = txtTemp & "AND RPR_COND = 'P'"
66            ElseIf Text(13) = "WAITING" Then
67                txtTemp = txtTemp & "AND RPR_COND = '' AND NOINV = ''"
68            ElseIf Text(13) = "FINISH" Then
69                txtTemp = txtTemp & "AND RPR_COND = 'S'"
70            ElseIf Text(13) = "CLOSING" Then
71                txtTemp = txtTemp & "AND RPR_COND = 'K'"
72            ElseIf Text(13) = "PAUSED" Then
73                txtTemp = txtTemp & "AND RPR_COND = 'T'"
74            End If
75        End If
76        If Text(14) <> "" Then
77            txt = txt & IIf(Trim(txt) <> "", " AND", "") & " TIPEINV = '" & Trim(Text(14)) & "'"
78        End If
79        If Text(15) <> "" Then
80            txtTemp = txtTemp & "AND A.FS_KD_PLG = '" & Trim(Text(15)) & "'"
81        End If
82        If chInvDealer = Checked Then
83            txt = txt & IIf(Trim(txt) <> "", " AND", "") & " RIGHT(A.NOINV,1) = 'B'"
84        End If
85        If chTHS = Checked Then
86            txtTemp = txtTemp & " and THS = '1'"
87        End If
88        If opLap(1) Then
89            If Nul(Trim(DtDOAkhir)) = "" Then
90                DtDOAkhir = Format(Layar.StatusBar1.Panels(2).Text, "DD/MM/yyyy")
91            End If
92            If Trim(DtDOAwal) <> "" Then
93                txtTemp = txtTemp & " and a.DATE BETWEEN '" & Format(Trim(DtDOAwal), "MM/DD/YYYY") & "' AND '" & Format(Trim(DtDOAkhir), "MM/DD/YYYY") & "' " & _
                                      " and A.TIME BETWEEN '" & Trim(Text(22)) & "' AND '" & Trim(Text(23)) & "' "
94            Else
95                txtTemp = txtTemp & " and a.DATE <= '" & Format(Trim(DtDOAkhir), "MM/DD/YYYY") & "' " & _
                                      " and A.TIME BETWEEN '" & Trim(Text(22)) & "' AND '" & Trim(Text(23)) & "' "
96            End If
97        End If
          
          
      '    If opSource(0) = 1 And opSource(1) = 0 Then
      '        txtTemp = txtTemp & " AND (A.JOB = '0' or A.JOB = '') "
      '    ElseIf opSource(0) = 0 And opSource(1) = 1 Then
      '        txtTemp = txtTemp & " AND A.JOB = '1' "
      '    ElseIf opSource(0) = 1 And opSource(1) = 1 Then
      '        txtTemp = txtTemp & " AND ((A.JOB = '0' or A.JOB = '') or A.JOB = '1') "
      '    Else
      '        MsgBox "Outlet unit harus dipilih"
      '        'Text(0).SetFocus
      '        Exit Sub
      '    End If
          
          
98        If opSource(0) = 1 And opSource(1) = 0 And opSource(2) = 0 Then
99            txtTemp = txtTemp & " AND A.JOB = '0' "
100       ElseIf opSource(0) = 1 And opSource(1) = 1 And opSource(2) = 0 Then
101           txtTemp = txtTemp & " AND (A.JOB = '0' or A.JOB = '1') "
102       ElseIf opSource(0) = 1 And opSource(1) = 1 And opSource(2) = 1 Then
103           txtTemp = txtTemp & " AND (A.JOB = '0' or A.JOB = '1' or A.JOB = '2' ) "
104       ElseIf opSource(0) = 0 And opSource(1) = 1 And opSource(2) = 0 Then
105           txtTemp = txtTemp & " AND A.JOB = '1' "
106       ElseIf opSource(0) = 0 And opSource(1) = 1 And opSource(2) = 1 Then
107           txtTemp = txtTemp & " AND (A.JOB = '1' or A.JOB = '2' ) "
108       ElseIf opSource(0) = 0 And opSource(1) = 0 And opSource(2) = 1 Then
109           txtTemp = txtTemp & " AND (A.JOB = '2' ) "
110       Else
111           MsgBox "Outlet unit harus dipilih"
              'Text(0).SetFocus
112           Exit Sub
113       End If
          
114       Call IsiTempRepo(txtTemp, vNamaTabel, txt)
          
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
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
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCari.show (vbModal)
            Case 2
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(0)
                    frmCari.show (vbModal)
                Else
                    MsgBox "Group Job Order Harus diisi"
                    Text(0).SetFocus
                    Exit Sub
                End If
            Case 3
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ALL"
                gBantuan4 = ""
                frmCari.show (vbModal)
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
            Case 7
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 9
                gMainKeyCari = Text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 11
                gMainKeyCari = Text(Index)
                gSearchOption = "TPSVC"
                frmCari.show (vbModal)
            Case 12
                gMainKeyCari = Text(Index)
                gSearchOption = "TP_ORD"
                frmCari.show (vbModal)
            Case 13
                gMainKeyCari = Text(Index)
                gSearchOption = "PROSES"
                frmCari.show (vbModal)
            Case 14
                gMainKeyCari = Text(Index)
                gSearchOption = "STATUSAKUN"
                frmCari.show (vbModal)
            Case 15
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.show (vbModal)
            Case 21
                gMainKeyCari = Text(Index)
                gSearchOption = "DLRSHIFT"
                gBantuan = "WHERE BERLAKU = 1 "
                frmCari.show (vbModal)
            Case 24
                ' PROGRAM
                gMainKeyCari = Text(Index)
                gSearchOption = "PROGRAM_ALL"
                gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset
      Dim rsOpsi As ADODB.Recordset

1         Text(Index) = Replace(Trim(Text(Index)), "'", "")
2         Select Case Index
              Case 0
3                 If Text(Index) <> "" Then
4                     Set rsSearch = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(Index) & "' and FS_LOG_DEL = ''")
5                     If rsSearch.EOF Then
6                         gMainKeyCari = Text(Index)
7                         gSearchOption = "JENIS_KERJA"
8                         frmCari.show (vbModal)
9                     Else
10                        Text(Index + 1) = rsSearch("REMARK")
11                    End If
12                Else
13                    Text(1) = ""
14                End If
15            Case 2
16                If Text(0) <> "" Then
17                    If Text(Index) <> "" Then
18                        Set rsSearch = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(0) & "' AND FS_ITEM = '" & Text(Index) & "' and FS_LOG_DEL = ''")
19                        If rsSearch.EOF Then
20                            gMainKeyCari = Text(Index)
21                            gSearchOption = "JOB_ORDER"
22                            gBantuan = Text(0)
23                            frmCari.show (vbModal)
24                        Else
25                            Text(Index + 1) = rsSearch("KETERANGAN")
26                        End If
27                    Else
28                        Text(Index + 1) = ""
29                    End If
30                Else
31                    MsgBox "Group Job Order Harus diisi"
32                    Text(0).SetFocus
33                    Exit Sub
34                End If
35            Case 3
36                If Text(Index) <> "" Then
37                    Set rsSearch = objO.FetchData("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Text(Index) & "'")
38                    If rsSearch.EOF Then
39                        gMainKeyCari = Text(Index)
40                        gSearchOption = "MEKANIK"
41                        gBantuan = "ALL"
42                        gBantuan4 = ""
43                        frmCari.show (vbModal)
44                    Else
45                        Text(Index + 1) = rsSearch("NAME")
46                    End If
47                Else
48                    Text(Index + 1) = ""
49                End If
50            Case 5
51                If Text(Index) <> "" Then
52                    Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
53                    If rsSearch.EOF Then
54                        gMainKeyCari = Text(Index)
55                        gSearchOption = "SA"
56                        frmCari.show (vbModal)
57                    Else
58                        Text(Index + 1) = rsSearch("FS_NM_PETUGAS")
59                    End If
60                Else
61                    Text(Index + 1) = ""
62                End If
63            Case 7
64                If Text(Index) <> "" Then
65                    If Log.RPLHistori = 0 Then
66                        Set rsSearch = objO.FetchData("SELECT * FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
67                    Else
68                        Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
69                    End If
70                    If rsSearch.EOF Then
71                        gMainKeyCari = Text(Index)
72                        gSearchOption = "MDL"
73                        frmCari.show (vbModal)
74                    Else
75                        Text(Index + 1) = rsSearch("DESC")
76                    End If
77                Else
78                    Text(Index + 1) = ""
79                End If
80            Case 9
81                If Text(Index) <> "" Then
82                    Set rsSearch = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Text(Index) & "' and FS_LOG_DEL = ''")
83                    If rsSearch.EOF Then
84                        gMainKeyCari = Text(Index)
85                        gBantuan = "ALL"
86                        gSearchOption = "DEALER"
87                        frmCari.show (vbModal)
88                    Else
89                        Text(Index + 1) = rsSearch("FS_NM_DEALER")
90                    End If
91                Else
92                    Text(Index + 1) = ""
93                End If
94            Case 11
95                chInvDealer.Visible = False
96                chInvDealer = Unchecked
97                If Text(Index) <> "" Then
98                    Set rsSearch = objO.FetchData("SELECT TPSERVICE FROM T_REPO WITH(NOLOCK) WHERE TPSERVICE = '" & Text(Index) & "'")
99                    If rsSearch.EOF Then
100                       gMainKeyCari = Text(Index)
101                       gSearchOption = "TPSVC"
102                       frmCari.show (vbModal)
103                   End If
                      
104                   chInvDealer.Visible = False
105                   Label19.Visible = False
106                   Text(24).Visible = False
107                   Text(25).Visible = False
108                   MaskEdBox11.Visible = False
109                   If Text(Index) = "PROGRAM" Then
110                       chInvDealer.Visible = True
111                       Label19.Visible = True
112                       Text(24).Visible = True
113                       Text(25).Visible = True
114                       MaskEdBox11.Visible = True
115                   End If
116               End If
117           Case 12
118               If Text(Index) <> "" Then
119                   Set rsSearch = objO.FetchData("SELECT TPORDER FROM T_REPO WITH(NOLOCK) WHERE TPORDER = '" & Text(Index) & "' AND TPORDER <> ''")
120                   If rsSearch.EOF Then
121                       gMainKeyCari = Text(Index)
122                       gSearchOption = "TP_ORD"
123                       frmCari.show (vbModal)
124                   End If
125               End If
126           Case 13
127               If Text(Index) <> "" Then
128                   Set rsSearch = objO.FetchData("SELECT PROSES FROM T_REPO WITH(NOLOCK) WHERE PROSES = '" & Text(Index) & "'")
129                   If rsSearch.EOF Then
130                       gMainKeyCari = Text(Index)
131                       gSearchOption = "PROSES"
132                       frmCari.show (vbModal)
133                   End If
134               End If
135           Case 14
136               If Text(Index) <> "" Then
137                   Set rsSearch = objO.FetchData("SELECT * FROM T_STATUSAKUN WITH(NOLOCK) WHERE TIPE = '" & Trim(Text(Index)) & "'")
138                   If rsSearch.EOF Then
139                       gMainKeyCari = Text(Index)
140                       gSearchOption = "STATUSAKUN"
141                       frmCari.show (vbModal)
142                   End If
143               End If
144           Case 15
145               If Text(Index) <> "" Then
146                       Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH(NOLOCK) WHERE kodeLgn = '" & Text(Index) & "'")
147                       If rsSearch.EOF Then
148                           gMainKeyCari = Text(Index)
149                           gSearchOption = "KODEPLG"
150                           frmCariAcess.show (vbModal)
151                       Else
152                           Text(16) = rsSearch("NAMALGN")
153                      End If
154               Else
155                   Text(16) = ""
156               End If
157           Case 17
158               If Text(Index) <> "" Then
159                   Set rsSearch = objO.FetchData("SELECT * FROM T_INVOICE WITH(NOLOCK) WHERE NOINV = '" & Trim(Text(Index)) & "'")
160                   If Not rsSearch.EOF Then
161                       Text(18) = Trim(rsSearch("CUSTOMER"))
162                       Text(19) = Trim(rsSearch("CUST_ADRES"))
163                       Text(20) = Trim(rsSearch("KOTA"))
164                   End If
165               End If
166           Case 21
167               If Text(Index) <> "" Then
168                   Set rsSearch = objO.FetchData("SELECT SHIFT, JAMMULAI, JAMSELESAI FROM DEALERSHIFT WITH(NOLOCK) WHERE shift = '" & Trim(Text(Index)) & "' AND BERLAKU = 1")
169                   If rsSearch.EOF Then
170                       gMainKeyCari = Text(Index)
171                       gSearchOption = "DLRSHIFT"
172                       gBantuan = "WHERE BERLAKU = 1 "
173                       frmCari.show (vbModal)
174                   Else
175                       Text(22) = Trim(rsSearch("JAMMULAI"))
176                       Text(23) = Trim(rsSearch("JAMSELESAI"))
177                   End If
178               End If
179           Case 24
180               If Text(Index) <> "" Then
                      
181                   If Log.RPLHistori = 0 Then
182                       Set rsOpsi = objO.FetchData("SELECT * FROM HO_program A WITH(NOLOCK) LEFT JOIN HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
183                   Else
184                       If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
185                           sSQL = "SELECT PROGRAM FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
186                           sSQL = sSQL & " UNION "
187                           sSQL = sSQL & " SELECT NAMAFP AS PROGRAM FROM [" & dbReplikasiHO & "]..HO_FIXPRICE A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_FIXPRICECBG B WITH(NOLOCK) ON A.id_FP = B.id_FP WHERE A.id_FP = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
188                           Set rsOpsi = objO.FetchData(sSQL)
189                       Else
190                           Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
191                       End If
                          ''Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
192                   End If
193                   If rsOpsi.EOF Then
194                       gMainKeyCari = Text(Index)
195                       gSearchOption = "PROGRAM_ALL"
196                       gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
197                       frmCari.show (vbModal)
198                   Else
199                       Text(25) = rsOpsi("PROGRAM")
200                   End If
201               End If
202       End Select
End Sub
