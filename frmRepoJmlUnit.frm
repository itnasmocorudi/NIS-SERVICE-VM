VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoJmlUnit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LAPORAN JUMLAH UNIT MASUK"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5685
      TabIndex        =   70
      Top             =   5745
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   8805
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   5715
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   5715
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   1875
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5715
      Width           =   1440
   End
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
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   5715
      Width           =   1050
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   4635
      TabIndex        =   71
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
      MICON           =   "frmRepoJmlUnit.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   675
      TabIndex        =   46
      Top             =   5325
      Visible         =   0   'False
      Width           =   13560
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
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton tblRefresh 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1890
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   661
   End
   Begin VB.CommandButton TblSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   480
      TabIndex        =   18
      Top             =   240
      Width           =   13455
      Begin VB.CheckBox opSource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TMS"
         Height          =   375
         Index           =   2
         Left            =   9600
         TabIndex        =   75
         Top             =   3960
         Width           =   975
      End
      Begin VB.CheckBox opSource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Small SP"
         Height          =   375
         Index           =   1
         Left            =   8520
         TabIndex        =   73
         Top             =   3960
         Width           =   975
      End
      Begin VB.CheckBox opSource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reguler"
         Height          =   375
         Index           =   0
         Left            =   7320
         TabIndex        =   72
         Top             =   3960
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ditinggal"
         Height          =   255
         Left            =   8520
         TabIndex        =   65
         Top             =   3720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ditunggu"
         Height          =   255
         Left            =   7320
         TabIndex        =   64
         Top             =   3720
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7320
         TabIndex        =   55
         Top             =   3360
         Width           =   2895
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
         Begin VB.OptionButton opLap 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan Shift"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   14
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         TabIndex        =   51
         Top             =   2295
         Width           =   3855
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1 Model"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Beberapa Model"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   52
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.CheckBox chNP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SB - NP"
         Height          =   255
         Left            =   2940
         TabIndex        =   50
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox chEM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EM"
         Height          =   255
         Left            =   2940
         TabIndex        =   49
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chTHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THS"
         Height          =   255
         Left            =   1500
         TabIndex        =   48
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox chRTJ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Return Job"
         Height          =   255
         Left            =   1500
         TabIndex        =   47
         Top             =   3480
         Width           =   1215
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
         Format          =   93716481
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1530
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2400
         TabIndex        =   19
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
         Index           =   2
         Left            =   1530
         TabIndex        =   3
         Top             =   1200
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   5400
         TabIndex        =   21
         Top             =   1200
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
         Left            =   1530
         TabIndex        =   4
         Top             =   1560
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
         Left            =   2400
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
         Index           =   1
         Left            =   2520
         TabIndex        =   25
         Top             =   840
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
         Left            =   2520
         TabIndex        =   26
         Top             =   1560
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
         Index           =   5
         Left            =   1530
         TabIndex        =   5
         Top             =   1920
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
         Left            =   2400
         TabIndex        =   27
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
         Index           =   7
         Left            =   1530
         TabIndex        =   6
         Top             =   2640
         Width           =   1875
         _ExtentX        =   3307
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
         Left            =   3480
         TabIndex        =   29
         Top             =   2640
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
         Left            =   2520
         TabIndex        =   31
         Top             =   1920
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
         Index           =   8
         Left            =   1530
         TabIndex        =   32
         Top             =   3000
         Width           =   3915
         _ExtentX        =   6906
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
         Left            =   7290
         TabIndex        =   8
         Top             =   840
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   10320
         TabIndex        =   33
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
         Index           =   10
         Left            =   7290
         TabIndex        =   35
         Top             =   1200
         Width           =   4635
         _ExtentX        =   8176
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
         Left            =   7290
         TabIndex        =   9
         Top             =   1560
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   8800
         TabIndex        =   36
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
         Index           =   12
         Left            =   7290
         TabIndex        =   10
         Top             =   2280
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   270
         Left            =   10320
         TabIndex        =   38
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
         Index           =   13
         Left            =   7290
         TabIndex        =   11
         Top             =   2640
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   10320
         TabIndex        =   40
         Top             =   2640
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
         Format          =   93585409
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   7
         Top             =   3000
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   7290
         TabIndex        =   12
         Top             =   3000
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   10440
         TabIndex        =   15
         Top             =   3360
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
         Left            =   11040
         TabIndex        =   56
         Top             =   3360
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
         Index           =   17
         Left            =   9600
         TabIndex        =   57
         Top             =   3720
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
         Index           =   18
         Left            =   10440
         TabIndex        =   58
         Top             =   3720
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
         Index           =   24
         Left            =   10425
         TabIndex        =   59
         Top             =   1560
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
         Left            =   11850
         TabIndex        =   60
         Top             =   1567
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
         Left            =   7290
         TabIndex        =   61
         Top             =   1920
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
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
         Left            =   6720
         TabIndex        =   74
         Top             =   4080
         Width           =   450
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Unit"
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
         TabIndex        =   63
         Top             =   3720
         Width           =   795
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
         Left            =   9300
         TabIndex        =   62
         Top             =   1605
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label13 
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
         Left            =   6360
         TabIndex        =   54
         Top             =   3390
         Width           =   840
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th Kend."
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
         Left            =   6555
         TabIndex        =   45
         Top             =   3045
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   13440
         Y1              =   600
         Y2              =   600
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
         TabIndex        =   43
         Top             =   240
         Width           =   60
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
         TabIndex        =   42
         Top             =   240
         Width           =   570
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
         Left            =   6720
         TabIndex        =   41
         Top             =   2685
         Width           =   480
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
         Left            =   6435
         TabIndex        =   39
         Top             =   2325
         Width           =   765
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
         Left            =   6330
         TabIndex        =   37
         Top             =   1605
         Width           =   870
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
         Left            =   6165
         TabIndex        =   34
         Top             =   885
         Width           =   1035
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
         Left            =   195
         TabIndex        =   30
         Top             =   2325
         Width           =   1245
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
         Left            =   1245
         TabIndex        =   28
         Top             =   1965
         Width           =   195
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
         Left            =   780
         TabIndex        =   24
         Top             =   1605
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
         Left            =   720
         TabIndex        =   22
         Top             =   1245
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
         Left            =   240
         TabIndex        =   20
         Top             =   885
         Width           =   1200
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   9480
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9255
      Left            =   0
      Picture         =   "frmRepoJmlUnit.frx":001C
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoJmlUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsTemp      As ADODB.Recordset
Dim Report      As Object
Dim vNamaTabel  As String
Dim isAdaData   As Boolean

Public Property Let passingMulai(VALUES As String)
    Text(17) = VALUES
End Property

Public Property Let passingSelesai(VALUES As String)
    Text(18) = VALUES
End Property

Private Sub hapusForm()
    For i = 0 To 18
        Select Case i
            Case 15
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    Tgl1 = isnow
    Tgl2 = isnow
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepo(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo(Ket As String, vNamaTabel As String, txt As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
          
          'DELIV <> 'BTL' AND
1         Set rsTemp = objO.FetchData("SELECT *, A.MECH AS MECH, A.CUSTOMER AS CUSTOMER, C.FS_KD_DEALER AS KDDLRJUAL, C.FS_NM_DEALER AS NMDLRJUAL, timeSerah,d.TUNGGUSVC FROM JOBDTA A WITH(NOLOCK) " & _
                       "INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN STALL B WITH(NOLOCK) ON REPLACE(SUBSTRING(B.KETERANGAN, CHARINDEX('|',B.KETERANGAN)+1,8),'|','') = A.NO_WO AND B.TANGGAL = A.[DATE] " & _
                       "left join jobdtl d WITH(NOLOCK) on a.no_chassis = d.no_chassis and a.no_wo = d.no_wo " & _
                       "WHERE [DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND A.FS_LOG_DEL = '' AND A.DELIV <> 'BTL' " & Ket & " ORDER BY a.NO_WO")
2         If Not rsTemp.EOF Then
3             Call BuatTabelRepoUnitMasuk(vNamaTabel)
4             Do While Not rsTemp.EOF
                  ' DARI JOBDTA
5                 For i = 1 To 8
6                     hlp = "REP" & i
7                     If NulSID(rsTemp(hlp)) <> "" Then
8                         hlp2 = "KLPKERJA" & i
9                         hlp3 = "ACTRATE" & i
10                        hlp4 = "MECH" '"MEKANIK" & i
11                        hlp5 = "RTE" & i
12                        If createTemp(vNamaTabel) Then
13                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("NO_CHASSIS") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                           & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("MDL") & "', '" & rsTemp("FMDL") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & rsTemp("NO_WO") & "','" & rsTemp("NOINV") & "', '" & IIf(rsTemp("TUNGGUSVC") = 0, "Ya", "Tidak") & "', '" & rsTemp("FRONT") & "', '" & rsTemp(hlp2) & "', " _
                                           & "'" & rsTemp(hlp) & "', '" & Trim(rsTemp(hlp4)) & "','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 1, Replace(CDbl(rsTemp(hlp5)) * -1, ",", "."), Replace(CDbl(rsTemp(hlp5)), ",", ".")) & "', '0', '0', '" & Trim(rsTemp("TPSVC")) & "', '" & rsTemp("THS") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("RPR_COND") & "', " _
                                           & "'" & rsTemp("NOFAK") & "', '" & IIf(IsNull(rsTemp("NOSTALL")), "", rsTemp("NOSTALL")) & "', '" & rsTemp("KDDLRJUAL") & "', '" & rsTemp("NMDLRJUAL") & "', '" & Trim(rsTemp("TIME")) & "', '" & Trim(rsTemp("TIMESERAH")) & "')")
14                        End If
15                    End If
16                Next
                  
                  ' DARI ADDITIONAL JOB
                  Dim rsTSub As ADODB.Recordset
17                Set rsTSub = objO.FetchData("SELECT * FROM JOBADD2 WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "'")
18                If Not rsTSub.EOF Then
19                    Do While Not rsTSub.EOF
20                        If createTemp(vNamaTabel) Then
21                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("NO_CHASSIS") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                           & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("MDL") & "', '" & rsTemp("FMDL") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & rsTemp("NO_WO") & "','" & rsTemp("NOINV") & "', '" & IIf(rsTemp("TUNGGUSVC") = 0, "Ya", "Tidak") & "', '" & rsTemp("FRONT") & "', '" & rsTSub("KLPKERJA") & "', " _
                                           & "'" & rsTSub("REP") & "', '" & Trim(rsTemp("MECH")) & "','" & Replace(CDbl(rsTSub("ACTRTE")), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 1, Replace((CDbl(rsTSub("RTE")) * -1), ",", "."), Replace(CDbl(rsTSub("RTE")), ",", ".")) & "', '0', '0', '" & Trim(rsTemp("TPSVC")) & "', '" & rsTemp("THS") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("RPR_COND") & "', " _
                                           & "'" & rsTemp("NOFAK") & "', '" & IIf(IsNull(rsTemp("NOSTALL")), "", rsTemp("NOSTALL")) & "', '" & rsTemp("KDDLRJUAL") & "', '" & rsTemp("NMDLRJUAL") & "', '" & Trim(rsTemp("TIME")) & "', '" & Trim(rsTemp("TIMESERAH")) & "')")
22                        End If
23                        rsTSub.MoveNext
24                    Loop
25                End If
                  
                  ' DARI SUBLET ' DEJ: DEALER EXTERNAL JOB
26                Set rsTSub = objO.FetchData("SELECT * FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "'")
27                If Not rsTSub.EOF Then
28                    Do While Not rsTSub.EOF
29                        If createTemp(vNamaTabel) Then
30                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("NO_CHASSIS") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                           & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("MDL") & "', '" & rsTemp("FMDL") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & rsTemp("NO_WO") & "','" & rsTemp("NOINV") & "','" & IIf(rsTemp("TUNGGUSVC") = 0, "Ya", "Tidak") & "', '" & rsTemp("FRONT") & "', ' SUB', " _
                                           & "'" & rsTSub("NAMABARANG") & "', '','0', '0', '0', '0', '" & Trim(rsTemp("TPSVC")) & "', '" & rsTemp("THS") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("RPR_COND") & "', " _
                                           & "'" & rsTemp("NOFAK") & "', '" & IIf(IsNull(rsTemp("NOSTALL")), "", rsTemp("NOSTALL")) & "', '" & rsTemp("KDDLRJUAL") & "', '" & rsTemp("NMDLRJUAL") & "', '" & Trim(rsTemp("TIME")) & "', '" & Trim(rsTemp("TIMESERAH")) & "')")
31                        End If
32                        rsTSub.MoveNext
33                    Loop
34                End If
35                rsTemp.MoveNext
36            Loop
37            Call Cetak(vNamaTabel, txt)
38        Else
39            isAdaData = False
40            MsgBox "Tidak ada data !!"
41            Exit Sub
42        End If
End Sub

Private Sub Cetak(vNamaTabel As String, txt As String)
Dim rsRepo As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptRepoJmlUnit
    If txt <> "" Then
        txt = " WHERE NO_WO IN (SELECT DISTINCT NO_WO FROM [TEMPDB]..[" & vNamaTabel & "] WHERE " & txt & ")"
    End If
    sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "]" & txt & " ORDER BY NO_WO, KLPKERJA DESC"
    Set rsRepo = objO.FetchData(sSQL)
    Report.database.SetDataSource rsRepo
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamat
    Report.txtShift.SetText IIf(opLap(0), "", "SHIFT : " & Trim(Text(16)) & " ~ Jam " & Trim(Text(17)) & " - " & Trim(Text(18)))
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdKeluar_Click()

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
Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
Report.PaperOrientation = POr
Report.PaperSize = Psize
Report.PrintOut

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
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
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

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    'sSQL = "DROP TABLE [TEMPDB]..[" & vNamaTabel & "]"
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.DeleteData(sSQL)
    Set Report = Nothing
    
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
        
    'adit disableexport 20161220 -begin
    'FraLockExport.Visible = False
    'adit disableexport 20161220 -end
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
    vNamaTabel = "JmlUnit" & DriveSerialNumber("c:")
    Call cmdTutup_Click
    'adit disableexport 20161220 -begin
    'FraLockExport.Visible = False
    'adit disableexport 20161220 -end
End Sub
'adit disableexport 20161220 -begin
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

Private Sub Command3_Click()
1     On Error GoTo Salah
      Dim nmFile As String
      Dim NamaInfo As String
2     NamaInfo = Me.Caption
3     CDialog.filter = "*.doc"
4     CDialog.FileName = Trim(NamaInfo)
5     CDialog.CancelError = True
6     CDialog.ShowSave
7     nmFile = CDialog.FileName & ".doc"
8     Report.ExportOptions.DestinationType = crEDTDiskFile
9     Report.ExportOptions.DiskFileName = nmFile
10    Report.ExportOptions.RTFExportAllPages = True
11    Report.ExportOptions.FormatType = crEFTExactRichText
12    Report.Export False
13    Exit Sub
Salah:
14    MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".export "), vbExclamation, "Err Number : " & Erl
15    MsgBox "Data batal export", vbInformation

End Sub
'adit disableexport 20161220 -end
Private Sub opLap_Click(Index As Integer)
    Text(16).Visible = False
    MaskEdBox10.Visible = False
    Text(16) = ""
    Text(17) = ""
    Text(18) = ""
    Select Case Index
        Case 1
            Text(16).Visible = True
            MaskEdBox10.Visible = True
            Text(16).SelLength = Len(Text(16))
            Text(16).SetFocus
    End Select
End Sub

Private Sub opVehicle_Click(Index As Integer)
    If opVehicle(0) Then
        Text(7).Visible = True
        Text(8).Visible = True
        MaskEdBox4.Visible = True
        Text(14).Visible = False
    Else
        Text(7).Visible = False
        Text(8).Visible = False
        MaskEdBox4.Visible = False
        Text(14).Visible = True
        Text(14).Left = Text(7).Left
        Text(14).SelLength = Len(Text(14))
        Text(14).SetFocus
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

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim txtTemp As String
Dim txt As String
    isAdaData = True
    boleh = False
    txtTemp = ""
    txt = ""
    If Text(0) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "KLPKERJA = '" & Text(0) & "' "
    End If
    If Text(2) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "REPS LIKE '%" & Text(2) & "%' "
    End If
    If Text(3) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "MECHS = '" & Text(3) & "' "
    End If
    If Text(5) <> "" Then
        txtTemp = txtTemp & "AND FRONT = '" & Text(5) & "' "
    End If
    If Text(7) <> "" Or Text(14) <> "" Then
        If opVehicle(0) Then
            txtTemp = txtTemp & "AND A.MDL = '" & Trim(Text(7)) & "' "
        Else
            txtTemp = txtTemp & "AND A.FMDL LIKE '%" & Trim(Text(14)) & "%' "
        End If
    End If
    If Text(9) <> "" Then
        txtTemp = txtTemp & "AND C.FS_KD_DEALER = '" & Text(9) & "' "
    End If
    If Text(11) <> "" Then
        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
            txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "WARRANTY", "W", "") & "' "
            ''20220715 revisi terkait data sebelum multiple marketing
            txtTemp = txtTemp & "AND TPSVC in ('" & IIf(Text(11) = "WARRANTY", "W", "','P") & "') "
        Else
            txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "REGULAR", "", IIf(Text(11) = "WARRANTY", "W", "P")) & "' "
        End If
        ''txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "REGULAR", "", IIf(Text(11) = "WARRANTY", "W", "P")) & "' "
        If Trim(Text(11)) = "PROGRAM" Then
            If Trim(Text(24)) <> "" Then
                If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
                    txtTemp = txtTemp & "AND (KODETIPE = '" & Trim(Text(24)) & "' OR A.no_wo in (select no_wo from jobdetail where no_wo = A.no_wo and id_program = '" & Trim(Text(24)) & "' )) "
                Else
                    txtTemp = txtTemp & "AND KODETIPE = '" & Trim(Text(24)) & "' "
                End If
                ''txtTemp = txtTemp & "AND KODETIPE = '" & Trim(Text(24)) & "' "
            End If
        End If
    End If
    
    If Text(12) <> "" Then
        If Text(12) = "BOOKING" Then
            txtTemp = txtTemp & "AND A.NOFAK <> ''"
        ElseIf Text(12) = "NON BOOKING" Then
            txtTemp = txtTemp & "AND A.NOFAK = ''"
        End If
    End If
    If Text(13) <> "" Then
        If Text(13) = "PROSES" Then
            txtTemp = txtTemp & "AND RPR_COND = 'P'" '"AND RPR_COND <> 'K' AND RPR_COND <> '' AND RPR_COND <> 'S' AND RPR_COND <> 'T'"
        ElseIf Text(13) = "WAITING" Then
            txtTemp = txtTemp & "AND RPR_COND = ''"
        ElseIf Text(13) = "FINISH" Then
            txtTemp = txtTemp & "AND RPR_COND = 'S'"
        ElseIf Text(13) = "CLOSING" Then
            txtTemp = txtTemp & "AND RPR_COND = 'K'"
        ElseIf Text(13) = "PAUSED" Then
            txtTemp = txtTemp & "AND RPR_COND = 'T'"
        End If
    End If
    If Text(15) > 0 Then
        txtTemp = txtTemp & "AND C.FS_TH_MOBIL = '" & CDbl(Text(15)) & "'"
    End If
    If chRTJ = Checked Then
        txtTemp = txtTemp & "AND FN_RETURN = '1'"
    End If
    If chTHS = Checked Then
        txtTemp = txtTemp & "AND THS = '1'"
    End If
    If chEM = Checked Then
        txtTemp = txtTemp & "AND a.code = '1'"
    End If
    If chNP = Checked Then
        txtTemp = txtTemp & "AND a.kd_k = '1'"
    End If
    If opLap(1) Then
        txtTemp = txtTemp & "AND a.TIME BETWEEN '" & Trim(Text(17)) & "' AND '" & Trim(Text(18)) & "' "
    End If
    
    If Check1 = 1 And Check2 = 0 Then
        txtTemp = txtTemp & "AND d.TUNGGUSVC = '0' "
    ElseIf Check2 = 1 And Check1 = 0 Then
        txtTemp = txtTemp & "AND d.TUNGGUSVC = '1' "
    ElseIf Check2 = 1 And Check1 = 1 Then
        txtTemp = txtTemp & "AND (d.TUNGGUSVC = '1' or d.TUNGGUSVC = '0') "
    Else
        MsgBox "Status unit harus dipilih"
        'Text(0).SetFocus
        Exit Sub
    End If
    
    
    If opSource(0) = 1 And opSource(1) = 0 And opSource(2) = 0 Then
        txtTemp = txtTemp & " AND A.JOB = '0' "
    ElseIf opSource(0) = 1 And opSource(1) = 1 And opSource(2) = 0 Then
        txtTemp = txtTemp & " AND (A.JOB = '0' or A.JOB = '1') "
    ElseIf opSource(0) = 1 And opSource(1) = 1 And opSource(2) = 1 Then
        txtTemp = txtTemp & " AND (A.JOB = '0' or A.JOB = '1' or A.JOB = '2' or A.TMS = '1') "
    ElseIf opSource(0) = 0 And opSource(1) = 1 And opSource(2) = 0 Then
        txtTemp = txtTemp & " AND A.JOB = '1' "
    ElseIf opSource(0) = 0 And opSource(1) = 1 And opSource(2) = 1 Then
        txtTemp = txtTemp & " AND (A.JOB = '1' or A.JOB = '2' or A.TMS = '1') "
    ElseIf opSource(0) = 0 And opSource(1) = 0 And opSource(2) = 1 Then
        txtTemp = txtTemp & " AND (A.JOB = '2' or A.TMS = '1') "
    Else
        MsgBox "Outlet unit harus dipilih"
        'Text(0).SetFocus
        Exit Sub
    End If
    
    Call IsiTempRepo(txtTemp, vNamaTabel, txt)
     
    'adit disableexport 20161220 -begin
    If isAdaData Then
        Call DisableExport
        Call setLokasiTombol
    End If
    'adit disableexport 20161220 -end
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
                    If Trim(Text(0)) <> "SUB" Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "JOB_ORDER"
                        gBantuan = Text(0)
                        frmCari.show (vbModal)
                    Else
                        gMainKeyCari = Text(Index)
                        gSearchOption = "SUBLET"
                        frmCari.show (vbModal)
                    End If
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
            Case 16
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


Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsHlp As ADODB.Recordset
Dim rsSearch As ADODB.Recordset
Dim rsOpsi As ADODB.Recordset

    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If rsSearch.EOF Then
                    If Trim(Text(Index)) <> "SUB" Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "JENIS_KERJA"
                        frmCari.show (vbModal)
                    Else
                        Text(Index + 1) = "SUBLET"
                    End If
                Else
                    Text(Index + 1) = rsSearch("REMARK")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 2
            If Text(0) <> "" Then
                If Text(Index) <> "" Then
                    Set rsSearch = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(0) & "' AND FS_ITEM = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                    If rsSearch.EOF Then
                        If Trim(Text(0)) <> "SUB" Then
                            gMainKeyCari = Text(Index)
                            gSearchOption = "JOB_ORDER"
                            gBantuan = Text(0)
                            frmCari.show (vbModal)
                        Else
                            Set rsHlp = objO.FetchData("SELECT DISTINCT NAMABARANG FROM JOBSUBLET WITH(NOLOCK) WHERE NAMABARANG = '" & Trim(Text(Index)) & "'")
                            If rsHlp.EOF Then
                                gMainKeyCari = Text(Index)
                                gSearchOption = "SUBLET"
                                frmCari.show (vbModal)
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox "Group Job Order Harus diisi"
                Text(0).SetFocus
                Exit Sub
            End If
        Case 3
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MEKANIK"
                    gBantuan = "ALL"
                    gBantuan4 = ""
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("NAME")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 5
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS  WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SA"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("FS_NM_PETUGAS")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 7
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("DESC")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 9
            If Text(Index) <> "" Then
                
                If Log.RPLHistori = 0 Then
                    Set rsSearch = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                Else
                    Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                End If
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "ALL"
                    gSearchOption = "DEALER"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("FS_NM_DEALER")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 11
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT TPSERVICE FROM T_REPO WITH(NOLOCK) WHERE TPSERVICE = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "TPSVC"
                    frmCari.show (vbModal)
                Else
                    Label19.Visible = False
                    MaskEdBox11.Visible = False
                    Text(25).Visible = False
                    Text(24).Visible = False
                    If Trim(Text(Index)) = "PROGRAM" Then
                        Label19.Visible = True
                        MaskEdBox11.Visible = True
                        Text(25).Visible = True
                        Text(24).Visible = True
                    End If
                End If
            End If
        Case 12
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT TPORDER FROM T_REPO WITH(NOLOCK) WHERE TPORDER = '" & Text(Index) & "' AND TPORDER <> ''")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "TP_ORD"
                    frmCari.show (vbModal)
                End If
            End If
        Case 13
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT PROSES FROM T_REPO WITH(NOLOCK) WHERE PROSES = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PROSES"
                    frmCari.show (vbModal)
                End If
            End If
        Case 16
            If Trim(Text(Index)) <> "" Then
                Set rsSearch = objO.FetchData("SELECT shift, jammulai, jamselesai FROM DEALERSHIFT WITH(NOLOCK) WHERE SHIFT = '" & Trim(Text(Index)) & "' AND BERLAKU = 1")
                If rsSearch.EOF Then
                    Text(16) = ""
                    Text(17) = ""
                    Text(18) = ""
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PROSES"
                    frmCari.show (vbModal)
                Else
                    Text(17) = Trim(rsSearch("JAMMULAI"))
                    Text(18) = Trim(rsSearch("JAMSELESAI"))
                End If
            End If
        Case 24
            If Trim(Text(Index)) <> "" Then
                
                If Log.RPLHistori = 0 Then
                    Set rsOpsi = objO.FetchData("SELECT * FROM HO_program A WITH(NOLOCK) LEFT JOIN HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
                Else
                    If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
                        sSQL = "SELECT PROGRAM FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
                        sSQL = sSQL & " UNION "
                        sSQL = sSQL & " SELECT NAMAFP AS PROGRAM FROM [" & dbReplikasiHO & "]..HO_FIXPRICE A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_FIXPRICECBG B WITH(NOLOCK) ON A.id_FP = B.id_FP WHERE A.id_FP = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
                        Set rsOpsi = objO.FetchData(sSQL)
                    Else
                        Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
                    End If
                    ''Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
                End If
                If rsOpsi.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PROGRAM_ALL"
                    gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
                    frmCari.show (vbModal)
                Else
                    Text(25) = rsOpsi("PROGRAM")
                End If
            End If
    End Select
End Sub
